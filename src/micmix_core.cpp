#include "micmix_core.h"

#include <audioclient.h>
#include <audioclientactivationparams.h>
#include <audiopolicy.h>
#include <endpointvolume.h>
#include <functiondiscoverykeys_devpkey.h>
#include <Shlwapi.h>
#include <bcrypt.h>
#include <tlhelp32.h>
#include <mmsystem.h>
#include <wrl/client.h>

#include <speex/speex_resampler.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <chrono>
#include <condition_variable>
#include <cmath>
#include <cstring>
#include <deque>
#include <cwctype>
#include <cstdlib>
#include <exception>
#include <filesystem>
#include <fstream>
#include <iomanip>
#include <sstream>
#include <unordered_map>
#include <unordered_set>

#include "settings_window.h"
#include "effects_window.h"

using Microsoft::WRL::ComPtr;

TS3Functions g_ts3Functions = {};
std::string g_pluginId;
extern "C" IMAGE_DOS_HEADER __ImageBase;

namespace {

constexpr char kLogChannel[] = "MicMix";
constexpr int  kTargetRate = 48000;
constexpr int  kResamplerQuality = 6; // Speex quality: 0 = fastest, 10 = best quality.
constexpr int  kVoiceTxPollMs = 40;
constexpr uint64_t kTalkEventFreshMs = 450ULL;
constexpr uint64_t kForceTxMusicWindowMs = 6000ULL;
constexpr uint64_t kVoiceTxReapplyMs = 12000ULL;
constexpr uint64_t kCaptureWatchdogSilenceMs = 4500ULL;
constexpr uint64_t kCaptureWatchdogCooldownMs = 12000ULL;
constexpr uint64_t kMusicMetaUpdateMinIntervalMs = 2000ULL;
constexpr uint64_t kClipIndicatorHoldMs = 2600ULL;
constexpr uint32_t kMinSupportedSourceRate = 8000;
constexpr uint32_t kMaxSupportedSourceRate = 384000;
constexpr uintmax_t kMaxConfigBytes = 1024U * 1024U;
constexpr size_t kMaxEffectsPerChain = 64;
constexpr size_t kMaxEffectPathLen = 1024;
constexpr size_t kMaxEffectNameLen = 192;
constexpr size_t kMaxEffectUidLen = 96;
constexpr size_t kMaxEffectStateBlobLen = 4096;
constexpr size_t kMaxEffectStatusLen = 64;
constexpr wchar_t kVstHostPipeName[] = L"micmix_vst_host";
constexpr wchar_t kVstAudioShmPrefix[] = L"Local\\MicMixVstAudioV2";
constexpr size_t kVstHostMaxPipeMessageBytes = 64U * 1024U;
constexpr size_t kVstHostSyncPayloadMaxBytes = 28U * 1024U;
constexpr uint64_t kVstOutputWaitMinUs = 6000ULL;
constexpr uint64_t kVstOutputWaitMaxUs = 24000ULL;
constexpr uint64_t kVstMicIpcAcquireSpinUs = 600ULL;
constexpr uint64_t kVstMicOutputWaitMinUs = 3000ULL;
constexpr uint64_t kVstMicOutputWaitMaxUs = 8000ULL;
constexpr size_t kVstPendingPacketLimit = 128;

uint64_t ComputeVstOutputWaitUs(uint32_t frames) {
    const uint64_t frameDurationUs = (static_cast<uint64_t>(frames) * 1000000ULL) / static_cast<uint64_t>(kTargetRate);
    const uint64_t suggested = frameDurationUs + 4000ULL;
    return std::clamp<uint64_t>(suggested, kVstOutputWaitMinUs, kVstOutputWaitMaxUs);
}

bool IsSupportedSourceRate(uint32_t sampleRate) {
    return sampleRate >= kMinSupportedSourceRate && sampleRate <= kMaxSupportedSourceRate;
}

uint32_t GetLogicalCpuCount() {
    SYSTEM_INFO si{};
    GetNativeSystemInfo(&si);
    return std::max<uint32_t>(1U, si.dwNumberOfProcessors);
}

int DetermineAutoResamplerQualityFromCpu(uint32_t logicalCpuCount) {
    // CPU-only heuristic with zero runtime benchmark overhead.
    if (logicalCpuCount <= 4U) return 5;
    if (logicalCpuCount <= 8U) return 6;
    if (logicalCpuCount <= 12U) return 7;
    return 8;
}

int DetermineAutoResamplerQuality() {
    return DetermineAutoResamplerQualityFromCpu(GetLogicalCpuCount());
}

int ResolveResamplerQualitySetting(int configuredValue) {
    if (configuredValue < 0) {
        return DetermineAutoResamplerQuality();
    }
    return std::clamp(configuredValue, 0, 10);
}

struct PendingVstPackets {
    std::deque<micmix::vstipc::AudioPacket> queue;

    void Clear() {
        queue.clear();
    }

    void Store(const micmix::vstipc::AudioPacket& packet) {
        if (queue.size() >= kVstPendingPacketLimit) {
            queue.pop_front();
        }
        queue.push_back(packet);
    }

    bool Take(uint32_t seq, micmix::vstipc::AudioPacket& out) {
        for (auto it = queue.begin(); it != queue.end(); ++it) {
            if (it->seq == seq) {
                out = *it;
                queue.erase(it);
                return true;
            }
        }
        return false;
    }
};

template <typename RingT>
bool TryReadMatchingVstPacket(
    RingT& ring,
    uint32_t seq,
    PendingVstPackets& pending,
    micmix::vstipc::AudioPacket& out,
    uint64_t waitUs) {
    if (pending.Take(seq, out)) {
        return true;
    }

    auto scanRing = [&]() -> bool {
        micmix::vstipc::AudioPacket packet{};
        while (micmix::vstipc::RingPop(ring, packet)) {
            if (packet.seq == seq) {
                out = packet;
                return true;
            }
            pending.Store(packet);
        }
        return false;
    };

    if (scanRing()) {
        return true;
    }
    if (waitUs == 0) {
        return false;
    }

    const auto deadline = std::chrono::steady_clock::now() + std::chrono::microseconds(waitUs);
    while (std::chrono::steady_clock::now() < deadline) {
        SwitchToThread();
        if (scanRing()) {
            return true;
        }
    }
    return pending.Take(seq, out);
}

std::mutex g_logMutex;
std::string g_logPath;
std::string g_lastLogPayload;
const char* g_lastLogLevel = nullptr;
uint32_t g_suppressedLogCount = 0;
uint32_t g_logWriteCount = 0;
std::atomic<bool> g_tsLogEnabled{true};

std::wstring Utf8ToWide(const std::string& text) {
    if (text.empty()) {
        return {};
    }
    int len = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, nullptr, 0);
    std::wstring out(static_cast<size_t>(len), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, out.data(), len);
    if (!out.empty() && out.back() == L'\0') {
        out.pop_back();
    }
    return out;
}

std::string WideToUtf8(const std::wstring& text) {
    if (text.empty()) {
        return {};
    }
    int len = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string out(static_cast<size_t>(len), '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, out.data(), len, nullptr, nullptr);
    if (!out.empty() && out.back() == '\0') {
        out.pop_back();
    }
    return out;
}

struct ComInit {
    HRESULT hr = E_FAIL;
    bool needsUninit = false;

    explicit ComInit(DWORD coinit = COINIT_MULTITHREADED) {
        hr = CoInitializeEx(nullptr, coinit);
        if (hr == RPC_E_CHANGED_MODE) {
            // COM is already initialized on this thread with a different apartment.
            // That is still usable for the APIs here; we just must not uninitialize.
            hr = S_OK;
            needsUninit = false;
            return;
        }
        needsUninit = SUCCEEDED(hr);
    }

    ~ComInit() {
        if (needsUninit) {
            CoUninitialize();
        }
    }
};

void AppendLogLine(const char* level, const std::string& text) {
    std::lock_guard<std::mutex> lock(g_logMutex);
    if (g_logPath.empty()) {
        return;
    }

    const uint64_t nowMs = GetTickCount64();
    if (g_lastLogLevel && std::strcmp(g_lastLogLevel, level) == 0 && g_lastLogPayload == text) {
        // Suppress short-burst duplicate lines to avoid hot loops spamming disk I/O.
        static uint64_t lastDuplicateMs = 0;
        if (nowMs <= (lastDuplicateMs + 400ULL)) {
            ++g_suppressedLogCount;
            lastDuplicateMs = nowMs;
            return;
        }
        lastDuplicateMs = nowMs;
    }

    if ((++g_logWriteCount % 128U) == 0U) {
        std::error_code ec;
        const std::filesystem::path logPath = std::filesystem::path(Utf8ToWide(g_logPath));
        const uintmax_t size = std::filesystem::file_size(logPath, ec);
        if (!ec && size > (4U * 1024U * 1024U)) {
            std::ofstream trunc(logPath, std::ios::trunc);
            if (trunc.is_open()) {
                trunc << "log rotated: exceeded 4 MiB\n";
            }
        }
    }

    std::ofstream out(std::filesystem::path(Utf8ToWide(g_logPath)), std::ios::app);
    if (!out.is_open()) {
        return;
    }
    SYSTEMTIME st{};
    GetLocalTime(&st);
    if (g_suppressedLogCount > 0) {
        out << st.wYear << "-" << st.wMonth << "-" << st.wDay << " "
            << st.wHour << ":" << st.wMinute << ":" << st.wSecond << "."
            << st.wMilliseconds << " [INFO] (suppressed duplicate lines: " << g_suppressedLogCount << ")\n";
        g_suppressedLogCount = 0;
    }
    out << st.wYear << "-" << st.wMonth << "-" << st.wDay << " "
        << st.wHour << ":" << st.wMinute << ":" << st.wSecond << "."
        << st.wMilliseconds << " [" << level << "] " << text << "\n";
    g_lastLogLevel = level;
    g_lastLogPayload = text;
}

std::string Trim(std::string value) {
    while (!value.empty() && std::isspace(static_cast<unsigned char>(value.back())) != 0) {
        value.pop_back();
    }
    size_t i = 0;
    while (i < value.size() && std::isspace(static_cast<unsigned char>(value[i])) != 0) {
        ++i;
    }
    return value.substr(i);
}

bool ParseBool(const std::string& value, bool fallback) {
    const std::string v = Trim(value);
    if (v == "1" || v == "true" || v == "TRUE") {
        return true;
    }
    if (v == "0" || v == "false" || v == "FALSE") {
        return false;
    }
    return fallback;
}

bool HasNonWhitespace(const std::string& text) {
    for (char ch : text) {
        if (std::isspace(static_cast<unsigned char>(ch)) == 0) {
            return true;
        }
    }
    return false;
}

size_t FirstPayloadContentOffset(const std::string& payload) {
    size_t pos = 0;
    if (payload.size() >= 3) {
        const unsigned char b0 = static_cast<unsigned char>(payload[0]);
        const unsigned char b1 = static_cast<unsigned char>(payload[1]);
        const unsigned char b2 = static_cast<unsigned char>(payload[2]);
        if (b0 == 0xEFU && b1 == 0xBBU && b2 == 0xBFU) {
            pos = 3;
        }
    }
    while (pos < payload.size() && std::isspace(static_cast<unsigned char>(payload[pos])) != 0) {
        ++pos;
    }
    return pos;
}

bool LooksLikeJsonObjectPayload(const std::string& payload) {
    const size_t pos = FirstPayloadContentOffset(payload);
    return pos < payload.size() && payload[pos] == '{';
}

bool ParseIniLikeConfigPayload(const std::string& payload, std::unordered_map<std::string, std::string>& kv) {
    bool parsedAny = false;
    std::istringstream in(payload);
    std::string line;
    while (std::getline(in, line)) {
        line = Trim(line);
        if (line.empty() || line[0] == '#' || line[0] == ';') {
            continue;
        }
        const size_t eq = line.find('=');
        if (eq == std::string::npos) {
            continue;
        }
        kv[Trim(line.substr(0, eq))] = Trim(line.substr(eq + 1));
        parsedAny = true;
    }
    return parsedAny;
}

size_t SkipJsonWs(const std::string& payload, size_t pos) {
    while (pos < payload.size() && std::isspace(static_cast<unsigned char>(payload[pos])) != 0) {
        ++pos;
    }
    return pos;
}

bool ParseJsonStringToken(const std::string& payload, size_t& pos, std::string& out) {
    out.clear();
    if (pos >= payload.size() || payload[pos] != '"') {
        return false;
    }
    ++pos;
    while (pos < payload.size()) {
        const char ch = payload[pos++];
        if (ch == '"') {
            return true;
        }
        if (ch == '\\') {
            if (pos >= payload.size()) {
                return false;
            }
            const char esc = payload[pos++];
            switch (esc) {
            case '"':
            case '\\':
            case '/':
                out.push_back(esc);
                break;
            case 'b':
                out.push_back('\b');
                break;
            case 'f':
                out.push_back('\f');
                break;
            case 'n':
                out.push_back('\n');
                break;
            case 'r':
                out.push_back('\r');
                break;
            case 't':
                out.push_back('\t');
                break;
            case 'u':
                if ((pos + 4) > payload.size()) {
                    return false;
                }
                for (size_t i = 0; i < 4; ++i) {
                    if (std::isxdigit(static_cast<unsigned char>(payload[pos + i])) == 0) {
                        return false;
                    }
                }
                pos += 4;
                out.push_back('?');
                break;
            default:
                return false;
            }
            continue;
        }
        if (static_cast<unsigned char>(ch) < 0x20U) {
            return false;
        }
        out.push_back(ch);
    }
    return false;
}

bool ConsumeJsonLiteral(const std::string& payload, size_t& pos, const char* literal) {
    const size_t n = std::strlen(literal);
    if ((pos + n) > payload.size()) {
        return false;
    }
    if (payload.compare(pos, n, literal) != 0) {
        return false;
    }
    pos += n;
    return true;
}

bool ParseJsonNumberToken(const std::string& payload, size_t& pos, std::string& out) {
    const size_t begin = pos;
    if (pos < payload.size() && payload[pos] == '-') {
        ++pos;
    }
    const size_t intBegin = pos;
    while (pos < payload.size() && std::isdigit(static_cast<unsigned char>(payload[pos])) != 0) {
        ++pos;
    }
    if (pos == intBegin) {
        return false;
    }
    if (pos < payload.size() && payload[pos] == '.') {
        ++pos;
        const size_t fracBegin = pos;
        while (pos < payload.size() && std::isdigit(static_cast<unsigned char>(payload[pos])) != 0) {
            ++pos;
        }
        if (pos == fracBegin) {
            return false;
        }
    }
    if (pos < payload.size() && (payload[pos] == 'e' || payload[pos] == 'E')) {
        ++pos;
        if (pos < payload.size() && (payload[pos] == '+' || payload[pos] == '-')) {
            ++pos;
        }
        const size_t expBegin = pos;
        while (pos < payload.size() && std::isdigit(static_cast<unsigned char>(payload[pos])) != 0) {
            ++pos;
        }
        if (pos == expBegin) {
            return false;
        }
    }
    out.assign(payload.substr(begin, pos - begin));
    return true;
}

bool SkipJsonComposite(const std::string& payload, size_t& pos, char openCh, char closeCh) {
    if (pos >= payload.size() || payload[pos] != openCh) {
        return false;
    }
    int depth = 0;
    while (pos < payload.size()) {
        const char ch = payload[pos];
        if (ch == '"') {
            std::string ignored;
            if (!ParseJsonStringToken(payload, pos, ignored)) {
                return false;
            }
            continue;
        }
        ++pos;
        if (ch == openCh) {
            ++depth;
            continue;
        }
        if (ch == closeCh) {
            --depth;
            if (depth == 0) {
                return true;
            }
        }
    }
    return false;
}

bool ParseJsonValueToken(const std::string& payload, size_t& pos, std::string& out) {
    out.clear();
    pos = SkipJsonWs(payload, pos);
    if (pos >= payload.size()) {
        return false;
    }
    const char ch = payload[pos];
    if (ch == '"') {
        return ParseJsonStringToken(payload, pos, out);
    }
    if (ch == '{') {
        return SkipJsonComposite(payload, pos, '{', '}');
    }
    if (ch == '[') {
        return SkipJsonComposite(payload, pos, '[', ']');
    }
    if (ch == 't') {
        if (!ConsumeJsonLiteral(payload, pos, "true")) {
            return false;
        }
        out = "true";
        return true;
    }
    if (ch == 'f') {
        if (!ConsumeJsonLiteral(payload, pos, "false")) {
            return false;
        }
        out = "false";
        return true;
    }
    if (ch == 'n') {
        if (!ConsumeJsonLiteral(payload, pos, "null")) {
            return false;
        }
        out.clear();
        return true;
    }
    return ParseJsonNumberToken(payload, pos, out);
}

bool ParseJsonObjectConfigPayload(const std::string& payload, std::unordered_map<std::string, std::string>& kv) {
    std::unordered_map<std::string, std::string> parsedKv;
    size_t pos = SkipJsonWs(payload, 0);
    if (pos >= payload.size() || payload[pos] != '{') {
        return false;
    }
    ++pos;
    for (;;) {
        pos = SkipJsonWs(payload, pos);
        if (pos >= payload.size()) {
            return false;
        }
        if (payload[pos] == '}') {
            ++pos;
            break;
        }
        std::string key;
        if (!ParseJsonStringToken(payload, pos, key)) {
            return false;
        }
        pos = SkipJsonWs(payload, pos);
        if (pos >= payload.size() || payload[pos] != ':') {
            return false;
        }
        ++pos;
        std::string value;
        if (!ParseJsonValueToken(payload, pos, value)) {
            return false;
        }
        parsedKv[Trim(key)] = Trim(value);
        pos = SkipJsonWs(payload, pos);
        if (pos >= payload.size()) {
            return false;
        }
        if (payload[pos] == ',') {
            ++pos;
            continue;
        }
        if (payload[pos] == '}') {
            ++pos;
            break;
        }
        return false;
    }
    pos = SkipJsonWs(payload, pos);
    if (pos != payload.size()) {
        return false;
    }
    kv = std::move(parsedKv);
    return true;
}

bool ReadPayloadWithLimit(std::istream& in, size_t maxBytes, std::string& out, bool& exceededLimit) {
    out.clear();
    exceededLimit = false;
    std::array<char, 8192> chunk{};
    while (in) {
        in.read(chunk.data(), static_cast<std::streamsize>(chunk.size()));
        const std::streamsize got = in.gcount();
        if (got <= 0) {
            break;
        }
        const size_t appendBytes = static_cast<size_t>(got);
        const size_t remaining = (out.size() < maxBytes) ? (maxBytes - out.size()) : 0U;
        if (appendBytes > remaining) {
            if (remaining > 0) {
                out.append(chunk.data(), remaining);
            }
            exceededLimit = true;
            return true;
        }
        out.append(chunk.data(), appendBytes);
    }
    return !in.bad();
}

std::string HrToHex(HRESULT hr) {
    std::ostringstream ss;
    ss << "0x" << std::hex << std::uppercase << static_cast<uint32_t>(hr);
    return ss.str();
}

void StripConfigControlChars(std::string& value, size_t maxLen) {
    value.erase(std::remove_if(value.begin(), value.end(), [](char ch) {
        return ch == '\r' || ch == '\n' || ch == '\t';
    }), value.end());
    if (value.size() > maxLen) {
        value.resize(maxLen);
    }
}

uint64_t Fnv1a64(const std::string& text) {
    uint64_t hash = 1469598103934665603ULL;
    for (unsigned char ch : text) {
        hash ^= static_cast<uint64_t>(ch);
        hash *= 1099511628211ULL;
    }
    return hash;
}

std::string GenerateHexToken(size_t byteCount);

std::string BuildEffectUid(const std::string& normalizedPath) {
    std::string folded = normalizedPath;
    std::transform(folded.begin(), folded.end(), folded.begin(), [](char ch) {
        return static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
    });
    std::ostringstream ss;
    ss << "mmx_" << std::hex << std::nouppercase << Fnv1a64(folded);
    return ss.str();
}

std::string BuildEffectUidWithSuffix(const std::string& normalizedPath, const std::string& suffix) {
    const std::string base = BuildEffectUid(normalizedPath);
    if (suffix.empty()) {
        return base;
    }
    const std::string sep = "_";
    if (base.size() + sep.size() + suffix.size() <= kMaxEffectUidLen) {
        return base + sep + suffix;
    }
    const size_t suffixBudget = std::min(suffix.size(), kMaxEffectUidLen > 1U ? (kMaxEffectUidLen - 1U) : 0U);
    const size_t baseBudget = (kMaxEffectUidLen > (1U + suffixBudget)) ? (kMaxEffectUidLen - 1U - suffixBudget) : 0U;
    std::string out = base.substr(0, baseBudget);
    out.push_back('_');
    out.append(suffix.substr(0, suffixBudget));
    return out;
}

template <typename ExistsFn>
std::string BuildUniqueEffectUidWithExists(const std::string& normalizedPath, ExistsFn exists) {
    const std::string baseUid = BuildEffectUid(normalizedPath);
    if (!exists(baseUid)) {
        return baseUid;
    }

    for (int attempt = 0; attempt < 64; ++attempt) {
        std::string token = GenerateHexToken(3);
        if (token.empty()) {
            std::ostringstream fallback;
            fallback << std::hex << std::nouppercase << GetTickCount64() << attempt;
            token = fallback.str();
        }
        const std::string candidate = BuildEffectUidWithSuffix(normalizedPath, token);
        if (!exists(candidate)) {
            return candidate;
        }
    }

    for (uint64_t counter = 0; counter < 2048; ++counter) {
        std::ostringstream ss;
        ss << std::hex << std::nouppercase << counter;
        const std::string candidate = BuildEffectUidWithSuffix(normalizedPath, ss.str());
        if (!exists(candidate)) {
            return candidate;
        }
    }
    return BuildEffectUidWithSuffix(normalizedPath, "uniq");
}

std::string GenerateHexToken(size_t byteCount) {
    static constexpr char kHex[] = "0123456789abcdef";
    std::string bytes;
    bytes.resize(byteCount);
    NTSTATUS status = BCryptGenRandom(
        nullptr,
        reinterpret_cast<PUCHAR>(bytes.data()),
        static_cast<ULONG>(bytes.size()),
        BCRYPT_USE_SYSTEM_PREFERRED_RNG);
    if (!BCRYPT_SUCCESS(status)) {
        BCRYPT_ALG_HANDLE rng = nullptr;
        if (BCRYPT_SUCCESS(BCryptOpenAlgorithmProvider(&rng, BCRYPT_RNG_ALGORITHM, nullptr, 0))) {
            status = BCryptGenRandom(
                rng,
                reinterpret_cast<PUCHAR>(bytes.data()),
                static_cast<ULONG>(bytes.size()),
                0);
            BCryptCloseAlgorithmProvider(rng, 0);
        }
    }
    if (!BCRYPT_SUCCESS(status)) {
        return {};
    }
    std::string out;
    out.reserve(byteCount * 2U);
    for (unsigned char b : bytes) {
        out.push_back(kHex[(b >> 4) & 0x0F]);
        out.push_back(kHex[b & 0x0F]);
    }
    return out;
}

std::wstring GenerateVstHostPipeName() {
    const std::string token = GenerateHexToken(8);
    if (token.empty()) {
        return {};
    }
    std::wostringstream ss;
    ss << kVstHostPipeName << L"_"
       << std::hex << std::nouppercase
       << static_cast<unsigned long>(GetCurrentProcessId())
       << L"_"
       << Utf8ToWide(token);
    return ss.str();
}

std::wstring GenerateVstAudioShmName() {
    const std::string token = GenerateHexToken(8);
    if (token.empty()) {
        return {};
    }
    std::wostringstream ss;
    ss << kVstAudioShmPrefix << L"_"
       << std::hex << std::nouppercase
       << static_cast<unsigned long>(GetCurrentProcessId())
       << L"_"
       << Utf8ToWide(token);
    return ss.str();
}

bool IsAmd64PeBinary(const std::filesystem::path& filePath) {
    HANDLE file = CreateFileW(filePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (file == INVALID_HANDLE_VALUE) {
        return false;
    }
    LARGE_INTEGER fileSize{};
    if (!GetFileSizeEx(file, &fileSize) || fileSize.QuadPart <= 0) {
        CloseHandle(file);
        return false;
    }
    const uint64_t bytes = static_cast<uint64_t>(fileSize.QuadPart);
    if (bytes < static_cast<uint64_t>(sizeof(IMAGE_DOS_HEADER))) {
        CloseHandle(file);
        return false;
    }
    HANDLE mapping = CreateFileMappingW(file, nullptr, PAGE_READONLY, 0, 0, nullptr);
    if (!mapping) {
        CloseHandle(file);
        return false;
    }
    void* view = MapViewOfFile(mapping, FILE_MAP_READ, 0, 0, 0);
    if (!view) {
        CloseHandle(mapping);
        CloseHandle(file);
        return false;
    }

    bool ok = false;
    const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(view);
    if (dos->e_magic == IMAGE_DOS_SIGNATURE &&
        dos->e_lfanew > 0) {
        const uint64_t ntOffset = static_cast<uint64_t>(dos->e_lfanew);
        const uint64_t ntEnd = ntOffset + static_cast<uint64_t>(sizeof(IMAGE_NT_HEADERS64));
        if (ntOffset < bytes && ntEnd <= bytes) {
        const auto* nt = reinterpret_cast<const IMAGE_NT_HEADERS64*>(
            reinterpret_cast<const BYTE*>(view) + dos->e_lfanew);
        if (nt->Signature == IMAGE_NT_SIGNATURE &&
            nt->FileHeader.Machine == IMAGE_FILE_MACHINE_AMD64 &&
            nt->OptionalHeader.Magic == IMAGE_NT_OPTIONAL_HDR64_MAGIC) {
            ok = true;
        }
        }
    }

    UnmapViewOfFile(view);
    CloseHandle(mapping);
    CloseHandle(file);
    return ok;
}

bool IsDigitsOnly(const std::string& value) {
    if (value.empty()) {
        return false;
    }
    return std::all_of(value.begin(), value.end(), [](char ch) {
        return std::isdigit(static_cast<unsigned char>(ch)) != 0;
    });
}

void SanitizeSettings(MicMixSettings& s) {
    s.configVersion = std::clamp(s.configVersion, 1, 8);
    if (!std::isfinite(s.musicGainDb)) { s.musicGainDb = -15.0f; }
    s.musicGainDb = std::clamp(s.musicGainDb, -30.0f, -2.0f);
    if (s.resamplerQuality < -1) {
        s.resamplerQuality = -1;
    } else if (s.resamplerQuality > 10) {
        s.resamplerQuality = 10;
    }
    s.bufferTargetMs = std::clamp(s.bufferTargetMs, 20, 250);
    if (!std::isfinite(s.micGateThresholdDbfs)) { s.micGateThresholdDbfs = -50.0f; }
    s.micGateThresholdDbfs = std::clamp(s.micGateThresholdDbfs, -90.0f, 0.0f);
    s.muteHotkeyModifiers &= (MOD_ALT | MOD_CONTROL | MOD_SHIFT | MOD_WIN);
    s.muteHotkeyVk = std::clamp(s.muteHotkeyVk, 0, 255);
    s.micInputMuteHotkeyModifiers &= (MOD_ALT | MOD_CONTROL | MOD_SHIFT | MOD_WIN);
    s.micInputMuteHotkeyVk = std::clamp(s.micInputMuteHotkeyVk, 0, 255);
    if (s.muteHotkeyVk != 0 &&
        s.muteHotkeyVk == s.micInputMuteHotkeyVk &&
        s.muteHotkeyModifiers == s.micInputMuteHotkeyModifiers) {
        // Keep music mute as primary and clear conflicting mic-input hotkey.
        s.micInputMuteHotkeyModifiers = 0;
        s.micInputMuteHotkeyVk = 0;
    }
    s.uiLastOpenTab = std::max(0, s.uiLastOpenTab);
    if (s.micGateMode != MicGateMode::AutoTs && s.micGateMode != MicGateMode::Custom) {
        s.micGateMode = MicGateMode::AutoTs;
    }

    StripConfigControlChars(s.loopbackDeviceId, 512);
    StripConfigControlChars(s.appProcessName, 128);
    StripConfigControlChars(s.appSessionId, 32);
    StripConfigControlChars(s.captureDeviceId, 512);

    if (!s.appSessionId.empty() && !IsDigitsOnly(s.appSessionId)) {
        s.appSessionId.clear();
    }

    MicMixApp::SanitizeEffectList(s.musicEffects);
    MicMixApp::SanitizeEffectList(s.micEffects);
}

void UpdateSlotStatusForUi(VstEffectSlot& slot, bool hostSynchronized) {
    if (!slot.enabled) {
        slot.lastStatus = "disabled";
        return;
    }
    if (slot.bypass) {
        slot.lastStatus = "bypassed";
        return;
    }
    slot.lastStatus = hostSynchronized ? "running" : "pending";
}

void UpdateAllSlotStatusesForUi(MicMixSettings& settings, bool hostSynchronized) {
    for (auto& slot : settings.musicEffects) {
        UpdateSlotStatusForUi(slot, hostSynchronized);
    }
    for (auto& slot : settings.micEffects) {
        UpdateSlotStatusForUi(slot, hostSynchronized);
    }
}

bool NearlyEqual(float a, float b) {
    return std::fabs(a - b) < 0.0001f;
}

bool IsSameEffectSlots(const std::vector<VstEffectSlot>& a, const std::vector<VstEffectSlot>& b) {
    if (a.size() != b.size()) {
        return false;
    }
    for (size_t i = 0; i < a.size(); ++i) {
        if (a[i].path != b[i].path ||
            a[i].name != b[i].name ||
            a[i].uid != b[i].uid ||
            a[i].stateBlob != b[i].stateBlob ||
            a[i].lastStatus != b[i].lastStatus ||
            a[i].enabled != b[i].enabled ||
            a[i].bypass != b[i].bypass) {
            return false;
        }
    }
    return true;
}

bool IsSameSettings(const MicMixSettings& a, const MicMixSettings& b) {
    return a.configVersion == b.configVersion &&
           a.sourceMode == b.sourceMode &&
           a.loopbackDeviceId == b.loopbackDeviceId &&
           a.appProcessName == b.appProcessName &&
           a.appSessionId == b.appSessionId &&
           a.autostartEnabled == b.autostartEnabled &&
           NearlyEqual(a.musicGainDb, b.musicGainDb) &&
           a.resamplerQuality == b.resamplerQuality &&
           a.forceTxEnabled == b.forceTxEnabled &&
           a.bufferTargetMs == b.bufferTargetMs &&
           a.musicMuted == b.musicMuted &&
           a.micInputMuted == b.micInputMuted &&
           a.uiLastOpenTab == b.uiLastOpenTab &&
           a.autoSwitchToLoopback == b.autoSwitchToLoopback &&
           a.muteHotkeyModifiers == b.muteHotkeyModifiers &&
           a.muteHotkeyVk == b.muteHotkeyVk &&
           a.micInputMuteHotkeyModifiers == b.micInputMuteHotkeyModifiers &&
           a.micInputMuteHotkeyVk == b.micInputMuteHotkeyVk &&
           a.captureDeviceId == b.captureDeviceId &&
           a.micGateMode == b.micGateMode &&
           NearlyEqual(a.micGateThresholdDbfs, b.micGateThresholdDbfs) &&
           a.micUseSmoothGate == b.micUseSmoothGate &&
           a.micUseKeyboardGuard == b.micUseKeyboardGuard &&
           a.micForceTsFilters == b.micForceTsFilters &&
           a.vstEffectsEnabled == b.vstEffectsEnabled &&
           a.vstHostAutostart == b.vstHostAutostart &&
           IsSameEffectSlots(a.musicEffects, b.musicEffects) &&
           IsSameEffectSlots(a.micEffects, b.micEffects);
}

bool IsBlockedUiProcess(uint32_t pid, const std::wstring& exeName) {
    if (pid <= 4) {
        return true;
    }
    static constexpr const wchar_t* kBlocked[] = {
        L"System",
        L"System Idle Process",
        L"[System Process]",
        L"Registry",
        L"Secure System",
        L"smss.exe",
        L"csrss.exe",
        L"wininit.exe",
        L"services.exe",
        L"lsass.exe",
        L"svchost.exe",
        L"audiodg.exe",
        L"ApplicationFrameHost.exe",
        L"RuntimeBroker.exe",
        L"SearchHost.exe",
        L"StartMenuExperienceHost.exe",
        L"TextInputHost.exe",
        L"ShellExperienceHost.exe",
        L"msedgewebview2.exe",
    };
    for (const auto* name : kBlocked) {
        if (_wcsicmp(exeName.c_str(), name) == 0) {
            return true;
        }
    }
    return false;
}

std::string DisplayNameFromExe(const std::string& exeName) {
    if (exeName.size() > 4) {
        const char* suffix = exeName.c_str() + (exeName.size() - 4);
        if (_stricmp(suffix, ".exe") == 0) {
            return exeName.substr(0, exeName.size() - 4);
        }
    }
    return exeName;
}

std::wstring ToLowerWide(std::wstring value) {
    std::transform(value.begin(), value.end(), value.begin(), [](wchar_t ch) { return static_cast<wchar_t>(std::towlower(ch)); });
    return value;
}

std::string ToLowerAscii(std::string value) {
    std::transform(value.begin(), value.end(), value.begin(), [](char ch) { return static_cast<char>(std::tolower(static_cast<unsigned char>(ch))); });
    return value;
}

bool StartsWithAscii(const std::string& value, const char* prefix) {
    if (!prefix) {
        return false;
    }
    const size_t prefixLen = std::strlen(prefix);
    if (value.size() < prefixLen) {
        return false;
    }
    return std::equal(prefix, prefix + prefixLen, value.begin());
}

std::string HexEncode(const std::string& input) {
    static constexpr char kHex[] = "0123456789ABCDEF";
    std::string out;
    out.reserve(input.size() * 2U);
    for (unsigned char ch : input) {
        out.push_back(kHex[(ch >> 4) & 0x0F]);
        out.push_back(kHex[ch & 0x0F]);
    }
    return out;
}

int HexNibble(char c) {
    if (c >= '0' && c <= '9') return c - '0';
    if (c >= 'a' && c <= 'f') return 10 + (c - 'a');
    if (c >= 'A' && c <= 'F') return 10 + (c - 'A');
    return -1;
}

bool HexDecode(const std::string& hex, std::string& out) {
    out.clear();
    if ((hex.size() % 2U) != 0U) {
        return false;
    }
    out.reserve(hex.size() / 2U);
    for (size_t i = 0; i < hex.size(); i += 2U) {
        const int hi = HexNibble(hex[i]);
        const int lo = HexNibble(hex[i + 1U]);
        if (hi < 0 || lo < 0) {
            out.clear();
            return false;
        }
        out.push_back(static_cast<char>((hi << 4) | lo));
    }
    return true;
}

bool NormalizeAndValidateEffectPath(const std::string& rawPath, std::string& normalizedPath, std::string& error);

std::string BuildHostSyncPayload(const MicMixSettings& settings) {
    auto sanitizeChain = [](const std::vector<VstEffectSlot>& slots) {
        std::vector<VstEffectSlot> sanitized;
        sanitized.reserve(slots.size());
        std::unordered_set<std::string> usedUids;
        for (const auto& slot : slots) {
            std::string normalizedPath;
            std::string pathError;
            if (!NormalizeAndValidateEffectPath(slot.path, normalizedPath, pathError)) {
                continue;
            }
            VstEffectSlot safe = slot;
            safe.path = normalizedPath;
            if (safe.uid.empty()) {
                safe.uid = BuildUniqueEffectUidWithExists(
                    safe.path,
                    [&usedUids](const std::string& uid) { return usedUids.find(uid) != usedUids.end(); });
            } else if (usedUids.find(safe.uid) != usedUids.end()) {
                safe.uid = BuildUniqueEffectUidWithExists(
                    safe.path,
                    [&usedUids](const std::string& uid) { return usedUids.find(uid) != usedUids.end(); });
            }
            usedUids.insert(safe.uid);
            sanitized.push_back(std::move(safe));
        }
        return sanitized;
    };

    const auto safeMusic = sanitizeChain(settings.musicEffects);
    const auto safeMic = sanitizeChain(settings.micEffects);

    std::ostringstream ss;
    ss << "effects_enabled=" << (settings.vstEffectsEnabled ? 1 : 0) << "\n";
    ss << "music.count=" << safeMusic.size() << "\n";
    for (size_t i = 0; i < safeMusic.size(); ++i) {
        const auto& slot = safeMusic[i];
        ss << "music." << i << ".path=" << slot.path << "\n";
        ss << "music." << i << ".name=" << slot.name << "\n";
        ss << "music." << i << ".uid=" << slot.uid << "\n";
        ss << "music." << i << ".last_status=" << slot.lastStatus << "\n";
        ss << "music." << i << ".enabled=" << (slot.enabled ? 1 : 0) << "\n";
        ss << "music." << i << ".bypass=" << (slot.bypass ? 1 : 0) << "\n";
    }
    ss << "mic.count=" << safeMic.size() << "\n";
    for (size_t i = 0; i < safeMic.size(); ++i) {
        const auto& slot = safeMic[i];
        ss << "mic." << i << ".path=" << slot.path << "\n";
        ss << "mic." << i << ".name=" << slot.name << "\n";
        ss << "mic." << i << ".uid=" << slot.uid << "\n";
        ss << "mic." << i << ".last_status=" << slot.lastStatus << "\n";
        ss << "mic." << i << ".enabled=" << (slot.enabled ? 1 : 0) << "\n";
        ss << "mic." << i << ".bypass=" << (slot.bypass ? 1 : 0) << "\n";
    }
    return ss.str();
}

bool NormalizeAndValidateEffectPath(const std::string& rawPath, std::string& normalizedPath, std::string& error) {
    error.clear();
    normalizedPath.clear();
    if (rawPath.empty()) {
        error = "effect path missing";
        return false;
    }
    std::filesystem::path path(Utf8ToWide(rawPath));
    if (path.empty()) {
        error = "effect path missing";
        return false;
    }
    if (!path.is_absolute()) {
        error = "effect path must be absolute";
        return false;
    }
    const std::wstring rawWide = path.wstring();
    if (rawWide.rfind(L"\\\\", 0) == 0) {
        error = "network paths are not allowed";
        return false;
    }
    std::error_code ec;
    std::filesystem::path canonical = std::filesystem::weakly_canonical(path, ec);
    if (ec) {
        canonical = path.lexically_normal();
        ec.clear();
    }
    if (canonical.empty()) {
        error = "invalid effect path";
        return false;
    }
    const std::wstring ext = ToLowerWide(canonical.extension().wstring());
    if (_wcsicmp(ext.c_str(), L".vst3") != 0) {
        error = "unsupported effect file extension (only .vst3)";
        return false;
    }
    if (!std::filesystem::exists(canonical, ec) || ec) {
        error = "effect path missing";
        return false;
    }
    const DWORD attrs = GetFileAttributesW(canonical.c_str());
    if (attrs == INVALID_FILE_ATTRIBUTES) {
        error = "effect file attributes unavailable";
        return false;
    }
    if ((attrs & FILE_ATTRIBUTE_DIRECTORY) != 0) {
        error = "effect path must point to a file";
        return false;
    }
    if ((attrs & FILE_ATTRIBUTE_REPARSE_POINT) != 0) {
        error = "reparse-point plugin paths are blocked";
        return false;
    }
    if (!IsAmd64PeBinary(canonical)) {
        error = "unsupported plugin binary architecture (x64 required)";
        return false;
    }
    normalizedPath = WideToUtf8(canonical.wstring());
    if (normalizedPath.empty()) {
        error = "invalid normalized effect path";
        return false;
    }
    return true;
}

bool IsSameSession(uint32_t pid, DWORD expectedSessionId) {
    DWORD sessionId = 0;
    if (!ProcessIdToSessionId(pid, &sessionId)) {
        return false;
    }
    return sessionId == expectedSessionId;
}

bool TryGetProcessImagePath(uint32_t pid, std::wstring& outPath) {
    outPath.clear();
    HANDLE process = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (!process) {
        return false;
    }
    std::wstring buffer(MAX_PATH, L'\0');
    DWORD size = static_cast<DWORD>(buffer.size());
    bool ok = false;
    while (QueryFullProcessImageNameW(process, 0, buffer.data(), &size) == FALSE) {
        if (GetLastError() != ERROR_INSUFFICIENT_BUFFER || buffer.size() >= 32768) {
            CloseHandle(process);
            return false;
        }
        buffer.resize(buffer.size() * 2);
        size = static_cast<DWORD>(buffer.size());
    }
    buffer.resize(size);
    outPath = std::move(buffer);
    ok = true;
    CloseHandle(process);
    return ok;
}

std::wstring GetWindowsDirLower() {
    wchar_t path[MAX_PATH]{};
    const UINT len = GetWindowsDirectoryW(path, static_cast<UINT>(std::size(path)));
    if (len == 0 || len >= std::size(path)) {
        return L"\\windows";
    }
    return ToLowerWide(std::wstring(path, path + len));
}

bool IsLikelyUserAppPath(const std::wstring& imagePathLower, const std::wstring& windowsDirLower) {
    if (imagePathLower.empty()) {
        return false;
    }
    if (imagePathLower.rfind(windowsDirLower, 0) == 0) {
        return false;
    }
    if (imagePathLower.find(L"\\program files\\") != std::wstring::npos ||
        imagePathLower.find(L"\\program files (x86)\\") != std::wstring::npos ||
        imagePathLower.find(L"\\users\\") != std::wstring::npos) {
        return true;
    }
    return false;
}

bool IsPreferredMediaProcessName(const std::string& exeNameLower) {
    static const std::unordered_set<std::string> kPreferred = {
        "spotify.exe",
        "chrome.exe",
        "msedge.exe",
        "firefox.exe",
        "brave.exe",
        "opera.exe",
        "discord.exe",
        "vlc.exe",
        "foobar2000.exe",
    };
    return kPreferred.find(exeNameLower) != kPreferred.end();
}

struct WindowPidEnumContext {
    std::unordered_set<uint32_t>* pids = nullptr;
};

BOOL CALLBACK EnumVisibleWindowPidsProc(HWND hwnd, LPARAM lParam) {
    auto* ctx = reinterpret_cast<WindowPidEnumContext*>(lParam);
    if (!ctx || !ctx->pids) {
        return FALSE;
    }
    if (!IsWindowVisible(hwnd)) {
        return TRUE;
    }
    if (GetWindowTextLengthW(hwnd) <= 0) {
        return TRUE;
    }
    if (GetWindow(hwnd, GW_OWNER) != nullptr) {
        return TRUE;
    }
    const LONG_PTR exStyle = GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
    if ((exStyle & WS_EX_TOOLWINDOW) != 0) {
        return TRUE;
    }
    DWORD pid = 0;
    GetWindowThreadProcessId(hwnd, &pid);
    if (pid > 4) {
        ctx->pids->insert(static_cast<uint32_t>(pid));
    }
    return TRUE;
}

std::unordered_set<uint32_t> EnumerateVisibleWindowPids() {
    std::unordered_set<uint32_t> out;
    WindowPidEnumContext ctx{ &out };
    EnumWindows(EnumVisibleWindowPidsProc, reinterpret_cast<LPARAM>(&ctx));
    return out;
}

std::unordered_set<uint32_t> EnumerateAudioSessionPids() {
    std::unordered_set<uint32_t> out;
    ComInit com;
    if (FAILED(com.hr)) {
        return out;
    }
    ComPtr<IMMDeviceEnumerator> enumerator;
    if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&enumerator))) || !enumerator) {
        return out;
    }
    ComPtr<IMMDeviceCollection> devices;
    if (FAILED(enumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &devices)) || !devices) {
        return out;
    }
    UINT deviceCount = 0;
    if (FAILED(devices->GetCount(&deviceCount))) {
        return out;
    }
    for (UINT i = 0; i < deviceCount; ++i) {
        ComPtr<IMMDevice> device;
        if (FAILED(devices->Item(i, &device)) || !device) {
            continue;
        }
        ComPtr<IAudioSessionManager2> manager;
        if (FAILED(device->Activate(__uuidof(IAudioSessionManager2), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(manager.GetAddressOf()))) || !manager) {
            continue;
        }
        ComPtr<IAudioSessionEnumerator> sessions;
        if (FAILED(manager->GetSessionEnumerator(&sessions)) || !sessions) {
            continue;
        }
        int count = 0;
        if (FAILED(sessions->GetCount(&count))) {
            continue;
        }
        for (int s = 0; s < count; ++s) {
            ComPtr<IAudioSessionControl> control;
            if (FAILED(sessions->GetSession(s, &control)) || !control) {
                continue;
            }
            ComPtr<IAudioSessionControl2> control2;
            if (FAILED(control.As(&control2)) || !control2) {
                continue;
            }
            DWORD pid = 0;
            if (SUCCEEDED(control2->GetProcessId(&pid)) && pid > 4) {
                out.insert(static_cast<uint32_t>(pid));
            }
        }
    }
    return out;
}

bool IsFloatFormat(const WAVEFORMATEX* wf) {
    if (wf->wFormatTag == WAVE_FORMAT_IEEE_FLOAT && wf->wBitsPerSample == 32) {
        return true;
    }
    if (wf->wFormatTag == WAVE_FORMAT_EXTENSIBLE && wf->cbSize >= 22) {
        auto* ext = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(wf);
        return ext->SubFormat == KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
    }
    return false;
}

bool IsPcm16Format(const WAVEFORMATEX* wf) {
    if (wf->wFormatTag == WAVE_FORMAT_PCM && wf->wBitsPerSample == 16) {
        return true;
    }
    if (wf->wFormatTag == WAVE_FORMAT_EXTENSIBLE && wf->cbSize >= 22) {
        auto* ext = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(wf);
        return ext->SubFormat == KSDATAFORMAT_SUBTYPE_PCM && wf->wBitsPerSample == 16;
    }
    return false;
}

class SpeexResamplerWrap {
public:
    ~SpeexResamplerWrap() {
        if (state_) {
            speex_resampler_destroy(state_);
            state_ = nullptr;
        }
    }

    bool Configure(uint32_t inRate, uint32_t outRate, int quality) {
        if (!IsSupportedSourceRate(inRate) || !IsSupportedSourceRate(outRate)) {
            return false;
        }
        if (state_ && inRate_ == inRate && outRate_ == outRate && quality_ == quality) {
            return true;
        }
        if (state_) {
            speex_resampler_destroy(state_);
            state_ = nullptr;
        }
        int err = RESAMPLER_ERR_SUCCESS;
        state_ = speex_resampler_init(1, inRate, outRate, quality, &err);
        if (!state_ || err != RESAMPLER_ERR_SUCCESS) {
            return false;
        }
        inRate_ = inRate;
        outRate_ = outRate;
        quality_ = quality;
        return true;
    }

    bool Process(const float* in, size_t inSamples, std::vector<float>& out) {
        if (!state_) {
            return false;
        }
        spx_uint32_t inLen = static_cast<spx_uint32_t>(inSamples);
        spx_uint32_t outLen = static_cast<spx_uint32_t>((static_cast<double>(inSamples) * outRate_ / inRate_) + 64);
        out.resize(outLen);
        const int res = speex_resampler_process_float(state_, 0, in, &inLen, out.data(), &outLen);
        if (res != RESAMPLER_ERR_SUCCESS) {
            out.clear();
            return false;
        }
        out.resize(outLen);
        return true;
    }

private:
    SpeexResamplerState* state_ = nullptr;
    uint32_t inRate_ = 0;
    uint32_t outRate_ = 0;
    int quality_ = kResamplerQuality;
};

} // namespace

void LogInfo(const std::string& text, uint64 schid) {
    if (g_tsLogEnabled.load(std::memory_order_acquire) && g_ts3Functions.logMessage) {
        g_ts3Functions.logMessage(text.c_str(), LogLevel_INFO, kLogChannel, schid);
    }
    AppendLogLine("INFO", text);
}

void LogWarn(const std::string& text, uint64 schid) {
    if (g_tsLogEnabled.load(std::memory_order_acquire) && g_ts3Functions.logMessage) {
        g_ts3Functions.logMessage(text.c_str(), LogLevel_WARNING, kLogChannel, schid);
    }
    AppendLogLine("WARN", text);
}

void LogError(const std::string& text, uint64 schid) {
    if (g_tsLogEnabled.load(std::memory_order_acquire) && g_ts3Functions.logMessage) {
        g_ts3Functions.logMessage(text.c_str(), LogLevel_ERROR, kLogChannel, schid);
    }
    AppendLogLine("ERROR", text);
}

void SetTsLoggingEnabled(bool enabled) {
    g_tsLogEnabled.store(enabled, std::memory_order_release);
}

ConfigStore::ConfigStore(std::string basePath)
    : basePath_(std::move(basePath)) {
    std::replace(basePath_.begin(), basePath_.end(), '/', '\\');
    if (!basePath_.empty() && basePath_.back() != '\\') {
        basePath_.push_back('\\');
    }
    const std::filesystem::path dir = std::filesystem::path(Utf8ToWide(basePath_)) / L"plugins" / L"micmix";
    std::error_code ec;
    std::filesystem::create_directories(dir, ec);
    if (ec) {
        std::error_code existsEc;
        if (!std::filesystem::exists(dir, existsEc)) {
            LogWarn("Config directory create failed: " + WideToUtf8(dir.wstring()) + " (" + ec.message() + ")", 0);
        }
    }
    configDir_ = WideToUtf8(dir.wstring());
    configPath_ = WideToUtf8((dir / L"config.ini").wstring());
    legacyConfigPath_ = WideToUtf8((dir / L"config.json").wstring());
    tmpPath_ = WideToUtf8((dir / L"config.tmp").wstring());
    lastGoodPath_ = WideToUtf8((dir / L"config.lastgood.ini").wstring());
    logPath_ = WideToUtf8((dir / L"micmix.log").wstring());
}

std::string ConfigStore::Trim(const std::string& value) {
    return ::Trim(value);
}

std::string ConfigStore::BoolToString(bool value) {
    return value ? "true" : "false";
}

std::string ConfigStore::SourceModeToString(SourceMode mode) {
    return mode == SourceMode::AppSession ? "app_session" : "loopback";
}

SourceMode ConfigStore::SourceModeFromString(const std::string& value) {
    if (value == "app_session" || value == "spotify_session") {
        return SourceMode::AppSession;
    }
    return SourceMode::Loopback;
}

std::string ConfigStore::MicGateModeToString(MicGateMode mode) {
    return mode == MicGateMode::Custom ? "custom" : "auto_ts";
}

MicGateMode ConfigStore::MicGateModeFromString(const std::string& value) {
    if (value == "custom" || value == "custom_threshold") {
        return MicGateMode::Custom;
    }
    return MicGateMode::AutoTs;
}

bool ConfigStore::Load(MicMixSettings& outSettings, std::string& warning) {
    warning.clear();
    auto appendWarning = [&](const char* msg) {
        if (!msg || msg[0] == '\0') {
            return;
        }
        if (!warning.empty()) {
            warning += " ";
        }
        warning += msg;
    };

    std::filesystem::path loadPath = std::filesystem::path(Utf8ToWide(configPath_));
    std::ifstream in(loadPath);
    bool usedLegacyConfig = false;
    if (!in.is_open()) {
        in.clear();
        loadPath = std::filesystem::path(Utf8ToWide(legacyConfigPath_));
        in.open(loadPath);
        usedLegacyConfig = in.is_open();
        if (!usedLegacyConfig) {
            return true;
        }
        appendWarning("Loaded legacy config.json file; settings will migrate to config.ini on next save.");
    }

    std::error_code sizeEc;
    const uintmax_t sizeBytes = std::filesystem::file_size(loadPath, sizeEc);
    if (sizeEc) {
        appendWarning("Config size check failed; file ignored.");
        return true;
    }
    if (sizeBytes > kMaxConfigBytes) {
        appendWarning("Config too large; file ignored.");
        return true;
    }

    std::unordered_map<std::string, std::string> kv;
    std::string payload;
    bool payloadTooLarge = false;
    if (!ReadPayloadWithLimit(in, static_cast<size_t>(kMaxConfigBytes), payload, payloadTooLarge)) {
        appendWarning("Config read failed; file ignored.");
        return true;
    }
    if (payloadTooLarge) {
        appendWarning("Config too large; file ignored.");
        return true;
    }
    bool parseIssue = false;
    bool parsed = false;
    if (LooksLikeJsonObjectPayload(payload)) {
        // Prefer strict JSON parsing for JSON-shaped payloads so values that
        // contain '=' are not accidentally interpreted as INI pairs.
        parsed = ParseJsonObjectConfigPayload(payload, kv);
        if (!parsed && HasNonWhitespace(payload)) {
            parseIssue = true;
        }
    } else {
        parsed = ParseIniLikeConfigPayload(payload, kv);
        if (!parsed) {
            parsed = ParseJsonObjectConfigPayload(payload, kv);
        }
        if (!parsed && HasNonWhitespace(payload)) {
            parseIssue = true;
        }
    }
    auto parseInt = [&](const char* key, int& outValue) {
        if (auto it = kv.find(key); it != kv.end()) {
            try {
                outValue = std::stoi(it->second);
            } catch (...) {
                parseIssue = true;
            }
        }
    };
    auto parseFloat = [&](const char* key, float& outValue) {
        if (auto it = kv.find(key); it != kv.end()) {
            try {
                outValue = std::stof(it->second);
            } catch (...) {
                parseIssue = true;
            }
        }
    };

    parseInt("config.version", outSettings.configVersion);
    if (auto it = kv.find("source.mode"); it != kv.end()) outSettings.sourceMode = SourceModeFromString(it->second);
    if (auto it = kv.find("source.loopback.device_id"); it != kv.end()) outSettings.loopbackDeviceId = it->second;
    if (auto it = kv.find("source.app.process_name"); it != kv.end()) {
        outSettings.appProcessName = it->second;
    } else if (auto itLegacy = kv.find("source.spotify.process_name"); itLegacy != kv.end()) {
        outSettings.appProcessName = itLegacy->second;
    }
    if (auto it = kv.find("source.app.session_id"); it != kv.end()) {
        outSettings.appSessionId = it->second;
    } else if (auto itLegacy = kv.find("source.spotify.session_id"); itLegacy != kv.end()) {
        outSettings.appSessionId = itLegacy->second;
    }
    if (auto it = kv.find("source.autostart_enabled"); it != kv.end()) outSettings.autostartEnabled = ::ParseBool(it->second, outSettings.autostartEnabled);
    if (auto it = kv.find("source.auto_switch_to_loopback"); it != kv.end()) outSettings.autoSwitchToLoopback = ::ParseBool(it->second, outSettings.autoSwitchToLoopback);
    parseFloat("mix.music_gain_db", outSettings.musicGainDb);
    parseInt("mix.resampler_quality", outSettings.resamplerQuality);
    if (auto it = kv.find("mix.force_tx_enabled"); it != kv.end()) outSettings.forceTxEnabled = ::ParseBool(it->second, outSettings.forceTxEnabled);
    parseInt("mix.buffer_target_ms", outSettings.bufferTargetMs);
    if (auto it = kv.find("mix.music_muted"); it != kv.end()) outSettings.musicMuted = ::ParseBool(it->second, outSettings.musicMuted);
    if (auto it = kv.find("mic.input_muted"); it != kv.end()) outSettings.micInputMuted = ::ParseBool(it->second, outSettings.micInputMuted);
    parseInt("hotkey.mute.modifiers", outSettings.muteHotkeyModifiers);
    parseInt("hotkey.mute.vk", outSettings.muteHotkeyVk);
    parseInt("hotkey.mic_input_mute.modifiers", outSettings.micInputMuteHotkeyModifiers);
    parseInt("hotkey.mic_input_mute.vk", outSettings.micInputMuteHotkeyVk);
    if (auto it = kv.find("capture.device_id"); it != kv.end()) outSettings.captureDeviceId = it->second;
    if (auto it = kv.find("mic.gate.mode"); it != kv.end()) outSettings.micGateMode = MicGateModeFromString(it->second);
    parseFloat("mic.gate.threshold_dbfs", outSettings.micGateThresholdDbfs);
    if (auto it = kv.find("mic.gate.smooth"); it != kv.end()) outSettings.micUseSmoothGate = ::ParseBool(it->second, outSettings.micUseSmoothGate);
    if (auto it = kv.find("mic.gate.keyboard_guard"); it != kv.end()) outSettings.micUseKeyboardGuard = ::ParseBool(it->second, outSettings.micUseKeyboardGuard);
    if (auto it = kv.find("mic.force_ts_filters"); it != kv.end()) outSettings.micForceTsFilters = ::ParseBool(it->second, outSettings.micForceTsFilters);
    if (auto it = kv.find("vst.effects_enabled"); it != kv.end()) outSettings.vstEffectsEnabled = ::ParseBool(it->second, outSettings.vstEffectsEnabled);
    if (auto it = kv.find("vst.host.autostart"); it != kv.end()) outSettings.vstHostAutostart = ::ParseBool(it->second, outSettings.vstHostAutostart);
    auto parseEffectChain = [&](const char* chainPrefix, std::vector<VstEffectSlot>& outList) {
        outList.clear();
        int count = -1;
        parseInt((std::string(chainPrefix) + ".count").c_str(), count);
        count = std::clamp(count, -1, static_cast<int>(kMaxEffectsPerChain));
        auto parseAt = [&](int index) {
            const std::string base = std::string(chainPrefix) + "." + std::to_string(index);
            const std::string keyPath = base + ".path";
            const auto pathIt = kv.find(keyPath);
            if (pathIt == kv.end()) {
                return false;
            }
            VstEffectSlot slot{};
            slot.path = pathIt->second;
            if (auto it = kv.find(base + ".name"); it != kv.end()) {
                slot.name = it->second;
            }
            if (auto it = kv.find(base + ".uid"); it != kv.end()) {
                slot.uid = it->second;
            }
            if (auto it = kv.find(base + ".state_blob"); it != kv.end()) {
                slot.stateBlob = it->second;
            }
            if (auto it = kv.find(base + ".last_status"); it != kv.end()) {
                slot.lastStatus = it->second;
            }
            if (auto it = kv.find(base + ".enabled"); it != kv.end()) {
                slot.enabled = ::ParseBool(it->second, true);
            }
            if (auto it = kv.find(base + ".bypass"); it != kv.end()) {
                slot.bypass = ::ParseBool(it->second, false);
            }
            outList.push_back(std::move(slot));
            return true;
        };
        if (count >= 0) {
            for (int i = 0; i < count; ++i) {
                parseAt(i);
            }
            return;
        }
        for (int i = 0; i < static_cast<int>(kMaxEffectsPerChain); ++i) {
            if (!parseAt(i)) {
                break;
            }
        }
    };
    parseEffectChain("vst.music", outSettings.musicEffects);
    parseEffectChain("vst.mic", outSettings.micEffects);
    parseInt("ui.last_open_tab", outSettings.uiLastOpenTab);
    if (parseIssue) {
        appendWarning("Config parse issue detected; fallback values were used.");
    }
    SanitizeSettings(outSettings);
    return true;
}

bool ConfigStore::Save(const MicMixSettings& settings, std::string& error) {
    error.clear();
    MicMixSettings safe = settings;
    SanitizeSettings(safe);
    std::ostringstream ss;
    ss << "config.version=" << safe.configVersion << "\n";
    ss << "source.mode=" << SourceModeToString(safe.sourceMode) << "\n";
    ss << "source.loopback.device_id=" << safe.loopbackDeviceId << "\n";
    ss << "source.app.process_name=" << safe.appProcessName << "\n";
    ss << "source.app.session_id=" << safe.appSessionId << "\n";
    ss << "source.autostart_enabled=" << BoolToString(safe.autostartEnabled) << "\n";
    ss << "source.auto_switch_to_loopback=" << BoolToString(safe.autoSwitchToLoopback) << "\n";
    ss << "mix.music_gain_db=" << safe.musicGainDb << "\n";
    ss << "mix.resampler_quality=" << safe.resamplerQuality << "\n";
    ss << "mix.force_tx_enabled=" << BoolToString(safe.forceTxEnabled) << "\n";
    ss << "mix.buffer_target_ms=" << safe.bufferTargetMs << "\n";
    ss << "mix.music_muted=" << BoolToString(safe.musicMuted) << "\n";
    ss << "mic.input_muted=" << BoolToString(safe.micInputMuted) << "\n";
    ss << "hotkey.mute.modifiers=" << safe.muteHotkeyModifiers << "\n";
    ss << "hotkey.mute.vk=" << safe.muteHotkeyVk << "\n";
    ss << "hotkey.mic_input_mute.modifiers=" << safe.micInputMuteHotkeyModifiers << "\n";
    ss << "hotkey.mic_input_mute.vk=" << safe.micInputMuteHotkeyVk << "\n";
    ss << "capture.device_id=" << safe.captureDeviceId << "\n";
    ss << "mic.gate.mode=" << MicGateModeToString(safe.micGateMode) << "\n";
    ss << "mic.gate.threshold_dbfs=" << safe.micGateThresholdDbfs << "\n";
    ss << "mic.gate.smooth=" << BoolToString(safe.micUseSmoothGate) << "\n";
    ss << "mic.gate.keyboard_guard=" << BoolToString(safe.micUseKeyboardGuard) << "\n";
    ss << "mic.force_ts_filters=" << BoolToString(safe.micForceTsFilters) << "\n";
    ss << "vst.effects_enabled=" << BoolToString(safe.vstEffectsEnabled) << "\n";
    ss << "vst.host.autostart=" << BoolToString(safe.vstHostAutostart) << "\n";
    ss << "vst.music.count=" << safe.musicEffects.size() << "\n";
    for (size_t i = 0; i < safe.musicEffects.size(); ++i) {
        const auto& slot = safe.musicEffects[i];
        ss << "vst.music." << i << ".path=" << slot.path << "\n";
        ss << "vst.music." << i << ".name=" << slot.name << "\n";
        ss << "vst.music." << i << ".uid=" << slot.uid << "\n";
        ss << "vst.music." << i << ".state_blob=" << slot.stateBlob << "\n";
        ss << "vst.music." << i << ".last_status=" << slot.lastStatus << "\n";
        ss << "vst.music." << i << ".enabled=" << BoolToString(slot.enabled) << "\n";
        ss << "vst.music." << i << ".bypass=" << BoolToString(slot.bypass) << "\n";
    }
    ss << "vst.mic.count=" << safe.micEffects.size() << "\n";
    for (size_t i = 0; i < safe.micEffects.size(); ++i) {
        const auto& slot = safe.micEffects[i];
        ss << "vst.mic." << i << ".path=" << slot.path << "\n";
        ss << "vst.mic." << i << ".name=" << slot.name << "\n";
        ss << "vst.mic." << i << ".uid=" << slot.uid << "\n";
        ss << "vst.mic." << i << ".state_blob=" << slot.stateBlob << "\n";
        ss << "vst.mic." << i << ".last_status=" << slot.lastStatus << "\n";
        ss << "vst.mic." << i << ".enabled=" << BoolToString(slot.enabled) << "\n";
        ss << "vst.mic." << i << ".bypass=" << BoolToString(slot.bypass) << "\n";
    }
    ss << "ui.last_open_tab=" << safe.uiLastOpenTab << "\n";

    {
        std::ofstream out(std::filesystem::path(Utf8ToWide(tmpPath_)), std::ios::binary | std::ios::trunc);
        if (!out.is_open()) {
            error = "open config tmp failed";
            return false;
        }
        const std::string payload = ss.str();
        out.write(payload.data(), static_cast<std::streamsize>(payload.size()));
        out.flush();
        if (!out.good()) {
            error = "write config tmp failed";
            return false;
        }
    }

    const std::wstring configW = Utf8ToWide(configPath_);
    const std::wstring legacyConfigW = Utf8ToWide(legacyConfigPath_);
    const std::wstring tmpW = Utf8ToWide(tmpPath_);
    const std::wstring lastGoodW = Utf8ToWide(lastGoodPath_);
    if (PathFileExistsW(configW.c_str())) {
        CopyFileW(configW.c_str(), lastGoodW.c_str(), FALSE);
    } else if (PathFileExistsW(legacyConfigW.c_str())) {
        CopyFileW(legacyConfigW.c_str(), lastGoodW.c_str(), FALSE);
    }
    if (!MoveFileExW(tmpW.c_str(), configW.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        error = "atomic config replace failed";
        return false;
    }
    return true;
}

AudioEngine::AudioEngine()
    : ring_(kTargetRate * 2) {
    const uint64_t nowMs = GetTickCount64();
    lastConsumeTickMs_.store(nowMs, std::memory_order_release);
    lastMusicSignalTickMs_.store(nowMs, std::memory_order_release);
}

float AudioEngine::DbToLinear(float db) {
    return std::pow(10.0f, db / 20.0f);
}

void AudioEngine::ApplySettings(const MicMixSettings& settings) {
    const float gainDb = std::clamp(settings.musicGainDb, -30.0f, -2.0f);
    const int bufferMs = std::clamp(settings.bufferTargetMs, 20, 250);
    const bool wasMuted = musicMuted_.exchange(settings.musicMuted, std::memory_order_acq_rel);
    if (settings.musicMuted && !wasMuted) {
        ring_.Reset();
    }
    micInputMuted_.store(settings.micInputMuted, std::memory_order_release);
    forceTxEnabled_.store(settings.forceTxEnabled, std::memory_order_release);
    musicGainLinear_.store(DbToLinear(gainDb), std::memory_order_release);
    bufferTargetMs_.store(bufferMs, std::memory_order_release);
    micUseSmoothGate_.store(settings.micUseSmoothGate, std::memory_order_release);
    micTalkDetected_.store(false, std::memory_order_release);
    micTailTxActive_.store(false, std::memory_order_release);
    lastMicTalkTickMs_.store(0, std::memory_order_release);
    micGateGain_.store(1.0f, std::memory_order_release);
    limiterGain_.store(1.0f, std::memory_order_release);
}

void AudioEngine::SetMusicSourceRunning(bool running) {
    sourceRunning_.store(running, std::memory_order_release);
    if (!running) {
        ring_.Reset();
        micTalkDetected_.store(false, std::memory_order_release);
        micTailTxActive_.store(false, std::memory_order_release);
        lastMicTalkTickMs_.store(0, std::memory_order_release);
        micRmsDbfs_.store(-120.0f, std::memory_order_release);
        micPeakDbfs_.store(-120.0f, std::memory_order_release);
        externalMicLinear_.store(0.0f, std::memory_order_release);
        musicRmsDbfs_.store(-120.0f, std::memory_order_release);
        musicPeakDbfs_.store(-120.0f, std::memory_order_release);
        musicSendPeakDbfs_.store(-120.0f, std::memory_order_release);
        lastSourceClipTickMs_.store(0, std::memory_order_release);
        sourceClipState_.store(false, std::memory_order_release);
        micGateGain_.store(1.0f, std::memory_order_release);
        limiterGain_.store(1.0f, std::memory_order_release);
    }
}
void AudioEngine::SetTalkState(bool talking) { talkState_.store(talking, std::memory_order_release); }
void AudioEngine::SetExternalMicLevel(float linear) {
    const float level = std::clamp(linear, 0.0f, 1.0f);
    externalMicLinear_.store(level, std::memory_order_release);
    const float db = 20.0f * std::log10(std::max(level, 0.000001f));
    const float prev = micRmsDbfs_.load(std::memory_order_relaxed);
    const float alpha = (db > prev) ? 0.30f : 0.12f;
    const float next = std::clamp(prev + ((db - prev) * alpha), -120.0f, 0.0f);
    micRmsDbfs_.store(next, std::memory_order_release);
}
void AudioEngine::SetMuted(bool muted) {
    const bool wasMuted = musicMuted_.exchange(muted, std::memory_order_acq_rel);
    if (muted && !wasMuted) {
        ring_.Reset();
    }
}
void AudioEngine::ToggleMute() {
    bool expected = musicMuted_.load(std::memory_order_relaxed);
    while (!musicMuted_.compare_exchange_weak(
        expected, !expected, std::memory_order_acq_rel, std::memory_order_relaxed)) {
    }
    if (!expected) {
        ring_.Reset();
    }
}

void AudioEngine::SetMicInputMuted(bool muted) {
    micInputMuted_.store(muted, std::memory_order_release);
}

void AudioEngine::ToggleMicInputMute() {
    bool expected = micInputMuted_.load(std::memory_order_relaxed);
    while (!micInputMuted_.compare_exchange_weak(
        expected, !expected, std::memory_order_acq_rel, std::memory_order_relaxed)) {
    }
}

void AudioEngine::ClearMusicBuffer() {
    ring_.Reset();
}

void AudioEngine::PushMusicSamples(const float* samples, size_t count) {
    if (!samples || count == 0) {
        return;
    }
    const bool muted = musicMuted_.load(std::memory_order_relaxed);
    // Source telemetry reflects source level before mixing to TS capture.
    const float gain = musicGainLinear_.load(std::memory_order_relaxed);
    float sq = 0.0f;
    float peak = 0.0f;
    bool sourceClipped = false;
    for (size_t i = 0; i < count; ++i) {
        const float v = muted ? 0.0f : (samples[i] * gain);
        sq += v * v;
        const float a = std::fabs(v);
        if (a > peak) {
            peak = a;
        }
        if (a >= 1.0f) {
            sourceClipped = true;
        }
    }
    const float rms = std::sqrt(sq / static_cast<float>(count));
    const float rmsDb = 20.0f * std::log10(std::max(rms, 0.000001f));
    const float peakDb = 20.0f * std::log10(std::max(peak, 0.000001f));

    const float prevRms = musicRmsDbfs_.load(std::memory_order_relaxed);
    const float rmsAlpha = (rmsDb > prevRms) ? 0.35f : 0.12f;
    const float nextRms = std::clamp(prevRms + ((rmsDb - prevRms) * rmsAlpha), -120.0f, 0.0f);
    musicRmsDbfs_.store(nextRms, std::memory_order_release);

    const float prevPeak = musicPeakDbfs_.load(std::memory_order_relaxed);
    float nextPeak = peakDb;
    if (peakDb < prevPeak) {
        nextPeak = prevPeak - 1.2f;
    }
    nextPeak = std::clamp(nextPeak, -120.0f, 0.0f);
    musicPeakDbfs_.store(nextPeak, std::memory_order_release);
    // Treat very low-level source noise as "no activity" to avoid extending
    // force-send windows when only near-silence is present.
    if (peakDb > -72.0f || rmsDb > -78.0f) {
        lastMusicSignalTickMs_.store(GetTickCount64(), std::memory_order_release);
    }
    const bool sourceWasClipped = sourceClipState_.exchange(sourceClipped, std::memory_order_acq_rel);
    if (sourceClipped && !sourceWasClipped) {
        const uint64_t nowMs = GetTickCount64();
        sourceClipEvents_.fetch_add(1, std::memory_order_relaxed);
        lastSourceClipTickMs_.store(nowMs, std::memory_order_release);
    }
    if (muted) {
        ring_.Reset();
        return;
    }

    const uint64_t nowMs = GetTickCount64();
    const uint64_t lastConsumeMs = lastConsumeTickMs_.load(std::memory_order_acquire);
    if (nowMs > lastConsumeMs + 1500ULL) {
        // No consumer activity for a while (e.g. no outgoing TS3 capture callback).
        // Drop stale buffered audio so producer cannot accumulate endless backlog.
        ring_.Reset();
    }
    const int targetMs = std::clamp(bufferTargetMs_.load(std::memory_order_relaxed), 20, 250);
    size_t maxQueued = (static_cast<size_t>(targetMs) * static_cast<size_t>(kTargetRate)) / 1000U;
    maxQueued += static_cast<size_t>(kTargetRate / 100); // ~10ms safety headroom.
    maxQueued = std::min(maxQueued, ring_.Capacity() > 0 ? (ring_.Capacity() - 1) : 0U);

    size_t allowed = 0;
    const size_t queued = ring_.Size();
    if (queued < maxQueued) {
        allowed = maxQueued - queued;
    }
    const size_t requestWrite = std::min(count, allowed);
    const size_t written = (requestWrite > 0) ? ring_.Write(samples, requestWrite) : 0;
    if (written < count) {
        overruns_.fetch_add(count - written, std::memory_order_relaxed);
    }
}

TelemetrySnapshot AudioEngine::SnapshotTelemetry() const {
    TelemetrySnapshot t{};
    const uint64_t nowMs = GetTickCount64();
    t.underruns = underruns_.load(std::memory_order_relaxed);
    t.overruns = overruns_.load(std::memory_order_relaxed);
    t.clippedSamples = clippedSamples_.load(std::memory_order_relaxed);
    t.sourceClipEvents = sourceClipEvents_.load(std::memory_order_relaxed);
    t.micClipEvents = micClipEvents_.load(std::memory_order_relaxed);
    t.reconnects = reconnectsMirror_.load(std::memory_order_relaxed);
    t.musicRmsDbfs = musicRmsDbfs_.load(std::memory_order_relaxed);
    t.musicPeakDbfs = musicPeakDbfs_.load(std::memory_order_relaxed);
    t.musicSendPeakDbfs = musicSendPeakDbfs_.load(std::memory_order_relaxed);
    t.talkStateActive = talkState_.load(std::memory_order_relaxed);
    t.micTalkDetected = micTalkDetected_.load(std::memory_order_relaxed);
    t.micRmsDbfs = micRmsDbfs_.load(std::memory_order_relaxed);
    t.micPeakDbfs = micPeakDbfs_.load(std::memory_order_relaxed);
    const uint64_t lastSignalMs = lastMusicSignalTickMs_.load(std::memory_order_relaxed);
    const uint64_t lastSourceClipMs = lastSourceClipTickMs_.load(std::memory_order_relaxed);
    const uint64_t lastMicClipMs = lastMicClipTickMs_.load(std::memory_order_relaxed);
    t.musicActive = sourceRunning_.load(std::memory_order_relaxed) && (nowMs <= (lastSignalMs + 1200ULL));
    t.sourceClipRecent = (lastSourceClipMs != 0ULL) && (nowMs <= (lastSourceClipMs + kClipIndicatorHoldMs));
    t.micClipRecent = (lastMicClipMs != 0ULL) && (nowMs <= (lastMicClipMs + kClipIndicatorHoldMs));
    return t;
}

void AudioEngine::NoteReconnect() {
    reconnectsMirror_.fetch_add(1, std::memory_order_relaxed);
}

void AudioEngine::EditCapturedVoice(short* samples, int sampleCount, int channels, int* edited) {
    if (!samples || sampleCount <= 0 || channels <= 0 || !edited) {
        return;
    }
    // Defensive bounds against malformed callback metadata.
    if (channels > 8 || sampleCount > kTargetRate) {
        return;
    }
    const int upstreamFlags = *edited;
    const uint64_t nowMs = GetTickCount64();
    lastConsumeTickMs_.store(nowMs, std::memory_order_release);
    const bool muted = musicMuted_.load(std::memory_order_relaxed);
    const bool micInputMuted = micInputMuted_.load(std::memory_order_relaxed);
    const bool sourceRunning = sourceRunning_.load(std::memory_order_relaxed);
    const bool forceTx = forceTxEnabled_.load(std::memory_order_relaxed);
    const bool talkOpen = talkState_.load(std::memory_order_relaxed);
    const bool smoothGate = micUseSmoothGate_.load(std::memory_order_relaxed);
    const float targetGateGain = talkOpen ? 1.0f : 0.0f;
    float gateGain = std::clamp(micGateGain_.load(std::memory_order_relaxed), 0.0f, 1.0f);
    float limiterGain = std::clamp(limiterGain_.load(std::memory_order_relaxed), 0.1f, 1.0f);
    // Smooth mic gate to avoid hard consonant cutoffs and start clicks while
    // still reacting fast enough for normal speech onset.
    constexpr float kGateAttackMs = 8.0f;
    constexpr float kGateReleaseMs = 90.0f;
    const float attackSamples = std::max(1.0f, (static_cast<float>(kTargetRate) * kGateAttackMs) / 1000.0f);
    const float releaseSamples = std::max(1.0f, (static_cast<float>(kTargetRate) * kGateReleaseMs) / 1000.0f);
    const float gateAttackStep = 1.0f / attackSamples;
    const float gateReleaseStep = 1.0f / releaseSamples;
    auto advanceGate = [&]() {
        if (!smoothGate) {
            gateGain = targetGateGain;
            return;
        }
        if (targetGateGain > gateGain) {
            gateGain = std::min(targetGateGain, gateGain + gateAttackStep);
        } else if (targetGateGain < gateGain) {
            gateGain = std::max(targetGateGain, gateGain - gateReleaseStep);
        }
    };
    // Limiter tuning:
    // - threshold slightly below full-scale leaves headroom against clipping,
    // - faster attack catches transient peaks,
    // - slower release avoids audible pumping.
    constexpr float kLimiterThreshold = 0.92f;
    constexpr float kLimiterAttackCoeff = 0.50f;
    constexpr float kLimiterReleaseCoeff = 0.0025f;
    auto advanceLimiter = [&](float preAbs) {
        float target = 1.0f;
        if (preAbs > kLimiterThreshold) {
            target = kLimiterThreshold / std::max(preAbs, 0.000001f);
        }
        const float coeff = (target < limiterGain) ? kLimiterAttackCoeff : kLimiterReleaseCoeff;
        limiterGain += (target - limiterGain) * coeff;
        limiterGain = std::clamp(limiterGain, 0.1f, 1.0f);
    };
    const bool hasQueuedMusic = ring_.Size() > 0;
    // Apply mic-input mute only while MicMix is actively running.
    // When MicMix is off, the dry microphone path must pass through unchanged.
    const bool applyMicInputMute = micInputMuted && sourceRunning;
    const uint64_t lastSignalMs = lastMusicSignalTickMs_.load(std::memory_order_relaxed);
    const bool recentMusicSignal = sourceRunning && !muted && (nowMs <= (lastSignalMs + kForceTxMusicWindowMs));
    // Mic tail TX is driven by actual outgoing post-effect PCM level:
    // - hysteresis thresholds avoid TX flutter around silence floor,
    // - short hangover avoids packet-boundary chatter.
    constexpr float kMicTailTxOnLevelLinear = 0.00050f;  // ~ -66 dBFS
    constexpr float kMicTailTxOffLevelLinear = 0.00025f; // ~ -72 dBFS
    constexpr uint64_t kMicTailTxHangoverMs = 180ULL;
    auto updateTailTxFromPeak = [&](float txPeakLinear) -> bool {
        bool micTailActive = micTailTxActive_.load(std::memory_order_relaxed);
        if (!applyMicInputMute) {
            if (micTailActive) {
                if (txPeakLinear >= kMicTailTxOffLevelLinear) {
                    lastMicTalkTickMs_.store(nowMs, std::memory_order_release);
                } else {
                    micTailActive = false;
                }
            } else if (txPeakLinear >= kMicTailTxOnLevelLinear) {
                micTailActive = true;
                lastMicTalkTickMs_.store(nowMs, std::memory_order_release);
            }
        } else {
            micTailActive = false;
        }
        micTailTxActive_.store(micTailActive, std::memory_order_release);
        const uint64_t lastMicTalkMs = lastMicTalkTickMs_.load(std::memory_order_relaxed);
        return !applyMicInputMute &&
               (micTailActive ||
                ((lastMicTalkMs != 0ULL) && (nowMs <= (lastMicTalkMs + kMicTailTxHangoverMs))));
    };

    float micRmsAcc = 0.0f;
    float micPeak = 0.0f;
    bool micClipped = false;
    for (int i = 0; i < sampleCount; ++i) {
        float micMono = 0.0f;
        for (int ch = 0; ch < channels; ++ch) {
            const int idx = (i * channels) + ch;
            const short inSample = samples[idx];
            if (std::abs(static_cast<int>(inSample)) >= 32760) {
                micClipped = true;
            }
            const float micLinear = static_cast<float>(inSample) / 32768.0f;
            micMono += micLinear;
            micPeak = std::max(micPeak, std::fabs(micLinear));
        }
        micMono /= static_cast<float>(channels);
        micRmsAcc += micMono * micMono;
    }
    const float micRms = std::sqrt(micRmsAcc / static_cast<float>(sampleCount));
    const float micRmsDb = 20.0f * std::log10(std::max(micRms, 0.000001f));
    const float prevMicRms = micRmsDbfs_.load(std::memory_order_relaxed);
    const float micAlpha = (micRmsDb > prevMicRms) ? 0.25f : 0.10f;
    const float nextMicRms = std::clamp(prevMicRms + ((micRmsDb - prevMicRms) * micAlpha), -120.0f, 0.0f);
    micRmsDbfs_.store(nextMicRms, std::memory_order_release);
    const float micPeakDb = 20.0f * std::log10(std::max(micPeak, 0.000001f));
    const float prevMicPeakDb = micPeakDbfs_.load(std::memory_order_relaxed);
    float nextMicPeakDb = micPeakDb;
    if (micPeakDb < prevMicPeakDb) {
        nextMicPeakDb = std::max(micPeakDb, prevMicPeakDb - 1.8f);
    }
    micPeakDbfs_.store(std::clamp(nextMicPeakDb, -120.0f, 0.0f), std::memory_order_release);
    const bool micWasClipped = micClipState_.exchange(micClipped, std::memory_order_acq_rel);
    if (micClipped && !micWasClipped) {
        micClipEvents_.fetch_add(1, std::memory_order_relaxed);
        lastMicClipTickMs_.store(nowMs, std::memory_order_release);
    }
    micTalkDetected_.store(talkOpen, std::memory_order_release);

    if (!sourceRunning || muted || !hasQueuedMusic) {
        const float prevSendPeakDb = musicSendPeakDbfs_.load(std::memory_order_relaxed);
        const float decayedSendPeakDb = std::max(-120.0f, prevSendPeakDb - 1.8f);
        musicSendPeakDbfs_.store(decayedSendPeakDb, std::memory_order_release);
        micTalkDetected_.store(talkOpen, std::memory_order_release);
        bool touched = false;

        // TS3 bitmask semantics:
        // bit 1 (value 1): audio buffer modified
        // bit 2 (value 2): packet should be sent
        // Keep upstream flags untouched when we did not mix anything.
        if (applyMicInputMute) {
            for (int i = 0; i < sampleCount; ++i) {
                for (int ch = 0; ch < channels; ++ch) {
                    const int idx = (i * channels) + ch;
                    const short next = 0;
                    if (next != samples[idx]) {
                        touched = true;
                        samples[idx] = next;
                    }
                }
            }
        } else {
            // In OFF/no-music path keep dry mic pass-through untouched.
            gateGain = 1.0f;
        }
        // Match limiter recovery behavior to the per-sample release path used
        // while mixing so transitions stay consistent when music stops.
        const float recoveryFactor = 1.0f - std::pow(1.0f - kLimiterReleaseCoeff, static_cast<float>(sampleCount));
        limiterGain += (1.0f - limiterGain) * recoveryFactor;
        limiterGain = std::clamp(limiterGain, 0.1f, 1.0f);
        micGateGain_.store(gateGain, std::memory_order_relaxed);
        limiterGain_.store(limiterGain, std::memory_order_relaxed);
        float txPeakLinear = 0.0f;
        for (int i = 0; i < sampleCount; ++i) {
            for (int ch = 0; ch < channels; ++ch) {
                const int idx = (i * channels) + ch;
                const float v = static_cast<float>(samples[idx]) / 32768.0f;
                txPeakLinear = std::max(txPeakLinear, std::fabs(v));
            }
        }
        const bool recentMicTailSignal = updateTailTxFromPeak(txPeakLinear);
        int outFlags = upstreamFlags;
        if (touched || applyMicInputMute) {
            outFlags |= 1;
        }
        if ((forceTx && recentMusicSignal) || recentMicTailSignal) {
            outFlags |= 2;
        }
        *edited = outFlags;
        return;
    }

    thread_local std::array<float, kCallbackScratch> music{};
    bool anyMusicSignal = false;
    bool touched = false;
    bool gateTouched = false;
    float mixPeak = 0.0f;
    float txPeakLinear = 0.0f;
    const float musicGain = musicGainLinear_.load(std::memory_order_relaxed);
    int offset = 0;
    while (offset < sampleCount) {
        const int chunk = std::min(sampleCount - offset, static_cast<int>(kCallbackScratch));
        const size_t pulled = ring_.Read(music.data(), static_cast<size_t>(chunk));
        if (pulled < static_cast<size_t>(chunk)) {
            std::fill(music.begin() + static_cast<std::ptrdiff_t>(pulled), music.begin() + chunk, 0.0f);
            underruns_.fetch_add(static_cast<uint64_t>(chunk - pulled), std::memory_order_relaxed);
        }

        for (int i = 0; i < chunk; ++i) {
            advanceGate();
            const float m = music[i] * musicGain;
            const float absM = std::fabs(m);
            if (absM > 0.0005f) {
                anyMusicSignal = true;
            }
            if (absM > mixPeak) {
                mixPeak = absM;
            }
            float framePrePeak = 0.0f;
            for (int ch = 0; ch < channels; ++ch) {
                const int idx = ((offset + i) * channels) + ch;
                const float dryInput = applyMicInputMute ? 0.0f : (static_cast<float>(samples[idx]) / 32768.0f);
                const float dry = dryInput * gateGain;
                const float pre = dry + m;
                framePrePeak = std::max(framePrePeak, std::fabs(pre));
            }
            advanceLimiter(framePrePeak);
            for (int ch = 0; ch < channels; ++ch) {
                const int idx = ((offset + i) * channels) + ch;
                const short prevSample = samples[idx];
                const float dryInput = applyMicInputMute ? 0.0f : (static_cast<float>(samples[idx]) / 32768.0f);
                const float dry = dryInput * gateGain;
                float out = dry + m;
                out *= limiterGain;
                const float absOut = std::fabs(out);
                if (absOut > txPeakLinear) {
                    txPeakLinear = absOut;
                }
                if (absOut > 1.0f) {
                    clippedSamples_.fetch_add(1, std::memory_order_relaxed);
                    out = std::copysign(1.0f, out);
                }
                const short nextSample = static_cast<short>(std::lrintf(std::clamp(out, -1.0f, 1.0f) * 32767.0f));
                if (nextSample != prevSample) {
                    gateTouched = true;
                }
                samples[idx] = nextSample;
            }
        }
        if (!muted && pulled > 0) {
            touched = true;
        }
        offset += chunk;
    }

    const float targetSendPeakDb = 20.0f * std::log10(std::max(mixPeak, 0.000001f));
    const float prevSendPeakDb = musicSendPeakDbfs_.load(std::memory_order_relaxed);
    float nextSendPeakDb = targetSendPeakDb;
    if (targetSendPeakDb < prevSendPeakDb) {
        nextSendPeakDb = std::max(targetSendPeakDb, prevSendPeakDb - 1.8f);
    }
    musicSendPeakDbfs_.store(std::clamp(nextSendPeakDb, -120.0f, 0.0f), std::memory_order_release);
    micTalkDetected_.store(talkOpen, std::memory_order_release);
    micGateGain_.store(gateGain, std::memory_order_relaxed);
    limiterGain_.store(limiterGain, std::memory_order_relaxed);
    const bool recentMicTailSignal = updateTailTxFromPeak(txPeakLinear);

    const bool mixedMusic = touched && anyMusicSignal && sourceRunning && !muted;
    const bool forceMusicTx = (forceTx && (recentMusicSignal || mixedMusic)) || recentMicTailSignal;

    int outFlags = upstreamFlags;
    if (mixedMusic || gateTouched || applyMicInputMute) {
        outFlags |= 1;
    }
    if (forceMusicTx) {
        outFlags |= 2;
    }
    *edited = outFlags;
}

namespace {

class ActivationHandler final : public IActivateAudioInterfaceCompletionHandler {
public:
    ActivationHandler() : event_(CreateEventW(nullptr, FALSE, FALSE, nullptr)) {}
    ~ActivationHandler() {
        if (event_) {
            CloseHandle(event_);
            event_ = nullptr;
        }
    }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppvObject) override {
        if (!ppvObject) return E_POINTER;
        if (riid == __uuidof(IUnknown) ||
            riid == __uuidof(IActivateAudioInterfaceCompletionHandler) ||
            riid == __uuidof(IAgileObject)) {
            *ppvObject = static_cast<IActivateAudioInterfaceCompletionHandler*>(this);
            AddRef();
            return S_OK;
        }
        *ppvObject = nullptr;
        return E_NOINTERFACE;
    }

    ULONG STDMETHODCALLTYPE AddRef() override {
        return static_cast<ULONG>(InterlockedIncrement(&refs_));
    }

    ULONG STDMETHODCALLTYPE Release() override {
        ULONG refs = static_cast<ULONG>(InterlockedDecrement(&refs_));
        if (refs == 0) {
            delete this;
        }
        return refs;
    }

    HRESULT STDMETHODCALLTYPE ActivateCompleted(IActivateAudioInterfaceAsyncOperation* operation) override {
        HRESULT hrActivate = E_FAIL;
        ComPtr<IUnknown> unk;
        const HRESULT hr = operation->GetActivateResult(&hrActivate, &unk);
        if (SUCCEEDED(hr) && SUCCEEDED(hrActivate) && unk) {
            unk.As(&client_);
            hr_ = client_ ? S_OK : E_NOINTERFACE;
        } else {
            hr_ = FAILED(hr) ? hr : hrActivate;
        }
        SetEvent(event_);
        return S_OK;
    }

    HRESULT Wait(ComPtr<IAudioClient>& outClient, DWORD timeoutMs) {
        if (!event_) return E_HANDLE;
        if (WaitForSingleObject(event_, timeoutMs) != WAIT_OBJECT_0) {
            return HRESULT_FROM_WIN32(ERROR_TIMEOUT);
        }
        outClient = client_;
        return hr_;
    }

private:
    LONG refs_ = 1;
    HANDLE event_ = nullptr;
    HRESULT hr_ = E_FAIL;
    ComPtr<IAudioClient> client_;
};

HRESULT ActivateProcessLoopbackClientWithMode(uint32_t pid, PROCESS_LOOPBACK_MODE mode, ComPtr<IAudioClient>& outClient) {
    AUDIOCLIENT_ACTIVATION_PARAMS params{};
    params.ActivationType = AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK;
    params.ProcessLoopbackParams.TargetProcessId = pid;
    params.ProcessLoopbackParams.ProcessLoopbackMode = mode;
    PROPVARIANT pv{};
    PropVariantInit(&pv);
    pv.vt = VT_BLOB;
    pv.blob.cbSize = sizeof(params);
    pv.blob.pBlobData = reinterpret_cast<BYTE*>(&params);

    ComPtr<IActivateAudioInterfaceAsyncOperation> op;
    ActivationHandler* handler = new ActivationHandler();
    HRESULT hr = ActivateAudioInterfaceAsync(VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK, __uuidof(IAudioClient), &pv, handler, &op);
    if (FAILED(hr)) {
        handler->Release();
        return hr;
    }
    hr = handler->Wait(outClient, 5000);
    handler->Release();
    return hr;
}

HRESULT ActivateProcessLoopbackClient(uint32_t pid, ComPtr<IAudioClient>& outClient) {
    outClient.Reset();
    return ActivateProcessLoopbackClientWithMode(pid, PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE, outClient);
}

bool IsProcessAlive(uint32_t pid) {
    HANDLE h = OpenProcess(SYNCHRONIZE, FALSE, pid);
    if (!h) return false;
    const DWORD wait = WaitForSingleObject(h, 0);
    CloseHandle(h);
    return wait == WAIT_TIMEOUT;
}

class SourceBase : public IAudioSource {
public:
    SourceBase(MicMixSettings settings, AudioSourceManager::AudioPushFn push, AudioSourceManager::StatusFn status)
        : settings_(std::move(settings)), push_(std::move(push)), status_(std::move(status)) {}

    ~SourceBase() override { Stop(); }

    bool Start() override {
        stop_.store(false, std::memory_order_release);
        try {
            thread_ = std::thread([this] { Run(); });
            return true;
        } catch (...) {
            return false;
        }
    }

    void Stop() override {
        stop_.store(true, std::memory_order_release);
        stopCv_.notify_all();
        if (thread_.joinable()) {
            thread_.join();
        }
    }

protected:
    MicMixSettings settings_;
    std::atomic<bool> stop_{false};
    AudioSourceManager::AudioPushFn push_;
    AudioSourceManager::StatusFn status_;
    std::vector<float> monoIn_;
    std::vector<float> monoOut_;
    SpeexResamplerWrap resampler_;

    bool StopRequested() const { return stop_.load(std::memory_order_acquire); }
    void Push(const float* data, size_t count) { if (push_) push_(data, count); }
    void SetStatus(SourceState st, const std::string& code, const std::string& msg, const std::string& detail) { if (status_) status_(st, code, msg, detail); }
    virtual bool CaptureOnce(std::string& code, std::string& msg, std::string& detail) = 0;

private:
    std::thread thread_;
    std::mutex stopMutex_;
    std::condition_variable stopCv_;

    static bool IsRetryableFailureCode(const std::string& code) {
        if (code == "activate_unsupported") {
            return false;
        }
        return true;
    }

    bool WaitReacquireBackoff(int seconds) {
        const auto timeout = std::chrono::seconds(std::max(0, seconds));
        std::unique_lock<std::mutex> lock(stopMutex_);
        return stopCv_.wait_for(lock, timeout, [this]() {
            return StopRequested();
        });
    }

    void Run() {
        ComInit com;
        if (FAILED(com.hr)) {
            SetStatus(SourceState::Error, "com_init_failed", "COM init failed", "");
            return;
        }
        int backoff = 1;
        while (!StopRequested()) {
            SetStatus(SourceState::Starting, "starting", "Starting source", "");
            std::string code;
            std::string msg;
            std::string detail;
            bool captureOk = false;
            try {
                captureOk = CaptureOnce(code, msg, detail);
            } catch (const std::exception& ex) {
                SetStatus(SourceState::Error, "capture_exception", "Source capture exception", ex.what());
                break;
            } catch (...) {
                SetStatus(SourceState::Error, "capture_exception", "Source capture exception", "");
                break;
            }
            if (captureOk) {
                backoff = 1;
                continue;
            }
            if (StopRequested()) break;
            if (!IsRetryableFailureCode(code)) {
                SetStatus(SourceState::Error, code.empty() ? "error" : code, msg.empty() ? "Source error" : msg, detail);
                break;
            }
            SetStatus(SourceState::Reacquiring, code.empty() ? "reacquire" : code, msg.empty() ? "Reacquiring source" : msg, detail);
            if (WaitReacquireBackoff(backoff)) {
                break;
            }
            backoff = std::min(backoff * 2, 15);
        }
        SetStatus(SourceState::Stopped, "stopped", "Source stopped", "");
    }
};

} // namespace

namespace {

class LoopbackSource final : public SourceBase {
public:
    LoopbackSource(const MicMixSettings& s, AudioSourceManager::AudioPushFn push, AudioSourceManager::StatusFn status)
        : SourceBase(s, std::move(push), std::move(status)) {}

private:
    bool CaptureOnce(std::string& code, std::string& msg, std::string& detail) override {
        (void)detail;
        ComPtr<IMMDeviceEnumerator> enumerator;
        HRESULT hr = CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&enumerator));
        if (FAILED(hr)) {
            code = "enumerator_failed"; msg = "MMDeviceEnumerator failed"; return false;
        }
        ComPtr<IMMDevice> device;
        if (!settings_.loopbackDeviceId.empty()) {
            std::wstring id = Utf8ToWide(settings_.loopbackDeviceId);
            hr = enumerator->GetDevice(id.c_str(), &device);
        } else {
            hr = enumerator->GetDefaultAudioEndpoint(eRender, eMultimedia, &device);
        }
        if (FAILED(hr) || !device) {
            code = "device_not_found"; msg = "Loopback device missing"; return false;
        }

        ComPtr<IAudioClient> client;
        hr = device->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(client.GetAddressOf()));
        if (FAILED(hr) || !client) {
            code = "open_failed"; msg = "Loopback activate failed"; return false;
        }

        WAVEFORMATEX* wf = nullptr;
        hr = client->GetMixFormat(&wf);
        if (FAILED(hr) || !wf) {
            code = "format_failed"; msg = "Loopback mix format failed"; return false;
        }
        auto releaseWf = [&]() { if (wf) { CoTaskMemFree(wf); wf = nullptr; } };
        const bool isFloat = IsFloatFormat(wf);
        const bool isPcm16 = IsPcm16Format(wf);
        if (!isFloat && !isPcm16) {
            code = "format_unsupported"; msg = "Unsupported loopback format"; releaseWf(); return false;
        }
        if (!IsSupportedSourceRate(wf->nSamplesPerSec)) {
            code = "format_unsupported";
            msg = "Unsupported loopback sample rate";
            detail = "sample_rate=" + std::to_string(wf->nSamplesPerSec);
            releaseWf();
            return false;
        }
        DWORD initFlags = AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK;
        hr = client->Initialize(AUDCLNT_SHAREMODE_SHARED, initFlags, 0, 0, wf, nullptr);
        if (FAILED(hr)) {
            initFlags = AUDCLNT_STREAMFLAGS_LOOPBACK;
            hr = client->Initialize(AUDCLNT_SHAREMODE_SHARED, initFlags, 0, 0, wf, nullptr);
            if (FAILED(hr)) {
                code = "client_init_failed"; msg = "Loopback init failed"; releaseWf(); return false;
            }
        }

        HANDLE event = nullptr;
        auto closeEvent = [&]() {
            if (event) {
                CloseHandle(event);
                event = nullptr;
            }
        };
        if ((initFlags & AUDCLNT_STREAMFLAGS_EVENTCALLBACK) != 0) {
            event = CreateEventW(nullptr, FALSE, FALSE, nullptr);
            if (!event) {
                code = "event_failed"; msg = "Loopback event create failed"; releaseWf(); return false;
            }
            hr = client->SetEventHandle(event);
            if (FAILED(hr)) {
                closeEvent();
                code = "event_set_failed"; msg = "Loopback event set failed"; releaseWf(); return false;
            }
        }

        ComPtr<IAudioCaptureClient> cap;
        hr = client->GetService(IID_PPV_ARGS(&cap));
        if (FAILED(hr) || !cap) {
            closeEvent();
            code = "capture_service_failed"; msg = "Loopback capture service failed"; releaseWf(); return false;
        }
        const int resamplerQuality = ResolveResamplerQualitySetting(settings_.resamplerQuality);
        if (!resampler_.Configure(wf->nSamplesPerSec, kTargetRate, resamplerQuality)) {
            closeEvent();
            code = "resampler_failed"; msg = "Resampler init failed"; releaseWf(); return false;
        }
        const int channels = std::max<int>(1, wf->nChannels);
        hr = client->Start();
        if (FAILED(hr)) {
            closeEvent();
            code = "start_failed"; msg = "Loopback start failed"; releaseWf(); return false;
        }
        SetStatus(SourceState::Running, "running", "Loopback running", "");

        while (!StopRequested()) {
            if (event) {
                const DWORD wait = WaitForSingleObject(event, 200);
                if (wait == WAIT_TIMEOUT) {
                    continue;
                }
                if (wait != WAIT_OBJECT_0) {
                    client->Stop();
                    closeEvent();
                    releaseWf();
                    code = "wait_failed"; msg = "Loopback wait failed"; return false;
                }
            }
            UINT32 nextPacket = 0;
            hr = cap->GetNextPacketSize(&nextPacket);
            if (FAILED(hr)) {
                client->Stop();
                closeEvent();
                releaseWf();
                code = "packet_failed"; msg = "Packet query failed"; return false;
            }
            if (nextPacket == 0) {
                if (!event) {
                    std::this_thread::sleep_for(std::chrono::milliseconds(4));
                }
                continue;
            }
            while (nextPacket > 0 && !StopRequested()) {
                BYTE* data = nullptr;
                UINT32 frames = 0;
                DWORD flags = 0;
                hr = cap->GetBuffer(&data, &frames, &flags, nullptr, nullptr);
                if (FAILED(hr)) {
                    client->Stop();
                    closeEvent();
                    releaseWf();
                    code = "buffer_failed"; msg = "Buffer fetch failed"; return false;
                }
                monoIn_.resize(frames);
                if (flags & AUDCLNT_BUFFERFLAGS_SILENT) {
                    std::fill(monoIn_.begin(), monoIn_.end(), 0.0f);
                } else if (isFloat) {
                    const float* in = reinterpret_cast<const float*>(data);
                    for (UINT32 i = 0; i < frames; ++i) {
                        float sum = 0.0f;
                        for (int ch = 0; ch < channels; ++ch) sum += in[(i * channels) + ch];
                        monoIn_[i] = sum / static_cast<float>(channels);
                    }
                } else {
                    const int16_t* in = reinterpret_cast<const int16_t*>(data);
                    for (UINT32 i = 0; i < frames; ++i) {
                        float sum = 0.0f;
                        for (int ch = 0; ch < channels; ++ch) sum += static_cast<float>(in[(i * channels) + ch]) / 32768.0f;
                        monoIn_[i] = sum / static_cast<float>(channels);
                    }
                }
                if (!resampler_.Process(monoIn_.data(), monoIn_.size(), monoOut_)) {
                    cap->ReleaseBuffer(frames);
                    client->Stop();
                    closeEvent();
                    releaseWf();
                    code = "resample_failed"; msg = "Resample failed"; return false;
                }
                if (!monoOut_.empty()) Push(monoOut_.data(), monoOut_.size());
                cap->ReleaseBuffer(frames);
                hr = cap->GetNextPacketSize(&nextPacket);
                if (FAILED(hr)) {
                    client->Stop();
                    closeEvent();
                    releaseWf();
                    code = "packet_failed"; msg = "Packet query failed"; return false;
                }
            }
        }

        client->Stop();
        closeEvent();
        releaseWf();
        return true;
    }
};

class AppSessionSource final : public SourceBase {
public:
    AppSessionSource(const MicMixSettings& s, AudioSourceManager::AudioPushFn push, AudioSourceManager::StatusFn status)
        : SourceBase(s, std::move(push), std::move(status)) {}

private:
    std::vector<uint32_t> ResolvePidCandidates() const {
        std::vector<uint32_t> out;
        std::unordered_set<uint32_t> seen;
        auto addCandidate = [&](uint32_t pid) {
            if (pid == 0 || !IsProcessAlive(pid)) {
                return;
            }
            if (seen.insert(pid).second) {
                out.push_back(pid);
            }
        };
        uint32_t stickyPid = 0;
        bool stickyParsed = false;
        if (!settings_.appSessionId.empty()) {
            char* endPtr = nullptr;
            const unsigned long parsed = strtoul(settings_.appSessionId.c_str(), &endPtr, 10);
            if (endPtr != settings_.appSessionId.c_str() && *endPtr == '\0') {
                stickyPid = static_cast<uint32_t>(parsed);
                stickyParsed = true;
            }
        }
        if (settings_.appProcessName.empty()) {
            if (stickyParsed) {
                addCandidate(stickyPid);
            }
            return out;
        }
        const auto audioPids = EnumerateAudioSessionPids();
        bool stickyAdded = false;
        if (stickyParsed && audioPids.find(stickyPid) != audioPids.end()) {
            addCandidate(stickyPid);
            stickyAdded = true;
        }
        const auto list = AudioSourceManager::EnumerateAppProcesses(settings_.appProcessName);
        // Prefer process instances with an active audio session first.
        for (const auto& item : list) {
            if (audioPids.find(item.pid) != audioPids.end()) {
                addCandidate(item.pid);
            }
        }
        for (const auto& item : list) {
            addCandidate(item.pid);
        }
        if (stickyParsed && !stickyAdded) {
            addCandidate(stickyPid);
        }
        return out;
    }

    bool TryActivateAnyCandidate(ComPtr<IAudioClient>& outClient, uint32_t& outPid, HRESULT& outHr, std::string& outDetail) const {
        const auto candidates = ResolvePidCandidates();
        if (candidates.empty()) {
            outHr = E_FAIL;
            outDetail.clear();
            return false;
        }

        std::string tried;
        HRESULT lastHr = E_FAIL;
        for (size_t i = 0; i < candidates.size(); ++i) {
            const uint32_t pid = candidates[i];
            ComPtr<IAudioClient> client;
            const HRESULT hr = ActivateProcessLoopbackClient(pid, client);
            if (SUCCEEDED(hr) && client) {
                outPid = pid;
                outClient = client;
                outHr = hr;
                outDetail = "pid=" + std::to_string(pid);
                return true;
            }
            lastHr = hr;
            if (i < 8) {
                if (!tried.empty()) tried += ",";
                tried += std::to_string(pid) + ":" + HrToHex(hr);
            }
        }
        outHr = lastHr;
        outDetail = "tried=" + tried + " last=" + HrToHex(lastHr);
        return false;
    }

    bool CaptureOnce(std::string& code, std::string& msg, std::string& detail) override {
        uint32_t pid = 0;
        ComPtr<IAudioClient> client;
        HRESULT hr = E_FAIL;
        if (!TryActivateAnyCandidate(client, pid, hr, detail)) {
            if (detail.empty()) {
                code = "session_not_found";
                msg = "Waiting for app";
                return false;
            }
            if (hr == E_NOTIMPL || hr == HRESULT_FROM_WIN32(ERROR_NOT_SUPPORTED)) {
                code = "activate_unsupported";
                msg = "Process capture not supported on this Windows build";
            } else {
                code = "activate_failed";
                msg = "Process loopback activation failed";
            }
            return false;
        }
        if (!pid || !client) {
            code = "session_not_found"; msg = "Waiting for app"; return false;
        }
        WAVEFORMATEX* wf = nullptr;
        bool ownsWf = false;
        WAVEFORMATEX fallbackWf{};
        auto useFallbackFormat = [&]() {
            fallbackWf = {};
            fallbackWf.wFormatTag = WAVE_FORMAT_PCM;
            fallbackWf.nChannels = 2;
            fallbackWf.nSamplesPerSec = 44100;
            fallbackWf.wBitsPerSample = 16;
            fallbackWf.nBlockAlign = static_cast<WORD>((fallbackWf.nChannels * fallbackWf.wBitsPerSample) / 8);
            fallbackWf.nAvgBytesPerSec = fallbackWf.nSamplesPerSec * fallbackWf.nBlockAlign;
            fallbackWf.cbSize = 0;
            wf = &fallbackWf;
            ownsWf = false;
        };
        hr = client->GetMixFormat(&wf);
        if (SUCCEEDED(hr) && wf) {
            ownsWf = true;
        } else {
            useFallbackFormat();
            detail = "pid=" + std::to_string(pid) + " mix_hr=" + HrToHex(hr) + " using=fallback_pcm_44100";
        }
        auto releaseWf = [&]() {
            if (ownsWf && wf) {
                CoTaskMemFree(wf);
            }
            wf = nullptr;
            ownsWf = false;
        };
        bool isFloat = false;
        bool isPcm16 = false;
        auto refreshFormatFlags = [&]() {
            isFloat = IsFloatFormat(wf);
            isPcm16 = IsPcm16Format(wf);
        };
        refreshFormatFlags();
        if (!isFloat && !isPcm16) {
            if (ownsWf) {
                releaseWf();
            }
            useFallbackFormat();
            refreshFormatFlags();
            detail = "pid=" + std::to_string(pid) + " using=fallback_pcm_44100";
        }
        if (!IsSupportedSourceRate(wf->nSamplesPerSec)) {
            if (ownsWf) {
                releaseWf();
                useFallbackFormat();
                refreshFormatFlags();
            }
            if (!IsSupportedSourceRate(wf->nSamplesPerSec)) {
                code = "format_unsupported";
                msg = "Unsupported app sample rate";
                if (!detail.empty()) {
                    detail += " ";
                }
                detail += "pid=" + std::to_string(pid) + " sample_rate=" + std::to_string(wf->nSamplesPerSec);
                return false;
            }
        }

        const DWORD flags = AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK | AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM | AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY;
        hr = client->Initialize(AUDCLNT_SHAREMODE_SHARED, flags, 0, 0, wf, nullptr);
        if (FAILED(hr)) {
            if (!detail.empty()) {
                detail += " ";
            }
            detail += "pid=" + std::to_string(pid) + " init_hr=" + HrToHex(hr);
            code = "client_init_failed"; msg = "App client init failed"; releaseWf(); return false;
        }

        HANDLE event = CreateEventW(nullptr, FALSE, FALSE, nullptr);
        auto closeEvent = [&]() {
            if (event) {
                CloseHandle(event);
                event = nullptr;
            }
        };
        if (event) {
            hr = client->SetEventHandle(event);
            if (FAILED(hr)) {
                if (!detail.empty()) {
                    detail += " ";
                }
                detail += "pid=" + std::to_string(pid) + " event_hr=" + HrToHex(hr) + " fallback=polling";
                closeEvent();
            }
        } else {
            if (!detail.empty()) {
                detail += " ";
            }
            detail += "pid=" + std::to_string(pid) + " event_create_failed fallback=polling";
        }
        ComPtr<IAudioCaptureClient> cap;
        hr = client->GetService(IID_PPV_ARGS(&cap));
        if (FAILED(hr) || !cap) {
            closeEvent();
            if (!detail.empty()) {
                detail += " ";
            }
            detail += "pid=" + std::to_string(pid) + " service_hr=" + HrToHex(hr);
            code = "capture_service_failed"; msg = "App capture service failed"; releaseWf(); return false;
        }
        const int resamplerQuality = ResolveResamplerQualitySetting(settings_.resamplerQuality);
        if (!resampler_.Configure(wf->nSamplesPerSec, kTargetRate, resamplerQuality)) {
            closeEvent();
            code = "resampler_failed"; msg = "Resampler init failed"; releaseWf(); return false;
        }
        const int channels = std::max<int>(1, wf->nChannels);
        hr = client->Start();
        if (FAILED(hr)) {
            closeEvent();
            if (!detail.empty()) {
                detail += " ";
            }
            detail += "pid=" + std::to_string(pid) + " start_hr=" + HrToHex(hr);
            code = "start_failed"; msg = "App start failed"; releaseWf(); return false;
        }
        SetStatus(SourceState::Running, "running", "App capture running", "");

        while (!StopRequested()) {
            if (event) {
                const DWORD wait = WaitForSingleObject(event, 200);
                if (wait == WAIT_TIMEOUT) {
                    if (!IsProcessAlive(pid)) {
                        client->Stop();
                        closeEvent();
                        releaseWf();
                        code = "session_lost"; msg = "App process exited"; return false;
                    }
                    continue;
                }
                if (wait != WAIT_OBJECT_0) {
                    client->Stop();
                    closeEvent();
                    releaseWf();
                    code = "wait_failed"; msg = "Event wait failed"; return false;
                }
            } else {
                if (!IsProcessAlive(pid)) {
                    client->Stop();
                    closeEvent();
                    releaseWf();
                    code = "session_lost"; msg = "App process exited"; return false;
                }
                std::this_thread::sleep_for(std::chrono::milliseconds(4));
            }
            UINT32 nextPacket = 0;
            hr = cap->GetNextPacketSize(&nextPacket);
            if (FAILED(hr)) {
                client->Stop();
                closeEvent();
                releaseWf();
                code = "packet_failed"; msg = "Packet query failed"; return false;
            }
            if (nextPacket == 0) {
                continue;
            }
            while (nextPacket > 0 && !StopRequested()) {
                BYTE* data = nullptr;
                UINT32 frames = 0;
                DWORD flagsRead = 0;
                hr = cap->GetBuffer(&data, &frames, &flagsRead, nullptr, nullptr);
                if (FAILED(hr)) {
                    client->Stop();
                    closeEvent();
                    releaseWf();
                    code = "buffer_failed"; msg = "Buffer read failed"; return false;
                }
                monoIn_.resize(frames);
                if (flagsRead & AUDCLNT_BUFFERFLAGS_SILENT) {
                    std::fill(monoIn_.begin(), monoIn_.end(), 0.0f);
                } else if (isFloat) {
                    const float* in = reinterpret_cast<const float*>(data);
                    for (UINT32 i = 0; i < frames; ++i) {
                        float sum = 0.0f;
                        for (int ch = 0; ch < channels; ++ch) sum += in[(i * channels) + ch];
                        monoIn_[i] = sum / static_cast<float>(channels);
                    }
                } else {
                    const int16_t* in = reinterpret_cast<const int16_t*>(data);
                    for (UINT32 i = 0; i < frames; ++i) {
                        float sum = 0.0f;
                        for (int ch = 0; ch < channels; ++ch) sum += static_cast<float>(in[(i * channels) + ch]) / 32768.0f;
                        monoIn_[i] = sum / static_cast<float>(channels);
                    }
                }
                if (!resampler_.Process(monoIn_.data(), monoIn_.size(), monoOut_)) {
                    cap->ReleaseBuffer(frames);
                    client->Stop();
                    closeEvent();
                    releaseWf();
                    code = "resample_failed"; msg = "Resample failed"; return false;
                }
                if (!monoOut_.empty()) Push(monoOut_.data(), monoOut_.size());
                cap->ReleaseBuffer(frames);
                hr = cap->GetNextPacketSize(&nextPacket);
                if (FAILED(hr)) {
                    client->Stop();
                    closeEvent();
                    releaseWf();
                    code = "packet_failed"; msg = "Packet query failed"; return false;
                }
            }
        }

        client->Stop();
        closeEvent();
        releaseWf();
        return true;
    }
};

} // namespace

class GlobalHotkeyManager {
public:
    using HotkeyExtractor = std::function<void(const MicMixSettings&, UINT&, UINT&)>;

    GlobalHotkeyManager(std::function<void()> onHotkey, HotkeyExtractor extractor, int hotkeyId, std::string logTag)
        : onHotkey_(std::move(onHotkey)),
          extractor_(std::move(extractor)),
          hotkeyId_(hotkeyId),
          logTag_(std::move(logTag)) {}

    ~GlobalHotkeyManager() {
        Stop();
    }

    void Start() {
        std::unique_lock<std::mutex> lock(mutex_);
        if (thread_.joinable()) {
            return;
        }
        stopRequested_.store(false, std::memory_order_release);
        controlToken_.store(GenerateControlToken(this), std::memory_order_release);
        threadReady_ = false;
        lastRejectedHotkeyLogTick_ = 0;
        lastAcceptedHotkeyTickMs_ = 0;
        try {
            thread_ = std::thread([this]() { ThreadMain(); });
        } catch (const std::exception& ex) {
            controlToken_.store(0, std::memory_order_release);
            LogError(logTag_ + " thread start failed: " + ex.what());
            return;
        } catch (...) {
            controlToken_.store(0, std::memory_order_release);
            LogError(logTag_ + " thread start failed: unknown");
            return;
        }
        if (!startedCv_.wait_for(lock, std::chrono::seconds(2), [this]() { return threadReady_; })) {
            LogWarn(logTag_ + " thread ready timeout");
        }
    }

    void Stop() {
        stopRequested_.store(true, std::memory_order_release);
        const uint64_t token = controlToken_.load(std::memory_order_acquire);
        DWORD threadId = 0;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            threadId = threadId_;
        }
        if (threadId != 0 && token != 0) {
            PostThreadMessageW(
                threadId,
                kMsgStopThread,
                kMsgStopTag,
                static_cast<LPARAM>(static_cast<UINT_PTR>(token)));
        }
        if (thread_.joinable()) {
            thread_.join();
        }
        stopRequested_.store(false, std::memory_order_release);
        controlToken_.store(0, std::memory_order_release);
    }

    void ApplySettings(const MicMixSettings& settings) {
        UINT newMods = 0;
        UINT newVk = 0;
        if (extractor_) {
            extractor_(settings, newMods, newVk);
        }
        bool changed = false;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            changed = (modifiers_ != newMods) || (vk_ != newVk);
            modifiers_ = newMods;
            vk_ = newVk;
        }
        if (!changed) {
            return;
        }
        const uint64_t token = controlToken_.load(std::memory_order_acquire);
        DWORD threadId = 0;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            threadId = threadId_;
        }
        if (threadId != 0 && token != 0) {
            PostThreadMessageW(
                threadId,
                kMsgApplySettings,
                kMsgApplyTag,
                static_cast<LPARAM>(static_cast<UINT_PTR>(token)));
        }
    }

    void SetCaptureBlocked(bool blocked) {
        captureBlocked_.store(blocked, std::memory_order_release);
    }

private:
    static constexpr UINT kMsgApplySettings = WM_APP + 0x11;
    static constexpr UINT kMsgStopThread = WM_APP + 0x12;
    static constexpr WPARAM kMsgApplyTag = static_cast<WPARAM>(0x0A11C0DEu);
    static constexpr WPARAM kMsgStopTag = static_cast<WPARAM>(0x05E0C0DEu);
    static constexpr uint64_t kRejectedHotkeyLogIntervalMs = 5000ULL;
    static constexpr uint64_t kHotkeyTriggerGuardMs = 180ULL;
    std::function<void()> onHotkey_;
    HotkeyExtractor extractor_;
    int hotkeyId_ = 0;
    std::string logTag_;
    std::thread thread_;
    std::mutex mutex_;
    std::condition_variable startedCv_;
    bool threadReady_ = false;
    DWORD threadId_ = 0;
    UINT modifiers_ = 0;
    UINT vk_ = 0;
    UINT registeredMods_ = 0;
    UINT registeredVk_ = 0;
    bool registeredValid_ = false;
    std::atomic<uint64_t> controlToken_{0};
    uint64_t lastRejectedHotkeyLogTick_ = 0;
    uint64_t lastAcceptedHotkeyTickMs_ = 0;
    bool hotkeyHeld_ = false;
    std::atomic<bool> captureBlocked_{false};
    std::atomic<bool> stopRequested_{false};

    static uint64_t GenerateControlToken(const GlobalHotkeyManager* self) {
        static std::atomic<uint64_t> sequence{0};
        uint64_t seed = static_cast<uint64_t>(reinterpret_cast<uintptr_t>(self));
        LARGE_INTEGER qpc{};
        QueryPerformanceCounter(&qpc);
        seed ^= static_cast<uint64_t>(qpc.QuadPart);
        seed ^= static_cast<uint64_t>(GetTickCount64());
        seed ^= static_cast<uint64_t>(GetCurrentProcessId()) << 32;
        seed ^= static_cast<uint64_t>(sequence.fetch_add(1, std::memory_order_relaxed) + 1);

        // splitmix64 finalizer for good bit diffusion.
        seed ^= seed >> 30;
        seed *= 0xBF58476D1CE4E5B9ULL;
        seed ^= seed >> 27;
        seed *= 0x94D049BB133111EBULL;
        seed ^= seed >> 31;
        if (seed == 0) {
            seed = 0x7B23D4E1A9135C4FULL;
        }
        return seed;
    }

    static bool IsVirtualKeyDown(UINT vk) {
        return (GetAsyncKeyState(static_cast<int>(vk)) & 0x8000) != 0;
    }

    static bool IsRequiredModifiersDown(UINT mods) {
        if ((mods & MOD_SHIFT) && !IsVirtualKeyDown(VK_SHIFT) && !IsVirtualKeyDown(VK_LSHIFT) && !IsVirtualKeyDown(VK_RSHIFT)) {
            return false;
        }
        if ((mods & MOD_CONTROL) && !IsVirtualKeyDown(VK_CONTROL) && !IsVirtualKeyDown(VK_LCONTROL) && !IsVirtualKeyDown(VK_RCONTROL)) {
            return false;
        }
        if ((mods & MOD_ALT) && !IsVirtualKeyDown(VK_MENU) && !IsVirtualKeyDown(VK_LMENU) && !IsVirtualKeyDown(VK_RMENU)) {
            return false;
        }
        if ((mods & MOD_WIN) && !IsVirtualKeyDown(VK_LWIN) && !IsVirtualKeyDown(VK_RWIN)) {
            return false;
        }
        return true;
    }

    bool IsAuthorizedControlMessage(const MSG& msg, UINT expectedMsg, WPARAM expectedTag) const {
        if (msg.message != expectedMsg || msg.wParam != expectedTag) {
            return false;
        }
        const uint64_t token = controlToken_.load(std::memory_order_acquire);
        if (token == 0) {
            return false;
        }
        return static_cast<uint64_t>(static_cast<UINT_PTR>(msg.lParam)) == token;
    }

    bool ShouldAcceptHotkeyEvent() {
        if (!registeredValid_ || registeredVk_ == 0) {
            return false;
        }
        if (captureBlocked_.load(std::memory_order_acquire)) {
            return false;
        }
        if (hotkeyHeld_) {
            return false;
        }
        const uint64_t now = static_cast<uint64_t>(GetTickCount64());
        if (now < (lastAcceptedHotkeyTickMs_ + kHotkeyTriggerGuardMs)) {
            return false;
        }
        if (!IsRequiredModifiersDown(registeredMods_) || !IsVirtualKeyDown(registeredVk_)) {
            if (now - lastRejectedHotkeyLogTick_ >= kRejectedHotkeyLogIntervalMs) {
                lastRejectedHotkeyLogTick_ = now;
                LogWarn(logTag_ + " ignored event: physical key state mismatch");
            }
            return false;
        }
        hotkeyHeld_ = true;
        lastAcceptedHotkeyTickMs_ = now;
        return true;
    }

    void RefreshHotkeyHoldState() {
        if (!hotkeyHeld_) {
            return;
        }
        if (!registeredValid_ || registeredVk_ == 0) {
            hotkeyHeld_ = false;
            return;
        }
        if (!IsRequiredModifiersDown(registeredMods_) || !IsVirtualKeyDown(registeredVk_)) {
            hotkeyHeld_ = false;
        }
    }

    void ApplyRegistration() {
        UINT mods = 0;
        UINT vk = 0;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            mods = modifiers_;
            vk = vk_;
        }

        if (hotkeyId_ != 0) {
            UnregisterHotKey(nullptr, hotkeyId_);
        }
        if (vk == 0) {
            registeredValid_ = false;
            hotkeyHeld_ = false;
            lastAcceptedHotkeyTickMs_ = 0;
            return;
        }

        UINT regMods = mods;
#ifdef MOD_NOREPEAT
        regMods |= MOD_NOREPEAT;
#endif
        if (hotkeyId_ == 0 || !RegisterHotKey(nullptr, hotkeyId_, regMods, vk)) {
            registeredValid_ = false;
            LogWarn(logTag_ + " register failed vk=" + std::to_string(vk) +
                    " mods=" + std::to_string(mods) +
                    " err=" + std::to_string(GetLastError()));
        } else {
            const bool changed = (!registeredValid_) || (registeredMods_ != mods) || (registeredVk_ != vk);
            if (changed) {
                LogInfo(logTag_ + " registered vk=" + std::to_string(vk) +
                        " mods=" + std::to_string(mods));
            }
            registeredMods_ = mods;
            registeredVk_ = vk;
            registeredValid_ = true;
            hotkeyHeld_ = false;
            lastAcceptedHotkeyTickMs_ = 0;
        }
    }

    void ThreadMain() {
        MSG msg{};
        PeekMessageW(&msg, nullptr, WM_USER, WM_USER, PM_NOREMOVE);
        {
            std::lock_guard<std::mutex> lock(mutex_);
            threadId_ = GetCurrentThreadId();
            threadReady_ = true;
        }
        startedCv_.notify_all();

        ApplyRegistration();

        while (!stopRequested_.load(std::memory_order_acquire)) {
            RefreshHotkeyHoldState();
            while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
                if (msg.message == kMsgApplySettings) {
                    if (!IsAuthorizedControlMessage(msg, kMsgApplySettings, kMsgApplyTag)) {
                        continue;
                    }
                    ApplyRegistration();
                    continue;
                }
                if (msg.message == kMsgStopThread) {
                    if (!IsAuthorizedControlMessage(msg, kMsgStopThread, kMsgStopTag)) {
                        continue;
                    }
                    stopRequested_.store(true, std::memory_order_release);
                    break;
                }
                if (msg.message == WM_QUIT) {
                    stopRequested_.store(true, std::memory_order_release);
                    break;
                }
                if (msg.message == WM_HOTKEY && static_cast<int>(msg.wParam) == hotkeyId_) {
                    if (!ShouldAcceptHotkeyEvent()) {
                        continue;
                    }
                    if (onHotkey_) {
                        try {
                            onHotkey_();
                        } catch (const std::exception& ex) {
                            LogError(logTag_ + " callback exception: " + ex.what());
                        } catch (...) {
                            LogError(logTag_ + " callback exception: unknown");
                        }
                    }
                }
            }
            if (stopRequested_.load(std::memory_order_acquire)) {
                break;
            }
            MsgWaitForMultipleObjectsEx(0, nullptr, 200, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
            RefreshHotkeyHoldState();
        }

        if (hotkeyId_ != 0) {
            UnregisterHotKey(nullptr, hotkeyId_);
        }
        {
            std::lock_guard<std::mutex> lock(mutex_);
            threadId_ = 0;
            threadReady_ = false;
        }
    }
};

class MicLevelMonitor {
public:
    explicit MicLevelMonitor(std::function<void(float)> onLevel)
        : onLevel_(std::move(onLevel)) {}

    ~MicLevelMonitor() { Stop(); }

    void Start() {
        std::lock_guard<std::mutex> lock(mutex_);
        if (thread_.joinable()) {
            return;
        }
        stop_ = false;
        dirty_ = true;
        try {
            thread_ = std::thread([this]() { ThreadMain(); });
        } catch (const std::exception& ex) {
            LogError(std::string("mic_monitor thread start failed: ") + ex.what());
        } catch (...) {
            LogError("mic_monitor thread start failed: unknown");
        }
    }

    void Stop() {
        {
            std::lock_guard<std::mutex> lock(mutex_);
            stop_ = true;
            dirty_ = true;
        }
        cv_.notify_all();
        if (thread_.joinable()) {
            thread_.join();
        }
    }

    void ApplySettings(const MicMixSettings& settings) {
        {
            std::lock_guard<std::mutex> lock(mutex_);
            captureDeviceId_ = settings.captureDeviceId;
            dirty_ = true;
        }
        cv_.notify_all();
    }

private:
    struct CaptureTap {
        ComPtr<IAudioClient> client;
        ComPtr<IAudioCaptureClient> capture;
        int channels = 1;
        bool isFloat = false;
        bool isPcm16 = false;
        float lastLevel = 0.0f;
    };

    std::function<void(float)> onLevel_;
    std::thread thread_;
    std::mutex mutex_;
    std::condition_variable cv_;
    bool stop_ = false;
    bool dirty_ = true;
    std::string captureDeviceId_;

    static std::wstring Utf8ToWideLocal(const std::string& text) {
        if (text.empty()) {
            return {};
        }
        int len = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, nullptr, 0);
        std::wstring out(static_cast<size_t>(len), L'\0');
        MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, out.data(), len);
        if (!out.empty() && out.back() == L'\0') {
            out.pop_back();
        }
        return out;
    }

    static bool ResolveCaptureDevice(const std::string& deviceId, ComPtr<IMMDevice>& outDev) {
        outDev.Reset();
        ComPtr<IMMDeviceEnumerator> enumerator;
        if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&enumerator))) || !enumerator) {
            return false;
        }
        if (!deviceId.empty()) {
            const std::wstring wId = Utf8ToWideLocal(deviceId);
            if (FAILED(enumerator->GetDevice(wId.c_str(), &outDev)) || !outDev) {
                return false;
            }
            return true;
        }
        // Better default for user mics in desktop apps.
        if (SUCCEEDED(enumerator->GetDefaultAudioEndpoint(eCapture, eMultimedia, &outDev)) && outDev) {
            return true;
        }
        if (SUCCEEDED(enumerator->GetDefaultAudioEndpoint(eCapture, eCommunications, &outDev)) && outDev) {
            return true;
        }
        return false;
    }

    static void CloseCaptureTap(CaptureTap& tap) {
        if (tap.client) {
            tap.client->Stop();
        }
        tap.capture.Reset();
        tap.client.Reset();
        tap.channels = 1;
        tap.isFloat = false;
        tap.isPcm16 = false;
        tap.lastLevel = 0.0f;
    }

    static bool OpenCaptureTapForDevice(const std::string& deviceId, CaptureTap& outTap) {
        CloseCaptureTap(outTap);
        ComPtr<IMMDevice> dev;
        if (!ResolveCaptureDevice(deviceId, dev) || !dev) {
            return false;
        }

        ComPtr<IAudioClient> client;
        if (FAILED(dev->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(client.GetAddressOf()))) || !client) {
            return false;
        }

        WAVEFORMATEX* wf = nullptr;
        if (FAILED(client->GetMixFormat(&wf)) || !wf) {
            return false;
        }
        const bool isFloat = IsFloatFormat(wf);
        const bool isPcm16 = IsPcm16Format(wf);
        const int channels = std::max(1, static_cast<int>(wf->nChannels));
        if (!isFloat && !isPcm16) {
            CoTaskMemFree(wf);
            return false;
        }

        HRESULT hr = client->Initialize(AUDCLNT_SHAREMODE_SHARED, 0, 0, 0, wf, nullptr);
        CoTaskMemFree(wf);
        if (FAILED(hr)) {
            return false;
        }

        ComPtr<IAudioCaptureClient> cap;
        if (FAILED(client->GetService(IID_PPV_ARGS(&cap))) || !cap) {
            return false;
        }
        if (FAILED(client->Start())) {
            return false;
        }

        outTap.client = client;
        outTap.capture = cap;
        outTap.channels = channels;
        outTap.isFloat = isFloat;
        outTap.isPcm16 = isPcm16;
        outTap.lastLevel = 0.0f;
        return true;
    }

    static bool ReadCaptureTapLevel(CaptureTap& tap, float& outLevel) {
        outLevel = tap.lastLevel * 0.92f;
        if (!tap.capture) {
            tap.lastLevel = outLevel;
            return true;
        }

        bool hadData = false;
        float sumSq = 0.0f;
        uint64_t sampleCount = 0;
        for (;;) {
            UINT32 packetFrames = 0;
            HRESULT hr = tap.capture->GetNextPacketSize(&packetFrames);
            if (FAILED(hr)) {
                return false;
            }
            if (packetFrames == 0) {
                break;
            }

            BYTE* data = nullptr;
            UINT32 frames = 0;
            DWORD flags = 0;
            hr = tap.capture->GetBuffer(&data, &frames, &flags, nullptr, nullptr);
            if (FAILED(hr)) {
                return false;
            }
            if ((flags & AUDCLNT_BUFFERFLAGS_SILENT) == 0 && data && frames > 0) {
                hadData = true;
                if (tap.isFloat) {
                    const float* in = reinterpret_cast<const float*>(data);
                    for (UINT32 i = 0; i < frames; ++i) {
                        float acc = 0.0f;
                        for (int ch = 0; ch < tap.channels; ++ch) {
                            const float v = in[(static_cast<size_t>(i) * static_cast<size_t>(tap.channels)) + static_cast<size_t>(ch)];
                            acc += v * v;
                        }
                        const float v = std::sqrt(acc / static_cast<float>(tap.channels));
                        sumSq += v * v;
                    }
                    sampleCount += frames;
                } else if (tap.isPcm16) {
                    const short* in = reinterpret_cast<const short*>(data);
                    for (UINT32 i = 0; i < frames; ++i) {
                        float acc = 0.0f;
                        for (int ch = 0; ch < tap.channels; ++ch) {
                            const float v = static_cast<float>(in[(static_cast<size_t>(i) * static_cast<size_t>(tap.channels)) + static_cast<size_t>(ch)]) / 32768.0f;
                            acc += v * v;
                        }
                        const float v = std::sqrt(acc / static_cast<float>(tap.channels));
                        sumSq += v * v;
                    }
                    sampleCount += frames;
                }
            }
            tap.capture->ReleaseBuffer(frames);
        }

        if (hadData && sampleCount > 0) {
            outLevel = std::sqrt(sumSq / static_cast<float>(sampleCount));
        }
        outLevel = std::clamp(outLevel, 0.0f, 1.0f);
        tap.lastLevel = outLevel;
        return true;
    }

    static bool OpenMeterForDevice(const std::string& deviceId, ComPtr<IAudioMeterInformation>& outMeter) {
        outMeter.Reset();
        ComPtr<IMMDevice> dev;
        if (!ResolveCaptureDevice(deviceId, dev) || !dev) {
            return false;
        }

        if (FAILED(dev->Activate(__uuidof(IAudioMeterInformation), CLSCTX_ALL, nullptr, reinterpret_cast<void**>(outMeter.GetAddressOf()))) || !outMeter) {
            outMeter.Reset();
            return false;
        }
        return true;
    }

    void ThreadMain() {
        ComInit com;
        if (FAILED(com.hr)) {
            return;
        }
        CaptureTap tap;
        ComPtr<IAudioMeterInformation> meter;
        std::string currentId;
        bool usingTap = false;
        uint64_t lastModeLogMs = 0;

        for (;;) {
            bool reopen = false;
            {
                std::unique_lock<std::mutex> lock(mutex_);
                cv_.wait_for(lock, std::chrono::milliseconds(35), [this]() { return stop_ || dirty_; });
                if (stop_) {
                    break;
                }
                if (dirty_) {
                    currentId = captureDeviceId_;
                    dirty_ = false;
                    reopen = true;
                }
            }

            if (reopen) {
                CloseCaptureTap(tap);
                meter.Reset();
                usingTap = OpenCaptureTapForDevice(currentId, tap);
                if (!usingTap) {
                    OpenMeterForDevice(currentId, meter);
                }
                const uint64_t now = GetTickCount64();
                if (now >= lastModeLogMs + 1500ULL) {
                    if (usingTap) {
                        LogInfo("mic_monitor source=capture_stream");
                    } else if (meter) {
                        LogInfo("mic_monitor source=endpoint_meter");
                    } else {
                        LogWarn("mic_monitor source=none");
                    }
                    lastModeLogMs = now;
                }
            }

            float level = 0.0f;
            bool ok = true;
            if (usingTap && tap.capture) {
                ok = ReadCaptureTapLevel(tap, level);
                if (!ok) {
                    CloseCaptureTap(tap);
                    usingTap = false;
                    OpenMeterForDevice(currentId, meter);
                }
            } else if (meter) {
                if (FAILED(meter->GetPeakValue(&level))) {
                    meter.Reset();
                    level = 0.0f;
                }
            }
            if (onLevel_) {
                onLevel_(level);
            }
        }
        CloseCaptureTap(tap);
        if (onLevel_) {
            onLevel_(0.0f);
        }
    }
};

class MixMonitorPlayer {
public:
    MixMonitorPlayer()
        : ring_(kTargetRate * 4) {}

    ~MixMonitorPlayer() {
        SetEnabled(false);
    }

    void SetEnabled(bool enabled) {
        std::lock_guard<std::mutex> lock(stateMutex_);
        if (enabled) {
            StartLocked();
        } else {
            StopLocked();
        }
    }

    bool IsEnabled() const {
        return enabled_.load(std::memory_order_acquire);
    }

    void PushCaptured(const short* samples, int sampleCount, int channels) {
        if (!IsEnabled() || !samples || sampleCount <= 0 || channels <= 0) {
            return;
        }
        const size_t frames = static_cast<size_t>(sampleCount);
        constexpr size_t kChunkFrames = 512;
        thread_local std::array<short, kChunkFrames * 2> stereoChunk{};
        size_t offset = 0;
        while (offset < frames) {
            const size_t chunkFrames = std::min(kChunkFrames, frames - offset);
            for (size_t i = 0; i < chunkFrames; ++i) {
                const size_t base = (offset + i) * static_cast<size_t>(channels);
                const short l = samples[base];
                const short r = (channels >= 2) ? samples[base + 1] : l;
                stereoChunk[(i * 2) + 0] = l;
                stereoChunk[(i * 2) + 1] = r;
            }
            const size_t chunkSamples = chunkFrames * 2;
            const size_t written = ring_.Write(stereoChunk.data(), chunkSamples);
            if (written < chunkSamples) {
                droppedSamples_.fetch_add(static_cast<uint64_t>(chunkSamples - written), std::memory_order_relaxed);
            }
            offset += chunkFrames;
        }
    }

private:
    struct Block {
        WAVEHDR header{};
        std::vector<short> data;
        bool prepared = false;
        bool queued = false;
    };

    static constexpr int kBlockFrames = 960; // 20ms @ 48kHz
    static constexpr int kBlockCount = 6;

    SpscRingBuffer<short> ring_;
    std::thread thread_;
    std::atomic<bool> enabled_{false};
    std::atomic<bool> stop_{false};
    std::atomic_uint64_t droppedSamples_{0};
    std::mutex stateMutex_;

    void StartLocked() {
        if (thread_.joinable()) {
            return;
        }
        ring_.Reset();
        droppedSamples_.store(0, std::memory_order_release);
        stop_.store(false, std::memory_order_release);
        enabled_.store(true, std::memory_order_release);
        try {
            thread_ = std::thread([this]() { ThreadMain(); });
        } catch (const std::exception& ex) {
            enabled_.store(false, std::memory_order_release);
            stop_.store(true, std::memory_order_release);
            LogError(std::string("mix_monitor thread start failed: ") + ex.what());
        } catch (...) {
            enabled_.store(false, std::memory_order_release);
            stop_.store(true, std::memory_order_release);
            LogError("mix_monitor thread start failed: unknown");
        }
    }

    void StopLocked() {
        enabled_.store(false, std::memory_order_release);
        stop_.store(true, std::memory_order_release);
        if (thread_.joinable()) {
            thread_.join();
        }
        ring_.Reset();
    }

    void ThreadMain() {
        WAVEFORMATEX wf{};
        wf.wFormatTag = WAVE_FORMAT_PCM;
        wf.nChannels = 2;
        wf.nSamplesPerSec = kTargetRate;
        wf.wBitsPerSample = 16;
        wf.nBlockAlign = static_cast<WORD>((wf.nChannels * wf.wBitsPerSample) / 8);
        wf.nAvgBytesPerSec = wf.nSamplesPerSec * wf.nBlockAlign;
        wf.cbSize = 0;

        HWAVEOUT wave = nullptr;
        const MMRESULT openRes = waveOutOpen(&wave, WAVE_MAPPER, &wf, 0, 0, CALLBACK_NULL);
        if (openRes != MMSYSERR_NOERROR || !wave) {
            enabled_.store(false, std::memory_order_release);
            LogWarn("mix_monitor open failed code=" + std::to_string(openRes));
            return;
        }

        std::vector<Block> blocks(static_cast<size_t>(kBlockCount));
        for (auto& block : blocks) {
            block.data.resize(static_cast<size_t>(kBlockFrames) * 2);
            block.header = {};
            block.header.lpData = reinterpret_cast<LPSTR>(block.data.data());
            block.header.dwBufferLength = static_cast<DWORD>(block.data.size() * sizeof(short));
            if (waveOutPrepareHeader(wave, &block.header, sizeof(WAVEHDR)) == MMSYSERR_NOERROR) {
                block.prepared = true;
            }
        }

        while (!stop_.load(std::memory_order_acquire)) {
            bool wroteAny = false;
            for (auto& block : blocks) {
                if (!block.prepared) {
                    continue;
                }
                if (block.queued && (block.header.dwFlags & WHDR_DONE) == 0) {
                    continue;
                }
                block.queued = false;
                const size_t need = block.data.size();
                const size_t got = ring_.Read(block.data.data(), need);
                if (got < need) {
                    std::fill(block.data.begin() + static_cast<std::ptrdiff_t>(got), block.data.end(), static_cast<short>(0));
                }
                block.header.dwBufferLength = static_cast<DWORD>(need * sizeof(short));
                block.header.dwFlags &= ~WHDR_DONE;
                const MMRESULT wr = waveOutWrite(wave, &block.header, sizeof(WAVEHDR));
                if (wr == MMSYSERR_NOERROR) {
                    block.queued = true;
                    wroteAny = true;
                }
            }
            if (!wroteAny) {
                std::this_thread::sleep_for(std::chrono::milliseconds(6));
            }
        }

        waveOutReset(wave);
        for (auto& block : blocks) {
            if (block.prepared) {
                waveOutUnprepareHeader(wave, &block.header, sizeof(WAVEHDR));
            }
            block.prepared = false;
            block.queued = false;
        }
        waveOutClose(wave);
        const uint64_t dropped = droppedSamples_.load(std::memory_order_relaxed);
        if (dropped > 0) {
            LogWarn("mix_monitor dropped_samples=" + std::to_string(dropped));
        }
    }
};

AudioSourceManager::AudioSourceManager(AudioPushFn pushFn, StatusFn statusFn)
    : pushFn_(std::move(pushFn)), statusFn_(std::move(statusFn)) {}

AudioSourceManager::~AudioSourceManager() { Stop(); }

void AudioSourceManager::ApplySettings(const MicMixSettings& settings) {
    std::lock_guard<std::mutex> lock(mutex_);
    settings_ = settings;
}

bool AudioSourceManager::Start() {
    std::lock_guard<std::mutex> lock(mutex_);
    if (running_) return true;
    source_ = CreateSourceLocked();
    if (!source_) {
        SetStatus(SourceState::Error, "create_failed", "Source creation failed", "");
        return false;
    }
    running_ = source_->Start();
    if (!running_) {
        source_.reset();
        SetStatus(SourceState::Error, "thread_start_failed", "Source thread start failed", "");
    }
    return running_;
}

void AudioSourceManager::Stop() {
    std::unique_ptr<IAudioSource> sourceToStop;
    {
        std::lock_guard<std::mutex> lock(mutex_);
        if (!running_ && !source_) {
            return;
        }
        running_ = false;
        sourceToStop = std::move(source_);
    }
    if (sourceToStop) {
        sourceToStop->Stop();
    }
    {
        std::lock_guard<std::mutex> lock(mutex_);
        SetStatus(SourceState::Stopped, "stopped", "Source stopped", "");
    }
}

void AudioSourceManager::Restart() {
    Stop();
    Start();
}

bool AudioSourceManager::IsRunning() const {
    std::lock_guard<std::mutex> lock(mutex_);
    return running_;
}

SourceStatus AudioSourceManager::GetStatus() const {
    std::lock_guard<std::mutex> lock(mutex_);
    return status_;
}

void AudioSourceManager::SetStatus(SourceState state, const std::string& code, const std::string& message, const std::string& detail) {
    status_.state = state;
    status_.code = code;
    status_.message = message;
    status_.detail = detail;
    if (state == SourceState::Reacquiring) {
        status_.reconnectCount += 1;
    }
    if (statusFn_) statusFn_(state, code, message, detail);
}

std::unique_ptr<IAudioSource> AudioSourceManager::CreateSourceLocked() {
    auto statusForward = [this](SourceState st, const std::string& c, const std::string& m, const std::string& d) {
        StatusFn callback;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            status_.state = st;
            status_.code = c;
            status_.message = m;
            status_.detail = d;
            if (st == SourceState::Reacquiring) {
                status_.reconnectCount += 1;
            }
            if (st == SourceState::Stopped) {
                running_ = false;
            }
            callback = statusFn_;
        }
        if (callback) {
            callback(st, c, m, d);
        }
    };
    if (settings_.sourceMode == SourceMode::AppSession) {
        return std::make_unique<AppSessionSource>(settings_, pushFn_, statusForward);
    }
    return std::make_unique<LoopbackSource>(settings_, pushFn_, statusForward);
}

template <typename DeviceInfoT>
std::vector<DeviceInfoT> EnumerateDevices(EDataFlow flow, ERole role, const char* fallbackName) {
    std::vector<DeviceInfoT> out;
    ComInit com;
    if (FAILED(com.hr)) return out;
    ComPtr<IMMDeviceEnumerator> enumerator;
    if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&enumerator)))) return out;

    std::wstring defaultId;
    ComPtr<IMMDevice> def;
    if (SUCCEEDED(enumerator->GetDefaultAudioEndpoint(flow, role, &def)) && def) {
        LPWSTR id = nullptr;
        if (SUCCEEDED(def->GetId(&id)) && id) {
            defaultId = id;
            CoTaskMemFree(id);
        }
    }

    ComPtr<IMMDeviceCollection> coll;
    if (FAILED(enumerator->EnumAudioEndpoints(flow, DEVICE_STATE_ACTIVE, &coll)) || !coll) return out;
    UINT count = 0;
    coll->GetCount(&count);
    out.reserve(count);
    for (UINT i = 0; i < count; ++i) {
        ComPtr<IMMDevice> dev;
        if (FAILED(coll->Item(i, &dev)) || !dev) continue;
        LPWSTR id = nullptr;
        if (FAILED(dev->GetId(&id)) || !id) continue;
        ComPtr<IPropertyStore> store;
        if (FAILED(dev->OpenPropertyStore(STGM_READ, &store)) || !store) {
            CoTaskMemFree(id);
            continue;
        }
        PROPVARIANT name{};
        PropVariantInit(&name);
        std::string display = fallbackName;
        if (SUCCEEDED(store->GetValue(PKEY_Device_FriendlyName, &name)) && name.vt == VT_LPWSTR && name.pwszVal) {
            display = WideToUtf8(name.pwszVal);
        }
        PropVariantClear(&name);
        DeviceInfoT info;
        info.id = WideToUtf8(id);
        info.name = display;
        info.isDefault = (_wcsicmp(id, defaultId.c_str()) == 0);
        out.push_back(std::move(info));
        CoTaskMemFree(id);
    }
    return out;
}

std::string BuildMusicMetaValue(const std::string& currentValue, bool active) {
    constexpr const char* kKeyPrefix = "micmix_active=";
    std::vector<std::string> parts;
    parts.reserve(8);

    size_t begin = 0;
    while (begin <= currentValue.size()) {
        size_t end = currentValue.find(';', begin);
        if (end == std::string::npos) {
            end = currentValue.size();
        }
        std::string token = currentValue.substr(begin, end - begin);
        size_t left = 0;
        while (left < token.size() && std::isspace(static_cast<unsigned char>(token[left])) != 0) {
            ++left;
        }
        size_t right = token.size();
        while (right > left && std::isspace(static_cast<unsigned char>(token[right - 1])) != 0) {
            --right;
        }
        if (right > left) {
            token = token.substr(left, right - left);
            if (token.rfind(kKeyPrefix, 0) != 0) {
                parts.push_back(std::move(token));
            }
        }
        if (end == currentValue.size()) {
            break;
        }
        begin = end + 1;
    }

    parts.push_back(std::string(kKeyPrefix) + (active ? "1" : "0"));
    std::string out;
    for (size_t i = 0; i < parts.size(); ++i) {
        if (i > 0) {
            out.push_back(';');
        }
        out += parts[i];
    }
    return out;
}

std::vector<LoopbackDeviceInfo> AudioSourceManager::EnumerateLoopbackDevices() {
    return EnumerateDevices<LoopbackDeviceInfo>(eRender, eMultimedia, "Device");
}

std::vector<CaptureDeviceInfo> AudioSourceManager::EnumerateCaptureDevices() {
    return EnumerateDevices<CaptureDeviceInfo>(eCapture, eCommunications, "Microphone");
}

std::vector<AppProcessInfo> AudioSourceManager::EnumerateAppProcesses(const std::string& processName) {
    std::vector<AppProcessInfo> out;
    std::wstring target;
    if (!processName.empty()) {
        target = Utf8ToWide(processName);
    }
    const bool filterByName = !target.empty();
    const auto audioPids = filterByName ? std::unordered_set<uint32_t>{} : EnumerateAudioSessionPids();
    const auto visiblePids = filterByName ? std::unordered_set<uint32_t>{} : EnumerateVisibleWindowPids();
    const std::wstring windowsDirLower = filterByName ? std::wstring{} : GetWindowsDirLower();
    DWORD currentSessionId = 0;
    ProcessIdToSessionId(GetCurrentProcessId(), &currentSessionId);
    const uint32_t currentPid = GetCurrentProcessId();
    std::unordered_map<std::string, std::pair<AppProcessInfo, int>> bestByExe;

    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snap == INVALID_HANDLE_VALUE) return out;
    PROCESSENTRY32W pe{};
    pe.dwSize = sizeof(pe);
    if (Process32FirstW(snap, &pe)) {
        do {
            const uint32_t pid = pe.th32ProcessID;
            const std::wstring exeNameW = pe.szExeFile;
            if (filterByName && _wcsicmp(exeNameW.c_str(), target.c_str()) != 0) {
                continue;
            }
            if (!IsSameSession(pid, currentSessionId) || pid == currentPid) {
                continue;
            }
            if (!filterByName) {
                if (IsBlockedUiProcess(pid, exeNameW)) {
                    continue;
                }
                const bool hasAudioSession = audioPids.find(pid) != audioPids.end();
                const bool hasVisibleWindow = visiblePids.find(pid) != visiblePids.end();
                std::wstring imagePath;
                const bool hasImagePath = TryGetProcessImagePath(pid, imagePath);
                const bool likelyUserApp = hasImagePath && IsLikelyUserAppPath(ToLowerWide(imagePath), windowsDirLower);
                const std::string exeName = WideToUtf8(exeNameW);
                const bool preferredMediaApp = IsPreferredMediaProcessName(ToLowerAscii(exeName));
                const bool desktopLike = (hasVisibleWindow && likelyUserApp) || hasAudioSession || preferredMediaApp;
                if (!desktopLike) {
                    continue;
                }
                const int score = (hasAudioSession ? 5 : 0) + (hasVisibleWindow ? 3 : 0) + (likelyUserApp ? 2 : 0) + (preferredMediaApp ? 3 : 0);
                AppProcessInfo info;
                info.pid = pid;
                info.exeName = exeName;
                info.displayName = DisplayNameFromExe(exeName);

                const std::string exeKey = ToLowerAscii(exeName);
                auto it = bestByExe.find(exeKey);
                if (it == bestByExe.end() || score > it->second.second || (score == it->second.second && pid < it->second.first.pid)) {
                    bestByExe[exeKey] = { std::move(info), score };
                }
                continue;
            }
            AppProcessInfo info;
            info.pid = pid;
            info.exeName = WideToUtf8(exeNameW);
            info.displayName = DisplayNameFromExe(info.exeName);
            out.push_back(std::move(info));
        } while (Process32NextW(snap, &pe));
    }
    CloseHandle(snap);

    if (!filterByName) {
        out.reserve(bestByExe.size());
        for (auto& kv : bestByExe) {
            out.push_back(std::move(kv.second.first));
        }
    }

    std::sort(out.begin(), out.end(), [](const AppProcessInfo& a, const AppProcessInfo& b) {
        if (_stricmp(a.exeName.c_str(), b.exeName.c_str()) == 0) {
            return a.pid < b.pid;
        }
        return _stricmp(a.exeName.c_str(), b.exeName.c_str()) < 0;
    });
    return out;
}

MicMixApp& MicMixApp::Instance() {
    static MicMixApp instance;
    return instance;
}

MicMixApp::MicMixApp() = default;
MicMixApp::~MicMixApp() = default;

void MicMixApp::SanitizeEffectList(std::vector<VstEffectSlot>& list) {
    if (list.size() > kMaxEffectsPerChain) {
        list.resize(kMaxEffectsPerChain);
    }
    std::unordered_set<std::string> usedUids;
    usedUids.reserve(list.size() * 2U);
    for (auto& slot : list) {
        StripConfigControlChars(slot.path, kMaxEffectPathLen);
        StripConfigControlChars(slot.name, kMaxEffectNameLen);
        StripConfigControlChars(slot.uid, kMaxEffectUidLen);
        StripConfigControlChars(slot.stateBlob, kMaxEffectStateBlobLen);
        StripConfigControlChars(slot.lastStatus, kMaxEffectStatusLen);
        if (slot.name.empty() && !slot.path.empty()) {
            const std::filesystem::path p(Utf8ToWide(slot.path));
            slot.name = WideToUtf8(p.stem().wstring());
        }
        if (slot.uid.empty() && !slot.path.empty()) {
            slot.uid = BuildUniqueEffectUidWithExists(
                slot.path,
                [&usedUids](const std::string& uid) { return usedUids.find(uid) != usedUids.end(); });
        } else if (!slot.uid.empty() && usedUids.find(slot.uid) != usedUids.end()) {
            const std::string uidSeed = !slot.path.empty() ? slot.path : slot.uid;
            slot.uid = BuildUniqueEffectUidWithExists(
                uidSeed,
                [&usedUids](const std::string& uid) { return usedUids.find(uid) != usedUids.end(); });
        }
        if (slot.lastStatus.empty()) {
            slot.lastStatus = "pending";
        }
        if (!slot.uid.empty()) {
            usedUids.insert(slot.uid);
        }
    }
    list.erase(std::remove_if(list.begin(), list.end(), [](const VstEffectSlot& slot) {
        return slot.path.empty();
    }), list.end());
}

std::wstring MicMixApp::ResolveVstHostPath() const {
    wchar_t pathBuf[MAX_PATH]{};
    const HMODULE module = reinterpret_cast<HMODULE>(&__ImageBase);
    const DWORD len = GetModuleFileNameW(module, pathBuf, static_cast<DWORD>(std::size(pathBuf)));
    if (len == 0 || len >= std::size(pathBuf)) {
        return {};
    }
    std::filesystem::path dllPath(pathBuf);
    return (dllPath.parent_path() / L"micmix_vst_host.exe").wstring();
}

std::wstring MicMixApp::BuildVstHostPipePath() const {
    const std::wstring pipeName = vstHostPipeName_.empty() ? std::wstring(kVstHostPipeName) : vstHostPipeName_;
    return std::wstring(L"\\\\.\\pipe\\") + pipeName;
}

bool MicMixApp::SendVstHostCommand(
    const std::string& command,
    std::string& response,
    DWORD timeoutMs,
    std::string& error) const {
    response.clear();
    error.clear();
    const std::wstring pipePath = BuildVstHostPipePath();
    const uint64_t startTick = GetTickCount64();
    const uint64_t effectiveTimeoutMs = std::max<uint64_t>(static_cast<uint64_t>(timeoutMs), 50ULL);
    const uint64_t deadlineTick = startTick + effectiveTimeoutMs;

    std::lock_guard<std::mutex> ipcLock(vstHostIpcMutex_);
    HANDLE pipe = INVALID_HANDLE_VALUE;
    for (;;) {
        if (WaitNamedPipeW(pipePath.c_str(), 60)) {
            pipe = CreateFileW(
                pipePath.c_str(),
                GENERIC_READ | GENERIC_WRITE,
                0,
                nullptr,
                OPEN_EXISTING,
                FILE_ATTRIBUTE_NORMAL,
                nullptr);
            if (pipe != INVALID_HANDLE_VALUE) {
                break;
            }
            const DWORD openErr = GetLastError();
            if (openErr != ERROR_PIPE_BUSY &&
                openErr != ERROR_FILE_NOT_FOUND &&
                openErr != ERROR_SEM_TIMEOUT) {
                error = "CreateFile(pipe) failed: " + std::to_string(openErr);
                return false;
            }
        } else {
            const DWORD waitErr = GetLastError();
            if (waitErr != ERROR_FILE_NOT_FOUND &&
                waitErr != ERROR_SEM_TIMEOUT &&
                waitErr != ERROR_PIPE_BUSY) {
                error = "WaitNamedPipe failed: " + std::to_string(waitErr);
                return false;
            }
        }
        if (GetTickCount64() >= deadlineTick) {
            error = "pipe connect timeout";
            return false;
        }
        Sleep(10);
    }

    DWORD mode = PIPE_READMODE_MESSAGE;
    SetNamedPipeHandleState(pipe, &mode, nullptr, nullptr);

    std::string wire;
    if (!vstHostAuthToken_.empty()) {
        wire = "AUTH " + vstHostAuthToken_ + " " + command;
    } else {
        wire = command;
    }
    if ((wire.size() + 1U) > kVstHostMaxPipeMessageBytes) {
        error = "pipe command too large";
        CloseHandle(pipe);
        return false;
    }
    if (wire.empty() || wire.back() != '\n') {
        wire.push_back('\n');
    }
    const DWORD expectedWriteBytes = static_cast<DWORD>(wire.size());
    DWORD written = 0;
    if (!WriteFile(pipe, wire.data(), expectedWriteBytes, &written, nullptr) || written != expectedWriteBytes) {
        error = "WriteFile(pipe) failed: " + std::to_string(GetLastError());
        CloseHandle(pipe);
        return false;
    }

    const uint64_t readDeadlineTick = GetTickCount64() + effectiveTimeoutMs;
    for (;;) {
        DWORD availBytes = 0;
        if (!PeekNamedPipe(pipe, nullptr, 0, nullptr, &availBytes, nullptr)) {
            error = "PeekNamedPipe(pipe) failed: " + std::to_string(GetLastError());
            CloseHandle(pipe);
            return false;
        }
        if (availBytes > 0) {
            break;
        }
        if (GetTickCount64() >= readDeadlineTick) {
            error = "pipe read timeout";
            CloseHandle(pipe);
            return false;
        }
        Sleep(5);
    }

    std::array<char, 32768> buffer{};
    DWORD readBytes = 0;
    if (!ReadFile(pipe, buffer.data(), static_cast<DWORD>(buffer.size() - 1), &readBytes, nullptr) || readBytes == 0) {
        error = "ReadFile(pipe) failed: " + std::to_string(GetLastError());
        CloseHandle(pipe);
        return false;
    }
    buffer[readBytes] = '\0';
    response.assign(buffer.data(), buffer.data() + readBytes);
    while (!response.empty() && (response.back() == '\r' || response.back() == '\n' || response.back() == '\0')) {
        response.pop_back();
    }
    CloseHandle(pipe);
    return true;
}

bool MicMixApp::PingVstHost(std::string& response, std::string& error) const {
    return SendVstHostCommand("PING", response, 2500, error);
}

bool MicMixApp::EnsureVstAudioIpc() {
    if (vstAudioShared_.load(std::memory_order_acquire) != nullptr) {
        return true;
    }
    if (vstAudioShmName_.empty()) {
        vstAudioShmName_ = GenerateVstAudioShmName();
    }
    if (vstAudioShmName_.empty()) {
        return false;
    }
    HANDLE map = CreateFileMappingW(
        INVALID_HANDLE_VALUE,
        nullptr,
        PAGE_READWRITE,
        0,
        static_cast<DWORD>(sizeof(micmix::vstipc::SharedMemory)),
        vstAudioShmName_.c_str());
    if (!map) {
        return false;
    }
    void* view = MapViewOfFile(map, FILE_MAP_ALL_ACCESS, 0, 0, sizeof(micmix::vstipc::SharedMemory));
    if (!view) {
        CloseHandle(map);
        return false;
    }
    auto* shm = reinterpret_cast<micmix::vstipc::SharedMemory*>(view);
    if (GetLastError() != ERROR_ALREADY_EXISTS ||
        shm->magic != micmix::vstipc::kMagic ||
        shm->version != micmix::vstipc::kVersion) {
        micmix::vstipc::InitializeSharedMemory(*shm);
    }
    micmix::vstipc::RingReset(shm->musicIn);
    micmix::vstipc::RingReset(shm->musicOut);
    micmix::vstipc::RingReset(shm->micIn);
    micmix::vstipc::RingReset(shm->micOut);
    InterlockedExchange(&shm->hostHeartbeat, 0);
    InterlockedExchange(&shm->pluginHeartbeat, 0);
    vstAudioMap_ = map;
    vstAudioShared_.store(shm, std::memory_order_release);
    vstMusicSeq_.store(1U, std::memory_order_release);
    vstMicSeq_.store(1U, std::memory_order_release);
    return true;
}

void MicMixApp::CloseVstAudioIpc() {
    if (!WaitForCaptureCallbacksToDrain(2000U)) {
        LogWarn("vst ipc close postponed: capture callback still active");
        return;
    }
    auto* shared = vstAudioShared_.exchange(nullptr, std::memory_order_acq_rel);
    if (shared) {
        UnmapViewOfFile(shared);
    }
    if (vstAudioMap_) {
        CloseHandle(vstAudioMap_);
        vstAudioMap_ = nullptr;
    }
    vstMusicSeq_.store(1U, std::memory_order_release);
    vstMicSeq_.store(1U, std::memory_order_release);
}

bool MicMixApp::SyncVstHostState(const MicMixSettings& settings, std::string& error) {
    error.clear();
    std::string payload = BuildHostSyncPayload(settings);
    if (payload.size() > kVstHostSyncPayloadMaxBytes) {
        error = "vst sync payload too large";
        return false;
    }
    std::string response;
    if (!SendVstHostCommand("SYNC " + HexEncode(payload), response, 220, error)) {
        return false;
    }
    if (!StartsWithAscii(response, "SYNC_OK")) {
        error = response.empty() ? "invalid host sync response" : response;
        return false;
    }
    {
        std::lock_guard<std::mutex> settingsLock(settingsMutex_);
        UpdateAllSlotStatusesForUi(settings_, true);
    }
    {
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = response;
    }
    return true;
}

bool MicMixApp::StartVstHostProcess(std::string& error) {
    std::lock_guard<std::recursive_mutex> lifecycleLock(vstHostLifecycleMutex_);
    error.clear();
    bool alreadyRunning = false;
    {
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        if (vstHostProcess_) {
            const DWORD wait = WaitForSingleObject(vstHostProcess_, 0);
            if (wait == WAIT_TIMEOUT) {
                vstHostRunning_.store(true, std::memory_order_release);
                alreadyRunning = (vstHostPid_.load(std::memory_order_acquire) != 0);
            } else {
                CloseHandle(vstHostProcess_);
                vstHostProcess_ = nullptr;
                if (vstHostThread_) {
                    CloseHandle(vstHostThread_);
                    vstHostThread_ = nullptr;
                }
                vstHostRunning_.store(false, std::memory_order_release);
                vstHostPid_.store(0, std::memory_order_release);
            }
        }
    }

    if (alreadyRunning) {
        return true;
    }

    if (!vstHostJob_) {
        HANDLE job = CreateJobObjectW(nullptr, nullptr);
        if (!job) {
            error = "CreateJobObject failed: " + std::to_string(GetLastError());
            std::lock_guard<std::mutex> lock(vstHostMutex_);
            vstHostMessage_ = error;
            return false;
        }
        JOBOBJECT_EXTENDED_LIMIT_INFORMATION limits{};
        limits.BasicLimitInformation.LimitFlags =
            JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE |
            JOB_OBJECT_LIMIT_DIE_ON_UNHANDLED_EXCEPTION;
        if (!SetInformationJobObject(
                job,
                JobObjectExtendedLimitInformation,
                &limits,
                static_cast<DWORD>(sizeof(limits)))) {
            error = "SetInformationJobObject failed: " + std::to_string(GetLastError());
            CloseHandle(job);
            std::lock_guard<std::mutex> lock(vstHostMutex_);
            vstHostMessage_ = error;
            return false;
        }
        vstHostJob_ = job;
    }

    const std::wstring hostPath = ResolveVstHostPath();
    if (hostPath.empty() || !PathFileExistsW(hostPath.c_str())) {
        error = "vst host binary missing: " + WideToUtf8(hostPath);
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = error;
        return false;
    }
    if (vstHostPipeName_.empty()) {
        vstHostPipeName_ = GenerateVstHostPipeName();
    }
    if (vstHostPipeName_.empty()) {
        error = "vst host pipe token generation failed";
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = error;
        return false;
    }
    if (vstHostAuthToken_.empty()) {
        vstHostAuthToken_ = GenerateHexToken(16);
    }
    if (vstHostAuthToken_.empty()) {
        error = "vst host auth token generation failed";
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = error;
        return false;
    }
    if (vstAudioShmName_.empty()) {
        vstAudioShmName_ = GenerateVstAudioShmName();
    }
    if (vstAudioShmName_.empty()) {
        error = "vst shared memory token generation failed";
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = error;
        return false;
    }
    if (!EnsureVstAudioIpc()) {
        error = "vst audio shared memory init failed";
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = error;
        return false;
    }

    std::wstring cmdLine = L"\"";
    cmdLine += hostPath;
    cmdLine += L"\" --pipe ";
    cmdLine += vstHostPipeName_;
    cmdLine += L" --auth ";
    cmdLine += Utf8ToWide(vstHostAuthToken_);
    cmdLine += L" --shm ";
    cmdLine += vstAudioShmName_;
    STARTUPINFOW si{};
    si.cb = sizeof(si);
    PROCESS_INFORMATION pi{};
    if (!CreateProcessW(
            nullptr,
            cmdLine.data(),
            nullptr,
            nullptr,
            FALSE,
            CREATE_NO_WINDOW,
            nullptr,
            nullptr,
            &si,
            &pi)) {
        error = "CreateProcess failed: " + std::to_string(GetLastError());
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = error;
        return false;
    }
    if (!AssignProcessToJobObject(vstHostJob_, pi.hProcess)) {
        error = "AssignProcessToJobObject failed: " + std::to_string(GetLastError());
        TerminateProcess(pi.hProcess, 0);
        WaitForSingleObject(pi.hProcess, 120);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = error;
        return false;
    }
    {
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostProcess_ = pi.hProcess;
        vstHostThread_ = pi.hThread;
        vstHostPid_.store(pi.dwProcessId, std::memory_order_release);
        vstHostRunning_.store(true, std::memory_order_release);
        vstHostMessage_ = "starting";
    }
    LogInfo("vst_host started pid=" + std::to_string(pi.dwProcessId));

    std::string pingResponse;
    std::string pingError;
    if (!PingVstHost(pingResponse, pingError) || !StartsWithAscii(pingResponse, "PONG")) {
        error = "vst host handshake failed: " + (pingError.empty() ? pingResponse : pingError);
        LogWarn(error);
        StopVstHostProcess();
        return false;
    }

    MicMixSettings snapshot;
    {
        std::lock_guard<std::mutex> settingsLock(settingsMutex_);
        snapshot = settings_;
    }
    std::string syncError;
    if (!SyncVstHostState(snapshot, syncError)) {
        error = "vst host sync failed: " + syncError;
        LogWarn(error);
        StopVstHostProcess();
        return false;
    }

    {
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = pingResponse;
    }
    if (!EnsureVstAudioIpc()) {
        error = "vst audio shared memory init failed";
        LogWarn(error);
        StopVstHostProcess();
        return false;
    }
    auto* shared = vstAudioShared_.load(std::memory_order_acquire);
    if (shared) {
        micmix::vstipc::RingReset(shared->musicIn);
        micmix::vstipc::RingReset(shared->musicOut);
        micmix::vstipc::RingReset(shared->micIn);
        micmix::vstipc::RingReset(shared->micOut);
        InterlockedExchange(&shared->hostHeartbeat, 0);
        InterlockedExchange(&shared->pluginHeartbeat, 0);
    }
    vstMusicSeq_.store(1U, std::memory_order_release);
    vstMicSeq_.store(1U, std::memory_order_release);
    vstHostLastHeartbeatTickMs_.store(GetTickCount64(), std::memory_order_release);
    vstHostStopPending_.store(false, std::memory_order_release);
    vstHostSyncPending_.store(false, std::memory_order_release);
    return true;
}

void MicMixApp::StopVstHostProcess() {
    std::lock_guard<std::recursive_mutex> lifecycleLock(vstHostLifecycleMutex_);
    {
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        if (!vstHostProcess_) {
            vstHostRunning_.store(false, std::memory_order_release);
            vstHostPid_.store(0, std::memory_order_release);
            if (vstHostThread_) {
                CloseHandle(vstHostThread_);
                vstHostThread_ = nullptr;
            }
            vstHostLastHeartbeatTickMs_.store(0, std::memory_order_release);
            vstHostStopPending_.store(false, std::memory_order_release);
            vstHostSyncPending_.store(false, std::memory_order_release);
            return;
        }
    }

    std::string quitResponse;
    std::string quitError;
    SendVstHostCommand("QUIT", quitResponse, 80, quitError);

    {
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        const DWORD wait = WaitForSingleObject(vstHostProcess_, 180);
        if (wait == WAIT_TIMEOUT) {
            TerminateProcess(vstHostProcess_, 0);
            WaitForSingleObject(vstHostProcess_, 120);
        }
        CloseHandle(vstHostProcess_);
        vstHostProcess_ = nullptr;
        if (vstHostThread_) {
            CloseHandle(vstHostThread_);
            vstHostThread_ = nullptr;
        }
        vstHostRunning_.store(false, std::memory_order_release);
        vstHostPid_.store(0, std::memory_order_release);
        vstHostMessage_ = "stopped";
    }
    vstHostLastHeartbeatTickMs_.store(0, std::memory_order_release);
    vstHostStopPending_.store(false, std::memory_order_release);
    vstHostSyncPending_.store(false, std::memory_order_release);
    LogInfo("vst_host stopped");
}

VstHostStatus MicMixApp::GetVstHostStatus() const {
    VstHostStatus status{};
    status.running = vstHostRunning_.load(std::memory_order_acquire);
    status.pid = vstHostPid_.load(std::memory_order_acquire);
    {
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        if (vstHostProcess_) {
            const DWORD wait = WaitForSingleObject(vstHostProcess_, 0);
            if (wait != WAIT_TIMEOUT) {
                status.running = false;
                status.pid = 0;
                status.message = "stopped";
            } else if (vstHostMessage_.empty()) {
                status.message = "running";
            }
        }
        if (status.message.empty()) {
            status.message = vstHostMessage_;
        }
    }
    return status;
}

void MicMixApp::MaintainVstHost(uint64_t nowMs) {
    MicMixSettings snapshot;
    bool hostWanted = false;
    {
        std::lock_guard<std::mutex> settingsLock(settingsMutex_);
        snapshot = settings_;
        hostWanted = snapshot.vstEffectsEnabled || snapshot.vstHostAutostart;
    }
    const uint64_t editorKeepAliveUntil = vstHostEditorKeepAliveUntilMs_.load(std::memory_order_acquire);
    if (!hostWanted && editorKeepAliveUntil != 0 && nowMs < editorKeepAliveUntil) {
        hostWanted = true;
    }

    if (!hostWanted) {
        vstHostEditorKeepAliveUntilMs_.store(0, std::memory_order_release);
        const bool stopPending = vstHostStopPending_.load(std::memory_order_acquire);
        if (stopPending || vstHostRunning_.load(std::memory_order_acquire) ||
            (vstHostPid_.load(std::memory_order_acquire) != 0U)) {
            LogInfo("vst_host stop requested: host_not_wanted");
            StopVstHostProcess();
        }
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostSyncPending_.store(false, std::memory_order_release);
        vstHostNextRestartTickMs_.store(0, std::memory_order_release);
        vstHostRestartAttempts_.store(0, std::memory_order_release);
        return;
    }

    bool alive = false;
    {
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        if (vstHostProcess_) {
            const DWORD wait = WaitForSingleObject(vstHostProcess_, 0);
            if (wait == WAIT_TIMEOUT) {
                alive = true;
            } else {
                DWORD exitCode = 0;
                GetExitCodeProcess(vstHostProcess_, &exitCode);
                CloseHandle(vstHostProcess_);
                vstHostProcess_ = nullptr;
                if (vstHostThread_) {
                    CloseHandle(vstHostThread_);
                    vstHostThread_ = nullptr;
                }
                vstHostRunning_.store(false, std::memory_order_release);
                vstHostPid_.store(0, std::memory_order_release);
                vstHostMessage_ = "exited:" + std::to_string(exitCode);
            }
        }
    }

    if (alive) {
        vstHostNextRestartTickMs_.store(0, std::memory_order_release);
        vstHostRestartAttempts_.store(0, std::memory_order_release);
        auto* shm = vstAudioShared_.load(std::memory_order_acquire);
        if (shm) {
            const LONG hostHeartbeat = micmix::vstipc::AtomicLoad(&shm->hostHeartbeat);
            if (hostHeartbeat != 0) {
                const uint32_t nowTick32 = static_cast<uint32_t>(GetTickCount());
                const uint32_t hbTick32 = static_cast<uint32_t>(hostHeartbeat);
                if ((nowTick32 - hbTick32) > 3000U) {
                    LogWarn("vst_host shared-memory heartbeat stale; restarting host");
                    StopVstHostProcess();
                    vstHostNextRestartTickMs_.store(nowMs + 200ULL, std::memory_order_release);
                    return;
                }
            }
        }
        const uint64_t lastHeartbeat = vstHostLastHeartbeatTickMs_.load(std::memory_order_acquire);
        if (lastHeartbeat == 0 || nowMs > (lastHeartbeat + 2500ULL)) {
            std::string hbResponse;
            std::string hbError;
            if (PingVstHost(hbResponse, hbError) && StartsWithAscii(hbResponse, "PONG")) {
                vstHostLastHeartbeatTickMs_.store(nowMs, std::memory_order_release);
                std::lock_guard<std::mutex> lock(vstHostMutex_);
                vstHostMessage_ = hbResponse;
            } else {
                LogWarn("vst_host heartbeat timeout: " + (hbError.empty() ? hbResponse : hbError));
                StopVstHostProcess();
                vstHostNextRestartTickMs_.store(nowMs + 200ULL, std::memory_order_release);
                return;
            }
        }
        if (!vstHostSyncPending_.load(std::memory_order_acquire)) {
            return;
        }
        std::string syncError;
        if (SyncVstHostState(snapshot, syncError)) {
            vstHostSyncPending_.store(false, std::memory_order_release);
            return;
        }
        const uint64_t lastLogMs = vstHostLastRestartLogTickMs_.load(std::memory_order_acquire);
        if (nowMs > (lastLogMs + 800ULL)) {
            vstHostLastRestartLogTickMs_.store(nowMs, std::memory_order_release);
            LogWarn("vst_host sync retry pending: " + syncError);
        }
        return;
    }

    const uint64_t nextRestart = vstHostNextRestartTickMs_.load(std::memory_order_acquire);
    if (nextRestart != 0 && nowMs < nextRestart) {
        return;
    }

    std::string startError;
    if (StartVstHostProcess(startError)) {
        vstHostSyncPending_.store(false, std::memory_order_release);
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostNextRestartTickMs_.store(0, std::memory_order_release);
        vstHostRestartAttempts_.store(0, std::memory_order_release);
        return;
    }

    vstHostSyncPending_.store(true, std::memory_order_release);
    const uint32_t prevAttempts = vstHostRestartAttempts_.load(std::memory_order_acquire);
    const uint32_t attempts = std::min<uint32_t>(prevAttempts + 1U, 10U);
    vstHostRestartAttempts_.store(attempts, std::memory_order_release);
    if (attempts >= 8U) {
        // Keep user settings intact; only throttle retries after repeated failures.
        constexpr uint64_t kBlockedRetryDelayMs = 10000ULL;
        {
            std::lock_guard<std::mutex> lock(vstHostMutex_);
            vstHostMessage_ = "blocked_after_retries";
        }
        vstHostSyncPending_.store(true, std::memory_order_release);
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostNextRestartTickMs_.store(nowMs + kBlockedRetryDelayMs, std::memory_order_release);

        const uint64_t lastLogMs = vstHostLastRestartLogTickMs_.load(std::memory_order_acquire);
        if (nowMs > (lastLogMs + 1200ULL)) {
            vstHostLastRestartLogTickMs_.store(nowMs, std::memory_order_release);
            LogWarn(
                "vst_host blocked_after_retries attempt=" + std::to_string(attempts) +
                " retry_in_ms=" + std::to_string(kBlockedRetryDelayMs) +
                " cause=" + startError);
        }
        return;
    }
    const uint32_t shift = std::min<uint32_t>(attempts, 6U);
    const uint64_t delayMs = std::min<uint64_t>(10000ULL, 150ULL << shift);
    vstHostNextRestartTickMs_.store(nowMs + delayMs, std::memory_order_release);
    {
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = "restart_in_ms=" + std::to_string(delayMs);
    }

    const uint64_t lastLogMs = vstHostLastRestartLogTickMs_.load(std::memory_order_acquire);
    if (nowMs > (lastLogMs + 800ULL)) {
        vstHostLastRestartLogTickMs_.store(nowMs, std::memory_order_release);
        LogWarn(
            "vst_host restart attempt=" + std::to_string(attempts) +
            " in_ms=" + std::to_string(delayMs) +
            " cause=" + startError);
    }
}

bool MicMixApp::IsEffectsEnabled() const {
    std::lock_guard<std::mutex> lock(settingsMutex_);
    return settings_.vstEffectsEnabled;
}

void MicMixApp::SetEffectsEnabled(bool enabled, bool saveConfig) {
    MicMixSettings snapshot;
    bool changed = false;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        if (settings_.vstEffectsEnabled != enabled) {
            settings_.vstEffectsEnabled = enabled;
            changed = true;
        }
        UpdateAllSlotStatusesForUi(settings_, false);
        snapshot = settings_;
    }
    vstEffectsEnabledCached_.store(snapshot.vstEffectsEnabled, std::memory_order_release);
    if (!changed) {
        return;
    }

    if (snapshot.vstEffectsEnabled || snapshot.vstHostAutostart) {
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostSyncPending_.store(true, std::memory_order_release);
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = "sync_pending";
    } else {
        vstHostSyncPending_.store(false, std::memory_order_release);
        vstHostStopPending_.store(true, std::memory_order_release);
        std::lock_guard<std::mutex> lock(vstHostMutex_);
        vstHostMessage_ = "stop_pending";
    }

    if (saveConfig && configStore_) {
        std::string err;
        if (!configStore_->Save(snapshot, err) && !err.empty()) {
            LogError("Config save failed: " + err);
        }
    }
}

std::vector<VstEffectSlot> MicMixApp::GetEffects(EffectChain chain) const {
    std::lock_guard<std::mutex> lock(settingsMutex_);
    if (chain == EffectChain::Music) {
        return settings_.musicEffects;
    }
    return settings_.micEffects;
}

bool MicMixApp::AddEffect(EffectChain chain, const VstEffectSlot& slot, std::string& error) {
    error.clear();
    VstEffectSlot safe = slot;
    std::vector<VstEffectSlot> tmp{safe};
    SanitizeEffectList(tmp);
    if (tmp.empty()) {
        error = "invalid effect slot";
        return false;
    }
    safe = tmp.front();
    std::string normalizedPath;
    if (!NormalizeAndValidateEffectPath(safe.path, normalizedPath, error)) {
        return false;
    }
    safe.path = normalizedPath;
    safe.lastStatus = "pending";
    MicMixSettings snapshot;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        auto& list = (chain == EffectChain::Music) ? settings_.musicEffects : settings_.micEffects;
        if (list.size() >= kMaxEffectsPerChain) {
            error = "effect limit reached";
            return false;
        }
        safe.uid = BuildUniqueEffectUidWithExists(
            safe.path,
            [&list](const std::string& uid) {
                return std::any_of(list.begin(), list.end(), [&uid](const VstEffectSlot& slot) {
                    return slot.uid == uid;
                });
            });
        list.push_back(safe);
        SanitizeEffectList(list);
        UpdateAllSlotStatusesForUi(settings_, false);
        snapshot = settings_;
    }
    if (snapshot.vstEffectsEnabled || snapshot.vstHostAutostart) {
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostSyncPending_.store(true, std::memory_order_release);
    }
    if (configStore_) {
        std::string saveErr;
        if (!configStore_->Save(snapshot, saveErr) && !saveErr.empty()) {
            LogError("Config save failed: " + saveErr);
        }
    }
    return true;
}

bool MicMixApp::RemoveEffect(EffectChain chain, size_t index, std::string& error) {
    error.clear();
    MicMixSettings snapshot;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        auto& list = (chain == EffectChain::Music) ? settings_.musicEffects : settings_.micEffects;
        if (index >= list.size()) {
            error = "invalid index";
            return false;
        }
        list.erase(list.begin() + static_cast<std::ptrdiff_t>(index));
        UpdateAllSlotStatusesForUi(settings_, false);
        snapshot = settings_;
    }
    if (snapshot.vstEffectsEnabled || snapshot.vstHostAutostart) {
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostSyncPending_.store(true, std::memory_order_release);
    }
    if (configStore_) {
        std::string saveErr;
        if (!configStore_->Save(snapshot, saveErr) && !saveErr.empty()) {
            LogError("Config save failed: " + saveErr);
        }
    }
    return true;
}

bool MicMixApp::MoveEffect(EffectChain chain, size_t fromIndex, size_t toIndex, std::string& error) {
    error.clear();
    MicMixSettings snapshot;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        auto& list = (chain == EffectChain::Music) ? settings_.musicEffects : settings_.micEffects;
        if (fromIndex >= list.size() || toIndex >= list.size()) {
            error = "invalid move index";
            return false;
        }
        if (fromIndex == toIndex) {
            return true;
        }
        VstEffectSlot moved = list[fromIndex];
        list.erase(list.begin() + static_cast<std::ptrdiff_t>(fromIndex));
        list.insert(list.begin() + static_cast<std::ptrdiff_t>(toIndex), std::move(moved));
        UpdateAllSlotStatusesForUi(settings_, false);
        snapshot = settings_;
    }
    if (snapshot.vstEffectsEnabled || snapshot.vstHostAutostart) {
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostSyncPending_.store(true, std::memory_order_release);
    }
    if (configStore_) {
        std::string saveErr;
        if (!configStore_->Save(snapshot, saveErr) && !saveErr.empty()) {
            LogError("Config save failed: " + saveErr);
        }
    }
    return true;
}

bool MicMixApp::SetEffectBypass(EffectChain chain, size_t index, bool bypass, std::string& error) {
    error.clear();
    MicMixSettings snapshot;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        auto& list = (chain == EffectChain::Music) ? settings_.musicEffects : settings_.micEffects;
        if (index >= list.size()) {
            error = "invalid index";
            return false;
        }
        list[index].bypass = bypass;
        UpdateAllSlotStatusesForUi(settings_, false);
        snapshot = settings_;
    }
    if (snapshot.vstEffectsEnabled || snapshot.vstHostAutostart) {
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostSyncPending_.store(true, std::memory_order_release);
    }
    if (configStore_) {
        std::string saveErr;
        if (!configStore_->Save(snapshot, saveErr) && !saveErr.empty()) {
            LogError("Config save failed: " + saveErr);
        }
    }
    return true;
}

bool MicMixApp::SetEffectEnabled(EffectChain chain, size_t index, bool enabled, std::string& error) {
    error.clear();
    MicMixSettings snapshot;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        auto& list = (chain == EffectChain::Music) ? settings_.musicEffects : settings_.micEffects;
        if (index >= list.size()) {
            error = "invalid index";
            return false;
        }
        list[index].enabled = enabled;
        UpdateAllSlotStatusesForUi(settings_, false);
        snapshot = settings_;
    }
    if (snapshot.vstEffectsEnabled || snapshot.vstHostAutostart) {
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostSyncPending_.store(true, std::memory_order_release);
    }
    if (configStore_) {
        std::string saveErr;
        if (!configStore_->Save(snapshot, saveErr) && !saveErr.empty()) {
            LogError("Config save failed: " + saveErr);
        }
    }
    return true;
}

bool MicMixApp::OpenEffectEditor(EffectChain chain, size_t index, std::string& error) {
    error.clear();
    const uint64_t nowMs = GetTickCount64();
    constexpr uint64_t kEditorKeepAliveMs = 900000ULL;
    vstHostEditorKeepAliveUntilMs_.store(nowMs + kEditorKeepAliveMs, std::memory_order_release);
    vstHostStopPending_.store(false, std::memory_order_release);

    MicMixSettings snapshot;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        const auto& list = (chain == EffectChain::Music) ? settings_.musicEffects : settings_.micEffects;
        if (index >= list.size()) {
            error = "invalid index";
            return false;
        }
        snapshot = settings_;
    }

    if (!vstHostRunning_.load(std::memory_order_acquire)) {
        std::string startError;
        if (!StartVstHostProcess(startError)) {
            error = startError.empty() ? "vst host start failed" : startError;
            return false;
        }
    }

    std::string syncError;
    if (!SyncVstHostState(snapshot, syncError)) {
        error = syncError.empty() ? "vst host sync failed" : syncError;
        return false;
    }

    const char* chainText = (chain == EffectChain::Music) ? "music" : "mic";
    const std::string command = std::string("EDITOR_OPEN ") + chainText + " " + std::to_string(index);
    std::string response;
    LogInfo("editor_open request chain=" + std::string(chainText) + " index=" + std::to_string(index));
    if (!SendVstHostCommand(command, response, 2500, error)) {
        LogWarn("editor_open transport failed chain=" + std::string(chainText) + " index=" + std::to_string(index) +
                " err=" + (error.empty() ? std::string("unknown") : error));
        return false;
    }
    if (!StartsWithAscii(response, "EDITOR_OK")) {
        error = response.empty() ? "editor_open_failed" : response;
        LogWarn("editor_open rejected chain=" + std::string(chainText) + " index=" + std::to_string(index) +
                " response=" + (response.empty() ? std::string("empty") : response));
        return false;
    }
    LogInfo("editor_open accepted chain=" + std::string(chainText) + " index=" + std::to_string(index) +
            " response=" + response);
    vstHostEditorKeepAliveUntilMs_.store(GetTickCount64() + kEditorKeepAliveMs, std::memory_order_release);
    return true;
}

namespace {
bool IsConnectedForTx(uint64 schid) {
    if (schid == 0 || !g_ts3Functions.getConnectionStatus) {
        return false;
    }
    int status = STATUS_DISCONNECTED;
    const unsigned int err = g_ts3Functions.getConnectionStatus(schid, &status);
    if (!(err == ERROR_ok || err == ERROR_ok_no_update)) {
        return false;
    }
    return status == STATUS_CONNECTED || status == STATUS_CONNECTION_ESTABLISHED;
}

// These keys are mirrored only if TS3 accepts them. Besides official keys, this
// includes community-observed idents used by some TS3 versions/builds.
constexpr std::array<const char*, 18> kMirroredPreprocessorIdents = {
    "agc",
    "agc_level",
    "agc_max_gain",
    "denoise",
    "echo_canceling",
    "typing_suppression",
    "echo_reduction",
    "echo_reduction_db",
    "vad_mode",
    "vad_rnn",
    "denoiser_level",
    "aec",
    "echo_cancellation",
    "vad_over_ptt",
    "delay_ptt",
    "delay_ptt_msecs",
    "continous_transmission",
    "continuous_transmission",
};

enum class VoiceRecordingState {
    Inactive = 0,
    ActiveSameServer = 1,
    ActiveOtherServer = 2,
};

VoiceRecordingState ResolveVoiceRecordingState(bool currentlyActive, uint64 currentSchid, uint64 targetSchid) {
    if (!currentlyActive || currentSchid == 0) {
        return VoiceRecordingState::Inactive;
    }
    if (currentSchid == targetSchid) {
        return VoiceRecordingState::ActiveSameServer;
    }
    return VoiceRecordingState::ActiveOtherServer;
}
} // namespace

bool MicMixApp::TryEnterCaptureCallback() {
    if (shutdownRequested_.load(std::memory_order_acquire)) {
        return false;
    }
    captureCallbacksInFlight_.fetch_add(1, std::memory_order_acq_rel);
    if (shutdownRequested_.load(std::memory_order_acquire)) {
        LeaveCaptureCallback();
        return false;
    }
    return true;
}

void MicMixApp::LeaveCaptureCallback() {
    uint32_t current = captureCallbacksInFlight_.load(std::memory_order_acquire);
    while (current > 0) {
        if (captureCallbacksInFlight_.compare_exchange_weak(
                current, current - 1, std::memory_order_acq_rel, std::memory_order_acquire)) {
            return;
        }
    }
}

bool MicMixApp::WaitForCaptureCallbacksToDrain(uint32_t maxWaitMs) const {
    const uint64_t deadline = GetTickCount64() + static_cast<uint64_t>(maxWaitMs);
    while (captureCallbacksInFlight_.load(std::memory_order_acquire) > 0U) {
        if (GetTickCount64() >= deadline) {
            return false;
        }
        Sleep(2);
    }
    return true;
}

void MicMixApp::SetVoiceRecordingState(bool active, uint64 schid) {
    std::lock_guard<std::mutex> lock(voiceTxMutex_);
    static uint64_t lastErrLogMs = 0;
    auto appendIdent = [](std::string& out, const char* ident) {
        if (!out.empty()) {
            out += ",";
        }
        out += ident;
    };

    const bool currentlyActive = voiceRecordingActive_.load(std::memory_order_relaxed);
    const uint64 currentSchid = voiceRecordingSchid_.load(std::memory_order_relaxed);
    auto resetSavedState = [this]() {
        savedInputStateValid_ = false;
        savedVadValid_ = false;
        savedVadThresholdValid_ = false;
        savedVadThresholdDbfs_ = -50.0f;
        savedVadExtraBufferValid_ = false;
        savedVadExtraBufferSize_ = 0;
        for (size_t i = 0; i < savedPreprocessorValuesValid_.size(); ++i) {
            savedPreprocessorValuesValid_[i] = false;
            savedPreprocessorValues_[i].clear();
        }
    };

    auto clearState = [this, &resetSavedState]() {
        resetSavedState();
        voiceRecordingActive_.store(false, std::memory_order_release);
        voiceRecordingSchid_.store(0, std::memory_order_release);
        voiceTxLastNudgeMs_.store(0, std::memory_order_release);
        forceTxHoldUntilMs_.store(0, std::memory_order_release);
        engine_.SetTalkState(true);
    };

    // During plugin shutdown, avoid any further TeamSpeak API calls from this path.
    // The client is tearing down and late preprocessor/input updates can crash TS3.
    if (shutdownRequested_.load(std::memory_order_acquire)) {
        clearState();
        return;
    }

    auto restoreStateForSchid = [this, &appendIdent](uint64 targetSchid) {
        bool ok = true;
        size_t restoredCount = 0;
        std::string restoredKeys;
        if (savedVadValid_ && g_ts3Functions.setPreProcessorConfigValue) {
            const unsigned int errVad = g_ts3Functions.setPreProcessorConfigValue(
                targetSchid, "vad", savedVadEnabled_ ? "true" : "false");
            ok = ok && (errVad == ERROR_ok || errVad == ERROR_ok_no_update);
            if (errVad == ERROR_ok || errVad == ERROR_ok_no_update) {
                ++restoredCount;
                appendIdent(restoredKeys, "vad");
            }
        }
        for (size_t i = 0; i < kMirroredPreprocessorIdents.size(); ++i) {
            if (!savedPreprocessorValuesValid_[i] || !g_ts3Functions.setPreProcessorConfigValue) {
                continue;
            }
            const unsigned int err = g_ts3Functions.setPreProcessorConfigValue(
                targetSchid, kMirroredPreprocessorIdents[i], savedPreprocessorValues_[i].c_str());
            if (err == ERROR_ok || err == ERROR_ok_no_update) {
                ++restoredCount;
                appendIdent(restoredKeys, kMirroredPreprocessorIdents[i]);
            } else {
                const uint64_t nowMs = GetTickCount64();
                if (nowMs > (lastErrLogMs + 4000ULL)) {
                    LogWarn("talk_mode restore preproc failed schid=" + std::to_string(targetSchid) +
                            " ident=" + kMirroredPreprocessorIdents[i] +
                            " err=" + std::to_string(err));
                    lastErrLogMs = nowMs;
                }
            }
        }
        if (savedInputStateValid_ && g_ts3Functions.setClientSelfVariableAsInt) {
            const unsigned int errInput = g_ts3Functions.setClientSelfVariableAsInt(
                targetSchid, CLIENT_INPUT_DEACTIVATED, savedInputDeactivated_);
            ok = ok && (errInput == ERROR_ok || errInput == ERROR_ok_no_update);
        }
        if (g_ts3Functions.flushClientSelfUpdates) {
            const unsigned int errFlush = g_ts3Functions.flushClientSelfUpdates(targetSchid, nullptr);
            ok = ok && (errFlush == ERROR_ok || errFlush == ERROR_ok_no_update);
        } else {
            ok = false;
        }
        if (!restoredKeys.empty()) {
            LogInfo("talk_mode restore preproc schid=" + std::to_string(targetSchid) +
                    " restored=" + std::to_string(restoredCount) +
                    " keys=" + restoredKeys);
        }
        return ok;
    };

    auto forceDeactivateInputForSchid = [](uint64 targetSchid) {
        if (!g_ts3Functions.setClientSelfVariableAsInt || !g_ts3Functions.flushClientSelfUpdates) {
            return false;
        }
        const unsigned int errInput = g_ts3Functions.setClientSelfVariableAsInt(
            targetSchid, CLIENT_INPUT_DEACTIVATED, INPUT_DEACTIVATED);
        const unsigned int errFlush = g_ts3Functions.flushClientSelfUpdates(targetSchid, nullptr);
        const bool okInput = (errInput == ERROR_ok || errInput == ERROR_ok_no_update);
        const bool okFlush = (errFlush == ERROR_ok || errFlush == ERROR_ok_no_update);
        return okInput && okFlush;
    };

    if (!active || schid == 0 || !IsConnectedForTx(schid)) {
        if (currentlyActive && currentSchid != 0 && IsConnectedForTx(currentSchid)) {
            const bool restored = restoreStateForSchid(currentSchid);
            if (!restored) {
                const bool forcedDeactivate = forceDeactivateInputForSchid(currentSchid);
                const uint64_t nowMs = GetTickCount64();
                if (nowMs > (lastErrLogMs + 2000ULL)) {
                    LogWarn("talk_mode restore failed schid=" + std::to_string(currentSchid) +
                            " fallback_input_deactivate=" + std::to_string(forcedDeactivate ? 1 : 0));
                    lastErrLogMs = nowMs;
                }
            } else {
                LogInfo("talk_mode restored schid=" + std::to_string(currentSchid));
            }
        }
        clearState();
        return;
    }

    if (!g_ts3Functions.getClientSelfVariableAsInt ||
        !g_ts3Functions.getPreProcessorConfigValue ||
        !g_ts3Functions.setClientSelfVariableAsInt ||
        !g_ts3Functions.setPreProcessorConfigValue ||
        !g_ts3Functions.flushClientSelfUpdates) {
        clearState();
        return;
    }

    const uint64_t nowMs = GetTickCount64();
    const uint64_t lastNudgeMs = voiceTxLastNudgeMs_.load(std::memory_order_relaxed);
    const VoiceRecordingState recordingState = ResolveVoiceRecordingState(currentlyActive, currentSchid, schid);
    const bool sameSchid = (recordingState == VoiceRecordingState::ActiveSameServer);
    const MicMixSettings runtimeSettings = GetSettings();
    const bool forceTsFilters = runtimeSettings.micForceTsFilters;
    if (sameSchid && nowMs <= (lastNudgeMs + kVoiceTxReapplyMs)) {
        return;
    }

    if (recordingState == VoiceRecordingState::ActiveOtherServer) {
        if (IsConnectedForTx(currentSchid)) {
            restoreStateForSchid(currentSchid);
        }
        resetSavedState();
    }

    int inputState = INPUT_DEACTIVATED;
    const unsigned int errGetInput = g_ts3Functions.getClientSelfVariableAsInt(
        schid, CLIENT_INPUT_DEACTIVATED, &inputState);
    auto readPreprocessorConfigValue = [schid](const char* ident, std::string& outValue) -> bool {
        if (!g_ts3Functions.getPreProcessorConfigValue) {
            return false;
        }
        char* value = nullptr;
        const unsigned int err = g_ts3Functions.getPreProcessorConfigValue(schid, ident, &value);
        const bool ok = (err == ERROR_ok || err == ERROR_ok_no_update) && value;
        if (ok) {
            outValue.assign(value);
        }
        if (value && g_ts3Functions.freeMemory) {
            g_ts3Functions.freeMemory(value);
        }
        return ok;
    };
    char* vadStr = nullptr;
    const unsigned int errGetVad = g_ts3Functions.getPreProcessorConfigValue(schid, "vad", &vadStr);
    bool vadEnabled = false;
    if ((errGetVad == ERROR_ok || errGetVad == ERROR_ok_no_update) && vadStr) {
        vadEnabled = (_stricmp(vadStr, "true") == 0);
    }
    if (vadStr && g_ts3Functions.freeMemory) {
        g_ts3Functions.freeMemory(vadStr);
    }
    float vadThreshold = -50.0f;
    bool vadThresholdValid = false;
    char* vadLevelStr = nullptr;
    const unsigned int errGetVadLevel = g_ts3Functions.getPreProcessorConfigValue(schid, "voiceactivation_level", &vadLevelStr);
    if ((errGetVadLevel == ERROR_ok || errGetVadLevel == ERROR_ok_no_update) && vadLevelStr) {
        char* endPtr = nullptr;
        const float parsed = std::strtof(vadLevelStr, &endPtr);
        if (endPtr && endPtr != vadLevelStr) {
            vadThreshold = std::clamp(parsed, -90.0f, 0.0f);
            vadThresholdValid = true;
        }
    }
    if (vadLevelStr && g_ts3Functions.freeMemory) {
        g_ts3Functions.freeMemory(vadLevelStr);
    }
    bool vadExtraBufferValid = false;
    int vadExtraBufferSize = 0;
    std::string vadExtraBufferStr;
    if (readPreprocessorConfigValue("vad_extrabuffersize", vadExtraBufferStr)) {
        char* endPtr = nullptr;
        const long parsed = std::strtol(vadExtraBufferStr.c_str(), &endPtr, 10);
        if (endPtr && endPtr != vadExtraBufferStr.c_str()) {
            vadExtraBufferSize = static_cast<int>(std::clamp(parsed, 0L, 50L));
            vadExtraBufferValid = true;
        }
    }

    std::array<std::string, kMirroredPreprocessorIdents.size()> mirroredValues{};
    std::array<bool, kMirroredPreprocessorIdents.size()> mirroredValueValid{};
    for (size_t i = 0; i < kMirroredPreprocessorIdents.size(); ++i) {
        mirroredValueValid[i] = readPreprocessorConfigValue(kMirroredPreprocessorIdents[i], mirroredValues[i]);
    }
    if (!sameSchid) {
        std::string supported;
        std::string unsupported;
        for (size_t i = 0; i < kMirroredPreprocessorIdents.size(); ++i) {
            if (mirroredValueValid[i]) {
                appendIdent(supported, kMirroredPreprocessorIdents[i]);
            } else {
                appendIdent(unsupported, kMirroredPreprocessorIdents[i]);
            }
        }
        LogInfo("talk_mode probe preproc schid=" + std::to_string(schid) +
                " supported=[" + supported + "] unsupported=[" + unsupported + "]");
    }

    if (!(errGetInput == ERROR_ok || errGetInput == ERROR_ok_no_update) ||
        !(errGetVad == ERROR_ok || errGetVad == ERROR_ok_no_update)) {
        if (nowMs > (lastErrLogMs + 2000ULL)) {
            LogWarn("talk_mode read failed schid=" + std::to_string(schid) +
                    " input_err=" + std::to_string(errGetInput) +
                    " vad_err=" + std::to_string(errGetVad));
            lastErrLogMs = nowMs;
        }
        return;
    }

    if (!currentlyActive || currentSchid != schid) {
        savedInputStateValid_ = true;
        savedInputDeactivated_ = inputState;
        savedVadValid_ = true;
        savedVadEnabled_ = vadEnabled;
        savedVadThresholdValid_ = vadThresholdValid;
        savedVadThresholdDbfs_ = vadThreshold;
        savedVadExtraBufferValid_ = vadExtraBufferValid;
        savedVadExtraBufferSize_ = vadExtraBufferSize;
        for (size_t i = 0; i < kMirroredPreprocessorIdents.size(); ++i) {
            savedPreprocessorValuesValid_[i] = mirroredValueValid[i];
            if (mirroredValueValid[i]) {
                savedPreprocessorValues_[i] = mirroredValues[i];
            } else {
                savedPreprocessorValues_[i].clear();
            }
        }
    }

    // Keep transport always open while force-send path is active, otherwise
    // TS may only start sending after first local voice trigger.
    const char* desiredVadValue = "false";
    const char* currentVadValue = vadEnabled ? "true" : "false";
    const bool needVadUpdate = !sameSchid || (_stricmp(currentVadValue, desiredVadValue) != 0);
    if (!sameSchid) {
        size_t keepAppliedCount = 0;
        std::string keepAppliedKeys;
        std::string forcedFilterKeys;
        for (size_t i = 0; i < kMirroredPreprocessorIdents.size(); ++i) {
            if (!savedPreprocessorValuesValid_[i] || !g_ts3Functions.setPreProcessorConfigValue) {
                continue;
            }
            std::string applyValue = savedPreprocessorValues_[i];
            const bool isTypingSuppression = (_stricmp(kMirroredPreprocessorIdents[i], "typing_suppression") == 0);
            const bool isDenoise = (_stricmp(kMirroredPreprocessorIdents[i], "denoise") == 0);
            if (forceTsFilters && (isTypingSuppression || isDenoise) && _stricmp(applyValue.c_str(), "true") != 0) {
                applyValue = "true";
                appendIdent(forcedFilterKeys, kMirroredPreprocessorIdents[i]);
            }
            const unsigned int err = g_ts3Functions.setPreProcessorConfigValue(
                schid, kMirroredPreprocessorIdents[i], applyValue.c_str());
            if (err == ERROR_ok || err == ERROR_ok_no_update) {
                ++keepAppliedCount;
                appendIdent(keepAppliedKeys, kMirroredPreprocessorIdents[i]);
            } else {
                if (nowMs > (lastErrLogMs + 4000ULL)) {
                    LogWarn("talk_mode keep preproc failed schid=" + std::to_string(schid) +
                            " ident=" + kMirroredPreprocessorIdents[i] +
                            " err=" + std::to_string(err));
                    lastErrLogMs = nowMs;
                }
            }
        }
        if (!keepAppliedKeys.empty()) {
            LogInfo("talk_mode keep preproc schid=" + std::to_string(schid) +
                    " applied=" + std::to_string(keepAppliedCount) +
                    " keys=" + keepAppliedKeys);
        }
        if (!forcedFilterKeys.empty()) {
            LogInfo("talk_mode forced filters schid=" + std::to_string(schid) +
                    " keys=" + forcedFilterKeys + " value=true");
        }
    }
    int desiredInputState = INPUT_ACTIVE;
    const bool needInputUpdate = !sameSchid || (inputState != desiredInputState);
    const bool needFlush = !sameSchid || needVadUpdate || needInputUpdate;
    if (sameSchid && !needFlush) {
        voiceTxLastNudgeMs_.store(nowMs, std::memory_order_release);
        return;
    }

    unsigned int errVad = ERROR_ok_no_update;
    if (needVadUpdate) {
        errVad = g_ts3Functions.setPreProcessorConfigValue(schid, "vad", desiredVadValue);
    }
    unsigned int errInput = ERROR_ok_no_update;
    if (needInputUpdate) {
        errInput = g_ts3Functions.setClientSelfVariableAsInt(
            schid, CLIENT_INPUT_DEACTIVATED, desiredInputState);
    }
    unsigned int errFlush = ERROR_ok_no_update;
    if (needFlush) {
        errFlush = g_ts3Functions.flushClientSelfUpdates(schid, nullptr);
    }
    const bool okVad = !needVadUpdate || (errVad == ERROR_ok || errVad == ERROR_ok_no_update);
    const bool okInput = !needInputUpdate || (errInput == ERROR_ok || errInput == ERROR_ok_no_update);
    const bool okFlush = !needFlush || (errFlush == ERROR_ok || errFlush == ERROR_ok_no_update);
    const bool ok = okVad && okInput && okFlush;
    if (!ok) {
        if (nowMs > (lastErrLogMs + 2000ULL)) {
            LogWarn("talk_mode force failed schid=" + std::to_string(schid) +
                    " vad_err=" + std::to_string(errVad) +
                    " input_err=" + std::to_string(errInput) +
                    " flush_err=" + std::to_string(errFlush) +
                    " need_vad=" + std::to_string(needVadUpdate ? 1 : 0) +
                    " need_input=" + std::to_string(needInputUpdate ? 1 : 0) +
                    " need_flush=" + std::to_string(needFlush ? 1 : 0));
            lastErrLogMs = nowMs;
        }
        clearState();
        return;
    }

    if (!sameSchid) {
        LogInfo("talk_mode forced continuous schid=" + std::to_string(schid) +
                " vad=" + desiredVadValue +
                " input_state=" + std::to_string(desiredInputState));
    } else if (needVadUpdate || needInputUpdate) {
        LogInfo("talk_mode refresh schid=" + std::to_string(schid) +
                " need_vad=" + std::to_string(needVadUpdate ? 1 : 0) +
                " need_input=" + std::to_string(needInputUpdate ? 1 : 0));
    }
    voiceRecordingActive_.store(true, std::memory_order_release);
    voiceRecordingSchid_.store(schid, std::memory_order_release);
    voiceTxLastNudgeMs_.store(nowMs, std::memory_order_release);
}

void MicMixApp::RefreshVoiceTxControl(uint64 schidHint) {
    uint64 schid = schidHint;
    if (schid == 0) {
        schid = activeSchid_.load(std::memory_order_acquire);
    }
    if (schid == 0 && g_ts3Functions.getCurrentServerConnectionHandlerID) {
        schid = g_ts3Functions.getCurrentServerConnectionHandlerID();
        if (schid != 0) {
            activeSchid_.store(schid, std::memory_order_release);
        }
    }
    if (!IsConnectedForTx(schid)) {
        SetVoiceRecordingState(false, 0);
        return;
    }

    const MicMixSettings s = GetSettings();
    const auto sourceManager = sourceManager_.load(std::memory_order_acquire);
    const SourceStatus st = sourceManager ? sourceManager->GetStatus() : SourceStatus{};
    const TelemetrySnapshot t = engine_.SnapshotTelemetry();
    const bool sourceUp = IsSourceStateActive(st.state);
    const uint64_t nowMs = GetTickCount64();
    const bool baseEligible = s.forceTxEnabled && sourceUp && !s.musicMuted;
    bool shouldKeepCaptureActive = false;
    if (baseEligible) {
        if (t.musicActive) {
            forceTxHoldUntilMs_.store(nowMs + kForceTxMusicWindowMs, std::memory_order_release);
            shouldKeepCaptureActive = true;
        } else {
            const uint64_t holdUntilMs = forceTxHoldUntilMs_.load(std::memory_order_acquire);
            shouldKeepCaptureActive = (holdUntilMs != 0ULL) && (nowMs <= holdUntilMs);
        }
    } else {
        forceTxHoldUntilMs_.store(0ULL, std::memory_order_release);
    }
    SetVoiceRecordingState(shouldKeepCaptureActive, schid);
}

void MicMixApp::ApplyMicInputTransportMute(bool muted) {
    if (!g_ts3Functions.setClientSelfVariableAsInt || !g_ts3Functions.flushClientSelfUpdates) {
        return;
    }

    // Transport-level input deactivation mutes the full outgoing capture path.
    // Only apply this while MicMix is active. In OFF state, always restore pass-through.
    bool shouldTransportMute = false;
    if (muted) {
        const MicMixSettings s = GetSettings();
        const auto sourceManager = sourceManager_.load(std::memory_order_acquire);
        const SourceStatus st = sourceManager ? sourceManager->GetStatus() : SourceStatus{};
        const bool sourceUp = IsSourceStateActive(st.state);
        const bool keepMusicPath = sourceUp && !s.musicMuted;
        shouldTransportMute = sourceUp && !keepMusicPath;
    }

    auto resolveSchid = [this]() -> uint64 {
        uint64 schid = activeSchid_.load(std::memory_order_acquire);
        if (schid == 0 && g_ts3Functions.getCurrentServerConnectionHandlerID) {
            schid = g_ts3Functions.getCurrentServerConnectionHandlerID();
            if (schid != 0) {
                activeSchid_.store(schid, std::memory_order_release);
            }
        }
        return schid;
    };

    std::lock_guard<std::mutex> lock(voiceTxMutex_);
    const uint64 activeSchid = resolveSchid();

    if (shouldTransportMute) {
        if (activeSchid == 0 || !IsConnectedForTx(activeSchid)) {
            return;
        }
        if (micInputTransportMuteActive_ && micInputTransportMuteSchid_ == activeSchid) {
            return;
        }

        int currentInputState = INPUT_DEACTIVATED;
        bool haveCurrentInputState = false;
        if (g_ts3Functions.getClientSelfVariableAsInt) {
            const unsigned int errGet = g_ts3Functions.getClientSelfVariableAsInt(
                activeSchid, CLIENT_INPUT_DEACTIVATED, &currentInputState);
            haveCurrentInputState = (errGet == ERROR_ok || errGet == ERROR_ok_no_update);
        }

        const unsigned int errSet = g_ts3Functions.setClientSelfVariableAsInt(
            activeSchid, CLIENT_INPUT_DEACTIVATED, INPUT_DEACTIVATED);
        const unsigned int errFlush = g_ts3Functions.flushClientSelfUpdates(activeSchid, nullptr);
        const bool okSet = (errSet == ERROR_ok || errSet == ERROR_ok_no_update);
        const bool okFlush = (errFlush == ERROR_ok || errFlush == ERROR_ok_no_update);
        if (okSet && okFlush) {
            micInputTransportMuteActive_ = true;
            micInputTransportMuteSchid_ = activeSchid;
            micInputTransportSavedValid_ = haveCurrentInputState;
            if (haveCurrentInputState) {
                micInputTransportSavedState_ = currentInputState;
            }
            LogInfo("mic_input transport muted schid=" + std::to_string(activeSchid));
        } else {
            LogWarn("mic_input transport mute failed schid=" + std::to_string(activeSchid) +
                    " set_err=" + std::to_string(errSet) +
                    " flush_err=" + std::to_string(errFlush));
        }
        return;
    }

    if (!micInputTransportMuteActive_) {
        return;
    }

    uint64 restoreSchid = micInputTransportMuteSchid_;
    if (restoreSchid == 0 || !IsConnectedForTx(restoreSchid)) {
        restoreSchid = activeSchid;
    }
    if (restoreSchid != 0 && IsConnectedForTx(restoreSchid)) {
        const int restoreState = micInputTransportSavedValid_ ? micInputTransportSavedState_ : INPUT_ACTIVE;
        const unsigned int errSet = g_ts3Functions.setClientSelfVariableAsInt(
            restoreSchid, CLIENT_INPUT_DEACTIVATED, restoreState);
        const unsigned int errFlush = g_ts3Functions.flushClientSelfUpdates(restoreSchid, nullptr);
        const bool okSet = (errSet == ERROR_ok || errSet == ERROR_ok_no_update);
        const bool okFlush = (errFlush == ERROR_ok || errFlush == ERROR_ok_no_update);
        if (!(okSet && okFlush)) {
            LogWarn("mic_input transport restore failed schid=" + std::to_string(restoreSchid) +
                    " set_err=" + std::to_string(errSet) +
                    " flush_err=" + std::to_string(errFlush));
        } else {
            LogInfo("mic_input transport restored schid=" + std::to_string(restoreSchid) +
                    " state=" + std::to_string(restoreState));
        }
    }

    micInputTransportMuteActive_ = false;
    micInputTransportMuteSchid_ = 0;
    micInputTransportSavedValid_ = false;
    micInputTransportSavedState_ = INPUT_DEACTIVATED;
}

void MicMixApp::SyncMusicActivityMeta(uint64 schid, bool musicActive, bool force) {
    // During plugin shutdown, avoid TS API calls from this path.
    // Late metadata updates can race client teardown in some TS3 builds.
    if (shutdownRequested_.load(std::memory_order_acquire)) {
        return;
    }
    if (!g_ts3Functions.setClientSelfVariableAsString || !g_ts3Functions.flushClientSelfUpdates) {
        return;
    }
    if (schid == 0 || !IsConnectedForTx(schid)) {
        return;
    }

    const uint64_t nowMs = GetTickCount64();
    std::lock_guard<std::mutex> lock(musicMetaMutex_);

    if (musicMetaSchid_ != schid) {
        musicMetaSchid_ = schid;
        musicMetaLastStateValid_ = false;
        musicMetaLastAttemptMs_ = 0;
    }

    if (!force && musicMetaLastStateValid_ && musicMetaLastState_ == musicActive) {
        return;
    }
    if (!force && nowMs <= (musicMetaLastAttemptMs_ + kMusicMetaUpdateMinIntervalMs)) {
        return;
    }

    std::string currentMeta;
    if (g_ts3Functions.getClientSelfVariableAsString) {
        char* currentMetaRaw = nullptr;
        const unsigned int errRead = g_ts3Functions.getClientSelfVariableAsString(
            schid, CLIENT_META_DATA, &currentMetaRaw);
        if ((errRead == ERROR_ok || errRead == ERROR_ok_no_update) && currentMetaRaw) {
            currentMeta.assign(currentMetaRaw);
        }
        if (currentMetaRaw && g_ts3Functions.freeMemory) {
            g_ts3Functions.freeMemory(currentMetaRaw);
        }
    }

    const std::string desiredMeta = BuildMusicMetaValue(currentMeta, musicActive);
    if (!force && desiredMeta == currentMeta) {
        musicMetaLastStateValid_ = true;
        musicMetaLastState_ = musicActive;
        return;
    }

    musicMetaLastAttemptMs_ = nowMs;
    const unsigned int errSet = g_ts3Functions.setClientSelfVariableAsString(
        schid, CLIENT_META_DATA, desiredMeta.c_str());
    const bool okSet = (errSet == ERROR_ok || errSet == ERROR_ok_no_update);
    unsigned int errFlush = ERROR_ok_no_update;
    if (okSet) {
        errFlush = g_ts3Functions.flushClientSelfUpdates(schid, nullptr);
    }
    const bool okFlush = (errFlush == ERROR_ok || errFlush == ERROR_ok_no_update);

    if (okSet && okFlush) {
        musicMetaLastStateValid_ = true;
        musicMetaLastState_ = musicActive;
        LogInfo(std::string("music_meta sync schid=") + std::to_string(schid) +
                " active=" + std::to_string(musicActive ? 1 : 0));
        return;
    }

    LogWarn(std::string("music_meta sync failed schid=") + std::to_string(schid) +
            " set_err=" + std::to_string(errSet) +
            " flush_err=" + std::to_string(errFlush));
}

void MicMixApp::VoiceTxThreadMain() {
    voiceTxThreadRunning_.store(true, std::memory_order_release);
    try {
        uint64_t lastLogMs = 0;
        uint64_t lastVadFallbackLogMs = 0;
        uint64_t vadGateHoldUntilMs = 0;
        uint64_t vadAboveThresholdSinceMs = 0;
        bool vadQualifiedOpen = false;
        uint64_t lastVstHostMaintainMs = 0;
        while (!voiceTxStop_.load(std::memory_order_acquire)) {
            if (shutdownRequested_.load(std::memory_order_acquire)) {
                break;
            }
            const uint64_t nowMs = GetTickCount64();
            if (nowMs > (lastVstHostMaintainMs + 200ULL)) {
                MaintainVstHost(nowMs);
                lastVstHostMaintainMs = nowMs;
            }
            uint64 schid = activeSchid_.load(std::memory_order_acquire);
            if (schid == 0 && g_ts3Functions.getCurrentServerConnectionHandlerID) {
                schid = g_ts3Functions.getCurrentServerConnectionHandlerID();
                if (schid != 0) {
                    activeSchid_.store(schid, std::memory_order_release);
                }
            }
            RefreshVoiceTxControl(schid);
            const bool connectedForTx = (schid != 0) && IsConnectedForTx(schid);
            if (!connectedForTx && schid != 0 &&
                activeSchid_.load(std::memory_order_acquire) == schid) {
                activeSchid_.store(0, std::memory_order_release);
            }
            if (connectedForTx) {
                const TelemetrySnapshot t = engine_.SnapshotTelemetry();
                SyncMusicActivityMeta(schid, t.musicActive, false);
            }

            if (connectedForTx && g_ts3Functions.getPreProcessorInfoValueFloat) {
                float micDb = -120.0f;
                const unsigned int errDb = g_ts3Functions.getPreProcessorInfoValueFloat(
                    schid, "decibel_last_period", &micDb);
                if (errDb == ERROR_ok || errDb == ERROR_ok_no_update) {
                    const float linear = std::pow(10.0f, std::clamp(micDb, -120.0f, 0.0f) / 20.0f);
                    engine_.SetExternalMicLevel(linear);
                }
            }

            if (g_ts3Functions.activateCaptureDevice) {
                if (schid == 0 && g_ts3Functions.getServerConnectionHandlerList) {
                    uint64* list = nullptr;
                    if (g_ts3Functions.getServerConnectionHandlerList(&list) == ERROR_ok && list) {
                        if (list[0] != 0) {
                            schid = list[0];
                            activeSchid_.store(schid, std::memory_order_release);
                        }
                        if (g_ts3Functions.freeMemory) {
                            g_ts3Functions.freeMemory(list);
                        }
                    }
                }

                if (schid != 0 && IsConnectedForTx(schid)) {
                    const bool shouldEnsureCapture = voiceRecordingActive_.load(std::memory_order_acquire) &&
                                                    (voiceRecordingSchid_.load(std::memory_order_acquire) == schid);
                    if (shouldEnsureCapture) {
                        bool gateByVad = false;
                        float vadThresholdDb = -50.0f;
                        int vadExtraBufferSize = 0;
                        bool keyboardGuardEnabled = true;
                        MicGateMode gateMode = MicGateMode::AutoTs;
                        {
                            const MicMixSettings runtimeSettings = GetSettings();
                            gateMode = runtimeSettings.micGateMode;
                            keyboardGuardEnabled = runtimeSettings.micUseKeyboardGuard;
                            if (gateMode == MicGateMode::Custom) {
                                vadThresholdDb = std::clamp(runtimeSettings.micGateThresholdDbfs, -90.0f, 0.0f);
                            }
                        }
                        {
                            std::lock_guard<std::mutex> lock(voiceTxMutex_);
                            gateByVad = (gateMode == MicGateMode::Custom) ? true : (savedVadValid_ && savedVadEnabled_);
                            if (gateMode != MicGateMode::Custom && savedVadThresholdValid_) {
                                vadThresholdDb = savedVadThresholdDbfs_;
                            }
                            if (savedVadExtraBufferValid_) {
                                vadExtraBufferSize = savedVadExtraBufferSize_;
                            }
                        }
                        bool talkOpen = true;
                        if (gateByVad) {
                            const uint64_t gateNowMs = GetTickCount64();
                            bool candidateOpen = false;
                            bool haveDecision = false;

                            const uint64_t talkEventMs = ownTalkStatusTickMs_.load(std::memory_order_acquire);
                            if (talkEventMs != 0 && gateNowMs <= (talkEventMs + kTalkEventFreshMs)) {
                                candidateOpen = ownTalkStatusActive_.load(std::memory_order_acquire);
                                haveDecision = true;
                            }

                            float micDb = -120.0f;
                            bool haveMicDb = false;
                            if (g_ts3Functions.getPreProcessorInfoValueFloat) {
                                const unsigned int errDb = g_ts3Functions.getPreProcessorInfoValueFloat(
                                    schid, "decibel_last_period", &micDb);
                                if (errDb == ERROR_ok || errDb == ERROR_ok_no_update) {
                                    haveMicDb = true;
                                }
                            }

                            if (gateMode == MicGateMode::AutoTs) {
                                // Strict TS auto mode: trust TS talk-status event for open/close
                                // instead of local dB heuristics.
                                if (talkEventMs != 0 && gateNowMs <= (talkEventMs + kTalkEventFreshMs)) {
                                    candidateOpen = ownTalkStatusActive_.load(std::memory_order_acquire);
                                    haveDecision = true;
                                } else {
                                    // Fail-open on stale talk events to avoid hard mic dropouts.
                                    candidateOpen = true;
                                    haveDecision = true;
                                    if (gateNowMs > (lastVadFallbackLogMs + 4000ULL)) {
                                        LogWarn("talk_gate auto_ts stale event schid=" + std::to_string(schid));
                                        lastVadFallbackLogMs = gateNowMs;
                                    }
                                }
                                vadAboveThresholdSinceMs = candidateOpen ? gateNowMs : 0;
                                vadQualifiedOpen = candidateOpen;
                            } else if (haveMicDb) {
                                // Custom mode: local hysteresis + qualification to tame transients.
                                const float openThresholdDb = vadThresholdDb + 1.0f;
                                const float closeThresholdDb = vadThresholdDb - 2.5f;
                                if (keyboardGuardEnabled) {
                                    if (candidateOpen) {
                                        vadAboveThresholdSinceMs = nowMs;
                                        vadQualifiedOpen = true;
                                    } else {
                                        if (vadQualifiedOpen) {
                                            candidateOpen = micDb >= closeThresholdDb;
                                        } else {
                                            if (micDb >= openThresholdDb) {
                                                if (vadAboveThresholdSinceMs == 0) {
                                                    vadAboveThresholdSinceMs = gateNowMs;
                                                }
                                                candidateOpen = gateNowMs >= (vadAboveThresholdSinceMs + 60ULL);
                                            } else {
                                                vadAboveThresholdSinceMs = 0;
                                                candidateOpen = false;
                                            }
                                        }
                                    }
                                } else {
                                    if (!candidateOpen) {
                                        candidateOpen = micDb >= vadThresholdDb;
                                    }
                                    vadAboveThresholdSinceMs = candidateOpen ? gateNowMs : 0;
                                    vadQualifiedOpen = candidateOpen;
                                }
                                haveDecision = true;
                            } else if (!haveDecision) {
                                const TelemetrySnapshot tel = engine_.SnapshotTelemetry();
                                if (tel.micRmsDbfs > -119.0f) {
                                    candidateOpen = tel.micRmsDbfs >= (vadThresholdDb - 1.0f);
                                } else {
                                    candidateOpen = tel.talkStateActive;
                                }
                                haveDecision = true;
                                if (gateNowMs > (lastVadFallbackLogMs + 4000ULL)) {
                                    LogWarn("talk_gate fallback used schid=" + std::to_string(schid));
                                    lastVadFallbackLogMs = gateNowMs;
                                }
                            }

                            if (!haveDecision) {
                                candidateOpen = true;
                            }

                            const uint64_t vadHoldMs = static_cast<uint64_t>(std::clamp(vadExtraBufferSize, 0, 50)) * 20ULL;
                            uint64_t effectiveHoldMs = vadHoldMs;
                            if (gateMode != MicGateMode::AutoTs) {
                                const uint64_t minHoldMs = keyboardGuardEnabled ? 160ULL : 40ULL;
                                effectiveHoldMs = std::max<uint64_t>(vadHoldMs, minHoldMs);
                            }
                            if (candidateOpen) {
                                vadGateHoldUntilMs = gateNowMs + effectiveHoldMs;
                                vadQualifiedOpen = true;
                                talkOpen = true;
                            } else if (effectiveHoldMs > 0 && gateNowMs <= vadGateHoldUntilMs) {
                                talkOpen = true;
                            } else {
                                talkOpen = false;
                                vadQualifiedOpen = false;
                            }
                        } else {
                            vadGateHoldUntilMs = 0;
                            vadAboveThresholdSinceMs = 0;
                            vadQualifiedOpen = false;
                        }
                        engine_.SetTalkState(talkOpen);

                        const uint64_t watchdogNowMs = GetTickCount64();
                        const uint64_t lastCaptureMs = lastCaptureEditTickMs_.load(std::memory_order_acquire);
                        const uint64_t lastNudgeMs = lastCaptureReopenTickMs_.load(std::memory_order_acquire);
                        if (watchdogNowMs > (lastCaptureMs + kCaptureWatchdogSilenceMs) &&
                            watchdogNowMs > (lastNudgeMs + kCaptureWatchdogCooldownMs)) {
                            const unsigned int err = g_ts3Functions.activateCaptureDevice(schid);
                            lastCaptureReopenTickMs_.store(watchdogNowMs, std::memory_order_release);
                            if (!(err == ERROR_ok || err == ERROR_ok_no_update)) {
                                if (watchdogNowMs > (lastLogMs + 2000ULL)) {
                                    LogWarn("capture_activate watchdog failed schid=" + std::to_string(schid) +
                                            " err=" + std::to_string(err));
                                    lastLogMs = watchdogNowMs;
                                }
                            } else if (watchdogNowMs > (lastLogMs + 3000ULL)) {
                                LogInfo("capture_activate watchdog schid=" + std::to_string(schid) +
                                        " err=" + std::to_string(err));
                                lastLogMs = watchdogNowMs;
                            }
                        }
                    } else {
                        engine_.SetTalkState(true);
                    }
                }
            }
            std::this_thread::sleep_for(std::chrono::milliseconds(kVoiceTxPollMs));
        }
    } catch (const std::exception& ex) {
        LogError(std::string("voice_tx thread exception: ") + ex.what());
    } catch (...) {
        LogError("voice_tx thread exception: unknown");
    }
    uint64 schid = voiceRecordingSchid_.load(std::memory_order_acquire);
    if (schid == 0) {
        schid = activeSchid_.load(std::memory_order_acquire);
    }
    if (schid != 0 &&
        !shutdownRequested_.load(std::memory_order_acquire) &&
        IsConnectedForTx(schid)) {
        SyncMusicActivityMeta(schid, false, true);
    }
    SetVoiceRecordingState(false, 0);
    voiceTxThreadRunning_.store(false, std::memory_order_release);
}

void MicMixApp::StartVoiceTxThread() {
    if (voiceTxThread_.joinable()) {
        return;
    }
    voiceTxStop_.store(false, std::memory_order_release);
    try {
        voiceTxThread_ = std::thread([this]() { VoiceTxThreadMain(); });
    } catch (const std::exception& ex) {
        voiceTxStop_.store(true, std::memory_order_release);
        LogError(std::string("voice_tx thread start failed: ") + ex.what());
    } catch (...) {
        voiceTxStop_.store(true, std::memory_order_release);
        LogError("voice_tx thread start failed: unknown");
    }
}

void MicMixApp::StopVoiceTxThread() {
    voiceTxStop_.store(true, std::memory_order_release);
    if (voiceTxThread_.joinable()) {
        voiceTxThread_.join();
    }
    SetVoiceRecordingState(false, 0);
}

bool MicMixApp::Initialize(const std::string& configBasePath) {
    if (initialized_.exchange(true)) return true;
    try {
        shutdownRequested_.store(false, std::memory_order_release);
        captureCallbacksInFlight_.store(0, std::memory_order_release);
        vstHostPipeName_ = GenerateVstHostPipeName();
        vstAudioShmName_ = GenerateVstAudioShmName();
        vstHostAuthToken_ = GenerateHexToken(16);
        {
            std::lock_guard<std::mutex> lock(musicMetaMutex_);
            musicMetaSchid_ = 0;
            musicMetaLastStateValid_ = false;
            musicMetaLastState_ = false;
            musicMetaLastAttemptMs_ = 0;
        }
        configStore_ = std::make_unique<ConfigStore>(configBasePath);
        g_logPath = configStore_->LogPath();
        std::string warn;
        configStore_->Load(settings_, warn);
        SanitizeSettings(settings_);
        if (!warn.empty()) LogWarn(warn);
        {
            std::lock_guard<std::mutex> lock(vstHostMutex_);
            vstHostMessage_ = "stopped";
        }
        vstEffectsEnabledCached_.store(settings_.vstEffectsEnabled, std::memory_order_release);
        vstHostNextRestartTickMs_.store(0, std::memory_order_release);
        vstHostRestartAttempts_.store(0, std::memory_order_release);
        vstHostLastRestartLogTickMs_.store(0, std::memory_order_release);
        vstHostLastHeartbeatTickMs_.store(0, std::memory_order_release);
        vstHostSyncPending_.store(false, std::memory_order_release);
        vstHostStopPending_.store(false, std::memory_order_release);
        engine_.ApplySettings(settings_);
        ApplyMicInputTransportMute(settings_.micInputMuted);
        // Disabled intentionally: an additional capture stream can interfere with
        // some driver/device combinations and TS microphone preview.
        // Mic activity is derived from TS3 capture callback path.
        micLevelMonitor_.reset();
        musicMuteHotkeyManager_ = std::make_unique<GlobalHotkeyManager>(
            [this]() { this->ToggleMute(); },
            [](const MicMixSettings& s, UINT& mods, UINT& vk) {
                mods = static_cast<UINT>(std::max(0, s.muteHotkeyModifiers));
                vk = static_cast<UINT>(std::max(0, s.muteHotkeyVk));
            },
            0x4D4D,
            "music_mute_hotkey");
        micInputMuteHotkeyManager_ = std::make_unique<GlobalHotkeyManager>(
            [this]() { this->ToggleMicInputMute(); },
            [](const MicMixSettings& s, UINT& mods, UINT& vk) {
                mods = static_cast<UINT>(std::max(0, s.micInputMuteHotkeyModifiers));
                vk = static_cast<UINT>(std::max(0, s.micInputMuteHotkeyVk));
            },
            0x4D4E,
            "mic_input_hotkey");
        musicMuteHotkeyManager_->Start();
        micInputMuteHotkeyManager_->Start();
        musicMuteHotkeyManager_->ApplySettings(settings_);
        micInputMuteHotkeyManager_->ApplySettings(settings_);
        auto sourceManager = std::make_shared<AudioSourceManager>(
            [this](const float* data, size_t count) { this->OnSourceSamples(data, count); },
            [this](SourceState st, const std::string& code, const std::string& msg, const std::string& detail) {
                engine_.SetMusicSourceRunning(st == SourceState::Running);
                if (st == SourceState::Reacquiring) {
                    engine_.NoteReconnect();
                }
                bool micInputMuted = false;
                {
                    std::lock_guard<std::mutex> lock(settingsMutex_);
                    micInputMuted = settings_.micInputMuted;
                }
                ApplyMicInputTransportMute(micInputMuted);
                std::string line = "source=" + SourceStateToString(st) + " code=" + code + " msg=" + msg;
                if (!detail.empty()) line += " detail=" + detail;
                LogInfo(line);
            });
        sourceManager->ApplySettings(settings_);
        sourceManager_.store(sourceManager, std::memory_order_release);
        auto mixMonitor = std::make_shared<MixMonitorPlayer>();
        mixMonitorPlayer_.store(mixMonitor, std::memory_order_release);
        if (settings_.autostartEnabled) sourceManager->Start();
        const int activeResamplerQuality = ResolveResamplerQualitySetting(settings_.resamplerQuality);
        const std::string mode = (settings_.resamplerQuality < 0) ? "auto_cpu" : "manual";
        LogInfo("resampler " + mode + " quality=" + std::to_string(activeResamplerQuality) +
                " logical_cpus=" + std::to_string(GetLogicalCpuCount()));
        const uint64_t nowMs = GetTickCount64();
        lastCaptureEditTickMs_.store(nowMs, std::memory_order_release);
        lastCaptureReopenTickMs_.store(0, std::memory_order_release);
        if (settings_.vstHostAutostart || settings_.vstEffectsEnabled) {
            std::string hostError;
            if (!StartVstHostProcess(hostError)) {
                LogWarn("vst_host startup warning: " + hostError);
            } else {
                std::string syncError;
                if (!SyncVstHostState(settings_, syncError)) {
                    LogWarn("vst_host initial sync warning: " + syncError);
                }
            }
        }
        StartVoiceTxThread();
        LogInfo("MicMix initialized");
        return true;
    } catch (const std::exception& ex) {
        LogError(std::string("MicMix initialize failed: ") + ex.what());
    } catch (...) {
        LogError("MicMix initialize failed: unknown");
    }

    shutdownRequested_.store(true, std::memory_order_release);
    StopVoiceTxThread();
    if (musicMuteHotkeyManager_) musicMuteHotkeyManager_->Stop();
    if (micInputMuteHotkeyManager_) micInputMuteHotkeyManager_->Stop();
    if (const auto sourceManager = sourceManager_.load(std::memory_order_acquire)) {
        sourceManager->Stop();
    }
    if (const auto mixMonitor = mixMonitorPlayer_.load(std::memory_order_acquire)) {
        mixMonitor->SetEnabled(false);
    }
    micLevelMonitor_.reset();
    musicMuteHotkeyManager_.reset();
    micInputMuteHotkeyManager_.reset();
    sourceManager_.store({}, std::memory_order_release);
    mixMonitorPlayer_.store({}, std::memory_order_release);
    StopVstHostProcess();
    CloseVstAudioIpc();
    vstHostNextRestartTickMs_.store(0, std::memory_order_release);
    vstHostRestartAttempts_.store(0, std::memory_order_release);
    vstHostLastRestartLogTickMs_.store(0, std::memory_order_release);
    vstHostLastHeartbeatTickMs_.store(0, std::memory_order_release);
    vstHostSyncPending_.store(false, std::memory_order_release);
    vstHostStopPending_.store(false, std::memory_order_release);
    if (vstHostJob_) {
        CloseHandle(vstHostJob_);
        vstHostJob_ = nullptr;
    }
    vstEffectsEnabledCached_.store(false, std::memory_order_release);
    vstHostPipeName_.clear();
    vstAudioShmName_.clear();
    vstHostAuthToken_.clear();
    configStore_.reset();
    initialized_.store(false, std::memory_order_release);
    return false;
}

void MicMixApp::Shutdown() {
    if (!initialized_.exchange(false)) return;
    {
        std::lock_guard<std::mutex> lock(musicMetaMutex_);
        musicMetaLastStateValid_ = false;
        musicMetaLastAttemptMs_ = 0;
    }
    SettingsWindowController::Instance().Close();
    EffectsWindowController::Instance().Close();
    if (musicMuteHotkeyManager_) musicMuteHotkeyManager_->Stop();
    if (micInputMuteHotkeyManager_) micInputMuteHotkeyManager_->Stop();
    // Keep shutdownRequested_ clear until voice-tx cleanup is done so we can
    // restore TeamSpeak input/VAD state before plugin teardown is finalized.
    StopVoiceTxThread();
    ApplyMicInputTransportMute(false);
    shutdownRequested_.store(true, std::memory_order_release);
    if (!WaitForCaptureCallbacksToDrain(1500U)) {
        LogWarn("capture callbacks still active during shutdown");
    }
    if (const auto sourceManager = sourceManager_.load(std::memory_order_acquire)) {
        sourceManager->Stop();
    }
    if (const auto mixMonitor = mixMonitorPlayer_.load(std::memory_order_acquire)) {
        mixMonitor->SetEnabled(false);
    }
    if (configStore_) {
        std::string err;
        configStore_->Save(settings_, err);
        if (!err.empty()) LogWarn("Config save warning: " + err);
    }
    musicMuteHotkeyManager_.reset();
    micInputMuteHotkeyManager_.reset();
    sourceManager_.store({}, std::memory_order_release);
    mixMonitorPlayer_.store({}, std::memory_order_release);
    StopVstHostProcess();
    CloseVstAudioIpc();
    vstHostNextRestartTickMs_.store(0, std::memory_order_release);
    vstHostRestartAttempts_.store(0, std::memory_order_release);
    vstHostLastRestartLogTickMs_.store(0, std::memory_order_release);
    vstHostLastHeartbeatTickMs_.store(0, std::memory_order_release);
    vstHostSyncPending_.store(false, std::memory_order_release);
    vstHostStopPending_.store(false, std::memory_order_release);
    if (vstHostJob_) {
        CloseHandle(vstHostJob_);
        vstHostJob_ = nullptr;
    }
    vstEffectsEnabledCached_.store(false, std::memory_order_release);
    vstHostPipeName_.clear();
    vstAudioShmName_.clear();
    vstHostAuthToken_.clear();
    configStore_.reset();
}

MicMixSettings MicMixApp::GetSettings() const {
    std::lock_guard<std::mutex> lock(settingsMutex_);
    return settings_;
}

void MicMixApp::ApplySettings(const MicMixSettings& settings, bool restartSource, bool saveConfig) {
    MicMixSettings safe = settings;
    SanitizeSettings(safe);
    bool changed = false;
    bool micInputMuteChanged = false;
    bool effectsEnabledChanged = false;
    bool hostAutostartChanged = false;
    bool micInputMuted = false;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        changed = !IsSameSettings(settings_, safe);
        micInputMuteChanged = (settings_.micInputMuted != safe.micInputMuted);
        effectsEnabledChanged = (settings_.vstEffectsEnabled != safe.vstEffectsEnabled);
        hostAutostartChanged = (settings_.vstHostAutostart != safe.vstHostAutostart);
        settings_ = safe;
        UpdateAllSlotStatusesForUi(settings_, false);
        safe = settings_;
        micInputMuted = settings_.micInputMuted;
    }
    vstEffectsEnabledCached_.store(safe.vstEffectsEnabled, std::memory_order_release);
    if (!changed && !restartSource) {
        return;
    }

    engine_.ApplySettings(safe);
    if (micInputMuteChanged) {
        ApplyMicInputTransportMute(micInputMuted);
    }
    if (effectsEnabledChanged || hostAutostartChanged) {
        if (safe.vstEffectsEnabled || safe.vstHostAutostart) {
            vstHostStopPending_.store(false, std::memory_order_release);
            vstHostSyncPending_.store(true, std::memory_order_release);
            std::lock_guard<std::mutex> lock(vstHostMutex_);
            vstHostMessage_ = "sync_pending";
        } else {
            vstHostSyncPending_.store(false, std::memory_order_release);
            vstHostStopPending_.store(true, std::memory_order_release);
            std::lock_guard<std::mutex> lock(vstHostMutex_);
            vstHostMessage_ = "stop_pending";
        }
    } else if (safe.vstEffectsEnabled || safe.vstHostAutostart) {
        vstHostStopPending_.store(false, std::memory_order_release);
        vstHostSyncPending_.store(true, std::memory_order_release);
    }
    if (musicMuteHotkeyManager_) {
        musicMuteHotkeyManager_->ApplySettings(safe);
    }
    if (micInputMuteHotkeyManager_) {
        micInputMuteHotkeyManager_->ApplySettings(safe);
    }
    if (const auto sourceManager = sourceManager_.load(std::memory_order_acquire)) {
        sourceManager->ApplySettings(safe);
        if (restartSource && sourceManager->IsRunning()) sourceManager->Restart();
    }
    if (saveConfig && changed && configStore_) {
        std::string err;
        if (!configStore_->Save(safe, err) && !err.empty()) LogError("Config save failed: " + err);
    }
}

void MicMixApp::StartSource() {
    if (const auto sourceManager = sourceManager_.load(std::memory_order_acquire)) {
        sourceManager->Start();
    }
}
void MicMixApp::StopSource() {
    if (const auto sourceManager = sourceManager_.load(std::memory_order_acquire)) {
        sourceManager->Stop();
    }
}

void MicMixApp::ToggleMute() {
    const auto sourceManager = sourceManager_.load(std::memory_order_acquire);
    const SourceStatus st = sourceManager ? sourceManager->GetStatus() : SourceStatus{};
    if (!IsSourceStateActive(st.state)) {
        LogInfo("music mute toggle ignored: micmix inactive");
        return;
    }
    engine_.ToggleMute();
    const bool newMuted = engine_.IsMuted();
    bool changed = false;
    MicMixSettings snapshot;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        if (settings_.musicMuted != newMuted) {
            settings_.musicMuted = newMuted;
            changed = true;
        }
        snapshot = settings_;
    }
    if (changed && configStore_) {
        std::string err;
        if (!configStore_->Save(snapshot, err) && !err.empty()) {
            LogError("Config save failed: " + err);
        }
    }
    LogInfo(std::string("music ") + (newMuted ? "muted" : "unmuted"));
}

void MicMixApp::ToggleMicInputMute() {
    const auto sourceManager = sourceManager_.load(std::memory_order_acquire);
    const SourceStatus st = sourceManager ? sourceManager->GetStatus() : SourceStatus{};
    if (!IsSourceStateActive(st.state)) {
        ApplyMicInputTransportMute(false);
        LogInfo("mic_input mute toggle ignored: micmix inactive");
        return;
    }
    engine_.ToggleMicInputMute();
    const bool newMuted = engine_.IsMicInputMuted();
    bool changed = false;
    MicMixSettings snapshot;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        if (settings_.micInputMuted != newMuted) {
            settings_.micInputMuted = newMuted;
            changed = true;
        }
        snapshot = settings_;
    }
    ApplyMicInputTransportMute(newMuted);
    if (changed && configStore_) {
        std::string err;
        if (!configStore_->Save(snapshot, err) && !err.empty()) {
            LogError("Config save failed: " + err);
        }
    }
    LogInfo(std::string("mic_input ") + (newMuted ? "muted" : "unmuted"));
}

void MicMixApp::SetMonitorEnabled(bool enabled) {
    if (const auto mixMonitor = mixMonitorPlayer_.load(std::memory_order_acquire)) {
        mixMonitor->SetEnabled(enabled);
        LogInfo(std::string("mix_monitor ") + (mixMonitor->IsEnabled() ? "enabled" : "disabled"));
    }
}

bool MicMixApp::IsMonitorEnabled() const {
    if (const auto mixMonitor = mixMonitorPlayer_.load(std::memory_order_acquire)) {
        return mixMonitor->IsEnabled();
    }
    return false;
}

void MicMixApp::ToggleMonitor() {
    SetMonitorEnabled(!IsMonitorEnabled());
}

void MicMixApp::OnSourceSamples(const float* data, size_t count) {
    if (!data || count == 0) {
        return;
    }
    thread_local PendingVstPackets pendingMusicPackets;
    thread_local uint32_t lastMusicSeq = 0;
    thread_local float musicWetMix = 1.0f;
    thread_local std::array<float, micmix::vstipc::kMaxFramesPerPacket> blended{};
    constexpr float kVstMixTransitionMs = 4.0f;
    constexpr float kVstMixTransitionSamples = (static_cast<float>(kTargetRate) * kVstMixTransitionMs) / 1000.0f;
    constexpr float kVstMixStep = 1.0f / kVstMixTransitionSamples;
    const bool effectsEnabled = vstEffectsEnabledCached_.load(std::memory_order_acquire);
    auto* shm = vstAudioShared_.load(std::memory_order_acquire);
    if (!effectsEnabled ||
        !shm ||
        !vstHostRunning_.load(std::memory_order_acquire)) {
        pendingMusicPackets.Clear();
        musicWetMix = 0.0f;
        engine_.PushMusicSamples(data, count);
        return;
    }
    const LONG hostHeartbeat = micmix::vstipc::AtomicLoad(&shm->hostHeartbeat);
    if (hostHeartbeat == 0 ||
        (static_cast<uint32_t>(GetTickCount()) - static_cast<uint32_t>(hostHeartbeat)) > 3000U) {
        pendingMusicPackets.Clear();
        musicWetMix = 0.0f;
        engine_.PushMusicSamples(data, count);
        return;
    }

    InterlockedExchange(&shm->pluginHeartbeat, static_cast<LONG>(GetTickCount64() & 0x7FFFFFFF));
    size_t offset = 0;
    while (offset < count) {
        const uint32_t frames = static_cast<uint32_t>(std::min<size_t>(
            static_cast<size_t>(micmix::vstipc::kMaxFramesPerPacket), count - offset));
        micmix::vstipc::AudioPacket packet{};
        packet.seq = vstMusicSeq_.fetch_add(1U, std::memory_order_relaxed);
        if (packet.seq < lastMusicSeq) {
            pendingMusicPackets.Clear();
        }
        lastMusicSeq = packet.seq;
        packet.frames = frames;
        packet.channels = 1;
        std::memcpy(packet.samples, data + offset, sizeof(float) * frames);

        bool processed = false;
        micmix::vstipc::AudioPacket out{};
        if (micmix::vstipc::RingPush(shm->musicIn, packet)) {
            processed = TryReadMatchingVstPacket(
                shm->musicOut,
                packet.seq,
                pendingMusicPackets,
                out,
                ComputeVstOutputWaitUs(frames));
        }

        const bool haveWet = processed && out.frames == frames;
        const float targetMix = haveWet ? 1.0f : 0.0f;
        for (uint32_t i = 0; i < frames; ++i) {
            if (targetMix > musicWetMix) {
                musicWetMix = std::min(targetMix, musicWetMix + kVstMixStep);
            } else if (targetMix < musicWetMix) {
                musicWetMix = std::max(targetMix, musicWetMix - kVstMixStep);
            }
            const float dry = data[offset + i];
            const float wet = haveWet ? (std::isfinite(out.samples[i]) ? out.samples[i] : dry) : dry;
            blended[i] = dry + ((wet - dry) * musicWetMix);
        }
        engine_.PushMusicSamples(blended.data(), frames);
        offset += frames;
    }
}

void MicMixApp::ProcessMicInputWithVst(short* samples, int sampleCount, int channels) {
    if (!samples || sampleCount <= 0 || channels <= 0) {
        return;
    }
    thread_local PendingVstPackets pendingMicPackets;
    thread_local uint32_t lastMicSeq = 0;
    thread_local float micWetMix = 1.0f;
    static std::atomic_flag micVstIpcBusy = ATOMIC_FLAG_INIT;
    bool acquired = !micVstIpcBusy.test_and_set(std::memory_order_acquire);
    if (!acquired) {
        const auto deadline = std::chrono::steady_clock::now() + std::chrono::microseconds(kVstMicIpcAcquireSpinUs);
        while (std::chrono::steady_clock::now() < deadline) {
            SwitchToThread();
            if (!micVstIpcBusy.test_and_set(std::memory_order_acquire)) {
                acquired = true;
                break;
            }
        }
    }
    if (!acquired) {
        pendingMicPackets.Clear();
        micWetMix = 0.0f;
        return;
    }
    struct MicIpcBusyGuard {
        std::atomic_flag& flag;
        ~MicIpcBusyGuard() {
            flag.clear(std::memory_order_release);
        }
    } busyGuard{micVstIpcBusy};
    constexpr float kVstMixTransitionMs = 4.0f;
    constexpr float kVstMixTransitionSamples = (static_cast<float>(kTargetRate) * kVstMixTransitionMs) / 1000.0f;
    constexpr float kVstMixStep = 1.0f / kVstMixTransitionSamples;
    const bool effectsEnabled = vstEffectsEnabledCached_.load(std::memory_order_acquire);
    auto* shm = vstAudioShared_.load(std::memory_order_acquire);
    if (!effectsEnabled ||
        !shm ||
        !vstHostRunning_.load(std::memory_order_acquire)) {
        pendingMicPackets.Clear();
        micWetMix = 0.0f;
        return;
    }
    const LONG hostHeartbeat = micmix::vstipc::AtomicLoad(&shm->hostHeartbeat);
    if (hostHeartbeat == 0 ||
        (static_cast<uint32_t>(GetTickCount()) - static_cast<uint32_t>(hostHeartbeat)) > 3000U) {
        pendingMicPackets.Clear();
        micWetMix = 0.0f;
        return;
    }

    InterlockedExchange(&shm->pluginHeartbeat, static_cast<LONG>(GetTickCount64() & 0x7FFFFFFF));
    thread_local std::array<float, micmix::vstipc::kMaxFramesPerPacket> mono{};

    int offset = 0;
    while (offset < sampleCount) {
        const uint32_t frames = static_cast<uint32_t>(std::min<int>(
            static_cast<int>(micmix::vstipc::kMaxFramesPerPacket), sampleCount - offset));
        for (uint32_t i = 0; i < frames; ++i) {
            float acc = 0.0f;
            for (int ch = 0; ch < channels; ++ch) {
                const int idx = ((offset + static_cast<int>(i)) * channels) + ch;
                acc += static_cast<float>(samples[idx]) / 32768.0f;
            }
            mono[i] = acc / static_cast<float>(channels);
        }

        micmix::vstipc::AudioPacket packet{};
        packet.seq = vstMicSeq_.fetch_add(1U, std::memory_order_relaxed);
        if (packet.seq < lastMicSeq) {
            pendingMicPackets.Clear();
        }
        lastMicSeq = packet.seq;
        packet.frames = frames;
        packet.channels = 1;
        std::memcpy(packet.samples, mono.data(), sizeof(float) * frames);

        bool processed = false;
        micmix::vstipc::AudioPacket out{};
        if (micmix::vstipc::RingPush(shm->micIn, packet)) {
            const uint64_t waitUs = std::clamp<uint64_t>(
                ComputeVstOutputWaitUs(frames),
                kVstMicOutputWaitMinUs,
                kVstMicOutputWaitMaxUs);
            processed = TryReadMatchingVstPacket(
                shm->micOut,
                packet.seq,
                pendingMicPackets,
                out,
                waitUs);
        }

        const bool haveWet = processed && out.frames == frames;
        const float targetMix = haveWet ? 1.0f : 0.0f;
        for (uint32_t i = 0; i < frames; ++i) {
            if (targetMix > micWetMix) {
                micWetMix = std::min(targetMix, micWetMix + kVstMixStep);
            } else if (targetMix < micWetMix) {
                micWetMix = std::max(targetMix, micWetMix - kVstMixStep);
            }
            const float dry = mono[i];
            const float wet = haveWet ? (std::isfinite(out.samples[i]) ? out.samples[i] : dry) : dry;
            const float mixed = dry + ((wet - dry) * micWetMix);
            const short s = static_cast<short>(std::lrintf(
                std::clamp(mixed, -1.0f, 1.0f) * 32767.0f));
            for (int ch = 0; ch < channels; ++ch) {
                const int idx = ((offset + static_cast<int>(i)) * channels) + ch;
                samples[idx] = s;
            }
        }
        offset += static_cast<int>(frames);
    }
}

void MicMixApp::SetGlobalHotkeyCaptureBlocked(bool blocked) {
    if (musicMuteHotkeyManager_) {
        musicMuteHotkeyManager_->SetCaptureBlocked(blocked);
    }
    if (micInputMuteHotkeyManager_) {
        micInputMuteHotkeyManager_->SetCaptureBlocked(blocked);
    }
}

void MicMixApp::SetTalkStateForOwnClient(uint64 schid, anyID clientId, int talkStatus) {
    if (!initialized_.load(std::memory_order_acquire) || schid == 0 || clientId == 0) {
        return;
    }

    anyID ownClientId = 0;
    if (g_ts3Functions.getClientID) {
        const unsigned int err = g_ts3Functions.getClientID(schid, &ownClientId);
        if (!(err == ERROR_ok || err == ERROR_ok_no_update)) {
            return;
        }
        if (ownClientId != 0 && ownClientId != clientId) {
            return;
        }
    }

    const bool talking = (talkStatus == STATUS_TALKING || talkStatus == STATUS_TALKING_WHILE_DISABLED);
    ownTalkStatusActive_.store(talking, std::memory_order_release);
    ownTalkStatusTickMs_.store(GetTickCount64(), std::memory_order_release);
}

void MicMixApp::SetActiveServer(uint64 schid) {
    if (!initialized_.load(std::memory_order_acquire) ||
        shutdownRequested_.load(std::memory_order_acquire)) {
        return;
    }
    activeSchid_.store(schid, std::memory_order_release);
    RefreshVoiceTxControl(schid);
}

void MicMixApp::OnConnectStatusChange(uint64 schid, int newStatus, unsigned int errorNumber) {
    if (!initialized_.load(std::memory_order_acquire) ||
        shutdownRequested_.load(std::memory_order_acquire)) {
        return;
    }
    (void)errorNumber;
    if (newStatus == STATUS_DISCONNECTED) {
        {
            std::lock_guard<std::mutex> lock(voiceTxMutex_);
            if (micInputTransportMuteActive_ && micInputTransportMuteSchid_ == schid) {
                micInputTransportMuteActive_ = false;
                micInputTransportMuteSchid_ = 0;
                micInputTransportSavedValid_ = false;
                micInputTransportSavedState_ = INPUT_DEACTIVATED;
            }
        }
        {
            std::lock_guard<std::mutex> lock(musicMetaMutex_);
            if (musicMetaSchid_ == schid) {
                musicMetaLastStateValid_ = false;
                musicMetaLastAttemptMs_ = 0;
            }
        }
        if (activeSchid_.load(std::memory_order_acquire) == schid) {
            activeSchid_.store(0, std::memory_order_release);
        }
        SetVoiceRecordingState(false, 0);
        return;
    }
    if (newStatus == STATUS_CONNECTION_ESTABLISHED || newStatus == STATUS_CONNECTION_ESTABLISHING || newStatus == STATUS_CONNECTED) {
        {
            std::lock_guard<std::mutex> lock(musicMetaMutex_);
            if (musicMetaSchid_ == schid) {
                musicMetaLastStateValid_ = false;
                musicMetaLastAttemptMs_ = 0;
            }
        }
        activeSchid_.store(schid, std::memory_order_release);
        RefreshVoiceTxControl(schid);
    }
}
std::vector<LoopbackDeviceInfo> MicMixApp::GetLoopbackDevices() const { return AudioSourceManager::EnumerateLoopbackDevices(); }
std::vector<CaptureDeviceInfo> MicMixApp::GetCaptureDevices() const { return AudioSourceManager::EnumerateCaptureDevices(); }
std::vector<AppProcessInfo> MicMixApp::GetAppProcesses() const {
    return AudioSourceManager::EnumerateAppProcesses();
}
SourceStatus MicMixApp::GetSourceStatus() const {
    if (const auto sourceManager = sourceManager_.load(std::memory_order_acquire)) {
        return sourceManager->GetStatus();
    }
    return {};
}
TelemetrySnapshot MicMixApp::GetTelemetry() const {
    TelemetrySnapshot t = engine_.SnapshotTelemetry();
    bool micInputMuted = false;
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        micInputMuted = settings_.micInputMuted;
    }
    if (micInputMuted) {
        t.micTalkDetected = false;
        t.micRmsDbfs = -120.0f;
        t.micPeakDbfs = -120.0f;
        t.micClipRecent = false;
        return t;
    }
    if (t.micRmsDbfs > -119.0f || !g_ts3Functions.getPreProcessorInfoValueFloat) {
        return t;
    }

    uint64 schid = activeSchid_.load(std::memory_order_acquire);
    if (schid == 0 && g_ts3Functions.getCurrentServerConnectionHandlerID) {
        schid = g_ts3Functions.getCurrentServerConnectionHandlerID();
    }
    if (schid == 0) {
        return t;
    }
    float micDb = -120.0f;
    const unsigned int errDb = g_ts3Functions.getPreProcessorInfoValueFloat(
        schid, "decibel_last_period", &micDb);
    if (errDb == ERROR_ok || errDb == ERROR_ok_no_update) {
        t.micRmsDbfs = std::clamp(micDb, -120.0f, 0.0f);
    }
    return t;
}
std::string MicMixApp::GetConfigDir() const { return configStore_ ? configStore_->ConfigDir() : std::string{}; }
std::string MicMixApp::GetPreferredTsCaptureDeviceName() const {
    uint64 schid = activeSchid_.load(std::memory_order_acquire);
    if (schid == 0 && g_ts3Functions.getServerConnectionHandlerList) {
        uint64* list = nullptr;
        if (g_ts3Functions.getServerConnectionHandlerList(&list) == ERROR_ok && list) {
            if (list[0] != 0) {
                schid = list[0];
            }
            if (g_ts3Functions.freeMemory) {
                g_ts3Functions.freeMemory(list);
            }
        }
    }
    if (schid == 0 || !g_ts3Functions.getCurrentCaptureDeviceName) {
        return {};
    }
    char* captureName = nullptr;
    int isDefault = 0;
    if (g_ts3Functions.getCurrentCaptureDeviceName(schid, &captureName, &isDefault) != ERROR_ok || !captureName) {
        return {};
    }
    std::string out = captureName;
    if (g_ts3Functions.freeMemory) {
        g_ts3Functions.freeMemory(captureName);
    }
    return out;
}

void MicMixApp::EditCapturedVoice(uint64 schid, short* samples, int sampleCount, int channels, int* edited) {
    if (!initialized_.load(std::memory_order_acquire) ||
        shutdownRequested_.load(std::memory_order_acquire)) {
        return;
    }
    if (!TryEnterCaptureCallback()) {
        return;
    }
    struct LeaveGuard {
        MicMixApp* app;
        ~LeaveGuard() {
            if (app) {
                app->LeaveCaptureCallback();
            }
        }
    } leaveGuard{this};
    const uint64_t nowMs = GetTickCount64();
    lastCaptureEditTickMs_.store(nowMs, std::memory_order_release);
    (void)schid;
    ProcessMicInputWithVst(samples, sampleCount, channels);
    engine_.EditCapturedVoice(samples, sampleCount, channels, edited);
    if (const auto mixMonitor = mixMonitorPlayer_.load(std::memory_order_acquire)) {
        mixMonitor->PushCaptured(samples, sampleCount, channels);
    }
}

void MicMixApp::OpenSettingsWindow() {
    SettingsWindowController::Instance().Open();
}

void MicMixApp::OpenEffectsWindow() {
    EffectsWindowController::Instance().Open();
}

std::string SourceStateToString(SourceState state) {
    switch (state) {
    case SourceState::Starting: return "Starting";
    case SourceState::Running: return "Running";
    case SourceState::Reacquiring: return "Reacquiring";
    case SourceState::Error: return "Error";
    default: return "Stopped";
    }
}
