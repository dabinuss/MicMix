#include "micmix_core.h"

#include <audioclient.h>
#include <audioclientactivationparams.h>
#include <audiopolicy.h>
#include <endpointvolume.h>
#include <functiondiscoverykeys_devpkey.h>
#include <Shlwapi.h>
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

using Microsoft::WRL::ComPtr;

TS3Functions g_ts3Functions = {};
std::string  g_pluginId;

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
constexpr uint32_t kMinSupportedSourceRate = 8000;
constexpr uint32_t kMaxSupportedSourceRate = 384000;

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

std::mutex g_logMutex;
std::string g_logPath;
std::string g_lastLogPayload;
const char* g_lastLogLevel = nullptr;
uint32_t g_suppressedLogCount = 0;
uint32_t g_logWriteCount = 0;

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
}

bool NearlyEqual(float a, float b) {
    return std::fabs(a - b) < 0.0001f;
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
           a.uiLastOpenTab == b.uiLastOpenTab &&
           a.autoSwitchToLoopback == b.autoSwitchToLoopback &&
           a.muteHotkeyModifiers == b.muteHotkeyModifiers &&
           a.muteHotkeyVk == b.muteHotkeyVk &&
           a.captureDeviceId == b.captureDeviceId &&
           a.micGateMode == b.micGateMode &&
           NearlyEqual(a.micGateThresholdDbfs, b.micGateThresholdDbfs) &&
           a.micUseSmoothGate == b.micUseSmoothGate &&
           a.micUseKeyboardGuard == b.micUseKeyboardGuard &&
           a.micForceTsFilters == b.micForceTsFilters;
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
    if (g_ts3Functions.logMessage) {
        g_ts3Functions.logMessage(text.c_str(), LogLevel_INFO, kLogChannel, schid);
    }
    AppendLogLine("INFO", text);
}

void LogWarn(const std::string& text, uint64 schid) {
    if (g_ts3Functions.logMessage) {
        g_ts3Functions.logMessage(text.c_str(), LogLevel_WARNING, kLogChannel, schid);
    }
    AppendLogLine("WARN", text);
}

void LogError(const std::string& text, uint64 schid) {
    if (g_ts3Functions.logMessage) {
        g_ts3Functions.logMessage(text.c_str(), LogLevel_ERROR, kLogChannel, schid);
    }
    AppendLogLine("ERROR", text);
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

    std::ifstream in(std::filesystem::path(Utf8ToWide(configPath_)));
    bool usedLegacyConfig = false;
    if (!in.is_open()) {
        in.clear();
        in.open(std::filesystem::path(Utf8ToWide(legacyConfigPath_)));
        usedLegacyConfig = in.is_open();
        if (!usedLegacyConfig) {
            return true;
        }
        appendWarning("Loaded legacy config.json file; settings will migrate to config.ini on next save.");
    }

    std::unordered_map<std::string, std::string> kv;
    std::string line;
    while (std::getline(in, line)) {
        line = Trim(line);
        if (line.empty() || line[0] == '#') {
            continue;
        }
        const size_t eq = line.find('=');
        if (eq == std::string::npos) {
            continue;
        }
        kv[Trim(line.substr(0, eq))] = Trim(line.substr(eq + 1));
    }

    bool parseIssue = false;
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
    parseInt("hotkey.mute.modifiers", outSettings.muteHotkeyModifiers);
    parseInt("hotkey.mute.vk", outSettings.muteHotkeyVk);
    if (auto it = kv.find("capture.device_id"); it != kv.end()) outSettings.captureDeviceId = it->second;
    if (auto it = kv.find("mic.gate.mode"); it != kv.end()) outSettings.micGateMode = MicGateModeFromString(it->second);
    parseFloat("mic.gate.threshold_dbfs", outSettings.micGateThresholdDbfs);
    if (auto it = kv.find("mic.gate.smooth"); it != kv.end()) outSettings.micUseSmoothGate = ::ParseBool(it->second, outSettings.micUseSmoothGate);
    if (auto it = kv.find("mic.gate.keyboard_guard"); it != kv.end()) outSettings.micUseKeyboardGuard = ::ParseBool(it->second, outSettings.micUseKeyboardGuard);
    if (auto it = kv.find("mic.force_ts_filters"); it != kv.end()) outSettings.micForceTsFilters = ::ParseBool(it->second, outSettings.micForceTsFilters);
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
    ss << "hotkey.mute.modifiers=" << safe.muteHotkeyModifiers << "\n";
    ss << "hotkey.mute.vk=" << safe.muteHotkeyVk << "\n";
    ss << "capture.device_id=" << safe.captureDeviceId << "\n";
    ss << "mic.gate.mode=" << MicGateModeToString(safe.micGateMode) << "\n";
    ss << "mic.gate.threshold_dbfs=" << safe.micGateThresholdDbfs << "\n";
    ss << "mic.gate.smooth=" << BoolToString(safe.micUseSmoothGate) << "\n";
    ss << "mic.gate.keyboard_guard=" << BoolToString(safe.micUseKeyboardGuard) << "\n";
    ss << "mic.force_ts_filters=" << BoolToString(safe.micForceTsFilters) << "\n";
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
    musicMuted_.store(settings.musicMuted, std::memory_order_release);
    forceTxEnabled_.store(settings.forceTxEnabled, std::memory_order_release);
    musicGainLinear_.store(DbToLinear(gainDb), std::memory_order_release);
    bufferTargetMs_.store(bufferMs, std::memory_order_release);
    micUseSmoothGate_.store(settings.micUseSmoothGate, std::memory_order_release);
    micTalkDetected_.store(false, std::memory_order_release);
    talkState_.store(false, std::memory_order_release);
    micGateGain_.store(1.0f, std::memory_order_release);
    limiterGain_.store(1.0f, std::memory_order_release);
}

void AudioEngine::SetMusicSourceRunning(bool running) {
    sourceRunning_.store(running, std::memory_order_release);
    if (!running) {
        ring_.Reset();
        micTalkDetected_.store(false, std::memory_order_release);
        micRmsDbfs_.store(-120.0f, std::memory_order_release);
        externalMicLinear_.store(0.0f, std::memory_order_release);
        musicRmsDbfs_.store(-120.0f, std::memory_order_release);
        musicPeakDbfs_.store(-120.0f, std::memory_order_release);
        musicSendPeakDbfs_.store(-120.0f, std::memory_order_release);
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
void AudioEngine::SetMuted(bool muted) { musicMuted_.store(muted, std::memory_order_release); }
void AudioEngine::ToggleMute() {
    bool expected = musicMuted_.load(std::memory_order_relaxed);
    while (!musicMuted_.compare_exchange_weak(
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
    for (size_t i = 0; i < count; ++i) {
        const float v = muted ? 0.0f : (samples[i] * gain);
        sq += v * v;
        const float a = std::fabs(v);
        if (a > peak) {
            peak = a;
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
    t.reconnects = reconnectsMirror_.load(std::memory_order_relaxed);
    t.musicRmsDbfs = musicRmsDbfs_.load(std::memory_order_relaxed);
    t.musicPeakDbfs = musicPeakDbfs_.load(std::memory_order_relaxed);
    t.musicSendPeakDbfs = musicSendPeakDbfs_.load(std::memory_order_relaxed);
    t.talkStateActive = talkState_.load(std::memory_order_relaxed);
    t.micTalkDetected = micTalkDetected_.load(std::memory_order_relaxed);
    t.micRmsDbfs = micRmsDbfs_.load(std::memory_order_relaxed);
    const uint64_t lastSignalMs = lastMusicSignalTickMs_.load(std::memory_order_relaxed);
    t.musicActive = sourceRunning_.load(std::memory_order_relaxed) && (nowMs <= (lastSignalMs + 1200ULL));
    return t;
}

void AudioEngine::NoteReconnect() {
    reconnectsMirror_.fetch_add(1, std::memory_order_relaxed);
}

void AudioEngine::EditCapturedVoice(short* samples, int sampleCount, int channels, int* edited) {
    if (!samples || sampleCount <= 0 || channels <= 0 || !edited) {
        return;
    }
    const int upstreamFlags = *edited;
    const uint64_t nowMs = GetTickCount64();
    lastConsumeTickMs_.store(nowMs, std::memory_order_release);
    const bool muted = musicMuted_.load(std::memory_order_relaxed);
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
    const uint64_t lastSignalMs = lastMusicSignalTickMs_.load(std::memory_order_relaxed);
    const bool recentMusicSignal = sourceRunning && !muted && (nowMs <= (lastSignalMs + kForceTxMusicWindowMs));

    float micRmsAcc = 0.0f;
    for (int i = 0; i < sampleCount; ++i) {
        float micMono = 0.0f;
        for (int ch = 0; ch < channels; ++ch) {
            const int idx = (i * channels) + ch;
            micMono += static_cast<float>(samples[idx]) / 32768.0f;
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
    micTalkDetected_.store(talkOpen, std::memory_order_release);

    if (!sourceRunning || muted || !hasQueuedMusic) {
        const float prevSendPeakDb = musicSendPeakDbfs_.load(std::memory_order_relaxed);
        const float decayedSendPeakDb = std::max(-120.0f, prevSendPeakDb - 1.8f);
        musicSendPeakDbfs_.store(decayedSendPeakDb, std::memory_order_release);
        micTalkDetected_.store(talkOpen, std::memory_order_release);
        bool gateTouched = false;

        // TS3 bitmask semantics:
        // bit 1 (value 1): audio buffer modified
        // bit 2 (value 2): packet should be sent
        // Keep upstream flags untouched when we did not mix anything.
        if (gateGain < 0.9999f || targetGateGain < 0.9999f) {
            for (int i = 0; i < sampleCount; ++i) {
                advanceGate();
                for (int ch = 0; ch < channels; ++ch) {
                    const int idx = (i * channels) + ch;
                    const float dry = (static_cast<float>(samples[idx]) / 32768.0f) * gateGain;
                    const float clipped = std::clamp(dry, -1.0f, 1.0f);
                    const short next = static_cast<short>(std::lrintf(clipped * 32767.0f));
                    if (next != samples[idx]) {
                        gateTouched = true;
                        samples[idx] = next;
                    }
                }
            }
        }
        // Match limiter recovery behavior to the per-sample release path used
        // while mixing so transitions stay consistent when music stops.
        const float recoveryFactor = 1.0f - std::pow(1.0f - kLimiterReleaseCoeff, static_cast<float>(sampleCount));
        limiterGain += (1.0f - limiterGain) * recoveryFactor;
        limiterGain = std::clamp(limiterGain, 0.1f, 1.0f);
        micGateGain_.store(gateGain, std::memory_order_relaxed);
        limiterGain_.store(limiterGain, std::memory_order_relaxed);
        int outFlags = upstreamFlags;
        if (gateTouched || !talkOpen) {
            outFlags |= 1;
        }
        if (forceTx && recentMusicSignal) {
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
                const float dry = (static_cast<float>(samples[idx]) / 32768.0f) * gateGain;
                const float pre = dry + m;
                framePrePeak = std::max(framePrePeak, std::fabs(pre));
            }
            advanceLimiter(framePrePeak);
            for (int ch = 0; ch < channels; ++ch) {
                const int idx = ((offset + i) * channels) + ch;
                const short prevSample = samples[idx];
                const float dry = (static_cast<float>(samples[idx]) / 32768.0f) * gateGain;
                float out = dry + m;
                out *= limiterGain;
                const float absOut = std::fabs(out);
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

    const bool mixedMusic = touched && anyMusicSignal && sourceRunning && !muted;
    const bool forceMusicTx = forceTx && (recentMusicSignal || mixedMusic);

    int outFlags = upstreamFlags;
    if (mixedMusic || gateTouched) {
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
        } catch (const std::exception& ex) {
            SetStatus(SourceState::Error, "thread_start_failed", "Source thread start failed", ex.what());
        } catch (...) {
            SetStatus(SourceState::Error, "thread_start_failed", "Source thread start failed", "");
        }
        return false;
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
        const bool isFloat = IsFloatFormat(wf);
        bool isPcm16 = IsPcm16Format(wf);
        if (!isFloat && !isPcm16) {
            if (ownsWf) {
                releaseWf();
            }
            useFallbackFormat();
            isPcm16 = true;
            detail = "pid=" + std::to_string(pid) + " using=fallback_pcm_44100";
        }
        if (!IsSupportedSourceRate(wf->nSamplesPerSec)) {
            if (ownsWf) {
                releaseWf();
                useFallbackFormat();
                isPcm16 = true;
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
    explicit GlobalHotkeyManager(std::function<void()> onHotkey)
        : onHotkey_(std::move(onHotkey)) {}

    ~GlobalHotkeyManager() {
        Stop();
    }

    void Start() {
        std::unique_lock<std::mutex> lock(mutex_);
        if (thread_.joinable()) {
            return;
        }
        threadReady_ = false;
        try {
            thread_ = std::thread([this]() { ThreadMain(); });
        } catch (const std::exception& ex) {
            LogError(std::string("mute_hotkey thread start failed: ") + ex.what());
            return;
        } catch (...) {
            LogError("mute_hotkey thread start failed: unknown");
            return;
        }
        if (!startedCv_.wait_for(lock, std::chrono::seconds(2), [this]() { return threadReady_; })) {
            LogWarn("mute_hotkey thread ready timeout");
        }
    }

    void Stop() {
        DWORD threadId = 0;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            threadId = threadId_;
        }
        if (threadId != 0) {
            PostThreadMessageW(threadId, WM_APP + 0x12, 0, 0);
        }
        if (thread_.joinable()) {
            thread_.join();
        }
    }

    void ApplySettings(const MicMixSettings& settings) {
        const UINT newMods = static_cast<UINT>(std::max(0, settings.muteHotkeyModifiers));
        const UINT newVk = static_cast<UINT>(std::max(0, settings.muteHotkeyVk));
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
        DWORD threadId = 0;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            threadId = threadId_;
        }
        if (threadId != 0) {
            PostThreadMessageW(threadId, WM_APP + 0x11, 0, 0);
        }
    }

private:
    static constexpr int kHotkeyId = 0x4D4D;
    std::function<void()> onHotkey_;
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

    void ApplyRegistration() {
        UINT mods = 0;
        UINT vk = 0;
        {
            std::lock_guard<std::mutex> lock(mutex_);
            mods = modifiers_;
            vk = vk_;
        }

        UnregisterHotKey(nullptr, kHotkeyId);
        if (vk == 0) {
            registeredValid_ = false;
            return;
        }

        UINT regMods = mods;
#ifdef MOD_NOREPEAT
        regMods |= MOD_NOREPEAT;
#endif
        if (!RegisterHotKey(nullptr, kHotkeyId, regMods, vk)) {
            registeredValid_ = false;
            LogWarn("mute_hotkey register failed vk=" + std::to_string(vk) +
                    " mods=" + std::to_string(mods) +
                    " err=" + std::to_string(GetLastError()));
        } else {
            const bool changed = (!registeredValid_) || (registeredMods_ != mods) || (registeredVk_ != vk);
            if (changed) {
                LogInfo("mute_hotkey registered vk=" + std::to_string(vk) +
                        " mods=" + std::to_string(mods));
            }
            registeredMods_ = mods;
            registeredVk_ = vk;
            registeredValid_ = true;
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

        while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
            if (msg.message == WM_APP + 0x11) {
                ApplyRegistration();
                continue;
            }
            if (msg.message == WM_APP + 0x12) {
                break;
            }
            if (msg.message == WM_HOTKEY && static_cast<int>(msg.wParam) == kHotkeyId) {
                if (onHotkey_) {
                    try {
                        onHotkey_();
                    } catch (const std::exception& ex) {
                        LogError(std::string("mute_hotkey callback exception: ") + ex.what());
                    } catch (...) {
                        LogError("mute_hotkey callback exception: unknown");
                    }
                }
            }
        }

        UnregisterHotKey(nullptr, kHotkeyId);
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
        Stop();
    }

    void SetEnabled(bool enabled) {
        if (enabled) {
            Start();
        } else {
            Stop();
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
        thread_local std::vector<short> stereo;
        stereo.resize(frames * 2);
        for (size_t i = 0; i < frames; ++i) {
            const size_t base = i * static_cast<size_t>(channels);
            const short l = samples[base];
            const short r = (channels >= 2) ? samples[base + 1] : l;
            stereo[(i * 2) + 0] = l;
            stereo[(i * 2) + 1] = r;
        }
        const size_t written = ring_.Write(stereo.data(), stereo.size());
        if (written < stereo.size()) {
            droppedSamples_.fetch_add(static_cast<uint64_t>(stereo.size() - written), std::memory_order_relaxed);
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

    void Start() {
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

    void Stop() {
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

    if (!active || schid == 0 || !IsConnectedForTx(schid)) {
        if (currentlyActive && currentSchid != 0) {
            const bool restored = restoreStateForSchid(currentSchid);
            if (!restored) {
                const uint64_t nowMs = GetTickCount64();
                if (nowMs > (lastErrLogMs + 2000ULL)) {
                    LogWarn("talk_mode restore failed schid=" + std::to_string(currentSchid));
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
        restoreStateForSchid(currentSchid);
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
    const SourceStatus st = sourceManager_ ? sourceManager_->GetStatus() : SourceStatus{};
    const TelemetrySnapshot t = engine_.SnapshotTelemetry();
    const bool sourceUp = (st.state == SourceState::Running) ||
                          (st.state == SourceState::Starting) ||
                          (st.state == SourceState::Reacquiring);
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

void MicMixApp::VoiceTxThreadMain() {
    voiceTxThreadRunning_.store(true, std::memory_order_release);
    try {
        uint64_t lastLogMs = 0;
        uint64_t lastVadFallbackLogMs = 0;
        uint64_t vadGateHoldUntilMs = 0;
        uint64_t vadAboveThresholdSinceMs = 0;
        bool vadQualifiedOpen = false;
        while (!voiceTxStop_.load(std::memory_order_acquire)) {
            uint64 schid = activeSchid_.load(std::memory_order_acquire);
            if (schid == 0 && g_ts3Functions.getCurrentServerConnectionHandlerID) {
                schid = g_ts3Functions.getCurrentServerConnectionHandlerID();
                if (schid != 0) {
                    activeSchid_.store(schid, std::memory_order_release);
                }
            }
            RefreshVoiceTxControl(schid);

            if (schid != 0 && g_ts3Functions.getPreProcessorInfoValueFloat) {
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

                if (schid != 0) {
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
                            const uint64_t nowMs = GetTickCount64();
                            bool candidateOpen = false;
                            bool haveDecision = false;

                            const uint64_t talkEventMs = ownTalkStatusTickMs_.load(std::memory_order_acquire);
                            if (talkEventMs != 0 && nowMs <= (talkEventMs + kTalkEventFreshMs)) {
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
                                if (talkEventMs != 0 && nowMs <= (talkEventMs + kTalkEventFreshMs)) {
                                    candidateOpen = ownTalkStatusActive_.load(std::memory_order_acquire);
                                    haveDecision = true;
                                } else {
                                    if (haveMicDb) {
                                        candidateOpen = micDb >= vadThresholdDb;
                                    } else {
                                        const TelemetrySnapshot tel = engine_.SnapshotTelemetry();
                                        candidateOpen = tel.micRmsDbfs >= (vadThresholdDb - 1.0f);
                                    }
                                    haveDecision = true;
                                    if (nowMs > (lastVadFallbackLogMs + 4000ULL)) {
                                        LogWarn("talk_gate auto_ts stale event schid=" + std::to_string(schid));
                                        lastVadFallbackLogMs = nowMs;
                                    }
                                }
                                vadAboveThresholdSinceMs = candidateOpen ? nowMs : 0;
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
                                                    vadAboveThresholdSinceMs = nowMs;
                                                }
                                                candidateOpen = nowMs >= (vadAboveThresholdSinceMs + 60ULL);
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
                                    vadAboveThresholdSinceMs = candidateOpen ? nowMs : 0;
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
                                if (nowMs > (lastVadFallbackLogMs + 4000ULL)) {
                                    LogWarn("talk_gate fallback used schid=" + std::to_string(schid));
                                    lastVadFallbackLogMs = nowMs;
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
                                vadGateHoldUntilMs = nowMs + effectiveHoldMs;
                                vadQualifiedOpen = true;
                                talkOpen = true;
                            } else if (effectiveHoldMs > 0 && nowMs <= vadGateHoldUntilMs) {
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

                        const uint64_t nowMs = GetTickCount64();
                        const uint64_t lastCaptureMs = lastCaptureEditTickMs_.load(std::memory_order_acquire);
                        const uint64_t lastNudgeMs = lastCaptureReopenTickMs_.load(std::memory_order_acquire);
                        if (nowMs > (lastCaptureMs + kCaptureWatchdogSilenceMs) &&
                            nowMs > (lastNudgeMs + kCaptureWatchdogCooldownMs)) {
                            const unsigned int err = g_ts3Functions.activateCaptureDevice(schid);
                            lastCaptureReopenTickMs_.store(nowMs, std::memory_order_release);
                            if (!(err == ERROR_ok || err == ERROR_ok_no_update)) {
                                if (nowMs > (lastLogMs + 2000ULL)) {
                                    LogWarn("capture_activate watchdog failed schid=" + std::to_string(schid) +
                                            " err=" + std::to_string(err));
                                    lastLogMs = nowMs;
                                }
                            } else if (nowMs > (lastLogMs + 3000ULL)) {
                                LogInfo("capture_activate watchdog schid=" + std::to_string(schid) +
                                        " err=" + std::to_string(err));
                                lastLogMs = nowMs;
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
        configStore_ = std::make_unique<ConfigStore>(configBasePath);
        g_logPath = configStore_->LogPath();
        std::string warn;
        configStore_->Load(settings_, warn);
        SanitizeSettings(settings_);
        if (!warn.empty()) LogWarn(warn);
        engine_.ApplySettings(settings_);
        // Disabled intentionally: an additional capture stream can interfere with
        // some driver/device combinations and TS microphone preview.
        // Mic activity is derived from TS3 capture callback path.
        micLevelMonitor_.reset();
        hotkeyManager_ = std::make_unique<GlobalHotkeyManager>([this]() {
            this->ToggleMute();
        });
        hotkeyManager_->Start();
        hotkeyManager_->ApplySettings(settings_);
        sourceManager_ = std::make_unique<AudioSourceManager>(
            [this](const float* data, size_t count) { engine_.PushMusicSamples(data, count); },
            [this](SourceState st, const std::string& code, const std::string& msg, const std::string& detail) {
                engine_.SetMusicSourceRunning(st == SourceState::Running);
                if (st == SourceState::Reacquiring) {
                    engine_.NoteReconnect();
                }
                std::string line = "source=" + SourceStateToString(st) + " code=" + code + " msg=" + msg;
                if (!detail.empty()) line += " detail=" + detail;
                LogInfo(line);
            });
        sourceManager_->ApplySettings(settings_);
        mixMonitorPlayer_ = std::make_unique<MixMonitorPlayer>();
        if (settings_.autostartEnabled) sourceManager_->Start();
        const int activeResamplerQuality = ResolveResamplerQualitySetting(settings_.resamplerQuality);
        const std::string mode = (settings_.resamplerQuality < 0) ? "auto_cpu" : "manual";
        LogInfo("resampler " + mode + " quality=" + std::to_string(activeResamplerQuality) +
                " logical_cpus=" + std::to_string(GetLogicalCpuCount()));
        const uint64_t nowMs = GetTickCount64();
        lastCaptureEditTickMs_.store(nowMs, std::memory_order_release);
        lastCaptureReopenTickMs_.store(0, std::memory_order_release);
        StartVoiceTxThread();
        LogInfo("MicMix initialized");
        return true;
    } catch (const std::exception& ex) {
        LogError(std::string("MicMix initialize failed: ") + ex.what());
    } catch (...) {
        LogError("MicMix initialize failed: unknown");
    }

    StopVoiceTxThread();
    if (hotkeyManager_) hotkeyManager_->Stop();
    if (sourceManager_) sourceManager_->Stop();
    if (mixMonitorPlayer_) mixMonitorPlayer_->SetEnabled(false);
    micLevelMonitor_.reset();
    hotkeyManager_.reset();
    sourceManager_.reset();
    mixMonitorPlayer_.reset();
    configStore_.reset();
    initialized_.store(false, std::memory_order_release);
    return false;
}

void MicMixApp::Shutdown() {
    if (!initialized_.exchange(false)) return;
    SettingsWindowController::Instance().Close();
    if (hotkeyManager_) hotkeyManager_->Stop();
    if (sourceManager_) sourceManager_->Stop();
    if (mixMonitorPlayer_) mixMonitorPlayer_->SetEnabled(false);
    StopVoiceTxThread();
    if (configStore_) {
        std::string err;
        configStore_->Save(settings_, err);
        if (!err.empty()) LogWarn("Config save warning: " + err);
    }
    hotkeyManager_.reset();
    sourceManager_.reset();
    mixMonitorPlayer_.reset();
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
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        changed = !IsSameSettings(settings_, safe);
        settings_ = safe;
    }
    if (!changed && !restartSource) {
        return;
    }

    engine_.ApplySettings(safe);
    if (hotkeyManager_) {
        hotkeyManager_->ApplySettings(safe);
    }
    if (sourceManager_) {
        sourceManager_->ApplySettings(safe);
        if (restartSource && sourceManager_->IsRunning()) sourceManager_->Restart();
    }
    if (saveConfig && changed && configStore_) {
        std::string err;
        if (!configStore_->Save(safe, err) && !err.empty()) LogError("Config save failed: " + err);
    }
}

void MicMixApp::StartSource() { if (sourceManager_) sourceManager_->Start(); }
void MicMixApp::StopSource() { if (sourceManager_) sourceManager_->Stop(); }

void MicMixApp::ToggleMute() {
    engine_.ToggleMute();
    auto s = GetSettings();
    s.musicMuted = engine_.IsMuted();
    ApplySettings(s, false, true);
}

void MicMixApp::SetMonitorEnabled(bool enabled) {
    if (mixMonitorPlayer_) {
        mixMonitorPlayer_->SetEnabled(enabled);
        LogInfo(std::string("mix_monitor ") + (mixMonitorPlayer_->IsEnabled() ? "enabled" : "disabled"));
    }
}

bool MicMixApp::IsMonitorEnabled() const {
    return mixMonitorPlayer_ ? mixMonitorPlayer_->IsEnabled() : false;
}

void MicMixApp::ToggleMonitor() {
    SetMonitorEnabled(!IsMonitorEnabled());
}

void MicMixApp::SetPushToPlayActive(bool active) {
    const bool wasActive = pushToPlayActive_.exchange(active, std::memory_order_acq_rel);
    if (active && !wasActive) {
        const bool currentMuted = engine_.IsMuted();
        pushToPlaySavedMuteState_.store(currentMuted, std::memory_order_release);
        pushToPlaySavedMuteValid_.store(true, std::memory_order_release);
        engine_.SetMuted(false);
    } else if (!active && wasActive) {
        bool restoreMuted = engine_.IsMuted();
        if (pushToPlaySavedMuteValid_.load(std::memory_order_acquire)) {
            restoreMuted = pushToPlaySavedMuteState_.load(std::memory_order_acquire);
        }
        pushToPlaySavedMuteValid_.store(false, std::memory_order_release);
        engine_.SetMuted(restoreMuted);
    }
    auto s = GetSettings();
    s.musicMuted = engine_.IsMuted();
    ApplySettings(s, false, true);
}

void MicMixApp::TogglePushToPlay() { SetPushToPlayActive(!pushToPlayActive_.load(std::memory_order_acquire)); }

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
    activeSchid_.store(schid, std::memory_order_release);
    RefreshVoiceTxControl(schid);
}

void MicMixApp::OnConnectStatusChange(uint64 schid, int newStatus, unsigned int errorNumber) {
    (void)errorNumber;
    if (newStatus == STATUS_DISCONNECTED) {
        if (activeSchid_.load(std::memory_order_acquire) == schid) {
            activeSchid_.store(0, std::memory_order_release);
        }
        SetVoiceRecordingState(false, 0);
        return;
    }
    if (newStatus == STATUS_CONNECTION_ESTABLISHED || newStatus == STATUS_CONNECTION_ESTABLISHING || newStatus == STATUS_CONNECTED) {
        activeSchid_.store(schid, std::memory_order_release);
        RefreshVoiceTxControl(schid);
    }
}
std::vector<LoopbackDeviceInfo> MicMixApp::GetLoopbackDevices() const { return AudioSourceManager::EnumerateLoopbackDevices(); }
std::vector<CaptureDeviceInfo> MicMixApp::GetCaptureDevices() const { return AudioSourceManager::EnumerateCaptureDevices(); }
std::vector<AppProcessInfo> MicMixApp::GetAppProcesses() const {
    return AudioSourceManager::EnumerateAppProcesses();
}
SourceStatus MicMixApp::GetSourceStatus() const { return sourceManager_ ? sourceManager_->GetStatus() : SourceStatus{}; }
TelemetrySnapshot MicMixApp::GetTelemetry() const {
    TelemetrySnapshot t = engine_.SnapshotTelemetry();
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
std::string MicMixApp::GetConfigDir() const { return configStore_ ? configStore_->ConfigPath() : std::string{}; }
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
    if (!initialized_.load(std::memory_order_acquire)) return;
    const uint64_t nowMs = GetTickCount64();
    lastCaptureEditTickMs_.store(nowMs, std::memory_order_release);
    (void)schid;
    engine_.EditCapturedVoice(samples, sampleCount, channels, edited);
    if (mixMonitorPlayer_) {
        mixMonitorPlayer_->PushCaptured(samples, sampleCount, channels);
    }
}

void MicMixApp::OpenSettingsWindow() {
    SettingsWindowController::Instance().Open();
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
