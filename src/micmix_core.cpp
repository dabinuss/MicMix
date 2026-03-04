#include "micmix_core.h"

#include <audioclient.h>
#include <audioclientactivationparams.h>
#include <audiopolicy.h>
#include <endpointvolume.h>
#include <functiondiscoverykeys_devpkey.h>
#include <Shlwapi.h>
#include <tlhelp32.h>
#include <wrl/client.h>

#include <speex/speex_resampler.h>

#include <algorithm>
#include <array>
#include <cctype>
#include <chrono>
#include <condition_variable>
#include <cmath>
#include <cwctype>
#include <cstdlib>
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

std::mutex g_logMutex;
std::string g_logPath;

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
    HRESULT hr = CoInitializeEx(nullptr, COINIT_MULTITHREADED);
    ~ComInit() {
        if (SUCCEEDED(hr)) {
            CoUninitialize();
        }
    }
};

void AppendLogLine(const char* level, const std::string& text) {
    std::lock_guard<std::mutex> lock(g_logMutex);
    if (g_logPath.empty()) {
        return;
    }
    std::ofstream out(g_logPath, std::ios::app);
    if (!out.is_open()) {
        return;
    }
    SYSTEMTIME st{};
    GetLocalTime(&st);
    out << st.wYear << "-" << st.wMonth << "-" << st.wDay << " "
        << st.wHour << ":" << st.wMinute << ":" << st.wSecond << "."
        << st.wMilliseconds << " [" << level << "] " << text << "\n";
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
    int quality_ = 6;
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
    const std::filesystem::path dir = std::filesystem::path(basePath_) / "plugins" / "micmix";
    std::filesystem::create_directories(dir);
    configPath_ = (dir / "config.json").string();
    tmpPath_ = (dir / "config.tmp").string();
    lastGoodPath_ = (dir / "config.lastgood.json").string();
    logPath_ = (dir / "micmix.log").string();
}

std::string ConfigStore::Trim(const std::string& value) {
    return ::Trim(value);
}

bool ConfigStore::ParseBool(const std::string& value, bool& out) {
    out = ::ParseBool(value, out);
    return true;
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

std::string ConfigStore::DuckingModeToString(DuckingMode mode) {
    if (mode == DuckingMode::PluginHotkey) {
        return "plugin_hotkey";
    }
    if (mode == DuckingMode::Ts3Talkstate) {
        return "ts3_talkstate";
    }
    return "mic_rms";
}

DuckingMode ConfigStore::DuckingModeFromString(const std::string& value) {
    if (value == "plugin_hotkey") {
        return DuckingMode::PluginHotkey;
    }
    if (value == "ts3_talkstate") {
        return DuckingMode::Ts3Talkstate;
    }
    return DuckingMode::MicRms;
}

bool ConfigStore::Load(MicMixSettings& outSettings, std::string& warning) {
    warning.clear();
    std::ifstream in(configPath_);
    if (!in.is_open()) {
        return true;
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

    try {
        if (auto it = kv.find("config.version"); it != kv.end()) outSettings.configVersion = std::stoi(it->second);
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
        if (auto it = kv.find("mix.music_gain_db"); it != kv.end()) outSettings.musicGainDb = std::stof(it->second);
        if (auto it = kv.find("mix.force_tx_enabled"); it != kv.end()) outSettings.forceTxEnabled = ::ParseBool(it->second, outSettings.forceTxEnabled);
        if (auto it = kv.find("mix.buffer_target_ms"); it != kv.end()) outSettings.bufferTargetMs = std::stoi(it->second);
        if (auto it = kv.find("mix.music_muted"); it != kv.end()) outSettings.musicMuted = ::ParseBool(it->second, outSettings.musicMuted);
        if (auto it = kv.find("hotkey.mute.modifiers"); it != kv.end()) outSettings.muteHotkeyModifiers = std::stoi(it->second);
        if (auto it = kv.find("hotkey.mute.vk"); it != kv.end()) outSettings.muteHotkeyVk = std::stoi(it->second);
        if (auto it = kv.find("capture.device_id"); it != kv.end()) outSettings.captureDeviceId = it->second;
        if (auto it = kv.find("ui.last_open_tab"); it != kv.end()) outSettings.uiLastOpenTab = std::stoi(it->second);
    } catch (...) {
        warning = "Config parse issue detected; fallback values were used.";
    }
    outSettings.duckingEnabled = false;
    outSettings.duckingMode = DuckingMode::MicRms;
    outSettings.duckingThresholdDbfs = -30.0f;
    outSettings.duckingAttackMs = 20.0f;
    outSettings.duckingReleaseMs = 220.0f;
    outSettings.duckingAmountDb = 0.0f;
    return true;
}

bool ConfigStore::Save(const MicMixSettings& settings, std::string& error) {
    error.clear();
    std::ostringstream ss;
    ss << "config.version=" << settings.configVersion << "\n";
    ss << "source.mode=" << SourceModeToString(settings.sourceMode) << "\n";
    ss << "source.loopback.device_id=" << settings.loopbackDeviceId << "\n";
    ss << "source.app.process_name=" << settings.appProcessName << "\n";
    ss << "source.app.session_id=" << settings.appSessionId << "\n";
    ss << "source.autostart_enabled=" << BoolToString(settings.autostartEnabled) << "\n";
    ss << "source.auto_switch_to_loopback=" << BoolToString(settings.autoSwitchToLoopback) << "\n";
    ss << "mix.music_gain_db=" << settings.musicGainDb << "\n";
    ss << "mix.force_tx_enabled=" << BoolToString(settings.forceTxEnabled) << "\n";
    ss << "mix.buffer_target_ms=" << settings.bufferTargetMs << "\n";
    ss << "mix.music_muted=" << BoolToString(settings.musicMuted) << "\n";
    ss << "hotkey.mute.modifiers=" << settings.muteHotkeyModifiers << "\n";
    ss << "hotkey.mute.vk=" << settings.muteHotkeyVk << "\n";
    ss << "capture.device_id=" << settings.captureDeviceId << "\n";
    ss << "ui.last_open_tab=" << settings.uiLastOpenTab << "\n";

    {
        std::ofstream out(tmpPath_, std::ios::binary | std::ios::trunc);
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

    if (PathFileExistsA(configPath_.c_str())) {
        CopyFileA(configPath_.c_str(), lastGoodPath_.c_str(), FALSE);
    }
    if (!MoveFileExA(tmpPath_.c_str(), configPath_.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
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
    musicMuted_.store(settings.musicMuted, std::memory_order_release);
    forceTxEnabled_.store(settings.forceTxEnabled, std::memory_order_release);
    musicGainLinear_.store(DbToLinear(settings.musicGainDb), std::memory_order_release);
    bufferTargetMs_.store(settings.bufferTargetMs, std::memory_order_release);
    duckingEnabled_.store(false, std::memory_order_release);
    duckingMode_.store(static_cast<int>(DuckingMode::MicRms), std::memory_order_release);
    duckThresholdLinear_.store(DbToLinear(-30.0f), std::memory_order_release);
    duckAmountLinear_.store(1.0f, std::memory_order_release);
    duckAttackMs_.store(20.0f, std::memory_order_release);
    duckReleaseMs_.store(220.0f, std::memory_order_release);
    duckCurrent_ = 1.0f;
    duckMeterLinear_.store(1.0f, std::memory_order_release);
    duckingNow_.store(false, std::memory_order_release);
    micTalkDetected_.store(false, std::memory_order_release);
    talkState_.store(false, std::memory_order_release);
    hotkeyDucking_.store(false, std::memory_order_release);
}

void AudioEngine::SetMusicSourceRunning(bool running) {
    sourceRunning_.store(running, std::memory_order_release);
    if (!running) {
        duckCurrent_ = 1.0f;
        duckMeterLinear_.store(1.0f, std::memory_order_release);
        duckingNow_.store(false, std::memory_order_release);
        micTalkDetected_.store(false, std::memory_order_release);
        micRmsDbfs_.store(-120.0f, std::memory_order_release);
        externalMicLinear_.store(0.0f, std::memory_order_release);
        musicRmsDbfs_.store(-120.0f, std::memory_order_release);
        musicPeakDbfs_.store(-120.0f, std::memory_order_release);
        musicSendPeakDbfs_.store(-120.0f, std::memory_order_release);
    }
}
void AudioEngine::SetTalkState(bool talking) { (void)talking; }
void AudioEngine::SetHotkeyDucking(bool active) { (void)active; }
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
void AudioEngine::ToggleMute() { musicMuted_.store(!musicMuted_.load(std::memory_order_acquire), std::memory_order_release); }
void AudioEngine::ClearMusicBuffer() { ring_.Reset(); }

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
    if (peakDb > -55.0f) {
        lastMusicSignalTickMs_.store(GetTickCount64(), std::memory_order_release);
    }

    const uint64_t nowMs = GetTickCount64();
    const uint64_t lastConsumeMs = lastConsumeTickMs_.load(std::memory_order_acquire);
    if (nowMs > lastConsumeMs + 1500ULL) {
        // No consumer activity for a while (e.g. no outgoing TS3 capture callback).
        // Drop stale buffered audio so producer cannot accumulate endless backlog.
        ring_.Reset();
    }
    const size_t written = ring_.Write(samples, count);
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
    t.duckingGainDb = 0.0f;
    t.duckingActive = false;
    t.talkStateActive = false;
    t.micTalkDetected = false;
    t.micRmsDbfs = micRmsDbfs_.load(std::memory_order_relaxed);
    const uint64_t lastSignalMs = lastMusicSignalTickMs_.load(std::memory_order_relaxed);
    t.musicActive = sourceRunning_.load(std::memory_order_relaxed) && (nowMs <= (lastSignalMs + 1200ULL));
    return t;
}

void AudioEngine::EditCapturedVoice(short* samples, int sampleCount, int channels, int* edited) {
    if (!samples || sampleCount <= 0 || channels <= 0 || !edited) {
        return;
    }
    const int upstreamEdited = *edited;
    lastConsumeTickMs_.store(GetTickCount64(), std::memory_order_release);
    const bool muted = musicMuted_.load(std::memory_order_relaxed);
    const bool sourceRunning = sourceRunning_.load(std::memory_order_relaxed);

    thread_local std::array<float, kCallbackScratch> music{};
    bool anyMusicSignal = false;
    bool touched = false;
    float mixPeak = 0.0f;
    int offset = 0;
    while (offset < sampleCount) {
        const int chunk = std::min(sampleCount - offset, static_cast<int>(kCallbackScratch));
        const size_t pulled = ring_.Read(music.data(), static_cast<size_t>(chunk));
        if (pulled < static_cast<size_t>(chunk)) {
            std::fill(music.begin() + static_cast<std::ptrdiff_t>(pulled), music.begin() + chunk, 0.0f);
            underruns_.fetch_add(static_cast<uint64_t>(chunk - pulled), std::memory_order_relaxed);
        }

        float rmsAcc = 0.0f;
        for (int i = 0; i < chunk; ++i) {
            const float mic = static_cast<float>(samples[(offset + i) * channels]) / 32768.0f;
            rmsAcc += mic * mic;
        }
        const float rms = std::sqrt(rmsAcc / static_cast<float>(chunk));
        const float rmsDb = 20.0f * std::log10(std::max(rms, 0.000001f));
        const float prevMicRms = micRmsDbfs_.load(std::memory_order_relaxed);
        const float micAlpha = (rmsDb > prevMicRms) ? 0.25f : 0.10f;
        const float nextMicRms = std::clamp(prevMicRms + ((rmsDb - prevMicRms) * micAlpha), -120.0f, 0.0f);
        micRmsDbfs_.store(nextMicRms, std::memory_order_release);
        const float musicGain = musicGainLinear_.load(std::memory_order_relaxed);
        for (int i = 0; i < chunk; ++i) {
            const float m = muted ? 0.0f : (music[i] * musicGain);
            const float absM = std::fabs(m);
            if (absM > 0.0005f) {
                anyMusicSignal = true;
            }
            if (absM > mixPeak) {
                mixPeak = absM;
            }
            for (int ch = 0; ch < channels; ++ch) {
                const int idx = ((offset + i) * channels) + ch;
                float out = (static_cast<float>(samples[idx]) / 32768.0f) + m;
                if (out > 1.0f) {
                    out = 1.0f;
                    clippedSamples_.fetch_add(1, std::memory_order_relaxed);
                } else if (out < -1.0f) {
                    out = -1.0f;
                    clippedSamples_.fetch_add(1, std::memory_order_relaxed);
                }
                samples[idx] = static_cast<short>(std::lrintf(out * 32767.0f));
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
    duckCurrent_ = 1.0f;
    duckMeterLinear_.store(1.0f, std::memory_order_release);
    duckingNow_.store(false, std::memory_order_release);
    micTalkDetected_.store(false, std::memory_order_release);

    const int mineEdited = (touched && anyMusicSignal && sourceRunning && !muted) ? 1 : 0;
    *edited = (upstreamEdited || mineEdited) ? 1 : 0;
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
        thread_ = std::thread([this] { Run(); });
        return true;
    }

    void Stop() override {
        stop_.store(true, std::memory_order_release);
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

    static bool IsRetryableFailureCode(const std::string& code) {
        if (code == "activate_unsupported") {
            return false;
        }
        return true;
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
            if (CaptureOnce(code, msg, detail)) {
                backoff = 1;
                continue;
            }
            if (StopRequested()) break;
            if (!IsRetryableFailureCode(code)) {
                SetStatus(SourceState::Error, code.empty() ? "error" : code, msg.empty() ? "Source error" : msg, detail);
                break;
            }
            SetStatus(SourceState::Reacquiring, code.empty() ? "reacquire" : code, msg.empty() ? "Reacquiring source" : msg, detail);
            std::this_thread::sleep_for(std::chrono::seconds(backoff));
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
        hr = client->Initialize(AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK, 0, 0, wf, nullptr);
        if (FAILED(hr)) {
            code = "client_init_failed"; msg = "Loopback init failed"; releaseWf(); return false;
        }

        ComPtr<IAudioCaptureClient> cap;
        hr = client->GetService(IID_PPV_ARGS(&cap));
        if (FAILED(hr) || !cap) {
            code = "capture_service_failed"; msg = "Loopback capture service failed"; releaseWf(); return false;
        }
        if (!resampler_.Configure(wf->nSamplesPerSec, kTargetRate, 6)) {
            code = "resampler_failed"; msg = "Resampler init failed"; releaseWf(); return false;
        }
        const int channels = std::max<int>(1, wf->nChannels);
        hr = client->Start();
        if (FAILED(hr)) {
            code = "start_failed"; msg = "Loopback start failed"; releaseWf(); return false;
        }
        SetStatus(SourceState::Running, "running", "Loopback running", "");

        while (!StopRequested()) {
            UINT32 nextPacket = 0;
            hr = cap->GetNextPacketSize(&nextPacket);
            if (FAILED(hr)) {
                client->Stop();
                releaseWf();
                code = "packet_failed"; msg = "Packet query failed"; return false;
            }
            if (nextPacket == 0) {
                std::this_thread::sleep_for(std::chrono::milliseconds(4));
                continue;
            }
            while (nextPacket > 0 && !StopRequested()) {
                BYTE* data = nullptr;
                UINT32 frames = 0;
                DWORD flags = 0;
                hr = cap->GetBuffer(&data, &frames, &flags, nullptr, nullptr);
                if (FAILED(hr)) {
                    client->Stop();
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
                    releaseWf();
                    code = "resample_failed"; msg = "Resample failed"; return false;
                }
                if (!monoOut_.empty()) Push(monoOut_.data(), monoOut_.size());
                cap->ReleaseBuffer(frames);
                hr = cap->GetNextPacketSize(&nextPacket);
                if (FAILED(hr)) {
                    client->Stop();
                    releaseWf();
                    code = "packet_failed"; msg = "Packet query failed"; return false;
                }
            }
        }

        client->Stop();
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
        if (!settings_.appSessionId.empty()) {
            uint32_t pid = static_cast<uint32_t>(strtoul(settings_.appSessionId.c_str(), nullptr, 10));
            addCandidate(pid);
        }
        if (settings_.appProcessName.empty()) {
            return out;
        }
        const auto list = AudioSourceManager::EnumerateAppProcesses(settings_.appProcessName);
        for (const auto& item : list) {
            addCandidate(item.pid);
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
        if (!event) {
            code = "event_failed"; msg = "Could not create event"; releaseWf(); return false;
        }
        hr = client->SetEventHandle(event);
        if (FAILED(hr)) {
            CloseHandle(event);
            if (!detail.empty()) {
                detail += " ";
            }
            detail += "pid=" + std::to_string(pid) + " event_hr=" + HrToHex(hr);
            code = "event_set_failed"; msg = "Could not set event"; releaseWf(); return false;
        }
        ComPtr<IAudioCaptureClient> cap;
        hr = client->GetService(IID_PPV_ARGS(&cap));
        if (FAILED(hr) || !cap) {
            CloseHandle(event);
            if (!detail.empty()) {
                detail += " ";
            }
            detail += "pid=" + std::to_string(pid) + " service_hr=" + HrToHex(hr);
            code = "capture_service_failed"; msg = "App capture service failed"; releaseWf(); return false;
        }
        if (!resampler_.Configure(wf->nSamplesPerSec, kTargetRate, 6)) {
            CloseHandle(event);
            code = "resampler_failed"; msg = "Resampler init failed"; releaseWf(); return false;
        }
        const int channels = std::max<int>(1, wf->nChannels);
        hr = client->Start();
        if (FAILED(hr)) {
            CloseHandle(event);
            if (!detail.empty()) {
                detail += " ";
            }
            detail += "pid=" + std::to_string(pid) + " start_hr=" + HrToHex(hr);
            code = "start_failed"; msg = "App start failed"; releaseWf(); return false;
        }
        SetStatus(SourceState::Running, "running", "App capture running", "");

        while (!StopRequested()) {
            DWORD wait = WaitForSingleObject(event, 200);
            if (wait == WAIT_TIMEOUT) {
                if (!IsProcessAlive(pid)) {
                    client->Stop();
                    CloseHandle(event);
                    releaseWf();
                    code = "session_lost"; msg = "App process exited"; return false;
                }
                continue;
            }
            if (wait != WAIT_OBJECT_0) {
                client->Stop();
                CloseHandle(event);
                releaseWf();
                code = "wait_failed"; msg = "Event wait failed"; return false;
            }
            UINT32 nextPacket = 0;
            hr = cap->GetNextPacketSize(&nextPacket);
            if (FAILED(hr)) {
                client->Stop();
                CloseHandle(event);
                releaseWf();
                code = "packet_failed"; msg = "Packet query failed"; return false;
            }
            while (nextPacket > 0 && !StopRequested()) {
                BYTE* data = nullptr;
                UINT32 frames = 0;
                DWORD flagsRead = 0;
                hr = cap->GetBuffer(&data, &frames, &flagsRead, nullptr, nullptr);
                if (FAILED(hr)) {
                    client->Stop();
                    CloseHandle(event);
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
                    CloseHandle(event);
                    releaseWf();
                    code = "resample_failed"; msg = "Resample failed"; return false;
                }
                if (!monoOut_.empty()) Push(monoOut_.data(), monoOut_.size());
                cap->ReleaseBuffer(frames);
                hr = cap->GetNextPacketSize(&nextPacket);
                if (FAILED(hr)) {
                    client->Stop();
                    CloseHandle(event);
                    releaseWf();
                    code = "packet_failed"; msg = "Packet query failed"; return false;
                }
            }
        }

        client->Stop();
        CloseHandle(event);
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
        thread_ = std::thread([this]() { ThreadMain(); });
        startedCv_.wait(lock, [this]() { return threadReady_; });
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
            return;
        }

        UINT regMods = mods;
#ifdef MOD_NOREPEAT
        regMods |= MOD_NOREPEAT;
#endif
        if (!RegisterHotKey(nullptr, kHotkeyId, regMods, vk)) {
            LogWarn("mute_hotkey register failed vk=" + std::to_string(vk) +
                    " mods=" + std::to_string(mods) +
                    " err=" + std::to_string(GetLastError()));
        } else {
            LogInfo("mute_hotkey registered vk=" + std::to_string(vk) +
                    " mods=" + std::to_string(mods));
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
                    onHotkey_();
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
        thread_ = std::thread([this]() { ThreadMain(); });
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
        std::lock_guard<std::mutex> lock(mutex_);
        SetStatus(st, c, m, d);
    };
    if (settings_.sourceMode == SourceMode::AppSession) {
        return std::make_unique<AppSessionSource>(settings_, pushFn_, statusForward);
    }
    return std::make_unique<LoopbackSource>(settings_, pushFn_, statusForward);
}

std::vector<LoopbackDeviceInfo> AudioSourceManager::EnumerateLoopbackDevices() {
    std::vector<LoopbackDeviceInfo> out;
    ComInit com;
    if (FAILED(com.hr)) return out;
    ComPtr<IMMDeviceEnumerator> enumerator;
    if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&enumerator)))) return out;

    std::wstring defaultId;
    ComPtr<IMMDevice> def;
    if (SUCCEEDED(enumerator->GetDefaultAudioEndpoint(eRender, eMultimedia, &def)) && def) {
        LPWSTR id = nullptr;
        if (SUCCEEDED(def->GetId(&id)) && id) {
            defaultId = id;
            CoTaskMemFree(id);
        }
    }

    ComPtr<IMMDeviceCollection> coll;
    if (FAILED(enumerator->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &coll)) || !coll) return out;
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
        std::string display = "Device";
        if (SUCCEEDED(store->GetValue(PKEY_Device_FriendlyName, &name)) && name.vt == VT_LPWSTR && name.pwszVal) {
            display = WideToUtf8(name.pwszVal);
        }
        PropVariantClear(&name);
        LoopbackDeviceInfo info;
        info.id = WideToUtf8(id);
        info.name = display;
        info.isDefault = (_wcsicmp(id, defaultId.c_str()) == 0);
        out.push_back(std::move(info));
        CoTaskMemFree(id);
    }
    return out;
}

std::vector<CaptureDeviceInfo> AudioSourceManager::EnumerateCaptureDevices() {
    std::vector<CaptureDeviceInfo> out;
    ComInit com;
    if (FAILED(com.hr)) return out;
    ComPtr<IMMDeviceEnumerator> enumerator;
    if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&enumerator)))) return out;

    std::wstring defaultId;
    ComPtr<IMMDevice> def;
    if (SUCCEEDED(enumerator->GetDefaultAudioEndpoint(eCapture, eCommunications, &def)) && def) {
        LPWSTR id = nullptr;
        if (SUCCEEDED(def->GetId(&id)) && id) {
            defaultId = id;
            CoTaskMemFree(id);
        }
    }

    ComPtr<IMMDeviceCollection> coll;
    if (FAILED(enumerator->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &coll)) || !coll) return out;
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
        std::string display = "Microphone";
        if (SUCCEEDED(store->GetValue(PKEY_Device_FriendlyName, &name)) && name.vt == VT_LPWSTR && name.pwszVal) {
            display = WideToUtf8(name.pwszVal);
        }
        PropVariantClear(&name);
        CaptureDeviceInfo info;
        info.id = WideToUtf8(id);
        info.name = display;
        info.isDefault = (_wcsicmp(id, defaultId.c_str()) == 0);
        out.push_back(std::move(info));
        CoTaskMemFree(id);
    }
    return out;
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

bool MicMixApp::Initialize(const std::string& configBasePath) {
    if (initialized_.exchange(true)) return true;
    configStore_ = std::make_unique<ConfigStore>(configBasePath);
    g_logPath = configStore_->LogPath();
    std::string warn;
    configStore_->Load(settings_, warn);
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
            std::string line = "source=" + SourceStateToString(st) + " code=" + code + " msg=" + msg;
            if (!detail.empty()) line += " detail=" + detail;
            LogInfo(line);
        });
    sourceManager_->ApplySettings(settings_);
    if (settings_.autostartEnabled) sourceManager_->Start();
    LogInfo("MicMix initialized");
    return true;
}

void MicMixApp::Shutdown() {
    if (!initialized_.exchange(false)) return;
    SettingsWindowController::Instance().Close();
    if (hotkeyManager_) hotkeyManager_->Stop();
    if (sourceManager_) sourceManager_->Stop();
    if (configStore_) {
        std::string err;
        configStore_->Save(settings_, err);
        if (!err.empty()) LogWarn("Config save warning: " + err);
    }
    hotkeyManager_.reset();
    sourceManager_.reset();
    configStore_.reset();
}

MicMixSettings MicMixApp::GetSettings() const {
    std::lock_guard<std::mutex> lock(settingsMutex_);
    return settings_;
}

void MicMixApp::ApplySettings(const MicMixSettings& settings, bool restartSource, bool saveConfig) {
    {
        std::lock_guard<std::mutex> lock(settingsMutex_);
        settings_ = settings;
    }
    engine_.ApplySettings(settings);
    if (hotkeyManager_) {
        hotkeyManager_->ApplySettings(settings);
    }
    if (sourceManager_) {
        sourceManager_->ApplySettings(settings);
        if (restartSource && sourceManager_->IsRunning()) sourceManager_->Restart();
    }
    if (saveConfig && configStore_) {
        std::string err;
        if (!configStore_->Save(settings, err) && !err.empty()) LogError("Config save failed: " + err);
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

void MicMixApp::SetPushToPlayActive(bool active) {
    pushToPlayActive_.store(active, std::memory_order_release);
    engine_.SetMuted(!active);
    auto s = GetSettings();
    s.musicMuted = !active;
    ApplySettings(s, false, true);
}

void MicMixApp::TogglePushToPlay() { SetPushToPlayActive(!pushToPlayActive_.load(std::memory_order_acquire)); }

void MicMixApp::SetTalkStateForOwnClient(uint64 schid, anyID clientId, int talkStatus) {
    (void)schid;
    (void)clientId;
    (void)talkStatus;
}

void MicMixApp::SetActiveServer(uint64 schid) {
    activeSchid_.store(schid, std::memory_order_release);
}
std::vector<LoopbackDeviceInfo> MicMixApp::GetLoopbackDevices() const { return AudioSourceManager::EnumerateLoopbackDevices(); }
std::vector<CaptureDeviceInfo> MicMixApp::GetCaptureDevices() const { return AudioSourceManager::EnumerateCaptureDevices(); }
std::vector<AppProcessInfo> MicMixApp::GetAppProcesses() const {
    auto apps = AudioSourceManager::EnumerateAppProcesses();
    std::string line = "app_list count=" + std::to_string(apps.size());
    const size_t limit = std::min<size_t>(apps.size(), 8);
    for (size_t i = 0; i < limit; ++i) {
        line += " | " + apps[i].exeName + ":" + std::to_string(apps[i].pid);
    }
    LogInfo(line);
    return apps;
}
SourceStatus MicMixApp::GetSourceStatus() const { return sourceManager_ ? sourceManager_->GetStatus() : SourceStatus{}; }
TelemetrySnapshot MicMixApp::GetTelemetry() const { return engine_.SnapshotTelemetry(); }
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
void MicMixApp::RefreshAppList() {}

void MicMixApp::EditCapturedVoice(uint64 schid, short* samples, int sampleCount, int channels, int* edited) {
    if (!initialized_.load(std::memory_order_acquire)) return;
    (void)schid;
    engine_.EditCapturedVoice(samples, sampleCount, channels, edited);
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
