#pragma once

#include <Windows.h>
#include <mmdeviceapi.h>
#include <string>
#include <vector>
#include <atomic>
#include <memory>
#include <mutex>
#include <thread>
#include <functional>

#include "teamspeak/public_definitions.h"
#include "teamspeak/public_errors.h"
#include "ts3_functions.h"

enum class SourceMode {
    Loopback = 0,
    AppSession = 1,
};

enum class SourceState {
    Stopped = 0,
    Starting = 1,
    Running = 2,
    Reacquiring = 3,
    Error = 4,
};

struct MicMixSettings {
    int         configVersion = 1;
    SourceMode  sourceMode = SourceMode::Loopback;
    std::string loopbackDeviceId;
    std::string appProcessName;
    std::string appSessionId;
    bool        autostartEnabled = false;
    float       musicGainDb = -15.0f;
    bool        forceTxEnabled = true;
    int         bufferTargetMs = 60;
    bool        musicMuted = false;
    int         uiLastOpenTab = 0;
    bool        autoSwitchToLoopback = false;
    int         muteHotkeyModifiers = 0;
    int         muteHotkeyVk = 0;
    std::string captureDeviceId;
};

struct LoopbackDeviceInfo {
    std::string id;
    std::string name;
    bool        isDefault = false;
};

struct AppProcessInfo {
    std::string displayName;
    uint32_t    pid = 0;
    std::string exeName;
};

struct CaptureDeviceInfo {
    std::string id;
    std::string name;
    bool        isDefault = false;
};

struct SourceStatus {
    SourceState  state = SourceState::Stopped;
    std::string  code;
    std::string  message;
    std::string  detail;
    uint64_t reconnectCount = 0;
};

struct TelemetrySnapshot {
    uint64_t underruns = 0;
    uint64_t overruns = 0;
    uint64_t clippedSamples = 0;
    uint64_t reconnects = 0;
    float musicRmsDbfs = -120.0f;
    float musicPeakDbfs = -120.0f;
    float musicSendPeakDbfs = -120.0f;
    bool  musicActive = false;
    bool  talkStateActive = false;
    bool  micTalkDetected = false;
    float micRmsDbfs = -120.0f;
};

extern TS3Functions g_ts3Functions;
extern std::string  g_pluginId;

void LogInfo(const std::string& text, uint64 schid = 0);
void LogWarn(const std::string& text, uint64 schid = 0);
void LogError(const std::string& text, uint64 schid = 0);

class ConfigStore {
public:
    explicit ConfigStore(std::string basePath);

    bool Load(MicMixSettings& outSettings, std::string& warning);
    bool Save(const MicMixSettings& settings, std::string& error);
    const std::string& ConfigPath() const { return configPath_; }
    const std::string& LogPath() const { return logPath_; }

private:
    std::string basePath_;
    std::string configPath_;
    std::string legacyConfigPath_;
    std::string tmpPath_;
    std::string lastGoodPath_;
    std::string logPath_;

    static std::string Trim(const std::string& value);
    static bool ParseBool(const std::string& value, bool& out);
    static std::string BoolToString(bool value);
    static std::string SourceModeToString(SourceMode mode);
    static SourceMode SourceModeFromString(const std::string& value);
};

template <typename T>
class SpscRingBuffer {
public:
    explicit SpscRingBuffer(size_t capacity)
        : capacity_(RoundUpPow2(capacity)), mask_(capacity_ - 1), buffer_(capacity_) {}

    size_t Capacity() const { return capacity_; }

    size_t Size() const {
        const size_t write = writeIndex_.load(std::memory_order_acquire);
        const size_t read = readIndex_.load(std::memory_order_acquire);
        return write - read;
    }

    size_t Write(const T* data, size_t count) {
        size_t write = writeIndex_.load(std::memory_order_relaxed);
        size_t read = readIndex_.load(std::memory_order_acquire);
        size_t freeCount = capacity_ - (write - read);
        size_t toWrite = count < freeCount ? count : freeCount;
        for (size_t i = 0; i < toWrite; ++i) {
            buffer_[(write + i) & mask_] = data[i];
        }
        writeIndex_.store(write + toWrite, std::memory_order_release);
        return toWrite;
    }

    size_t Read(T* out, size_t count) {
        size_t read = readIndex_.load(std::memory_order_relaxed);
        size_t write = writeIndex_.load(std::memory_order_acquire);
        size_t avail = write - read;
        size_t toRead = count < avail ? count : avail;
        for (size_t i = 0; i < toRead; ++i) {
            out[i] = buffer_[(read + i) & mask_];
        }
        readIndex_.store(read + toRead, std::memory_order_release);
        return toRead;
    }

    void Reset() {
        // Keep monotonic counters and just drop buffered data. This avoids
        // unsigned underflow races if producer/consumer are active concurrently.
        const size_t write = writeIndex_.load(std::memory_order_acquire);
        readIndex_.store(write, std::memory_order_release);
    }

private:
    static size_t RoundUpPow2(size_t value) {
        size_t out = 1;
        while (out < value) {
            out <<= 1;
        }
        return out;
    }

    size_t               capacity_;
    size_t               mask_;
    std::vector<T>       buffer_;
    std::atomic<size_t>  readIndex_{0};
    std::atomic<size_t>  writeIndex_{0};
};

class AudioEngine {
public:
    AudioEngine();

    void ApplySettings(const MicMixSettings& settings);
    void SetMusicSourceRunning(bool running);
    void SetTalkState(bool talking);
    void SetExternalMicLevel(float linear);
    void ToggleMute();
    void SetMuted(bool muted);
    bool IsMuted() const { return musicMuted_.load(std::memory_order_relaxed); }

    void PushMusicSamples(const float* samples, size_t count);
    void ClearMusicBuffer();
    void EditCapturedVoice(short* samples, int sampleCount, int channels, int* edited);

    TelemetrySnapshot SnapshotTelemetry() const;

private:
    static constexpr int kTargetRate = 48000;
    static constexpr size_t kCallbackScratch = 4096;

    SpscRingBuffer<float> ring_;

    std::atomic<bool>  musicMuted_{false};
    std::atomic<bool>  sourceRunning_{false};
    std::atomic<bool>  forceTxEnabled_{true};
    std::atomic<float> musicGainLinear_{0.5f};
    std::atomic<int>   bufferTargetMs_{60};

    std::atomic<bool>  talkState_{true};
    std::atomic<bool>  micTalkDetected_{false};
    std::atomic<float> micRmsDbfs_{-120.0f};
    std::atomic<float> externalMicLinear_{0.0f};

    std::atomic_uint64_t underruns_{0};
    std::atomic_uint64_t overruns_{0};
    std::atomic_uint64_t clippedSamples_{0};
    std::atomic_uint64_t reconnectsMirror_{0};
    std::atomic_uint64_t lastConsumeTickMs_{0};
    std::atomic_uint64_t lastMusicSignalTickMs_{0};
    std::atomic_uint64_t lastMicTalkTickMs_{0};
    std::atomic<float> musicRmsDbfs_{-120.0f};
    std::atomic<float> musicPeakDbfs_{-120.0f};
    std::atomic<float> musicSendPeakDbfs_{-120.0f};

    static float DbToLinear(float db);
};

class IAudioSource {
public:
    virtual ~IAudioSource() = default;
    virtual bool Start() = 0;
    virtual void Stop() = 0;
};

class AudioSourceManager {
public:
    using AudioPushFn = std::function<void(const float*, size_t)>;
    using StatusFn    = std::function<void(SourceState, const std::string&, const std::string&, const std::string&)>;

    AudioSourceManager(AudioPushFn pushFn, StatusFn statusFn);
    ~AudioSourceManager();

    void ApplySettings(const MicMixSettings& settings);
    bool Start();
    void Stop();
    void Restart();
    bool IsRunning() const;
    SourceStatus GetStatus() const;

    static std::vector<LoopbackDeviceInfo> EnumerateLoopbackDevices();
    static std::vector<AppProcessInfo> EnumerateAppProcesses(const std::string& processName = "");
    static std::vector<CaptureDeviceInfo> EnumerateCaptureDevices();

private:
    mutable std::mutex           mutex_;
    MicMixSettings               settings_;
    std::unique_ptr<IAudioSource> source_;
    AudioPushFn                  pushFn_;
    StatusFn                     statusFn_;
    SourceStatus                 status_;
    bool                         running_ = false;

    void SetStatus(SourceState state, const std::string& code, const std::string& message, const std::string& detail);
    std::unique_ptr<IAudioSource> CreateSourceLocked();
};

class GlobalHotkeyManager;
class MicLevelMonitor;
class MixMonitorPlayer;

class MicMixApp {
public:
    static MicMixApp& Instance();

    bool Initialize(const std::string& configBasePath);
    void Shutdown();

    MicMixSettings GetSettings() const;
    void ApplySettings(const MicMixSettings& settings, bool restartSource, bool saveConfig);
    void StartSource();
    void StopSource();
    void ToggleMute();
    void SetMonitorEnabled(bool enabled);
    bool IsMonitorEnabled() const;
    void ToggleMonitor();
    void SetPushToPlayActive(bool active);
    void TogglePushToPlay();
    void SetTalkStateForOwnClient(uint64 schid, anyID clientId, int talkStatus);
    void SetActiveServer(uint64 schid);
    void OnConnectStatusChange(uint64 schid, int newStatus, unsigned int errorNumber);

    std::vector<LoopbackDeviceInfo> GetLoopbackDevices() const;
    std::vector<CaptureDeviceInfo> GetCaptureDevices() const;
    std::vector<AppProcessInfo> GetAppProcesses() const;
    SourceStatus GetSourceStatus() const;
    TelemetrySnapshot GetTelemetry() const;
    std::string GetConfigDir() const;
    std::string GetPreferredTsCaptureDeviceName() const;

    void EditCapturedVoice(uint64 schid, short* samples, int sampleCount, int channels, int* edited);
    bool IsInitialized() const { return initialized_.load(std::memory_order_relaxed); }

    void OpenSettingsWindow();
    void RefreshAppList();

private:
    MicMixApp();
    ~MicMixApp();
    void StartVoiceTxThread();
    void StopVoiceTxThread();
    void VoiceTxThreadMain();
    void SetVoiceRecordingState(bool active, uint64 schid);
    void RefreshVoiceTxControl(uint64 schidHint);
    void NudgeCapturePath(uint64 schid);

    std::atomic<bool> initialized_{false};
    std::unique_ptr<ConfigStore> configStore_;
    MicMixSettings settings_;
    mutable std::mutex settingsMutex_;
    AudioEngine engine_;
    std::unique_ptr<AudioSourceManager> sourceManager_;
    std::unique_ptr<GlobalHotkeyManager> hotkeyManager_;
    std::unique_ptr<MicLevelMonitor> micLevelMonitor_;
    std::unique_ptr<MixMonitorPlayer> mixMonitorPlayer_;
    std::atomic<uint64> activeSchid_{0};
    std::atomic<bool> pushToPlayActive_{false};
    std::thread voiceTxThread_;
    std::atomic<bool> voiceTxStop_{false};
    std::atomic<bool> voiceTxThreadRunning_{false};
    std::atomic<bool> voiceRecordingActive_{false};
    std::atomic<uint64> voiceRecordingSchid_{0};
    std::atomic_uint64_t voiceTxLastNudgeMs_{0};
    std::atomic_uint64_t voiceControlLastEvalMs_{0};
    std::atomic_uint64_t lastCaptureEditTickMs_{0};
    std::atomic_uint64_t lastCaptureReopenTickMs_{0};
    std::mutex voiceTxMutex_;
    bool savedInputStateValid_ = false;
    int savedInputDeactivated_ = INPUT_DEACTIVATED;
    bool savedVadValid_ = false;
    bool savedVadEnabled_ = false;
    bool savedVadThresholdValid_ = false;
    float savedVadThresholdDbfs_ = -50.0f;
};

std::string SourceStateToString(SourceState state);
