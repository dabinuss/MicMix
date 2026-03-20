#pragma once

#include <Windows.h>
#include <mmdeviceapi.h>
#include <array>
#include <string>
#include <vector>
#include <atomic>
#include <memory>
#include <mutex>
#include <thread>
#include <functional>
#include <cstdint>

#include "vst_audio_ipc.h"
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

inline constexpr bool IsSourceStateActive(SourceState state) noexcept {
    return state == SourceState::Running ||
           state == SourceState::Starting ||
           state == SourceState::Reacquiring;
}

enum class MicGateMode {
    AutoTs = 0,
    Custom = 1,
};

enum class EffectChain {
    Music = 0,
    Mic = 1,
};

struct VstEffectSlot {
    std::string path;
    std::string name;
    std::string uid;
    std::string stateBlob;
    std::string lastStatus;
    bool        enabled = true;
    bool        bypass = false;
};

struct VstHostStatus {
    bool        running = false;
    uint32_t    pid = 0;
    std::string message;
};

struct MicMixSettings {
    int         configVersion = 1;
    SourceMode  sourceMode = SourceMode::Loopback;
    std::string loopbackDeviceId;
    std::string appProcessName;
    std::string appSessionId;
    bool        autostartEnabled = false;
    float       musicGainDb = -15.0f;
    int         resamplerQuality = -1;
    bool        forceTxEnabled = true;
    int         bufferTargetMs = 60;
    bool        musicMuted = false;
    bool        micInputMuted = false;
    int         uiLastOpenTab = 0;
    bool        autoSwitchToLoopback = false;
    int         muteHotkeyModifiers = 0;
    int         muteHotkeyVk = 0;
    int         micInputMuteHotkeyModifiers = 0;
    int         micInputMuteHotkeyVk = 0;
    std::string captureDeviceId;
    MicGateMode micGateMode = MicGateMode::AutoTs;
    float       micGateThresholdDbfs = -50.0f;
    bool        micUseSmoothGate = true;
    bool        micUseKeyboardGuard = true;
    bool        micForceTsFilters = true;
    bool        vstEffectsEnabled = false;
    bool        vstHostAutostart = true;
    std::vector<VstEffectSlot> musicEffects;
    std::vector<VstEffectSlot> micEffects;
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
    uint64_t sourceClipEvents = 0;
    uint64_t micClipEvents = 0;
    uint64_t reconnects = 0;
    float musicRmsDbfs = -120.0f;
    float musicPeakDbfs = -120.0f;
    float musicSendPeakDbfs = -120.0f;
    bool  musicActive = false;
    bool  sourceClipRecent = false;
    bool  talkStateActive = false;
    bool  micTalkDetected = false;
    bool  micClipRecent = false;
    float micRmsDbfs = -120.0f;
    float micPeakDbfs = -120.0f;
};

extern TS3Functions g_ts3Functions;
extern std::string g_pluginId;

void SetTsLoggingEnabled(bool enabled);
void LogInfo(const std::string& text, uint64 schid = 0);
void LogWarn(const std::string& text, uint64 schid = 0);
void LogError(const std::string& text, uint64 schid = 0);

class ConfigStore {
public:
    explicit ConfigStore(std::string basePath);

    bool Load(MicMixSettings& outSettings, std::string& warning);
    bool Save(const MicMixSettings& settings, std::string& error);
    const std::string& ConfigPath() const { return configPath_; }
    const std::string& ConfigDir() const { return configDir_; }
    const std::string& LogPath() const { return logPath_; }

private:
    std::string basePath_;
    std::string configDir_;
    std::string configPath_;
    std::string legacyConfigPath_;
    std::string tmpPath_;
    std::string lastGoodPath_;
    std::string logPath_;

    static std::string Trim(const std::string& value);
    static std::string BoolToString(bool value);
    static std::string SourceModeToString(SourceMode mode);
    static SourceMode SourceModeFromString(const std::string& value);
    static std::string MicGateModeToString(MicGateMode mode);
    static MicGateMode MicGateModeFromString(const std::string& value);
};

template <typename T>
// Lock-free ring buffer for exactly one producer thread and one consumer thread.
// Memory ordering guarantees:
// - producer publishes writes via writeIndex_ (release),
// - consumer observes published data via writeIndex_ (acquire),
// - consumer publishes reads via readIndex_ (release),
// - producer observes consumed data via readIndex_ (acquire).
// Using additional producers/consumers is unsupported and will race.
class SpscRingBuffer {
public:
    explicit SpscRingBuffer(size_t capacity)
        : capacity_(RoundUpPow2(capacity)), mask_(capacity_ - 1), buffer_(capacity_) {}

    size_t Capacity() const { return capacity_; }

    size_t Size() const {
        const size_t write = writeIndex_.load(std::memory_order_acquire);
        const size_t read = readIndex_.load(std::memory_order_relaxed);
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
    static constexpr size_t kIndexPaddingBytes = 64;

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
    std::array<unsigned char, kIndexPaddingBytes> indexPadding_{};
    std::atomic<size_t>  writeIndex_{0};
};

// Threading model:
// - PushMusicSamples: source/capture thread.
// - EditCapturedVoice: TeamSpeak real-time capture callback thread.
// - ApplySettings/Set* control methods: UI/hotkey/control threads.
//
// Real-time safety:
// - EditCapturedVoice is designed for RT use (no locks, no allocations on hot path).
// - PushMusicSamples is lock-free but may touch telemetry atomics and ring buffer.
//
// Contract of EditCapturedVoice:
// - Mixes buffered music into captured PCM16 voice frames in-place.
// - Preserves upstream edited flags unless it actually modifies outgoing audio.
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
    void ToggleMicInputMute();
    void SetMicInputMuted(bool muted);
    bool IsMicInputMuted() const { return micInputMuted_.load(std::memory_order_relaxed); }

    void PushMusicSamples(const float* samples, size_t count);
    void ClearMusicBuffer();
    void EditCapturedVoice(short* samples, int sampleCount, int channels, int* edited);
    void NoteReconnect();

    TelemetrySnapshot SnapshotTelemetry() const;

private:
    static constexpr int kTargetRate = 48000;
    static constexpr size_t kCallbackScratch = 4096;

    SpscRingBuffer<float> ring_;

    std::atomic<bool>  musicMuted_{false};
    std::atomic<bool>  micInputMuted_{false};
    std::atomic<bool>  sourceRunning_{false};
    std::atomic<bool>  forceTxEnabled_{true};
    std::atomic<float> musicGainLinear_{0.5f};
    std::atomic<int>   bufferTargetMs_{60};

    std::atomic<bool>  talkState_{true};
    std::atomic<bool>  micTalkDetected_{false};
    std::atomic<bool>  micTailTxActive_{false};
    std::atomic<float> micRmsDbfs_{-120.0f};
    std::atomic<float> micPeakDbfs_{-120.0f};
    std::atomic<float> externalMicLinear_{0.0f};
    std::atomic<float> micGateGain_{1.0f};
    std::atomic<float> limiterGain_{1.0f};
    std::atomic<bool>  micUseSmoothGate_{true};

    std::atomic_uint64_t underruns_{0};
    std::atomic_uint64_t overruns_{0};
    std::atomic_uint64_t clippedSamples_{0};
    std::atomic_uint64_t sourceClipEvents_{0};
    std::atomic_uint64_t micClipEvents_{0};
    std::atomic_uint64_t reconnectsMirror_{0};
    std::atomic_uint64_t lastConsumeTickMs_{0};
    std::atomic_uint64_t lastMusicSignalTickMs_{0};
    std::atomic_uint64_t lastSourceClipTickMs_{0};
    std::atomic_uint64_t lastMicClipTickMs_{0};
    std::atomic_uint64_t lastMicTalkTickMs_{0};
    std::atomic<bool> sourceClipState_{false};
    std::atomic<bool> micClipState_{false};
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
    void ToggleMicInputMute();
    void SetMonitorEnabled(bool enabled);
    bool IsMonitorEnabled() const;
    void ToggleMonitor();
    void SetGlobalHotkeyCaptureBlocked(bool blocked);
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
    VstHostStatus GetVstHostStatus() const;
    bool IsEffectsEnabled() const;
    void SetEffectsEnabled(bool enabled, bool saveConfig = true);
    std::vector<VstEffectSlot> GetEffects(EffectChain chain) const;
    bool AddEffect(EffectChain chain, const VstEffectSlot& slot, std::string& error);
    bool RemoveEffect(EffectChain chain, size_t index, std::string& error);
    bool MoveEffect(EffectChain chain, size_t fromIndex, size_t toIndex, std::string& error);
    bool SetEffectBypass(EffectChain chain, size_t index, bool bypass, std::string& error);
    bool SetEffectEnabled(EffectChain chain, size_t index, bool enabled, std::string& error);
    bool OpenEffectEditor(EffectChain chain, size_t index, std::string& error);
    static void SanitizeEffectList(std::vector<VstEffectSlot>& list);

    void EditCapturedVoice(uint64 schid, short* samples, int sampleCount, int channels, int* edited);
    bool IsInitialized() const { return initialized_.load(std::memory_order_relaxed); }

    void OpenSettingsWindow();
    void OpenEffectsWindow();

private:
    static constexpr size_t kSavedPreprocessorSlots = 18;

    MicMixApp();
    ~MicMixApp();
    void StartVoiceTxThread();
    void StopVoiceTxThread();
    void VoiceTxThreadMain();
    void SetVoiceRecordingState(bool active, uint64 schid);
    void RefreshVoiceTxControl(uint64 schidHint);
    void SyncMusicActivityMeta(uint64 schid, bool musicActive, bool force);
    void ApplyMicInputTransportMute(bool muted);
    void OnSourceSamples(const float* data, size_t count);
    void ProcessMicInputWithVst(short* samples, int sampleCount, int channels);

    std::atomic<bool> initialized_{false};
    std::unique_ptr<ConfigStore> configStore_;
    MicMixSettings settings_;
    mutable std::mutex settingsMutex_;
    AudioEngine engine_;
    std::atomic<std::shared_ptr<AudioSourceManager>> sourceManager_{};
    std::unique_ptr<GlobalHotkeyManager> musicMuteHotkeyManager_;
    std::unique_ptr<GlobalHotkeyManager> micInputMuteHotkeyManager_;
    std::unique_ptr<MicLevelMonitor> micLevelMonitor_;
    std::atomic<std::shared_ptr<MixMonitorPlayer>> mixMonitorPlayer_{};
    std::atomic<uint64> activeSchid_{0};
    std::atomic<bool> ownTalkStatusActive_{false};
    std::atomic_uint64_t ownTalkStatusTickMs_{0};
    std::thread voiceTxThread_;
    std::atomic<bool> voiceTxStop_{false};
    std::atomic<bool> voiceTxThreadRunning_{false};
    std::atomic<bool> voiceRecordingActive_{false};
    std::atomic<uint64> voiceRecordingSchid_{0};
    std::atomic_uint64_t voiceTxLastNudgeMs_{0};
    std::atomic_uint64_t voiceControlLastEvalMs_{0};
    std::atomic_uint64_t forceTxHoldUntilMs_{0};
    std::atomic_uint64_t lastCaptureEditTickMs_{0};
    std::atomic_uint64_t lastCaptureReopenTickMs_{0};
    std::atomic<bool> shutdownRequested_{false};
    std::atomic<uint32_t> captureCallbacksInFlight_{0};
    std::mutex voiceTxMutex_;
    std::mutex musicMetaMutex_;
    uint64 musicMetaSchid_ = 0;
    bool musicMetaLastStateValid_ = false;
    bool musicMetaLastState_ = false;
    uint64_t musicMetaLastAttemptMs_ = 0;
    bool savedInputStateValid_ = false;
    int savedInputDeactivated_ = INPUT_DEACTIVATED;
    bool savedVadValid_ = false;
    bool savedVadEnabled_ = false;
    bool savedVadThresholdValid_ = false;
    float savedVadThresholdDbfs_ = -50.0f;
    bool savedVadExtraBufferValid_ = false;
    int savedVadExtraBufferSize_ = 0;
    std::array<std::string, kSavedPreprocessorSlots> savedPreprocessorValues_{};
    std::array<bool, kSavedPreprocessorSlots> savedPreprocessorValuesValid_{};
    bool micInputTransportMuteActive_ = false;
    uint64 micInputTransportMuteSchid_ = 0;
    bool micInputTransportSavedValid_ = false;
    int micInputTransportSavedState_ = INPUT_DEACTIVATED;
    std::atomic<bool> vstHostRunning_{false};
    std::atomic<bool> vstEffectsEnabledCached_{false};
    std::atomic<bool> vstHostSyncPending_{false};
    std::atomic<bool> vstHostStopPending_{false};
    std::atomic<uint32_t> vstHostPid_{0};
    mutable std::mutex vstHostMutex_;
    mutable std::mutex vstHostIpcMutex_;
    std::recursive_mutex vstHostLifecycleMutex_;
    HANDLE vstHostProcess_ = nullptr;
    HANDLE vstHostThread_ = nullptr;
    HANDLE vstHostJob_ = nullptr;
    HANDLE vstAudioMap_ = nullptr;
    std::atomic<micmix::vstipc::SharedMemory*> vstAudioShared_{nullptr};
    std::wstring vstHostPipeName_;
    std::string vstHostAuthToken_;
    std::string vstHostMessage_;
    std::atomic<uint32_t> vstMusicSeq_{1};
    std::atomic<uint32_t> vstMicSeq_{1};

    bool TryEnterCaptureCallback();
    void LeaveCaptureCallback();
    bool WaitForCaptureCallbacksToDrain(uint32_t maxWaitMs) const;
    bool StartVstHostProcess(std::string& error);
    void StopVstHostProcess();
    std::wstring ResolveVstHostPath() const;
    std::wstring BuildVstHostPipePath() const;
    bool SendVstHostCommand(const std::string& command, std::string& response, DWORD timeoutMs, std::string& error) const;
    bool SyncVstHostState(const MicMixSettings& settings, std::string& error);
    bool PingVstHost(std::string& response, std::string& error) const;
    bool EnsureVstAudioIpc();
    void CloseVstAudioIpc();
    void MaintainVstHost(uint64_t nowMs);

    std::atomic_uint64_t vstHostNextRestartTickMs_{0};
    std::atomic<uint32_t> vstHostRestartAttempts_{0};
    std::atomic_uint64_t vstHostLastRestartLogTickMs_{0};
    std::atomic_uint64_t vstHostLastHeartbeatTickMs_{0};
    std::atomic_uint64_t vstHostEditorKeepAliveUntilMs_{0};
};

std::string SourceStateToString(SourceState state);
