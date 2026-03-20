#include "settings_window.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <CommCtrl.h>
#include <shellapi.h>
#include <uxtheme.h>

#include <atomic>
#include <algorithm>
#include <cstdio>
#include <cwctype>
#include <exception>
#include <limits>
#include <mutex>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <string>

#include "micmix_core.h"
#include "micmix_version.h"
#include "resource.h"
#include "ui_shared.h"

#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "UxTheme.lib")

namespace {

enum ControlId {
    IDC_TITLE = 1000,
    IDC_SUBTITLE = 1001,
    IDC_SOURCE = 1002,
    IDC_RESCAN = 1005,
    IDC_GAIN = 1006,
    IDC_GAIN_VALUE = 1007,
    IDC_FORCE_TX = 1008,
    IDC_MIC_INPUT_MUTE = 1009,
    IDC_MUTE = 1010,
    IDC_AUTOSTART = 1011,
    IDC_START = 1012,
    IDC_STOP = 1013,
    IDC_SAVE = 1014,
    IDC_STATUS = 1015,
    IDC_METER_TEXT = 1019,
    IDC_MUTE_HOTKEY_SET = 1020,
    IDC_MUTE_HOTKEY_TEXT = 1021,
    IDC_MIC_INPUT_HOTKEY_SET = 1022,
    IDC_MIC_INPUT_HOTKEY_TEXT = 1023,
    IDC_MIC_DEVICE = 1024,
    IDC_MIC_METER_TEXT = 1025,
    IDC_MONITOR = 1026,
    IDC_MONITOR_HINT = 1027,
    IDC_MIC_METER_HINT = 1028,
    IDC_GAIN_HINT = 1029,
    IDC_FORCE_TX_HINT = 1030,
    IDC_VERSION = 1031,
    IDC_REPO_LINK = 1032,
    IDC_MUSIC_CLIP_EVENTS = 1033,
    IDC_MIC_CLIP_EVENTS = 1034,
};

enum class SourceChoiceType {
    Loopback,
    App,
};

enum class HotkeyCaptureTarget {
    None,
    MusicMute,
    MicInputMute,
};

enum class HeaderStatusBadgeState {
    Active,
    Muted,
    Off,
    Error,
};

struct SourceChoice {
    SourceChoiceType type = SourceChoiceType::Loopback;
    std::string id;
    std::string processName;
    uint32_t pid = 0;
};

std::mutex g_mutex;
std::mutex g_enumMutex;
std::mutex g_sourceRefreshThreadMutex;
std::thread g_thread;
std::thread g_sourceRefreshThread;
std::atomic<bool> g_running{false};
std::atomic<HWND> g_hwnd{nullptr};
std::vector<LoopbackDeviceInfo> g_loopbacks;
std::vector<CaptureDeviceInfo> g_captureDevices;
std::vector<AppProcessInfo> g_apps;
std::vector<SourceChoice> g_sourceChoices;
int g_lastValidSourceSel = -1;
UiTheme g_theme = DefaultUiTheme();
int g_dpi = 96;
bool g_loadingUi = false;
int g_hotkeyRefreshTick = 0;
HotkeyCaptureTarget g_hotkeyCaptureTarget = HotkeyCaptureTarget::None;
UINT g_muteHotkeyModifiers = 0;
UINT g_muteHotkeyVk = 0;
UINT g_micInputMuteHotkeyModifiers = 0;
UINT g_micInputMuteHotkeyVk = 0;
std::vector<LoopbackDeviceInfo> g_pendingLoopbacks;
std::vector<CaptureDeviceInfo> g_pendingCaptureDevices;
std::vector<AppProcessInfo> g_pendingApps;
std::atomic_uint64_t g_sourceRefreshSeq{0};
std::atomic_uint64_t g_sourceRefreshActiveSeq{0};
std::atomic<bool> g_sourceRefreshInFlight{false};
std::atomic<bool> g_sourceRefreshPending{false};
std::atomic<bool> g_sourceRefreshPendingReload{false};
std::atomic<bool> g_saveDebouncePending{false};
std::unordered_map<uint32_t, HICON> g_appIconsByPid;
HICON g_loopbackFallbackIcon = nullptr;
HICON g_appFallbackIcon = nullptr;
bool g_ownerAutostartChecked = false;
bool g_ownerForceTxChecked = false;
bool g_ownerMicInputMuteChecked = false;
bool g_ownerMuteChecked = false;
HeaderStatusBadgeState g_headerStatusBadgeState = HeaderStatusBadgeState::Off;
bool g_windowTitleActive = false;
bool g_windowTitleStateValid = false;

constexpr INT_PTR kSourceItemHeader = -10;
constexpr INT_PTR kSourceItemDivider = -11;
constexpr INT_PTR kSourceItemPlaceholder = -12;
constexpr UINT kMsgSourceRefreshDone = WM_APP + 0x31;
constexpr UINT kTimerStatusUpdate = 1;
constexpr UINT kTimerDebouncedSave = 2;
constexpr UINT kStatusTimerIntervalMs = 40;
constexpr int kStatusDetailsEveryTicks = 3;
constexpr int kHotkeyLabelRefreshEveryTicks = 6;
constexpr UINT kDebouncedSaveMs = 300;
constexpr float kMusicGainMinDb = -30.0f;
constexpr float kMusicGainMaxDb = -2.0f;
constexpr int kMusicGainStepPerDb = 10;
constexpr int kMusicGainSliderMax = static_cast<int>((kMusicGainMaxDb - kMusicGainMinDb) * static_cast<float>(kMusicGainStepPerDb));
constexpr int kSourceIconSizePx = 16;
constexpr int kClientWidthPx = kUiClientWidthPx;
constexpr int kClientHeightPx = kUiClientHeightPx;
constexpr int kCardMarginPx = kUiCardMarginPx;
constexpr int kCardGapPx = kUiCardGapPx;
constexpr int kCardInnerPaddingPx = kUiCardInnerPaddingPx;
constexpr wchar_t kRepoUrl[] = L"https://github.com/dabinuss/MicMix";
constexpr wchar_t kSettingsHeaderTitle[] = L"MicMix - Mixing";

HFONT g_fontBody = nullptr;
HFONT g_fontSmall = nullptr;
HFONT g_fontHint = nullptr;
HFONT g_fontTiny = nullptr;
HFONT g_fontClipInfo = nullptr;
HFONT g_fontControlLarge = nullptr;
HFONT g_fontTitle = nullptr;
HFONT g_fontMono = nullptr;
HBRUSH g_brushBg = nullptr;
HBRUSH g_brushCard = nullptr;
UiCommonResources g_commonUi{};

RECT g_rcSource{};
RECT g_rcMix{};
RECT g_rcSession{};
RECT g_rcHeaderBadge{};
RECT g_rcMeter{};
RECT g_rcMicMeter{};
RECT g_rcMusicClip{};
RECT g_rcMicClip{};
RECT g_rcStatus{};
float g_meterVisualDb = -60.0f;
float g_meterHoldDb = -60.0f;
float g_micMeterVisualDb = -60.0f;
float g_micMeterHoldDb = -60.0f;
int g_lastMusicMeterLevelPx = -1;
int g_lastMusicMeterHoldPx = -1;
bool g_lastMusicMeterActive = false;
int g_lastMicMeterLevelPx = -1;
int g_lastMicMeterHoldPx = -1;
bool g_lastMicMeterActive = false;
int g_lastMusicClipLitSegments = -1;
bool g_lastMusicClipRecent = false;
int g_lastMicClipLitSegments = -1;
bool g_lastMicClipRecent = false;
std::wstring g_lastMusicMeterText;
std::wstring g_lastMicMeterText;
TelemetrySnapshot g_lastTelemetry{};
uint64_t g_lastMusicClipEvents = std::numeric_limits<uint64_t>::max();
uint64_t g_lastMicClipEvents = std::numeric_limits<uint64_t>::max();
std::wstring g_lastStatusText;

int S(int px) {
    return MulDiv(px, g_dpi, 96);
}

bool IsOwnerCheckboxControlId(int id) {
    return id == IDC_AUTOSTART || id == IDC_FORCE_TX || id == IDC_MIC_INPUT_MUTE || id == IDC_MUTE;
}

bool IsHandCursorControlId(int id) {
    switch (id) {
    case IDC_REPO_LINK:
    case IDC_SOURCE:
    case IDC_MIC_DEVICE:
    case IDC_RESCAN:
    case IDC_START:
    case IDC_STOP:
    case IDC_SAVE:
    case IDC_MONITOR:
    case IDC_AUTOSTART:
    case IDC_FORCE_TX:
    case IDC_MIC_INPUT_MUTE:
    case IDC_MUTE:
    case IDC_MUTE_HOTKEY_SET:
    case IDC_MIC_INPUT_HOTKEY_SET:
        return true;
    default:
        return false;
    }
}

bool IsPointOverGainThumb(HWND gainHwnd, const POINT& screenPt) {
    if (!gainHwnd) {
        return false;
    }
    RECT thumbRc{};
    SendMessageW(gainHwnd, TBM_GETTHUMBRECT, 0, reinterpret_cast<LPARAM>(&thumbRc));
    POINT clientPt = screenPt;
    if (!ScreenToClient(gainHwnd, &clientPt)) {
        return false;
    }
    return PtInRect(&thumbRc, clientPt) != FALSE;
}

bool GetOwnerCheckboxValue(int id) {
    switch (id) {
    case IDC_AUTOSTART: return g_ownerAutostartChecked;
    case IDC_FORCE_TX: return g_ownerForceTxChecked;
    case IDC_MIC_INPUT_MUTE: return g_ownerMicInputMuteChecked;
    case IDC_MUTE: return g_ownerMuteChecked;
    default: return false;
    }
}

void SetOwnerCheckboxValue(int id, bool value) {
    switch (id) {
    case IDC_AUTOSTART: g_ownerAutostartChecked = value; break;
    case IDC_FORCE_TX: g_ownerForceTxChecked = value; break;
    case IDC_MIC_INPUT_MUTE: g_ownerMicInputMuteChecked = value; break;
    case IDC_MUTE: g_ownerMuteChecked = value; break;
    default: break;
    }
}

void DestroyIconHandle(HICON& icon) {
    if (icon) {
        DestroyIcon(icon);
        icon = nullptr;
    }
}

void ClearAppIconCache() {
    for (auto& entry : g_appIconsByPid) {
        DestroyIconHandle(entry.second);
    }
    g_appIconsByPid.clear();
}

HICON CopySharedIcon(HICON sharedIcon) {
    if (!sharedIcon) {
        return nullptr;
    }
    return CopyIcon(sharedIcon);
}

HICON LoadStockSmallIcon(SHSTOCKICONID stockId) {
    SHSTOCKICONINFO sii{};
    sii.cbSize = sizeof(sii);
    if (SUCCEEDED(SHGetStockIconInfo(stockId, SHGSI_ICON | SHGSI_SMALLICON, &sii)) && sii.hIcon) {
        return sii.hIcon;
    }
    return nullptr;
}

bool TryGetProcessImagePath(uint32_t pid, std::wstring& outPath) {
    outPath.clear();
    HANDLE process = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, pid);
    if (!process) {
        return false;
    }
    std::wstring buffer(MAX_PATH, L'\0');
    DWORD size = static_cast<DWORD>(buffer.size());
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
    CloseHandle(process);
    return true;
}

HICON TryLoadExeSmallIcon(const std::wstring& exePath) {
    if (exePath.empty()) {
        return nullptr;
    }
    SHFILEINFOW sfi{};
    if (SHGetFileInfoW(exePath.c_str(), 0, &sfi, sizeof(sfi), SHGFI_ICON | SHGFI_SMALLICON) == 0) {
        return nullptr;
    }
    return sfi.hIcon;
}

void EnsureFallbackIcons() {
    if (!g_loopbackFallbackIcon) {
        g_loopbackFallbackIcon = LoadStockSmallIcon(SIID_AUDIOFILES);
        if (!g_loopbackFallbackIcon) {
            g_loopbackFallbackIcon = CopySharedIcon(LoadIconW(nullptr, IDI_INFORMATION));
        }
    }
    if (!g_appFallbackIcon) {
        g_appFallbackIcon = LoadStockSmallIcon(SIID_APPLICATION);
        if (!g_appFallbackIcon) {
            g_appFallbackIcon = CopySharedIcon(LoadIconW(nullptr, IDI_APPLICATION));
        }
    }
}

void EnsureAppIconForPid(uint32_t pid) {
    if (pid == 0 || g_appIconsByPid.find(pid) != g_appIconsByPid.end()) {
        return;
    }
    std::wstring exePath;
    if (TryGetProcessImagePath(pid, exePath)) {
        if (HICON icon = TryLoadExeSmallIcon(exePath)) {
            g_appIconsByPid.emplace(pid, icon);
            return;
        }
    }
    if (g_appFallbackIcon) {
        if (HICON iconCopy = CopyIcon(g_appFallbackIcon)) {
            g_appIconsByPid.emplace(pid, iconCopy);
        }
    }
}

void PruneAppIconCache(const std::unordered_set<uint32_t>& activePids) {
    for (auto it = g_appIconsByPid.begin(); it != g_appIconsByPid.end();) {
        if (activePids.find(it->first) == activePids.end()) {
            DestroyIconHandle(it->second);
            it = g_appIconsByPid.erase(it);
        } else {
            ++it;
        }
    }
}

std::wstring ToLowerWide(std::wstring text) {
    std::transform(text.begin(), text.end(), text.begin(), [](wchar_t c) {
        return static_cast<wchar_t>(towlower(c));
    });
    return text;
}

std::wstring NormalizeDeviceName(std::wstring text) {
    text = ToLowerWide(std::move(text));
    const size_t defPos = text.find(L"(default)");
    if (defPos != std::wstring::npos) {
        text = text.substr(0, defPos);
    }
    while (!text.empty() && iswspace(text.back()) != 0) {
        text.pop_back();
    }
    while (!text.empty() && iswspace(text.front()) != 0) {
        text.erase(text.begin());
    }
    return text;
}

void EnsureUiResources() {
    EnsureCommonUiResources(g_commonUi, g_theme, g_dpi);
    g_brushBg = g_commonUi.bgBrush;
    g_brushCard = g_commonUi.cardBrush;
    g_fontBody = g_commonUi.bodyFont;
    g_fontSmall = g_commonUi.smallFont;
    g_fontHint = g_commonUi.hintFont;
    g_fontTiny = g_commonUi.tinyFont;
    g_fontControlLarge = g_commonUi.controlLargeFont;
    g_fontTitle = g_commonUi.titleFont;
    g_fontMono = g_commonUi.monoFont;
    if (!g_fontClipInfo) g_fontClipInfo = CreateFontW(-S(10), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    EnsureFallbackIcons();
}

void ReleaseUiResources() {
    ClearAppIconCache();
    DestroyIconHandle(g_loopbackFallbackIcon);
    DestroyIconHandle(g_appFallbackIcon);
    if (g_fontClipInfo) { DeleteObject(g_fontClipInfo); g_fontClipInfo = nullptr; }
    ReleaseCommonUiResources(g_commonUi);
    g_fontBody = nullptr;
    g_fontSmall = nullptr;
    g_fontHint = nullptr;
    g_fontTiny = nullptr;
    g_fontControlLarge = nullptr;
    g_fontTitle = nullptr;
    g_fontMono = nullptr;
    g_brushBg = nullptr;
    g_brushCard = nullptr;
}

void SetStatusText(HWND hwnd, const std::wstring& text) {
    if (text == g_lastStatusText) {
        return;
    }
    g_lastStatusText = text;
    if (hwnd && g_rcStatus.right > g_rcStatus.left && g_rcStatus.bottom > g_rcStatus.top) {
        InvalidateRect(hwnd, &g_rcStatus, FALSE);
    }
}

void SetControlEnabled(HWND hwnd, int controlId, bool enabled) {
    HWND ctl = GetDlgItem(hwnd, controlId);
    if (!ctl) {
        return;
    }
    const BOOL currentlyEnabled = IsWindowEnabled(ctl);
    if ((currentlyEnabled != FALSE) == enabled) {
        return;
    }
    EnableWindow(ctl, enabled ? TRUE : FALSE);
    InvalidateRect(ctl, nullptr, TRUE);
}

HeaderStatusBadgeState ResolveHeaderStatusBadgeState(const MicMixSettings& settings, const SourceStatus& status) {
    if (status.state == SourceState::Error) {
        return HeaderStatusBadgeState::Error;
    }
    if (!IsSourceStateActive(status.state)) {
        return HeaderStatusBadgeState::Off;
    }
    if (settings.musicMuted) {
        return HeaderStatusBadgeState::Muted;
    }
    return HeaderStatusBadgeState::Active;
}

HeaderBadgeVisual GetHeaderBadgeVisual(HeaderStatusBadgeState state) {
    switch (state) {
    case HeaderStatusBadgeState::Active:
        return { L"ACTIVE", RGB(21, 171, 88), RGB(9, 133, 63), RGB(255, 255, 255), RGB(224, 255, 234) };
    case HeaderStatusBadgeState::Muted:
        return { L"MUTED", RGB(227, 149, 19), RGB(181, 113, 6), RGB(255, 255, 255), RGB(255, 240, 213) };
    case HeaderStatusBadgeState::Error:
        return { L"ERROR", RGB(214, 61, 53), RGB(166, 36, 29), RGB(255, 255, 255), RGB(255, 226, 223) };
    default:
        return {};
    }
}

void UpdateWindowTitleState(HWND hwnd, bool active) {
    if (!hwnd) {
        return;
    }
    if (g_windowTitleStateValid && g_windowTitleActive == active) {
        return;
    }
    g_windowTitleStateValid = true;
    g_windowTitleActive = active;
    SetWindowTextW(hwnd, active ? L"MicMix Settings - ACTIVE" : L"MicMix Settings - OFF");
}

void UpdateHeaderStatusBadgeState(HWND hwnd, const MicMixSettings& settings, const SourceStatus& status) {
    const HeaderStatusBadgeState nextBadge = ResolveHeaderStatusBadgeState(settings, status);
    if (nextBadge != g_headerStatusBadgeState) {
        g_headerStatusBadgeState = nextBadge;
        if (g_rcHeaderBadge.right > g_rcHeaderBadge.left) {
            InvalidateRect(hwnd, &g_rcHeaderBadge, FALSE);
        } else {
            InvalidateRect(hwnd, nullptr, FALSE);
        }
    }
    UpdateWindowTitleState(hwnd, IsSourceStateActive(status.state) && status.state != SourceState::Error);
}

int GainDbToSlider(float db) {
    const float clamped = std::clamp(db, kMusicGainMinDb, kMusicGainMaxDb);
    const int value = static_cast<int>((clamped - kMusicGainMinDb) * static_cast<float>(kMusicGainStepPerDb));
    return std::max(0, std::min(kMusicGainSliderMax, value));
}

float SliderToGainDb(int slider) {
    const int clamped = std::max(0, std::min(kMusicGainSliderMax, slider));
    return kMusicGainMinDb + (static_cast<float>(clamped) / static_cast<float>(kMusicGainStepPerDb));
}

float MeterDbToUnit(float dbfs) {
    const float clamped = std::clamp(dbfs, -60.0f, 0.0f);
    return (clamped + 60.0f) / 60.0f;
}

void SetStaticTextIfChanged(HWND ctl, std::wstring& cache, const std::wstring& nextText) {
    if (!ctl || cache == nextText) {
        return;
    }
    cache = nextText;
    SetWindowTextW(ctl, cache.c_str());
}

void UpdateMusicMeter(HWND hwnd) {
    HWND txt = GetDlgItem(hwnd, IDC_METER_TEXT);
    if (!txt) {
        return;
    }
    const TelemetrySnapshot t = g_lastTelemetry;
    const float shownPeakDb = (t.musicSendPeakDbfs > -119.0f) ? t.musicSendPeakDbfs : t.musicPeakDbfs;
    const bool meterActive = IsSignalMeterActive(
        t.musicActive,
        t.musicSendPeakDbfs,
        t.musicPeakDbfs,
        t.musicRmsDbfs);

    float targetDb = -60.0f;
    if (meterActive) {
        targetDb = std::clamp(shownPeakDb, -60.0f, 0.0f);
    }
    if (targetDb > g_meterVisualDb) {
        g_meterVisualDb += (targetDb - g_meterVisualDb) * 0.70f;
    } else {
        g_meterVisualDb += (targetDb - g_meterVisualDb) * 0.30f;
    }
    g_meterVisualDb = std::clamp(g_meterVisualDb, -60.0f, 0.0f);

    if (targetDb >= g_meterHoldDb) {
        g_meterHoldDb = targetDb;
    } else {
        g_meterHoldDb = std::max(-60.0f, g_meterHoldDb - 0.6f);
    }

    if (!meterActive) {
        SetStaticTextIfChanged(txt, g_lastMusicMeterText, L"No signal");
    } else {
        const wchar_t* grade = L"OK";
        if (shownPeakDb > -8.0f) {
            grade = L"Hot";
        } else if (shownPeakDb > -16.0f) {
            grade = L"Warm";
        } else if (shownPeakDb < -35.0f) {
            grade = L"Quiet";
        }
        wchar_t meter[96];
        swprintf_s(meter, L"%.1f dBFS  |  %s", shownPeakDb, grade);
        SetStaticTextIfChanged(txt, g_lastMusicMeterText, meter);
    }

    const int meterWidth = std::max(1, static_cast<int>(g_rcMeter.right - g_rcMeter.left));
    const int levelPx = std::clamp(::MeterDbToPixels(g_meterVisualDb, meterWidth), 0, meterWidth);
    const int holdPx = std::clamp(::MeterDbToPixels(g_meterHoldDb, meterWidth), 0, meterWidth - 1);
    const bool meterChanged =
        (levelPx != g_lastMusicMeterLevelPx) ||
        (holdPx != g_lastMusicMeterHoldPx) ||
        (meterActive != g_lastMusicMeterActive);
    g_lastMusicMeterLevelPx = levelPx;
    g_lastMusicMeterHoldPx = holdPx;
    g_lastMusicMeterActive = meterActive;
    if (meterChanged) {
        InvalidateRect(hwnd, &g_rcMeter, FALSE);
    }

    const int clipLitSegments = ::ClipLitSegmentsFromDb(shownPeakDb);
    const bool clipChanged =
        (clipLitSegments != g_lastMusicClipLitSegments) ||
        (t.sourceClipRecent != g_lastMusicClipRecent);
    g_lastMusicClipLitSegments = clipLitSegments;
    g_lastMusicClipRecent = t.sourceClipRecent;
    if (clipChanged) {
        InvalidateRect(hwnd, &g_rcMusicClip, FALSE);
    }
}

void UpdateMicMeter(HWND hwnd) {
    HWND txt = GetDlgItem(hwnd, IDC_MIC_METER_TEXT);
    if (!txt) {
        return;
    }
    const TelemetrySnapshot t = g_lastTelemetry;
    float targetDb = -60.0f;
    if (t.micRmsDbfs > -119.0f) {
        targetDb = std::clamp(t.micRmsDbfs, -60.0f, 0.0f);
    }
    if (targetDb > g_micMeterVisualDb) {
        g_micMeterVisualDb += (targetDb - g_micMeterVisualDb) * 0.68f;
    } else {
        g_micMeterVisualDb += (targetDb - g_micMeterVisualDb) * 0.28f;
    }
    g_micMeterVisualDb = std::clamp(g_micMeterVisualDb, -60.0f, 0.0f);

    if (targetDb >= g_micMeterHoldDb) {
        g_micMeterHoldDb = targetDb;
    } else {
        g_micMeterHoldDb = std::max(-60.0f, g_micMeterHoldDb - 0.7f);
    }

    if (t.micRmsDbfs <= -119.0f) {
        SetStaticTextIfChanged(txt, g_lastMicMeterText, L"No signal");
    } else {
        const wchar_t* grade = L"OK";
        if (t.micPeakDbfs > -8.0f) {
            grade = L"Hot";
        } else if (t.micPeakDbfs > -16.0f) {
            grade = L"Warm";
        } else if (t.micPeakDbfs < -35.0f) {
            grade = L"Quiet";
        }
        wchar_t meter[96];
        swprintf_s(meter, L"%.1f dBFS  |  %s", t.micRmsDbfs, grade);
        SetStaticTextIfChanged(txt, g_lastMicMeterText, meter);
    }

    const int meterWidth = std::max(1, static_cast<int>(g_rcMicMeter.right - g_rcMicMeter.left));
    const int levelPx = std::clamp(::MeterDbToPixels(g_micMeterVisualDb, meterWidth), 0, meterWidth);
    const int holdPx = std::clamp(::MeterDbToPixels(g_micMeterHoldDb, meterWidth), 0, meterWidth - 1);
    const bool meterActive = t.micRmsDbfs > -119.0f;
    const bool meterChanged =
        (levelPx != g_lastMicMeterLevelPx) ||
        (holdPx != g_lastMicMeterHoldPx) ||
        (meterActive != g_lastMicMeterActive);
    g_lastMicMeterLevelPx = levelPx;
    g_lastMicMeterHoldPx = holdPx;
    g_lastMicMeterActive = meterActive;
    if (meterChanged) {
        InvalidateRect(hwnd, &g_rcMicMeter, FALSE);
    }

    const int clipLitSegments = ::ClipLitSegmentsFromDb(t.micPeakDbfs);
    const bool clipChanged =
        (clipLitSegments != g_lastMicClipLitSegments) ||
        (t.micClipRecent != g_lastMicClipRecent);
    g_lastMicClipLitSegments = clipLitSegments;
    g_lastMicClipRecent = t.micClipRecent;
    if (clipChanged) {
        InvalidateRect(hwnd, &g_rcMicClip, FALSE);
    }
}

void UpdateMonitorButton(HWND hwnd) {
    const bool enabled = MicMixApp::Instance().IsMonitorEnabled();
    SetMonitorButtonText(hwnd, IDC_MONITOR, enabled);
}

void UpdateGainLabel(HWND hwnd) {
    const int slider = static_cast<int>(SendMessageW(GetDlgItem(hwnd, IDC_GAIN), TBM_GETPOS, 0, 0));
    const float db = SliderToGainDb(slider);
    wchar_t buf[64];
    swprintf_s(buf, L"%+.1f dB", db);
    SetWindowTextW(GetDlgItem(hwnd, IDC_GAIN_VALUE), buf);
}

bool IsModifierOnlyKey(UINT vk) {
    return vk == VK_SHIFT || vk == VK_LSHIFT || vk == VK_RSHIFT ||
           vk == VK_CONTROL || vk == VK_LCONTROL || vk == VK_RCONTROL ||
           vk == VK_MENU || vk == VK_LMENU || vk == VK_RMENU ||
           vk == VK_LWIN || vk == VK_RWIN;
}

UINT ReadCurrentModifiers() {
    UINT mods = 0;
    if ((GetKeyState(VK_CONTROL) & 0x8000) != 0) mods |= MOD_CONTROL;
    if ((GetKeyState(VK_MENU) & 0x8000) != 0) mods |= MOD_ALT;
    if ((GetKeyState(VK_SHIFT) & 0x8000) != 0) mods |= MOD_SHIFT;
    if ((GetKeyState(VK_LWIN) & 0x8000) != 0 || (GetKeyState(VK_RWIN) & 0x8000) != 0) mods |= MOD_WIN;
    return mods;
}

bool IsExtendedVk(UINT vk) {
    switch (vk) {
    case VK_INSERT:
    case VK_DELETE:
    case VK_HOME:
    case VK_END:
    case VK_PRIOR:
    case VK_NEXT:
    case VK_LEFT:
    case VK_RIGHT:
    case VK_UP:
    case VK_DOWN:
    case VK_NUMLOCK:
    case VK_DIVIDE:
    case VK_RCONTROL:
    case VK_RMENU:
        return true;
    default:
        return false;
    }
}

std::wstring VkToText(UINT vk) {
    UINT scanCode = MapVirtualKeyW(vk, MAPVK_VK_TO_VSC);
    LONG lParam = static_cast<LONG>(scanCode << 16);
    if (IsExtendedVk(vk)) {
        lParam |= (1 << 24);
    }
    wchar_t keyName[128]{};
    if (GetKeyNameTextW(lParam, keyName, static_cast<int>(std::size(keyName))) > 0) {
        return keyName;
    }
    wchar_t fallback[32]{};
    swprintf_s(fallback, L"VK %u", vk);
    return fallback;
}

std::wstring FormatHotkeyText(UINT mods, UINT vk) {
    if (vk == 0) {
        return L"Not set";
    }
    std::wstring out;
    if ((mods & MOD_CONTROL) != 0) out += L"Ctrl + ";
    if ((mods & MOD_ALT) != 0) out += L"Alt + ";
    if ((mods & MOD_SHIFT) != 0) out += L"Shift + ";
    if ((mods & MOD_WIN) != 0) out += L"Win + ";
    out += VkToText(vk);
    return out;
}

bool IsHotkeyConflict(UINT lhsMods, UINT lhsVk, UINT rhsMods, UINT rhsVk) {
    return lhsVk != 0 && rhsVk != 0 && lhsVk == rhsVk && lhsMods == rhsMods;
}

bool IsCapturingHotkey() {
    return g_hotkeyCaptureTarget != HotkeyCaptureTarget::None;
}

void UpdateMuteHotkeyLabel(HWND hwnd) {
    if (g_hotkeyCaptureTarget == HotkeyCaptureTarget::MusicMute) {
        return;
    }
    HWND label = GetDlgItem(hwnd, IDC_MUTE_HOTKEY_TEXT);
    if (!label) {
        return;
    }
    const std::wstring text = L"Current: " + FormatHotkeyText(g_muteHotkeyModifiers, g_muteHotkeyVk);
    SetWindowTextW(label, text.c_str());
}

void UpdateMicInputMuteHotkeyLabel(HWND hwnd) {
    if (g_hotkeyCaptureTarget == HotkeyCaptureTarget::MicInputMute) {
        return;
    }
    HWND label = GetDlgItem(hwnd, IDC_MIC_INPUT_HOTKEY_TEXT);
    if (!label) {
        return;
    }
    const std::wstring text = L"Current: " + FormatHotkeyText(g_micInputMuteHotkeyModifiers, g_micInputMuteHotkeyVk);
    SetWindowTextW(label, text.c_str());
}

void UpdateHotkeyLabels(HWND hwnd) {
    UpdateMuteHotkeyLabel(hwnd);
    UpdateMicInputMuteHotkeyLabel(hwnd);
}

void BeginHotkeyCapture(HWND hwnd, HotkeyCaptureTarget target) {
    g_hotkeyCaptureTarget = target;
    MicMixApp::Instance().SetGlobalHotkeyCaptureBlocked(true);
    int labelId = 0;
    if (target == HotkeyCaptureTarget::MusicMute) {
        labelId = IDC_MUTE_HOTKEY_TEXT;
    } else if (target == HotkeyCaptureTarget::MicInputMute) {
        labelId = IDC_MIC_INPUT_HOTKEY_TEXT;
    }
    if (labelId != 0) {
        SetWindowTextW(GetDlgItem(hwnd, labelId), L"Current: Press key... (Esc=Cancel, Del=Clear)");
    }
    SetFocus(hwnd);
}

void JoinSourceRefreshThread() {
    std::thread worker;
    {
        std::lock_guard<std::mutex> lock(g_sourceRefreshThreadMutex);
        if (!g_sourceRefreshThread.joinable()) {
            return;
        }
        worker = std::move(g_sourceRefreshThread);
    }
    worker.join();
}

bool StartSourceRefreshWorker(HWND hwnd, uint64_t seq) {
    if (!hwnd) {
        return false;
    }

    JoinSourceRefreshThread();
    try {
        std::thread worker([hwnd, seq]() {
            auto& app = MicMixApp::Instance();
            std::vector<LoopbackDeviceInfo> loopbacks = app.GetLoopbackDevices();
            std::vector<CaptureDeviceInfo> captureDevices = app.GetCaptureDevices();
            std::vector<AppProcessInfo> apps = app.GetAppProcesses();
            {
                std::lock_guard<std::mutex> lock(g_enumMutex);
                g_pendingLoopbacks = std::move(loopbacks);
                g_pendingCaptureDevices = std::move(captureDevices);
                g_pendingApps = std::move(apps);
            }
            if (!PostMessageW(hwnd, kMsgSourceRefreshDone, static_cast<WPARAM>(seq), 0)) {
                g_sourceRefreshInFlight.store(false, std::memory_order_release);
                g_sourceRefreshActiveSeq.store(0, std::memory_order_release);
                const HWND currentHwnd = g_hwnd.load(std::memory_order_acquire);
                if (currentHwnd != nullptr) {
                    LogWarn("settings_window refresh completion post failed");
                }
            }
        });
        {
            std::lock_guard<std::mutex> lock(g_sourceRefreshThreadMutex);
            g_sourceRefreshThread = std::move(worker);
        }
        return true;
    } catch (const std::exception& ex) {
        g_sourceRefreshInFlight.store(false, std::memory_order_release);
        g_sourceRefreshActiveSeq.store(0, std::memory_order_release);
        LogError(std::string("settings_window refresh thread start failed: ") + ex.what());
    } catch (...) {
        g_sourceRefreshInFlight.store(false, std::memory_order_release);
        g_sourceRefreshActiveSeq.store(0, std::memory_order_release);
        LogError("settings_window refresh thread start failed: unknown");
    }
    return false;
}

void RequestSourceRefresh(HWND hwnd, bool reloadSettings) {
    if (!hwnd) {
        return;
    }
    if (reloadSettings) {
        g_sourceRefreshPendingReload.store(true, std::memory_order_release);
    }
    HWND refreshBtn = GetDlgItem(hwnd, IDC_RESCAN);
    if (refreshBtn) {
        EnableWindow(refreshBtn, FALSE);
    }
    SetStatusText(hwnd, L"Refreshing source list...");

    if (g_sourceRefreshInFlight.exchange(true, std::memory_order_acq_rel)) {
        g_sourceRefreshPending.store(true, std::memory_order_release);
        return;
    }
    const uint64_t seq = g_sourceRefreshSeq.fetch_add(1, std::memory_order_acq_rel) + 1ULL;
    g_sourceRefreshActiveSeq.store(seq, std::memory_order_release);
    if (!StartSourceRefreshWorker(hwnd, seq)) {
        g_sourceRefreshInFlight.store(false, std::memory_order_release);
        g_sourceRefreshActiveSeq.store(0, std::memory_order_release);
        if (refreshBtn) {
            EnableWindow(refreshBtn, TRUE);
        }
        SetStatusText(hwnd, L"Refresh failed to start");
    }
}

void PopulateCombos(HWND hwnd) {
    g_sourceChoices.clear();
    g_lastValidSourceSel = -1;
    EnsureFallbackIcons();
    std::unordered_set<uint32_t> activeAppPids;

    auto addStaticItem = [](HWND combo, const wchar_t* text, INT_PTR data) {
        const LRESULT idx = SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(text));
        if (idx >= 0) {
            SendMessageW(combo, CB_SETITEMDATA, static_cast<WPARAM>(idx), static_cast<LPARAM>(data));
        }
    };

    auto addSelectableItem = [](HWND combo, const std::wstring& text, const SourceChoice& choice) {
        const LRESULT idx = SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(text.c_str()));
        if (idx >= 0) {
            g_sourceChoices.push_back(choice);
            const INT_PTR itemData = static_cast<INT_PTR>(g_sourceChoices.size() - 1);
            SendMessageW(combo, CB_SETITEMDATA, static_cast<WPARAM>(idx), static_cast<LPARAM>(itemData));
        }
    };

    HWND source = GetDlgItem(hwnd, IDC_SOURCE);
    SendMessageW(source, CB_RESETCONTENT, 0, 0);

    addStaticItem(source, L"Audio Channels:", kSourceItemHeader);
    if (g_loopbacks.empty()) {
        addStaticItem(source, L"  (no audio channels found)", kSourceItemPlaceholder);
    } else {
        for (const auto& d : g_loopbacks) {
            std::wstring label = L"  ";
            label += Utf8ToWide(d.name);
            if (d.isDefault) {
                label += L" (Default)";
            }
            SourceChoice choice{};
            choice.type = SourceChoiceType::Loopback;
            choice.id = d.id;
            addSelectableItem(source, label, choice);
        }
    }

    addStaticItem(source, L"------------------------------", kSourceItemDivider);
    addStaticItem(source, L"Running Apps:", kSourceItemHeader);
    if (g_apps.empty()) {
        addStaticItem(source, L"  (no running app found)", kSourceItemPlaceholder);
    } else {
        for (const auto& p : g_apps) {
            std::wstring label = L"  ";
            label += Utf8ToWide(p.displayName);
            SourceChoice choice{};
            choice.type = SourceChoiceType::App;
            choice.id = std::to_string(p.pid);
            choice.processName = p.exeName;
            choice.pid = p.pid;
            addSelectableItem(source, label, choice);
            activeAppPids.insert(p.pid);
            EnsureAppIconForPid(p.pid);
        }
    }
    PruneAppIconCache(activeAppPids);

    HWND mic = GetDlgItem(hwnd, IDC_MIC_DEVICE);
    if (mic) {
        SendMessageW(mic, CB_RESETCONTENT, 0, 0);
        for (const auto& d : g_captureDevices) {
            std::wstring label = Utf8ToWide(d.name);
            if (d.isDefault) {
                label += L" (Default)";
            }
            SendMessageW(mic, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(label.c_str()));
        }
    }
}

void LoadSettings(HWND hwnd) {
    g_loadingUi = true;
    auto& app = MicMixApp::Instance();
    const MicMixSettings s = app.GetSettings();
    const std::wstring tsCaptureNameNorm = NormalizeDeviceName(Utf8ToWide(app.GetPreferredTsCaptureDeviceName()));

    int selectedComboIdx = -1;
    int firstLoopbackIdx = -1;
    int firstSelectableIdx = -1;

    HWND source = GetDlgItem(hwnd, IDC_SOURCE);
    const int count = static_cast<int>(SendMessageW(source, CB_GETCOUNT, 0, 0));
    for (int i = 0; i < count; ++i) {
        const LRESULT data = SendMessageW(source, CB_GETITEMDATA, static_cast<WPARAM>(i), 0);
        if (data < 0 || data == CB_ERR || static_cast<size_t>(data) >= g_sourceChoices.size()) {
            continue;
        }
        if (firstSelectableIdx < 0) {
            firstSelectableIdx = i;
        }
        const SourceChoice& choice = g_sourceChoices[static_cast<size_t>(data)];
        if (choice.type == SourceChoiceType::Loopback && firstLoopbackIdx < 0) {
            firstLoopbackIdx = i;
        }
        if (s.sourceMode == SourceMode::Loopback && choice.type == SourceChoiceType::Loopback && choice.id == s.loopbackDeviceId) {
            selectedComboIdx = i;
            break;
        }
        if (s.sourceMode == SourceMode::AppSession && choice.type == SourceChoiceType::App) {
            if (choice.id == s.appSessionId) {
                selectedComboIdx = i;
                break;
            }
            if (selectedComboIdx < 0 && !s.appProcessName.empty() && _stricmp(choice.processName.c_str(), s.appProcessName.c_str()) == 0) {
                selectedComboIdx = i;
            }
        }
    }
    if (selectedComboIdx < 0) {
        if (s.sourceMode == SourceMode::Loopback && firstLoopbackIdx >= 0) {
            selectedComboIdx = firstLoopbackIdx;
        } else {
            selectedComboIdx = firstSelectableIdx;
        }
    }
    if (selectedComboIdx >= 0) {
        SendMessageW(source, CB_SETCURSEL, selectedComboIdx, 0);
        g_lastValidSourceSel = selectedComboIdx;
    }

    int micSel = -1;
    int micDefaultIdx = -1;
    HWND micCombo = GetDlgItem(hwnd, IDC_MIC_DEVICE);
    for (size_t i = 0; i < g_captureDevices.size(); ++i) {
        const auto& d = g_captureDevices[i];
        if (d.isDefault && micDefaultIdx < 0) {
            micDefaultIdx = static_cast<int>(i);
        }
        if (!s.captureDeviceId.empty() && d.id == s.captureDeviceId) {
            micSel = static_cast<int>(i);
            break;
        }
        if (micSel < 0 && !tsCaptureNameNorm.empty()) {
            const std::wstring devNorm = NormalizeDeviceName(Utf8ToWide(d.name));
            if (devNorm == tsCaptureNameNorm) {
                micSel = static_cast<int>(i);
            }
        }
    }
    if (micSel < 0) {
        micSel = micDefaultIdx;
    }
    if (micCombo && micSel >= 0) {
        SendMessageW(micCombo, CB_SETCURSEL, micSel, 0);
    }

    SendMessageW(GetDlgItem(hwnd, IDC_GAIN), TBM_SETPOS, TRUE, GainDbToSlider(s.musicGainDb));
    SetOwnerCheckboxValue(IDC_FORCE_TX, s.forceTxEnabled);
    SetOwnerCheckboxValue(IDC_MIC_INPUT_MUTE, s.micInputMuted);
    SetOwnerCheckboxValue(IDC_MUTE, s.musicMuted);
    SetOwnerCheckboxValue(IDC_AUTOSTART, s.autostartEnabled);
    InvalidateRect(GetDlgItem(hwnd, IDC_FORCE_TX), nullptr, TRUE);
    InvalidateRect(GetDlgItem(hwnd, IDC_MIC_INPUT_MUTE), nullptr, TRUE);
    InvalidateRect(GetDlgItem(hwnd, IDC_MUTE), nullptr, TRUE);
    InvalidateRect(GetDlgItem(hwnd, IDC_AUTOSTART), nullptr, TRUE);
    g_muteHotkeyModifiers = static_cast<UINT>(std::max(0, s.muteHotkeyModifiers));
    g_muteHotkeyVk = static_cast<UINT>(std::max(0, s.muteHotkeyVk));
    g_micInputMuteHotkeyModifiers = static_cast<UINT>(std::max(0, s.micInputMuteHotkeyModifiers));
    g_micInputMuteHotkeyVk = static_cast<UINT>(std::max(0, s.micInputMuteHotkeyVk));
    UpdateGainLabel(hwnd);
    UpdateHotkeyLabels(hwnd);
    g_loadingUi = false;
}

MicMixSettings CollectSettings(HWND hwnd) {
    auto s = MicMixApp::Instance().GetSettings();

    const int sourceSel = static_cast<int>(SendMessageW(GetDlgItem(hwnd, IDC_SOURCE), CB_GETCURSEL, 0, 0));
    if (sourceSel >= 0) {
        const LRESULT data = SendMessageW(GetDlgItem(hwnd, IDC_SOURCE), CB_GETITEMDATA, static_cast<WPARAM>(sourceSel), 0);
        if (data >= 0 && data != CB_ERR && static_cast<size_t>(data) < g_sourceChoices.size()) {
            const SourceChoice& choice = g_sourceChoices[static_cast<size_t>(data)];
            if (choice.type == SourceChoiceType::Loopback) {
                s.sourceMode = SourceMode::Loopback;
                s.loopbackDeviceId = choice.id;
            } else {
                s.sourceMode = SourceMode::AppSession;
                s.appSessionId = choice.id;
                s.appProcessName = choice.processName;
            }
        }
    }

    const int micSel = static_cast<int>(SendMessageW(GetDlgItem(hwnd, IDC_MIC_DEVICE), CB_GETCURSEL, 0, 0));
    if (micSel >= 0 && static_cast<size_t>(micSel) < g_captureDevices.size()) {
        s.captureDeviceId = g_captureDevices[static_cast<size_t>(micSel)].id;
    }

    s.musicGainDb = SliderToGainDb(static_cast<int>(SendMessageW(GetDlgItem(hwnd, IDC_GAIN), TBM_GETPOS, 0, 0)));
    s.forceTxEnabled = GetOwnerCheckboxValue(IDC_FORCE_TX);
    s.micInputMuted = GetOwnerCheckboxValue(IDC_MIC_INPUT_MUTE);
    s.musicMuted = GetOwnerCheckboxValue(IDC_MUTE);
    s.autostartEnabled = GetOwnerCheckboxValue(IDC_AUTOSTART);
    s.muteHotkeyModifiers = static_cast<int>(g_muteHotkeyModifiers);
    s.muteHotkeyVk = static_cast<int>(g_muteHotkeyVk);
    s.micInputMuteHotkeyModifiers = static_cast<int>(g_micInputMuteHotkeyModifiers);
    s.micInputMuteHotkeyVk = static_cast<int>(g_micInputMuteHotkeyVk);
    return s;
}

void UpdateStatus(HWND hwnd, bool includeDetails = true) {
    auto& app = MicMixApp::Instance();
    const TelemetrySnapshot t = app.GetTelemetry();
    g_lastTelemetry = t;

    if (t.sourceClipEvents != g_lastMusicClipEvents) {
        wchar_t clipMusic[64]{};
        swprintf_s(clipMusic, L"CLIP Events: %llu", static_cast<unsigned long long>(t.sourceClipEvents));
        SetWindowTextW(GetDlgItem(hwnd, IDC_MUSIC_CLIP_EVENTS), clipMusic);
        g_lastMusicClipEvents = t.sourceClipEvents;
    }
    if (t.micClipEvents != g_lastMicClipEvents) {
        wchar_t clipMic[64]{};
        swprintf_s(clipMic, L"CLIP Events: %llu", static_cast<unsigned long long>(t.micClipEvents));
        SetWindowTextW(GetDlgItem(hwnd, IDC_MIC_CLIP_EVENTS), clipMic);
        g_lastMicClipEvents = t.micClipEvents;
    }

    UpdateMusicMeter(hwnd);
    UpdateMicMeter(hwnd);
    if (!includeDetails) {
        return;
    }

    const MicMixSettings s = app.GetSettings();
    const SourceStatus st = app.GetSourceStatus();
    const bool sourceUp = IsSourceStateActive(st.state);
    SetControlEnabled(hwnd, IDC_MIC_INPUT_MUTE, sourceUp);
    SetControlEnabled(hwnd, IDC_MUTE, sourceUp);
    UpdateHeaderStatusBadgeState(hwnd, s, st);
    if (GetOwnerCheckboxValue(IDC_MUTE) != s.musicMuted) {
        SetOwnerCheckboxValue(IDC_MUTE, s.musicMuted);
        InvalidateRect(GetDlgItem(hwnd, IDC_MUTE), nullptr, TRUE);
    }
    if (GetOwnerCheckboxValue(IDC_MIC_INPUT_MUTE) != s.micInputMuted) {
        SetOwnerCheckboxValue(IDC_MIC_INPUT_MUTE, s.micInputMuted);
        InvalidateRect(GetDlgItem(hwnd, IDC_MIC_INPUT_MUTE), nullptr, TRUE);
    }
    std::string line1 = "State: " + SourceStateToString(st.state) + "  |  " + (st.message.empty() ? "idle" : st.message);
    if (!st.detail.empty()) {
        line1 += "  (" + st.detail + ")";
    }
    const float shownSendPeakDb = (t.musicSendPeakDbfs > -119.0f) ? t.musicSendPeakDbfs : t.musicPeakDbfs;
    char micBuf[32]{};
    sprintf_s(micBuf, "mic=%.1fdB", t.micRmsDbfs);
    char sendBuf[64]{};
    sprintf_s(sendBuf, "send=%.1fdB src=%.1fdB", shownSendPeakDb, t.musicPeakDbfs);
    const std::string resamplerText = (s.resamplerQuality < 0) ? "auto-cpu" : std::to_string(s.resamplerQuality);
    std::string line2 = "underrun=" + std::to_string(t.underruns) +
        "  overrun=" + std::to_string(t.overruns) +
        "  clip=" + std::to_string(t.clippedSamples) +
        "  resampler=" + resamplerText +
        "  " + micBuf +
        "  " + sendBuf +
        "  tag.micmix_active=" + std::to_string(t.musicActive ? 1 : 0);
    SetStatusText(hwnd, Utf8ToWide(line1 + "\r\n" + line2));
    UpdateMonitorButton(hwnd);
}

void FlushDebouncedSettingsSave(HWND hwnd) {
    if (!hwnd) {
        return;
    }
    KillTimer(hwnd, kTimerDebouncedSave);
    if (!g_saveDebouncePending.exchange(false, std::memory_order_acq_rel) || g_loadingUi) {
        return;
    }
    auto s = CollectSettings(hwnd);
    MicMixApp::Instance().ApplySettings(s, false, true);
}

void ApplyLiveSettings(HWND hwnd, bool restartSource, bool saveImmediately = true) {
    if (g_loadingUi) return;
    auto s = CollectSettings(hwnd);
    if (saveImmediately) {
        KillTimer(hwnd, kTimerDebouncedSave);
        g_saveDebouncePending.store(false, std::memory_order_release);
        MicMixApp::Instance().ApplySettings(s, restartSource, true);
    } else {
        MicMixApp::Instance().ApplySettings(s, restartSource, false);
        g_saveDebouncePending.store(true, std::memory_order_release);
        SetTimer(hwnd, kTimerDebouncedSave, kDebouncedSaveMs, nullptr);
    }
    UpdateStatus(hwnd);
}

void ApplyControlTheme(HWND hwnd) {
    const int unthemeOwnerDrawIds[] = {
        IDC_AUTOSTART,
        IDC_FORCE_TX,
        IDC_MIC_INPUT_MUTE,
        IDC_MUTE,
    };
    ApplyMicMixControlTheme(
        hwnd,
        kUiThemeButtons | kUiThemeComboBoxes | kUiThemeTrackBars,
        true,
        true,
        unthemeOwnerDrawIds,
        static_cast<int>(std::size(unthemeOwnerDrawIds)));
}

void ApplyFonts(HWND hwnd) {
    EnumChildWindows(hwnd, [](HWND child, LPARAM) -> BOOL {
        SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(g_fontBody), TRUE);
        return TRUE;
    }, 0);
    SetControlFontById(hwnd, IDC_TITLE, g_fontTitle);
    SetControlFontById(hwnd, IDC_SUBTITLE, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFontById(hwnd, IDC_VERSION, g_fontSmall);
    SetControlFontById(hwnd, IDC_REPO_LINK, g_fontSmall);
    SetControlFontById(hwnd, IDC_METER_TEXT, g_fontSmall);
    SetControlFontById(hwnd, IDC_MIC_METER_TEXT, g_fontSmall);
    SetControlFontById(hwnd, IDC_MUSIC_CLIP_EVENTS, g_fontClipInfo ? g_fontClipInfo : (g_fontHint ? g_fontHint : g_fontSmall));
    SetControlFontById(hwnd, IDC_MIC_CLIP_EVENTS, g_fontClipInfo ? g_fontClipInfo : (g_fontHint ? g_fontHint : g_fontSmall));
    SetControlFontById(hwnd, IDC_AUTOSTART, g_fontControlLarge ? g_fontControlLarge : g_fontBody);
    SetControlFontById(hwnd, IDC_FORCE_TX, g_fontControlLarge ? g_fontControlLarge : g_fontBody);
    SetControlFontById(hwnd, IDC_MIC_INPUT_MUTE, g_fontControlLarge ? g_fontControlLarge : g_fontBody);
    SetControlFontById(hwnd, IDC_MUTE, g_fontControlLarge ? g_fontControlLarge : g_fontBody);
    SetControlFontById(hwnd, IDC_MONITOR_HINT, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFontById(hwnd, IDC_MIC_METER_HINT, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFontById(hwnd, IDC_GAIN_HINT, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFontById(hwnd, IDC_FORCE_TX_HINT, g_fontHint ? g_fontHint : g_fontSmall);
}

void ComputeLayout() {
    const int left = S(kCardMarginPx);
    const int right = S(kClientWidthPx - kCardMarginPx);
    const int gap = S(kCardGapPx);
    g_rcSession = { left, S(82), right, S(210) };
    g_rcSource = { left, g_rcSession.bottom + gap, right, S(342) };
    g_rcMix = { left, g_rcSource.bottom + gap, right, S(kClientHeightPx - kCardMarginPx) };
}

void DrawCard(HDC hdc, const RECT& rc) {
    DrawFlatCardFrame(hdc, rc, g_brushCard, g_theme.border);
}

void FillSolidRect(HDC hdc, const RECT& rc, COLORREF color) {
    const COLORREF oldColor = SetDCBrushColor(hdc, color);
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(DC_BRUSH));
    FillRect(hdc, &rc, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
    SelectObject(hdc, oldBrush);
    if (oldColor != CLR_INVALID) {
        SetDCBrushColor(hdc, oldColor);
    }
}

void DrawHeaderStatusBadge(HDC hdc) {
    if (g_rcHeaderBadge.right <= g_rcHeaderBadge.left || g_rcHeaderBadge.bottom <= g_rcHeaderBadge.top) {
        return;
    }
    const HeaderBadgeVisual visual = GetHeaderBadgeVisual(g_headerStatusBadgeState);
    DrawHeaderBadge(hdc, g_rcHeaderBadge, visual, g_fontSmall, g_fontBody, g_dpi);
}

void DrawLevelMeter(HDC hdc, const RECT& meterRect, bool active, float visualDb, float holdDb) {
    const MeterFonts fonts{
        g_fontTiny,
        g_fontHint,
        g_fontSmall,
        g_fontBody,
    };
    DrawLevelMeterShared(hdc, meterRect, active, visualDb, holdDb, fonts, g_dpi);
}

void EndHotkeyCapture(HWND hwnd, bool refreshLabels = true) {
    g_hotkeyCaptureTarget = HotkeyCaptureTarget::None;
    MicMixApp::Instance().SetGlobalHotkeyCaptureBlocked(false);
    if (refreshLabels) {
        UpdateHotkeyLabels(hwnd);
    }
}

LRESULT DrawGainSliderCustom(const NMCUSTOMDRAW* cd) {
    if (!cd) {
        return CDRF_DODEFAULT;
    }
    if (cd->dwDrawStage == CDDS_PREPAINT) {
        return CDRF_NOTIFYITEMDRAW;
    }
    if (cd->dwDrawStage != CDDS_ITEMPREPAINT) {
        return CDRF_DODEFAULT;
    }

    const DWORD item = static_cast<DWORD>(cd->dwItemSpec);
    const RECT rc = cd->rc;
    if (item == TBCD_CHANNEL) {
        RECT line = rc;
        const int centerY = (rc.top + rc.bottom) / 2;
        line.top = centerY;
        line.bottom = centerY + 1;
        FillSolidRect(cd->hdc, line, RGB(204, 212, 223));
        return CDRF_SKIPDEFAULT;
    }
    if (item == TBCD_THUMB) {
        RECT thumb = rc;
        FillSolidRect(cd->hdc, thumb, RGB(245, 248, 252));
        const COLORREF borderColor = GetSysColor(COLOR_BTNSHADOW);
        const COLORREF oldPenColor = SetDCPenColor(cd->hdc, borderColor);
        HGDIOBJ oldPen = SelectObject(cd->hdc, GetStockObject(DC_PEN));
        HGDIOBJ oldBrush = SelectObject(cd->hdc, GetStockObject(HOLLOW_BRUSH));
        Rectangle(cd->hdc, thumb.left, thumb.top, thumb.right, thumb.bottom);
        SelectObject(cd->hdc, oldBrush);
        SelectObject(cd->hdc, oldPen);
        if (oldPenColor != CLR_INVALID) {
            SetDCPenColor(cd->hdc, oldPenColor);
        }
        return CDRF_SKIPDEFAULT;
    }
    return CDRF_DODEFAULT;
}

void DrawClipStrip(HDC hdc, const RECT& rcStrip, float dangerUnit, bool clipRecent) {
    const MeterFonts fonts{
        g_fontTiny,
        g_fontHint,
        g_fontSmall,
        g_fontBody,
    };
    DrawClipStripShared(hdc, rcStrip, dangerUnit, clipRecent, fonts, g_dpi);
}

void DrawSourceComboItem(const DRAWITEMSTRUCT* dis) {
    if (!dis || dis->itemID == static_cast<UINT>(-1)) {
        return;
    }
    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool disabled = (dis->itemState & ODS_DISABLED) != 0;

    std::wstring itemText;
    const LRESULT textLen = SendMessageW(dis->hwndItem, CB_GETLBTEXTLEN, dis->itemID, 0);
    if (textLen > 0 && textLen != CB_ERR) {
        itemText.resize(static_cast<size_t>(textLen) + 1U);
        const LRESULT copied = SendMessageW(
            dis->hwndItem,
            CB_GETLBTEXT,
            dis->itemID,
            reinterpret_cast<LPARAM>(itemText.data()));
        if (copied != CB_ERR && copied >= 0) {
            itemText.resize(static_cast<size_t>(copied));
        } else {
            itemText.clear();
        }
    }

    RECT rc = dis->rcItem;
    if (selected) {
        FillSolidRect(dis->hDC, rc, RGB(218, 234, 255));
    } else {
        FillSolidRect(dis->hDC, rc, RGB(255, 255, 255));
    }

    const LRESULT itemData = SendMessageW(dis->hwndItem, CB_GETITEMDATA, dis->itemID, 0);
    const bool isSelectableSource = !(itemData < 0 || itemData == CB_ERR || static_cast<size_t>(itemData) >= g_sourceChoices.size());

    HICON icon = nullptr;
    if (isSelectableSource) {
        const SourceChoice& choice = g_sourceChoices[static_cast<size_t>(itemData)];
        if (choice.type == SourceChoiceType::Loopback) {
            icon = g_loopbackFallbackIcon;
        } else {
            EnsureAppIconForPid(choice.pid);
            if (auto it = g_appIconsByPid.find(choice.pid); it != g_appIconsByPid.end()) {
                icon = it->second;
            } else {
                icon = g_appFallbackIcon;
            }
        }
    }

    int textLeft = rc.left + S(10);
    const int iconSize = S(kSourceIconSizePx);
    if (icon) {
        const int iconY = rc.top + ((rc.bottom - rc.top - iconSize) / 2);
        DrawIconEx(dis->hDC, rc.left + S(8), iconY, icon, iconSize, iconSize, 0, nullptr, DI_NORMAL);
        textLeft = rc.left + S(8) + iconSize + S(8);
    }

    SetBkMode(dis->hDC, TRANSPARENT);
    if (disabled) {
        SetTextColor(dis->hDC, RGB(150, 156, 166));
    } else if (itemData == kSourceItemHeader) {
        SetTextColor(dis->hDC, RGB(104, 112, 124));
    } else if (itemData == kSourceItemDivider) {
        SetTextColor(dis->hDC, RGB(184, 190, 200));
    } else if (itemData == kSourceItemPlaceholder) {
        SetTextColor(dis->hDC, RGB(132, 138, 148));
    } else {
        SetTextColor(dis->hDC, RGB(32, 36, 44));
    }
    SelectObject(dis->hDC, g_fontBody ? g_fontBody : GetStockObject(DEFAULT_GUI_FONT));

    RECT textRc = rc;
    textRc.left = textLeft;
    DrawTextW(dis->hDC, itemText.c_str(), -1, &textRc, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);

    if ((dis->itemState & ODS_FOCUS) != 0) {
        RECT focus = rc;
        focus.left += S(2);
        focus.right -= S(2);
        DrawFocusRect(dis->hDC, &focus);
    }
}

void DrawOwnerCheckbox(const DRAWITEMSTRUCT* dis) {
    if (!dis) {
        return;
    }
    DrawOwnerCheckboxShared(
        dis,
        GetOwnerCheckboxValue(dis->CtlID),
        g_fontControlLarge,
        g_fontBody,
        g_theme.text,
        g_theme.card,
        g_dpi);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        g_dpi = GetDpiForWindow(hwnd);
        EnsureUiResources();
        ComputeLayout();
        g_hotkeyRefreshTick = 0;
        g_hotkeyCaptureTarget = HotkeyCaptureTarget::None;
        g_lastMusicClipEvents = std::numeric_limits<uint64_t>::max();
        g_lastMicClipEvents = std::numeric_limits<uint64_t>::max();
        g_lastMusicMeterLevelPx = -1;
        g_lastMusicMeterHoldPx = -1;
        g_lastMusicMeterActive = false;
        g_lastMicMeterLevelPx = -1;
        g_lastMicMeterHoldPx = -1;
        g_lastMicMeterActive = false;
        g_lastMusicClipLitSegments = -1;
        g_lastMusicClipRecent = false;
        g_lastMicClipLitSegments = -1;
        g_lastMicClipRecent = false;
        g_lastMusicMeterText.clear();
        g_lastMicMeterText.clear();
        g_headerStatusBadgeState = HeaderStatusBadgeState::Off;
        g_windowTitleActive = false;
        g_windowTitleStateValid = false;
        const std::wstring versionText = Utf8ToWide(std::string("v") + MICMIX_VERSION + "  by dabinuss");

        const int contentLeft = S(kCardMarginPx + kCardInnerPaddingPx);
        const int contentRight = S(kClientWidthPx - kCardMarginPx - kCardInnerPaddingPx);
        const int controlGap = S(12);
        const int labelX = contentLeft;
        const int fieldX = S(180);
        const int actionButtonW = S(142);
        const int monitorButtonW = S(122);
        const int refreshButtonW = S(130);
        const int gainSliderW = S(360);
        const int gainSliderLeftNudge = S(6);
        const int meterW = S(340);
        const int sourceW = contentRight - fieldX - controlGap - refreshButtonW;
        const int micComboW = contentRight - fieldX;
        const int statusTextX = fieldX + meterW + S(4);
        const int statusTextW = contentRight - statusTextX;
        const int valueTextX = fieldX + gainSliderW;
        const int forceTxToggleW = S(240);
        const int muteToggleW = S(156);
        const int muteToggleGap = S(10);
        const int muteMusicToggleX = fieldX + muteToggleW + muteToggleGap;
        const int hintTopOffset = S(38);
        const int titleX = contentLeft;
        const int titleY = S(10);
        const int titleH = S(42);
        HDC headerMeasureDc = GetDC(hwnd);
        const HeaderLayout headerLayout = ComputeHeaderLayout(
            headerMeasureDc,
            g_fontTitle,
            g_fontSmall,
            g_fontBody,
            g_dpi,
            kSettingsHeaderTitle,
            contentLeft,
            contentRight,
            titleY,
            titleH,
            300);
        if (headerMeasureDc) {
            ReleaseDC(hwnd, headerMeasureDc);
        }
        const int titleW = headerLayout.titleWidthPx;
        const int headerMetaX = headerLayout.metaX;
        const int headerMetaActualW = headerLayout.metaWidth;
        const int statusX = contentLeft;
        const int statusY = S(704);
        const int statusW = contentRight - statusX;
        const int statusH = S(44);
        g_rcHeaderBadge = headerLayout.badgeRect;
        g_rcStatus = { statusX, statusY, statusX + statusW, statusY + statusH };

        CreateWindowW(L"STATIC", kSettingsHeaderTitle, WS_CHILD | WS_VISIBLE, titleX, titleY, titleW, titleH, hwnd, reinterpret_cast<HMENU>(IDC_TITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Configure MicMix and route one audio source to your mic output", WS_CHILD | WS_VISIBLE, contentLeft, S(43), headerMetaX - contentLeft - S(8), S(24), hwnd, reinterpret_cast<HMENU>(IDC_SUBTITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", versionText.c_str(), WS_CHILD | WS_VISIBLE | SS_RIGHT, headerMetaX, S(18), headerMetaActualW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_VERSION), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"github.com/dabinuss/MicMix", WS_CHILD | WS_VISIBLE | SS_RIGHT | SS_NOTIFY, headerMetaX, S(36), headerMetaActualW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_REPO_LINK), nullptr, nullptr);

        CreateWindowW(L"BUTTON", L"Enable MicMix", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, contentLeft, S(108), actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_START), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Disable MicMix", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, contentLeft + actionButtonW + controlGap, S(108), actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_STOP), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Restart Source", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, contentLeft + (actionButtonW + controlGap) * 2, S(108), actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_SAVE), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Monitor Mix: Off", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, contentLeft + (actionButtonW + controlGap) * 3, S(108), monitorButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_MONITOR), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Only while connected", WS_CHILD | WS_VISIBLE, contentLeft + (actionButtonW + controlGap) * 3, S(108) + hintTopOffset, S(162), S(18), hwnd, reinterpret_cast<HMENU>(IDC_MONITOR_HINT), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Auto-enable when TeamSpeak starts", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP, contentLeft, S(170), S(320), S(28), hwnd, reinterpret_cast<HMENU>(IDC_AUTOSTART), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"Audio Source", WS_CHILD | WS_VISIBLE, labelX, S(252), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        HWND source = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS, fieldX, S(248), sourceW, S(360), hwnd, reinterpret_cast<HMENU>(IDC_SOURCE), nullptr, nullptr);
        SendMessageW(source, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), S(24));
        SendMessageW(source, CB_SETITEMHEIGHT, 0, S(22));
        SendMessageW(source, CB_SETMINVISIBLE, 18, 0);
        CreateWindowW(L"BUTTON", L"Refresh", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, fieldX + sourceW + controlGap, S(248), refreshButtonW, S(30), hwnd, reinterpret_cast<HMENU>(IDC_RESCAN), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"Audio Output (Mic)", WS_CHILD | WS_VISIBLE, labelX, S(296), S(140), S(24), hwnd, nullptr, nullptr, nullptr);
        HWND mic = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | WS_VSCROLL | CBS_DROPDOWNLIST, fieldX, S(292), micComboW, S(240), hwnd, reinterpret_cast<HMENU>(IDC_MIC_DEVICE), nullptr, nullptr);
        SendMessageW(mic, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), S(24));
        SendMessageW(mic, CB_SETITEMHEIGHT, 0, S(22));
        SendMessageW(mic, CB_SETMINVISIBLE, 12, 0);

        CreateWindowW(L"STATIC", L"Music Volume", WS_CHILD | WS_VISIBLE, labelX, S(388), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        HWND gain = CreateWindowW(
            TRACKBAR_CLASSW,
            L"",
            WS_CHILD | WS_VISIBLE | WS_CLIPSIBLINGS | TBS_AUTOTICKS,
            fieldX - gainSliderLeftNudge,
            S(384),
            gainSliderW + gainSliderLeftNudge,
            S(30),
            hwnd,
            reinterpret_cast<HMENU>(IDC_GAIN),
            nullptr,
            nullptr);
        SendMessageW(gain, TBM_SETRANGE, TRUE, MAKELONG(0, kMusicGainSliderMax));
        SendMessageW(gain, TBM_SETTICFREQ, 20, 0);
        SendMessageW(gain, TBM_SETPAGESIZE, 0, 10);
        SendMessageW(gain, TBM_SETLINESIZE, 0, 1);
        SendMessageW(gain, TBM_SETTHUMBLENGTH, static_cast<WPARAM>(S(18)), 0);
        SendMessageW(gain, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
        CreateWindowW(L"STATIC", L"-15.0 dB", WS_CHILD | WS_VISIBLE, valueTextX, S(388), S(102), S(24), hwnd, reinterpret_cast<HMENU>(IDC_GAIN_VALUE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Max is -2 dB to reduce clipping risk", WS_CHILD | WS_VISIBLE, fieldX, S(414), S(330), S(20), hwnd, reinterpret_cast<HMENU>(IDC_GAIN_HINT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Music Meter", WS_CHILD | WS_VISIBLE, labelX, S(438), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        g_rcMeter = { fieldX, S(436), fieldX + meterW, S(436) + S(28) };
        g_rcMusicClip = { fieldX, S(466), fieldX + meterW, S(466) + S(12) };
        CreateWindowW(L"STATIC", L"No signal", WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS, statusTextX, S(438), statusTextW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_METER_TEXT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"CLIP Events: 0", WS_CHILD | WS_VISIBLE, statusTextX, S(466), statusTextW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_CLIP_EVENTS), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Send music when mic is silent", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP, fieldX, S(480), forceTxToggleW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_FORCE_TX), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Mute mic input", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP, fieldX, S(512), muteToggleW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_INPUT_MUTE), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Mute music", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP, muteMusicToggleX, S(512), muteToggleW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUTE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Music keeps sending even when you are not speaking", WS_CHILD | WS_VISIBLE, fieldX, S(544), S(360), S(18), hwnd, reinterpret_cast<HMENU>(IDC_FORCE_TX_HINT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Mic Meter", WS_CHILD | WS_VISIBLE, labelX, S(568), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        g_rcMicMeter = { fieldX, S(566), fieldX + meterW, S(566) + S(28) };
        g_rcMicClip = { fieldX, S(596), fieldX + meterW, S(596) + S(12) };
        CreateWindowW(L"STATIC", L"No signal", WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS, statusTextX, S(568), statusTextW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_MIC_METER_TEXT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"CLIP Events: 0", WS_CHILD | WS_VISIBLE, statusTextX, S(596), statusTextW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_MIC_CLIP_EVENTS), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Only while connected", WS_CHILD | WS_VISIBLE, fieldX, S(610), S(180), S(18), hwnd, reinterpret_cast<HMENU>(IDC_MIC_METER_HINT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Music Mute Hotkey", WS_CHILD | WS_VISIBLE, labelX, S(632), S(150), S(24), hwnd, nullptr, nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Set...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, fieldX, S(628), S(160), S(30), hwnd, reinterpret_cast<HMENU>(IDC_MUTE_HOTKEY_SET), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Current: Not set", WS_CHILD | WS_VISIBLE, fieldX + S(172), S(632), S(290), S(24), hwnd, reinterpret_cast<HMENU>(IDC_MUTE_HOTKEY_TEXT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Mic Input Mute Hotkey", WS_CHILD | WS_VISIBLE, labelX, S(666), S(160), S(24), hwnd, nullptr, nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Set...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, fieldX, S(662), S(160), S(30), hwnd, reinterpret_cast<HMENU>(IDC_MIC_INPUT_HOTKEY_SET), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Current: Not set", WS_CHILD | WS_VISIBLE, fieldX + S(172), S(666), S(290), S(24), hwnd, reinterpret_cast<HMENU>(IDC_MIC_INPUT_HOTKEY_TEXT), nullptr, nullptr);

        ApplyControlTheme(hwnd);
        ApplyFonts(hwnd);
        PopulateCombos(hwnd);
        LoadSettings(hwnd);
        RequestSourceRefresh(hwnd, true);
        UpdateMonitorButton(hwnd);
        SendMessageW(hwnd, WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
        SetTimer(hwnd, kTimerStatusUpdate, kStatusTimerIntervalMs, nullptr);
        UpdateStatus(hwnd);
        return 0;
    }
    case WM_HSCROLL:
        if (reinterpret_cast<HWND>(lParam) == GetDlgItem(hwnd, IDC_GAIN)) {
            SendMessageW(GetDlgItem(hwnd, IDC_GAIN), WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEFOCUS), 0);
            UpdateGainLabel(hwnd);
            InvalidateRect(GetDlgItem(hwnd, IDC_GAIN_HINT), nullptr, TRUE);
            ApplyLiveSettings(hwnd, false, false);
        }
        return 0;
    case WM_MEASUREITEM: {
        auto* mis = reinterpret_cast<MEASUREITEMSTRUCT*>(lParam);
        if (!mis) {
            return FALSE;
        }
        if (mis->CtlID == IDC_SOURCE) {
            mis->itemHeight = static_cast<UINT>(S(24));
            return TRUE;
        }
        return FALSE;
    }
    case WM_DRAWITEM: {
        auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
        if (!dis) {
            return FALSE;
        }
        if (dis->CtlID == IDC_SOURCE) {
            DrawSourceComboItem(dis);
            return TRUE;
        }
        if (dis->CtlID == IDC_AUTOSTART || dis->CtlID == IDC_FORCE_TX ||
            dis->CtlID == IDC_MIC_INPUT_MUTE || dis->CtlID == IDC_MUTE) {
            DrawOwnerCheckbox(dis);
            return TRUE;
        }
        return FALSE;
    }
    case WM_NOTIFY: {
        auto* hdr = reinterpret_cast<NMHDR*>(lParam);
        if (!hdr) {
            return 0;
        }
        if (hdr->idFrom == IDC_GAIN && hdr->code == NM_CUSTOMDRAW) {
            return DrawGainSliderCustom(reinterpret_cast<NMCUSTOMDRAW*>(lParam));
        }
        return 0;
    }
    case kMsgSourceRefreshDone: {
        const uint64_t completedSeq = static_cast<uint64_t>(wParam);
        const uint64_t activeSeq = g_sourceRefreshActiveSeq.load(std::memory_order_acquire);
        if (completedSeq != 0 && activeSeq != 0 && completedSeq != activeSeq) {
            return 0;
        }
        JoinSourceRefreshThread();
        {
            std::lock_guard<std::mutex> lock(g_enumMutex);
            g_loopbacks = g_pendingLoopbacks;
            g_captureDevices = g_pendingCaptureDevices;
            g_apps = g_pendingApps;
        }
        PopulateCombos(hwnd);
        if (g_sourceRefreshPendingReload.exchange(false, std::memory_order_acq_rel)) {
            LoadSettings(hwnd);
        }
        g_sourceRefreshInFlight.store(false, std::memory_order_release);
        g_sourceRefreshActiveSeq.store(0, std::memory_order_release);

        if (g_sourceRefreshPending.exchange(false, std::memory_order_acq_rel)) {
            const uint64_t nextSeq = g_sourceRefreshSeq.fetch_add(1, std::memory_order_acq_rel) + 1ULL;
            g_sourceRefreshInFlight.store(true, std::memory_order_release);
            g_sourceRefreshActiveSeq.store(nextSeq, std::memory_order_release);
            SetStatusText(hwnd, L"Refreshing source list...");
            if (!StartSourceRefreshWorker(hwnd, nextSeq)) {
                g_sourceRefreshInFlight.store(false, std::memory_order_release);
                g_sourceRefreshActiveSeq.store(0, std::memory_order_release);
                HWND refreshBtn = GetDlgItem(hwnd, IDC_RESCAN);
                if (refreshBtn) {
                    EnableWindow(refreshBtn, TRUE);
                }
                SetStatusText(hwnd, L"Refresh failed to start");
            }
            return 0;
        }

        HWND refreshBtn = GetDlgItem(hwnd, IDC_RESCAN);
        if (refreshBtn) {
            EnableWindow(refreshBtn, TRUE);
        }
        UpdateStatus(hwnd);
        return 0;
    }
    case WM_TIMER:
        if (wParam == kTimerDebouncedSave) {
            FlushDebouncedSettingsSave(hwnd);
            return 0;
        }
        if (wParam == kTimerStatusUpdate) {
            const int tick = ++g_hotkeyRefreshTick;
            if ((tick % kHotkeyLabelRefreshEveryTicks) == 0) {
                UpdateHotkeyLabels(hwnd);
            }
            UpdateStatus(hwnd, (tick % kStatusDetailsEveryTicks) == 0);
            return 0;
        }
        return 0;
    case WM_KEYDOWN:
    case WM_SYSKEYDOWN:
        if (IsCapturingHotkey()) {
            const UINT vk = static_cast<UINT>(wParam);
            UINT* activeMods = nullptr;
            UINT* activeVk = nullptr;
            if (g_hotkeyCaptureTarget == HotkeyCaptureTarget::MusicMute) {
                activeMods = &g_muteHotkeyModifiers;
                activeVk = &g_muteHotkeyVk;
            } else if (g_hotkeyCaptureTarget == HotkeyCaptureTarget::MicInputMute) {
                activeMods = &g_micInputMuteHotkeyModifiers;
                activeVk = &g_micInputMuteHotkeyVk;
            }
            if (vk == VK_ESCAPE) {
                EndHotkeyCapture(hwnd);
                return 0;
            }
            if (vk == VK_DELETE || vk == VK_BACK) {
                EndHotkeyCapture(hwnd, false);
                if (activeMods) {
                    *activeMods = 0;
                }
                if (activeVk) {
                    *activeVk = 0;
                }
                ApplyLiveSettings(hwnd, false);
                UpdateHotkeyLabels(hwnd);
                return 0;
            }
            if (IsModifierOnlyKey(vk)) {
                return 0;
            }
            const UINT mods = ReadCurrentModifiers();
            if (g_hotkeyCaptureTarget == HotkeyCaptureTarget::MusicMute &&
                IsHotkeyConflict(mods, vk, g_micInputMuteHotkeyModifiers, g_micInputMuteHotkeyVk)) {
                EndHotkeyCapture(hwnd, false);
                MessageBeep(MB_ICONWARNING);
                SetStatusText(hwnd, L"Hotkey conflict: already assigned to Mic Input Mute.");
                UpdateHotkeyLabels(hwnd);
                return 0;
            }
            if (g_hotkeyCaptureTarget == HotkeyCaptureTarget::MicInputMute &&
                IsHotkeyConflict(mods, vk, g_muteHotkeyModifiers, g_muteHotkeyVk)) {
                EndHotkeyCapture(hwnd, false);
                MessageBeep(MB_ICONWARNING);
                SetStatusText(hwnd, L"Hotkey conflict: already assigned to Music Mute.");
                UpdateHotkeyLabels(hwnd);
                return 0;
            }
            EndHotkeyCapture(hwnd, false);
            if (activeMods) {
                *activeMods = mods;
            }
            if (activeVk) {
                *activeVk = vk;
            }
            ApplyLiveSettings(hwnd, false);
            UpdateHotkeyLabels(hwnd);
            return 0;
        }
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    case WM_ACTIVATE:
        if (LOWORD(wParam) == WA_INACTIVE && IsCapturingHotkey()) {
            EndHotkeyCapture(hwnd);
            SetStatusText(hwnd, L"Hotkey capture cancelled.");
            return 0;
        }
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_REPO_LINK:
            if (HIWORD(wParam) == STN_CLICKED) {
                ShellExecuteW(hwnd, L"open", kRepoUrl, nullptr, nullptr, SW_SHOWNORMAL);
                return 0;
            }
            break;
        case IDC_RESCAN:
            RequestSourceRefresh(hwnd, true);
            return 0;
        case IDC_START:
            MicMixApp::Instance().StartSource();
            UpdateStatus(hwnd);
            return 0;
        case IDC_STOP:
            MicMixApp::Instance().StopSource();
            UpdateStatus(hwnd);
            return 0;
        case IDC_SAVE:
            ApplyLiveSettings(hwnd, false, true);
            MicMixApp::Instance().StopSource();
            MicMixApp::Instance().StartSource();
            UpdateStatus(hwnd);
            return 0;
        case IDC_MONITOR:
            if (HIWORD(wParam) == BN_CLICKED) {
                MicMixApp::Instance().ToggleMonitor();
                UpdateMonitorButton(hwnd);
                return 0;
            }
            break;
        case IDC_MUTE_HOTKEY_SET:
            if (HIWORD(wParam) == BN_CLICKED) {
                BeginHotkeyCapture(hwnd, HotkeyCaptureTarget::MusicMute);
                return 0;
            }
            break;
        case IDC_MIC_INPUT_HOTKEY_SET:
            if (HIWORD(wParam) == BN_CLICKED) {
                BeginHotkeyCapture(hwnd, HotkeyCaptureTarget::MicInputMute);
                return 0;
            }
            break;
        case IDC_SOURCE:
            if (HIWORD(wParam) == CBN_SELCHANGE) {
                const int sourceSel = static_cast<int>(SendMessageW(GetDlgItem(hwnd, IDC_SOURCE), CB_GETCURSEL, 0, 0));
                const LRESULT data = SendMessageW(GetDlgItem(hwnd, IDC_SOURCE), CB_GETITEMDATA, static_cast<WPARAM>(sourceSel), 0);
                if (data < 0 || data == CB_ERR || static_cast<size_t>(data) >= g_sourceChoices.size()) {
                    if (g_lastValidSourceSel >= 0) {
                        SendMessageW(GetDlgItem(hwnd, IDC_SOURCE), CB_SETCURSEL, g_lastValidSourceSel, 0);
                    }
                    return 0;
                }
                g_lastValidSourceSel = sourceSel;
                ApplyLiveSettings(hwnd, true);
                return 0;
            }
            break;
        case IDC_MIC_DEVICE:
            if (HIWORD(wParam) == CBN_SELCHANGE) {
                ApplyLiveSettings(hwnd, false);
                return 0;
            }
            break;
        case IDC_FORCE_TX:
        case IDC_MIC_INPUT_MUTE:
        case IDC_MUTE:
        case IDC_AUTOSTART:
            if (HIWORD(wParam) == BN_CLICKED) {
                const int id = LOWORD(wParam);
                if ((id == IDC_MIC_INPUT_MUTE || id == IDC_MUTE) &&
                    IsWindowEnabled(GetDlgItem(hwnd, id)) == FALSE) {
                    return 0;
                }
                if (IsOwnerCheckboxControlId(id)) {
                    SetOwnerCheckboxValue(id, !GetOwnerCheckboxValue(id));
                    InvalidateRect(GetDlgItem(hwnd, id), nullptr, TRUE);
                }
                ApplyLiveSettings(hwnd, false);
                return 0;
            }
            break;
        default:
            break;
        }
        return 0;
    case WM_CTLCOLORSTATIC: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        HWND ctl = reinterpret_cast<HWND>(lParam);
        const int id = GetDlgCtrlID(ctl);
        RECT rc{};
        GetWindowRect(ctl, &rc);
        MapWindowPoints(nullptr, hwnd, reinterpret_cast<LPPOINT>(&rc), 2);
        POINT pt{ rc.left, rc.top };
        const bool inCard = PtInRect(&g_rcSource, pt) || PtInRect(&g_rcMix, pt) || PtInRect(&g_rcSession, pt);
        SetBkMode(hdc, TRANSPARENT);
        if (id == IDC_TITLE) {
            SetTextColor(hdc, g_theme.text);
        } else if (id == IDC_SUBTITLE || id == IDC_VERSION) {
            SetTextColor(hdc, g_theme.muted);
        } else if (id == IDC_REPO_LINK) {
            SetTextColor(hdc, RGB(25, 100, 200));
        } else if (id == IDC_MONITOR_HINT || id == IDC_MIC_METER_HINT || id == IDC_GAIN_HINT || id == IDC_FORCE_TX_HINT ||
                   id == IDC_MUSIC_CLIP_EVENTS || id == IDC_MIC_CLIP_EVENTS) {
            SetTextColor(hdc, g_theme.muted);
        } else {
            SetTextColor(hdc, g_theme.text);
        }
        return reinterpret_cast<INT_PTR>(inCard ? g_brushCard : g_brushBg);
    }
    case WM_SETCURSOR: {
        const HWND target = reinterpret_cast<HWND>(wParam);
        const UINT hit = LOWORD(lParam);
        if (target && hit == HTCLIENT) {
            const int id = GetDlgCtrlID(target);
            if (id == IDC_GAIN) {
                POINT cursorPt{};
                GetCursorPos(&cursorPt);
                const bool overThumb = IsPointOverGainThumb(target, cursorPt);
                const bool draggingThumb = (GetCapture() == target) && ((GetKeyState(VK_LBUTTON) & 0x8000) != 0);
                if (overThumb || draggingThumb) {
                    SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
                    return TRUE;
                }
            }
            if (IsHandCursorControlId(id)) {
                SetCursor(LoadCursorW(nullptr, IDC_HAND));
                return TRUE;
            }
        }
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    case WM_PAINT: {
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT clientRc{};
        GetClientRect(hwnd, &clientRc);

        HDC paintDc = hdc;
        HDC memDc = nullptr;
        HBITMAP memBmp = nullptr;
        HGDIOBJ oldBmp = nullptr;
        bool bufferedPaint = false;
        const int clientW = std::max(0, static_cast<int>(clientRc.right - clientRc.left));
        const int clientH = std::max(0, static_cast<int>(clientRc.bottom - clientRc.top));
        if (clientW > 0 && clientH > 0) {
            memDc = CreateCompatibleDC(hdc);
            if (memDc) {
                memBmp = CreateCompatibleBitmap(hdc, clientW, clientH);
                if (memBmp) {
                    oldBmp = SelectObject(memDc, memBmp);
                    paintDc = memDc;
                    bufferedPaint = true;
                } else {
                    DeleteDC(memDc);
                    memDc = nullptr;
                }
            }
        }

        const RECT bgFillRc = bufferedPaint ? clientRc : ps.rcPaint;
        FillRect(paintDc, &bgFillRc, g_brushBg);

        RECT accent{ S(0), S(0), S(kClientWidthPx), S(4) };
        HBRUSH accentBrush = CreateSolidBrush(g_theme.accent);
        FillRect(paintDc, &accent, accentBrush);
        DeleteObject(accentBrush);
        DrawHeaderStatusBadge(paintDc);

        DrawCard(paintDc, g_rcSource);
        DrawCard(paintDc, g_rcMix);
        DrawCard(paintDc, g_rcSession);
        DrawLevelMeter(
            paintDc,
            g_rcMeter,
            IsSignalMeterActive(
                g_lastTelemetry.musicActive,
                g_lastTelemetry.musicSendPeakDbfs,
                g_lastTelemetry.musicPeakDbfs,
                g_lastTelemetry.musicRmsDbfs),
            g_meterVisualDb,
            g_meterHoldDb);
        DrawLevelMeter(paintDc, g_rcMicMeter, g_lastTelemetry.micRmsDbfs > -119.0f, g_micMeterVisualDb, g_micMeterHoldDb);
        const float shownMusicPeak = (g_lastTelemetry.musicSendPeakDbfs > -119.0f) ? g_lastTelemetry.musicSendPeakDbfs : g_lastTelemetry.musicPeakDbfs;
        DrawClipStrip(
            paintDc, g_rcMusicClip, ::ClipDangerUnitFromDb(shownMusicPeak),
            g_lastTelemetry.sourceClipRecent);
        DrawClipStrip(
            paintDc, g_rcMicClip, ::ClipDangerUnitFromDb(g_lastTelemetry.micPeakDbfs),
            g_lastTelemetry.micClipRecent);

        const SectionHeading sectionHeadings[] = {
            { g_rcSession, L"MICMIX SETTINGS" },
            { g_rcSource, L"AUDIO ROUTING" },
            { g_rcMix, L"MIX BEHAVIOR" },
        };
        DrawSectionHeadings(
            paintDc,
            sectionHeadings,
            static_cast<int>(std::size(sectionHeadings)),
            g_fontSmall,
            g_fontBody,
            g_theme.muted,
            g_dpi);

        SetBkMode(paintDc, TRANSPARENT);
        SetTextColor(paintDc, RGB(40, 48, 61));
        HGDIOBJ oldStatusFont = SelectObject(paintDc, g_fontMono ? g_fontMono : (g_fontBody ? g_fontBody : GetStockObject(DEFAULT_GUI_FONT)));
        RECT statusRc = g_rcStatus;
        DrawTextW(
            paintDc,
            g_lastStatusText.empty() ? L"State: Stopped" : g_lastStatusText.c_str(),
            -1,
            &statusRc,
            DT_LEFT | DT_TOP | DT_NOPREFIX | DT_WORDBREAK);
        SelectObject(paintDc, oldStatusFont);

        if (bufferedPaint) {
            const int copyW = std::max(0, static_cast<int>(ps.rcPaint.right - ps.rcPaint.left));
            const int copyH = std::max(0, static_cast<int>(ps.rcPaint.bottom - ps.rcPaint.top));
            if (copyW > 0 && copyH > 0) {
                BitBlt(hdc, ps.rcPaint.left, ps.rcPaint.top, copyW, copyH, memDc, ps.rcPaint.left, ps.rcPaint.top, SRCCOPY);
            }
            SelectObject(memDc, oldBmp);
            DeleteObject(memBmp);
            DeleteDC(memDc);
        }
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;
    case WM_CLOSE:
        EndHotkeyCapture(hwnd, false);
        FlushDebouncedSettingsSave(hwnd);
        if (MicMixApp::Instance().IsMonitorEnabled()) {
            MicMixApp::Instance().SetMonitorEnabled(false);
        }
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
        EndHotkeyCapture(hwnd, false);
        if (MicMixApp::Instance().IsMonitorEnabled()) {
            MicMixApp::Instance().SetMonitorEnabled(false);
        }
        g_hwnd.store(nullptr, std::memory_order_release);
        g_sourceRefreshInFlight.store(false, std::memory_order_release);
        g_sourceRefreshActiveSeq.store(0, std::memory_order_release);
        g_sourceRefreshPending.store(false, std::memory_order_release);
        g_sourceRefreshPendingReload.store(false, std::memory_order_release);
        g_saveDebouncePending.store(false, std::memory_order_release);
        g_hotkeyRefreshTick = 0;
        g_hotkeyCaptureTarget = HotkeyCaptureTarget::None;
        g_lastMusicClipEvents = std::numeric_limits<uint64_t>::max();
        g_lastMicClipEvents = std::numeric_limits<uint64_t>::max();
        g_lastMusicMeterLevelPx = -1;
        g_lastMusicMeterHoldPx = -1;
        g_lastMusicMeterActive = false;
        g_lastMicMeterLevelPx = -1;
        g_lastMicMeterHoldPx = -1;
        g_lastMicMeterActive = false;
        g_lastMusicClipLitSegments = -1;
        g_lastMusicClipRecent = false;
        g_lastMicClipLitSegments = -1;
        g_lastMicClipRecent = false;
        g_lastMusicMeterText.clear();
        g_lastMicMeterText.clear();
        g_headerStatusBadgeState = HeaderStatusBadgeState::Off;
        g_windowTitleActive = false;
        g_windowTitleStateValid = false;
        g_lastStatusText.clear();
        KillTimer(hwnd, kTimerStatusUpdate);
        KillTimer(hwnd, kTimerDebouncedSave);
        ReleaseUiResources();
        PostQuitMessage(0);
        return 0;
    default:
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
}

void WindowThreadMain() {
    INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_STANDARD_CLASSES | ICC_BAR_CLASSES };
    InitCommonControlsEx(&icc);

    HINSTANCE hInst = ResolveModuleHandleFromAddress(reinterpret_cast<const void*>(&WindowThreadMain));
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.lpszClassName = L"MicMixSettingsWindow";
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APP_ICON));
    wc.hIconSm = LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APP_ICON_SMALL));
    wc.hbrBackground = nullptr;
    UnregisterClassW(wc.lpszClassName, hInst);
    RegisterClassExW(&wc);

    HWND hwnd = CreateWindowExW(
        0,
        wc.lpszClassName,
        L"MicMix Settings",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        []() {
            RECT rc{ 0, 0, S(kClientWidthPx), S(kClientHeightPx) };
            AdjustWindowRectEx(&rc, WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX, FALSE, 0);
            return rc.right - rc.left;
        }(),
        []() {
            RECT rc{ 0, 0, S(kClientWidthPx), S(kClientHeightPx) };
            AdjustWindowRectEx(&rc, WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX, FALSE, 0);
            return rc.bottom - rc.top;
        }(),
        nullptr,
        nullptr,
        hInst,
        nullptr);
    if (!hwnd) {
        g_running.store(false, std::memory_order_release);
        return;
    }
    SendMessageW(hwnd, WM_SETICON, ICON_BIG, reinterpret_cast<LPARAM>(LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APP_ICON))));
    SendMessageW(hwnd, WM_SETICON, ICON_SMALL, reinterpret_cast<LPARAM>(LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APP_ICON_SMALL))));
    g_hwnd.store(hwnd, std::memory_order_release);
    ShowWindow(hwnd, SW_SHOW);
    UpdateWindow(hwnd);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
        HWND mainHwnd = g_hwnd.load(std::memory_order_acquire);
        if (mainHwnd &&
            IsCapturingHotkey() &&
            msg.hwnd != mainHwnd &&
            (msg.message == WM_KEYDOWN || msg.message == WM_SYSKEYDOWN)) {
            SendMessageW(mainHwnd, msg.message, msg.wParam, msg.lParam);
            continue;
        }
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    g_running.store(false, std::memory_order_release);
}

} // namespace

SettingsWindowController& SettingsWindowController::Instance() {
    static SettingsWindowController instance;
    return instance;
}

void SettingsWindowController::Open() {
    std::lock_guard<std::mutex> lock(g_mutex);
    if (g_thread.joinable() && !g_running.load(std::memory_order_acquire)) {
        g_thread.join();
    }
    if (g_running.load(std::memory_order_acquire)) {
        HWND hwnd = g_hwnd.load(std::memory_order_acquire);
        if (hwnd) {
            ShowWindow(hwnd, SW_RESTORE);
            SetForegroundWindow(hwnd);
        }
        return;
    }
    g_running.store(true, std::memory_order_release);
    try {
        g_thread = std::thread(WindowThreadMain);
    } catch (const std::exception& ex) {
        g_running.store(false, std::memory_order_release);
        LogError(std::string("settings_window thread start failed: ") + ex.what());
    } catch (...) {
        g_running.store(false, std::memory_order_release);
        LogError("settings_window thread start failed: unknown");
    }
}

void SettingsWindowController::Close() {
    std::lock_guard<std::mutex> lock(g_mutex);
    if (MicMixApp::Instance().IsMonitorEnabled()) {
        MicMixApp::Instance().SetMonitorEnabled(false);
    }
    if (!g_running.load(std::memory_order_acquire)) {
        if (g_thread.joinable()) g_thread.join();
        JoinSourceRefreshThread();
        return;
    }
    HWND hwnd = g_hwnd.load(std::memory_order_acquire);
    if (hwnd) {
        PostMessageW(hwnd, WM_CLOSE, 0, 0);
    }
    if (g_thread.joinable()) {
        g_thread.join();
    }
    JoinSourceRefreshThread();
}
