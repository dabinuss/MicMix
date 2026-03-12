#include "settings_window.h"

#include <Windows.h>
#include <CommCtrl.h>
#include <shellapi.h>
#include <uxtheme.h>

#include <atomic>
#include <algorithm>
#include <cstdio>
#include <cwctype>
#include <exception>
#include <mutex>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <string>

#include "micmix_core.h"
#include "micmix_version.h"
#include "resource.h"

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
    IDC_MUTE = 1010,
    IDC_AUTOSTART = 1011,
    IDC_START = 1012,
    IDC_STOP = 1013,
    IDC_SAVE = 1014,
    IDC_STATUS = 1015,
    IDC_METER_TEXT = 1019,
    IDC_MUTE_HOTKEY_SET = 1020,
    IDC_MUTE_HOTKEY_TEXT = 1021,
    IDC_MIC_DEVICE = 1024,
    IDC_MIC_METER_TEXT = 1025,
    IDC_MONITOR = 1026,
    IDC_MONITOR_HINT = 1027,
    IDC_MIC_METER_HINT = 1028,
    IDC_GAIN_HINT = 1029,
    IDC_FORCE_TX_HINT = 1030,
    IDC_VERSION = 1031,
};

enum class SourceChoiceType {
    Loopback,
    App,
};

struct SourceChoice {
    SourceChoiceType type = SourceChoiceType::Loopback;
    std::string id;
    std::string processName;
    uint32_t pid = 0;
};

struct UiTheme {
    COLORREF bg = RGB(241, 244, 248);
    COLORREF card = RGB(255, 255, 255);
    COLORREF border = RGB(212, 218, 229);
    COLORREF text = RGB(30, 33, 41);
    COLORREF muted = RGB(96, 105, 119);
    COLORREF accent = RGB(33, 112, 218);
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
UiTheme g_theme;
int g_dpi = 96;
bool g_loadingUi = false;
int g_hotkeyRefreshTick = 0;
bool g_waitingForHotkey = false;
UINT g_muteHotkeyModifiers = 0;
UINT g_muteHotkeyVk = 0;
std::vector<LoopbackDeviceInfo> g_pendingLoopbacks;
std::vector<CaptureDeviceInfo> g_pendingCaptureDevices;
std::vector<AppProcessInfo> g_pendingApps;
std::atomic_uint64_t g_sourceRefreshSeq{0};
std::atomic<bool> g_sourceRefreshReloadSettings{false};
std::unordered_map<uint32_t, HICON> g_appIconsByPid;
HICON g_loopbackFallbackIcon = nullptr;
HICON g_appFallbackIcon = nullptr;

constexpr INT_PTR kSourceItemHeader = -10;
constexpr INT_PTR kSourceItemDivider = -11;
constexpr INT_PTR kSourceItemPlaceholder = -12;
constexpr UINT kMsgSourceRefreshDone = WM_APP + 0x31;
constexpr float kMusicGainMinDb = -30.0f;
constexpr float kMusicGainMaxDb = -2.0f;
constexpr int kMusicGainStepPerDb = 10;
constexpr int kMusicGainSliderMax = static_cast<int>((kMusicGainMaxDb - kMusicGainMinDb) * static_cast<float>(kMusicGainStepPerDb));
constexpr int kSourceIconSizePx = 16;
constexpr int kClientWidthPx = 680;
constexpr int kClientHeightPx = 690;
constexpr int kCardMarginPx = 16;
constexpr int kCardGapPx = 12;
constexpr int kCardInnerPaddingPx = 20;

HFONT g_fontBody = nullptr;
HFONT g_fontSmall = nullptr;
HFONT g_fontHint = nullptr;
HFONT g_fontTitle = nullptr;
HFONT g_fontMono = nullptr;
HBRUSH g_brushBg = nullptr;
HBRUSH g_brushCard = nullptr;

RECT g_rcSource{};
RECT g_rcMix{};
RECT g_rcSession{};
RECT g_rcMeter{};
RECT g_rcMicMeter{};
float g_meterVisualDb = -60.0f;
float g_meterHoldDb = -60.0f;
float g_micMeterVisualDb = -60.0f;
float g_micMeterHoldDb = -60.0f;
TelemetrySnapshot g_lastTelemetry{};
std::wstring g_lastStatusText;

int S(int px) {
    return MulDiv(px, g_dpi, 96);
}

std::wstring Utf8ToWide(const std::string& text) {
    if (text.empty()) return {};
    int len = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, nullptr, 0);
    std::wstring out(static_cast<size_t>(len), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, out.data(), len);
    if (!out.empty() && out.back() == L'\0') out.pop_back();
    return out;
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

HINSTANCE GetPluginModuleHandle() {
    HMODULE mod = nullptr;
    if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                           GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                           reinterpret_cast<LPCWSTR>(&GetPluginModuleHandle),
                           &mod) == 0 || mod == nullptr) {
        return GetModuleHandleW(nullptr);
    }
    return mod;
}

void EnsureUiResources() {
    if (!g_brushBg) g_brushBg = CreateSolidBrush(g_theme.bg);
    if (!g_brushCard) g_brushCard = CreateSolidBrush(g_theme.card);
    if (!g_fontBody) g_fontBody = CreateFontW(-S(13), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!g_fontSmall) g_fontSmall = CreateFontW(-S(12), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!g_fontHint) g_fontHint = CreateFontW(-S(11), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!g_fontTitle) g_fontTitle = CreateFontW(-S(24), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI Semibold");
    if (!g_fontMono) g_fontMono = CreateFontW(-S(13), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, FIXED_PITCH, L"Consolas");
    EnsureFallbackIcons();
}

void ReleaseUiResources() {
    ClearAppIconCache();
    DestroyIconHandle(g_loopbackFallbackIcon);
    DestroyIconHandle(g_appFallbackIcon);
    if (g_fontBody) { DeleteObject(g_fontBody); g_fontBody = nullptr; }
    if (g_fontSmall) { DeleteObject(g_fontSmall); g_fontSmall = nullptr; }
    if (g_fontHint) { DeleteObject(g_fontHint); g_fontHint = nullptr; }
    if (g_fontTitle) { DeleteObject(g_fontTitle); g_fontTitle = nullptr; }
    if (g_fontMono) { DeleteObject(g_fontMono); g_fontMono = nullptr; }
    if (g_brushBg) { DeleteObject(g_brushBg); g_brushBg = nullptr; }
    if (g_brushCard) { DeleteObject(g_brushCard); g_brushCard = nullptr; }
}

void SetControlFont(HWND hwnd, int id, HFONT font) {
    HWND ctl = GetDlgItem(hwnd, id);
    if (ctl && font) SendMessageW(ctl, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
}

void SetStatusText(HWND hwnd, const std::wstring& text) {
    if (text == g_lastStatusText) {
        return;
    }
    g_lastStatusText = text;
    SetWindowTextW(GetDlgItem(hwnd, IDC_STATUS), text.c_str());
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

int MeterDbToPixels(float dbfs, int widthPx) {
    const float clamped = std::clamp(dbfs, -60.0f, 0.0f);
    return static_cast<int>(((clamped + 60.0f) / 60.0f) * static_cast<float>(widthPx));
}

bool IsMusicMeterActive(const TelemetrySnapshot& t) {
    const float shownPeakDb = (t.musicSendPeakDbfs > -119.0f) ? t.musicSendPeakDbfs : t.musicPeakDbfs;
    const bool levelSuggestsSignal = (shownPeakDb > -96.0f) || (t.musicRmsDbfs > -100.0f);
    return t.musicActive || levelSuggestsSignal;
}

void UpdateMusicMeter(HWND hwnd) {
    HWND txt = GetDlgItem(hwnd, IDC_METER_TEXT);
    if (!txt) {
        return;
    }
    const TelemetrySnapshot t = g_lastTelemetry;
    const float shownPeakDb = (t.musicSendPeakDbfs > -119.0f) ? t.musicSendPeakDbfs : t.musicPeakDbfs;
    const bool meterActive = IsMusicMeterActive(t);

    float targetDb = -60.0f;
    if (meterActive) {
        targetDb = std::clamp(shownPeakDb, -60.0f, 0.0f);
    }
    if (targetDb > g_meterVisualDb) {
        g_meterVisualDb += (targetDb - g_meterVisualDb) * 0.55f;
    } else {
        g_meterVisualDb += (targetDb - g_meterVisualDb) * 0.22f;
    }
    g_meterVisualDb = std::clamp(g_meterVisualDb, -60.0f, 0.0f);

    if (targetDb >= g_meterHoldDb) {
        g_meterHoldDb = targetDb;
    } else {
        g_meterHoldDb = std::max(-60.0f, g_meterHoldDb - 1.2f);
    }

    if (!meterActive) {
        SetWindowTextW(txt, L"No signal");
    } else {
        const wchar_t* grade = L"ok";
        if (shownPeakDb > -8.0f) {
            grade = L"loud";
        } else if (shownPeakDb > -16.0f) {
            grade = L"hot";
        } else if (shownPeakDb < -35.0f) {
            grade = L"quiet";
        }
        wchar_t meter[96];
        swprintf_s(meter, L"%.1f dBFS (%s)", shownPeakDb, grade);
        SetWindowTextW(txt, meter);
    }

    InvalidateRect(hwnd, &g_rcMeter, FALSE);
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
        g_micMeterVisualDb += (targetDb - g_micMeterVisualDb) * 0.52f;
    } else {
        g_micMeterVisualDb += (targetDb - g_micMeterVisualDb) * 0.20f;
    }
    g_micMeterVisualDb = std::clamp(g_micMeterVisualDb, -60.0f, 0.0f);

    if (targetDb >= g_micMeterHoldDb) {
        g_micMeterHoldDb = targetDb;
    } else {
        g_micMeterHoldDb = std::max(-60.0f, g_micMeterHoldDb - 1.4f);
    }

    if (t.micRmsDbfs <= -119.0f) {
        SetWindowTextW(txt, L"No signal");
    } else {
        wchar_t meter[64];
        swprintf_s(meter, L"%.1f dBFS", t.micRmsDbfs);
        SetWindowTextW(txt, meter);
    }
    InvalidateRect(hwnd, &g_rcMicMeter, FALSE);
}

void UpdateMonitorButton(HWND hwnd) {
    HWND btn = GetDlgItem(hwnd, IDC_MONITOR);
    if (!btn) {
        return;
    }
    const bool enabled = MicMixApp::Instance().IsMonitorEnabled();
    SetWindowTextW(btn, enabled ? L"Monitor Mix: On" : L"Monitor Mix: Off");
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

void UpdateMuteHotkeyLabel(HWND hwnd) {
    if (g_waitingForHotkey) {
        return;
    }
    HWND label = GetDlgItem(hwnd, IDC_MUTE_HOTKEY_TEXT);
    if (!label) {
        return;
    }
    const std::wstring text = L"Current: " + FormatHotkeyText(g_muteHotkeyModifiers, g_muteHotkeyVk);
    SetWindowTextW(label, text.c_str());
}

void BeginHotkeyCapture(HWND hwnd) {
    g_waitingForHotkey = true;
    SetWindowTextW(GetDlgItem(hwnd, IDC_MUTE_HOTKEY_TEXT), L"Current: Press key... (Esc=Cancel, Del=Clear)");
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

void RequestSourceRefresh(HWND hwnd, bool reloadSettings) {
    if (!hwnd) {
        return;
    }
    const uint64_t seq = g_sourceRefreshSeq.fetch_add(1, std::memory_order_acq_rel) + 1ULL;
    if (reloadSettings) {
        g_sourceRefreshReloadSettings.store(true, std::memory_order_release);
    }
    HWND refreshBtn = GetDlgItem(hwnd, IDC_RESCAN);
    if (refreshBtn) {
        EnableWindow(refreshBtn, FALSE);
    }
    SetStatusText(hwnd, L"Refreshing source list...");

    JoinSourceRefreshThread();
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
    SendMessageW(GetDlgItem(hwnd, IDC_FORCE_TX), BM_SETCHECK, s.forceTxEnabled ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(hwnd, IDC_MUTE), BM_SETCHECK, s.musicMuted ? BST_CHECKED : BST_UNCHECKED, 0);
    SendMessageW(GetDlgItem(hwnd, IDC_AUTOSTART), BM_SETCHECK, s.autostartEnabled ? BST_CHECKED : BST_UNCHECKED, 0);
    g_muteHotkeyModifiers = static_cast<UINT>(std::max(0, s.muteHotkeyModifiers));
    g_muteHotkeyVk = static_cast<UINT>(std::max(0, s.muteHotkeyVk));
    UpdateGainLabel(hwnd);
    UpdateMuteHotkeyLabel(hwnd);
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
    s.forceTxEnabled = SendMessageW(GetDlgItem(hwnd, IDC_FORCE_TX), BM_GETCHECK, 0, 0) == BST_CHECKED;
    s.musicMuted = SendMessageW(GetDlgItem(hwnd, IDC_MUTE), BM_GETCHECK, 0, 0) == BST_CHECKED;
    s.autostartEnabled = SendMessageW(GetDlgItem(hwnd, IDC_AUTOSTART), BM_GETCHECK, 0, 0) == BST_CHECKED;
    s.muteHotkeyModifiers = static_cast<int>(g_muteHotkeyModifiers);
    s.muteHotkeyVk = static_cast<int>(g_muteHotkeyVk);
    return s;
}

void UpdateStatus(HWND hwnd) {
    auto& app = MicMixApp::Instance();
    const MicMixSettings s = app.GetSettings();
    const SourceStatus st = app.GetSourceStatus();
    const TelemetrySnapshot t = app.GetTelemetry();
    HWND muteCtl = GetDlgItem(hwnd, IDC_MUTE);
    if (muteCtl) {
        const LRESULT current = SendMessageW(muteCtl, BM_GETCHECK, 0, 0);
        const LRESULT want = s.musicMuted ? BST_CHECKED : BST_UNCHECKED;
        if (current != want) {
            g_loadingUi = true;
            SendMessageW(muteCtl, BM_SETCHECK, static_cast<WPARAM>(want), 0);
            g_loadingUi = false;
        }
    }
    g_lastTelemetry = t;
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
        "  " + sendBuf;
    SetStatusText(hwnd, Utf8ToWide(line1 + "\r\n" + line2));
    UpdateMusicMeter(hwnd);
    UpdateMicMeter(hwnd);
    UpdateMonitorButton(hwnd);
}

void ApplyLiveSettings(HWND hwnd, bool restartSource) {
    if (g_loadingUi) return;
    auto s = CollectSettings(hwnd);
    MicMixApp::Instance().ApplySettings(s, restartSource, true);
    UpdateStatus(hwnd);
}

void ApplyControlTheme(HWND hwnd) {
    EnumChildWindows(hwnd, [](HWND child, LPARAM) -> BOOL {
        wchar_t className[32]{};
        GetClassNameW(child, className, static_cast<int>(std::size(className)));
        if (_wcsicmp(className, WC_COMBOBOXW) == 0 ||
            _wcsicmp(className, WC_BUTTONW) == 0 ||
            _wcsicmp(className, TRACKBAR_CLASSW) == 0) {
            // Keep controls on native Windows visual style (no Explorer override)
            // to stay closer to TS3's neutral UI widgets.
            SetWindowTheme(child, nullptr, nullptr);
        }
        return TRUE;
    }, 0);
}

void ApplyFonts(HWND hwnd) {
    EnumChildWindows(hwnd, [](HWND child, LPARAM) -> BOOL {
        SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(g_fontBody), TRUE);
        return TRUE;
    }, 0);
    SetControlFont(hwnd, IDC_TITLE, g_fontTitle);
    SetControlFont(hwnd, IDC_SUBTITLE, g_fontSmall);
    SetControlFont(hwnd, IDC_VERSION, g_fontSmall);
    SetControlFont(hwnd, IDC_METER_TEXT, g_fontSmall);
    SetControlFont(hwnd, IDC_MIC_METER_TEXT, g_fontSmall);
    SetControlFont(hwnd, IDC_MONITOR_HINT, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFont(hwnd, IDC_MIC_METER_HINT, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFont(hwnd, IDC_GAIN_HINT, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFont(hwnd, IDC_FORCE_TX_HINT, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFont(hwnd, IDC_STATUS, g_fontMono);
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
    FillRect(hdc, &rc, g_brushCard);
    HPEN pen = CreatePen(PS_SOLID, 1, g_theme.border);
    HGDIOBJ oldPen = SelectObject(hdc, pen);
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(HOLLOW_BRUSH));
    Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
    SelectObject(hdc, oldBrush);
    SelectObject(hdc, oldPen);
    DeleteObject(pen);
}

void FillSolidRect(HDC hdc, const RECT& rc, COLORREF color) {
    HBRUSH b = CreateSolidBrush(color);
    FillRect(hdc, &rc, b);
    DeleteObject(b);
}

void DrawLevelMeter(HDC hdc, const RECT& meterRect, bool active, float visualDb, float holdDb) {
    if (meterRect.right <= meterRect.left || meterRect.bottom <= meterRect.top) {
        return;
    }
    RECT rc = meterRect;
    FillSolidRect(hdc, rc, RGB(235, 240, 246));

    const int width = rc.right - rc.left;
    const int levelPx = std::clamp(MeterDbToPixels(visualDb, width), 0, width);
    const int holdPx = std::clamp(MeterDbToPixels(holdDb, width), 0, width - 1);
    // Intentionally stricter: show warning zones earlier to encourage lower music levels.
    const int greenEnd = static_cast<int>(width * 0.55f);
    const int yellowEnd = static_cast<int>(width * 0.78f);

    if (active && levelPx > 0) {
        RECT seg = rc;
        seg.right = rc.left + std::min(levelPx, greenEnd);
        if (seg.right > seg.left) FillSolidRect(hdc, seg, RGB(54, 181, 78));

        if (levelPx > greenEnd) {
            seg.left = rc.left + greenEnd;
            seg.right = rc.left + std::min(levelPx, yellowEnd);
            if (seg.right > seg.left) FillSolidRect(hdc, seg, RGB(232, 191, 58));
        }
        if (levelPx > yellowEnd) {
            seg.left = rc.left + yellowEnd;
            seg.right = rc.left + levelPx;
            if (seg.right > seg.left) FillSolidRect(hdc, seg, RGB(214, 75, 63));
        }

        HPEN holdPen = CreatePen(PS_SOLID, 1, RGB(255, 255, 255));
        HGDIOBJ oldPen = SelectObject(hdc, holdPen);
        MoveToEx(hdc, rc.left + holdPx, rc.top, nullptr);
        LineTo(hdc, rc.left + holdPx, rc.bottom);
        SelectObject(hdc, oldPen);
        DeleteObject(holdPen);
    }

    HPEN border = CreatePen(PS_SOLID, 1, RGB(184, 193, 207));
    HGDIOBJ oldPen = SelectObject(hdc, border);
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(HOLLOW_BRUSH));
    Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
    SelectObject(hdc, oldBrush);
    SelectObject(hdc, oldPen);
    DeleteObject(border);
}

void DrawSourceComboItem(const DRAWITEMSTRUCT* dis) {
    if (!dis || dis->itemID == static_cast<UINT>(-1)) {
        return;
    }
    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool disabled = (dis->itemState & ODS_DISABLED) != 0;

    wchar_t textBuf[512]{};
    SendMessageW(dis->hwndItem, CB_GETLBTEXT, dis->itemID, reinterpret_cast<LPARAM>(textBuf));

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
    DrawTextW(dis->hDC, textBuf, -1, &textRc, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);

    if ((dis->itemState & ODS_FOCUS) != 0) {
        RECT focus = rc;
        focus.left += S(2);
        focus.right -= S(2);
        DrawFocusRect(dis->hDC, &focus);
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        g_dpi = GetDpiForWindow(hwnd);
        EnsureUiResources();
        ComputeLayout();
        const std::wstring versionText = Utf8ToWide(std::string("v") + MICMIX_VERSION);

        const int contentLeft = S(kCardMarginPx + kCardInnerPaddingPx);
        const int contentRight = S(kClientWidthPx - kCardMarginPx - kCardInnerPaddingPx);
        const int controlGap = S(12);
        const int labelX = contentLeft;
        const int fieldX = S(180);
        const int actionButtonW = S(142);
        const int monitorButtonW = S(122);
        const int refreshButtonW = S(130);
        const int gainSliderW = S(360);
        const int meterW = S(340);
        const int sourceW = contentRight - fieldX - controlGap - refreshButtonW;
        const int micComboW = contentRight - fieldX;
        const int statusTextX = fieldX + gainSliderW;
        const int statusTextW = contentRight - statusTextX;
        const int valueTextX = fieldX + gainSliderW;
        const int muteToggleX = fieldX + S(262);
        const int hintTopOffset = S(38);
        const int headerVersionW = S(120);
        const int headerVersionX = contentRight - headerVersionW;
        const int statusX = contentLeft;
        const int statusY = S(610);
        const int statusW = contentRight - statusX;
        const int statusH = S(44);

        CreateWindowW(L"STATIC", L"MicMix", WS_CHILD | WS_VISIBLE, contentLeft, S(14), S(220), S(36), hwnd, reinterpret_cast<HMENU>(IDC_TITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Configure MicMix and route one audio source to your mic output", WS_CHILD | WS_VISIBLE, contentLeft, S(44), headerVersionX - contentLeft - S(8), S(20), hwnd, reinterpret_cast<HMENU>(IDC_SUBTITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", versionText.c_str(), WS_CHILD | WS_VISIBLE | SS_RIGHT, headerVersionX, S(18), headerVersionW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_VERSION), nullptr, nullptr);

        CreateWindowW(L"BUTTON", L"Enable MicMix", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, contentLeft, S(108), actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_START), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Disable MicMix", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, contentLeft + actionButtonW + controlGap, S(108), actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_STOP), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Restart Source", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, contentLeft + (actionButtonW + controlGap) * 2, S(108), actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_SAVE), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Monitor Mix: Off", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, contentLeft + (actionButtonW + controlGap) * 3, S(108), monitorButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_MONITOR), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Only while connected", WS_CHILD | WS_VISIBLE, contentLeft + (actionButtonW + controlGap) * 3, S(108) + hintTopOffset, S(162), S(18), hwnd, reinterpret_cast<HMENU>(IDC_MONITOR_HINT), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Auto-enable when TeamSpeak starts", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, contentLeft, S(170), S(320), S(24), hwnd, reinterpret_cast<HMENU>(IDC_AUTOSTART), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"Audio Source", WS_CHILD | WS_VISIBLE, labelX, S(252), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        HWND source = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS, fieldX, S(248), sourceW, S(360), hwnd, reinterpret_cast<HMENU>(IDC_SOURCE), nullptr, nullptr);
        SendMessageW(source, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), S(24));
        SendMessageW(source, CB_SETITEMHEIGHT, 0, S(22));
        SendMessageW(source, CB_SETMINVISIBLE, 18, 0);
        CreateWindowW(L"BUTTON", L"Refresh", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, fieldX + sourceW + controlGap, S(248), refreshButtonW, S(30), hwnd, reinterpret_cast<HMENU>(IDC_RESCAN), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"Audio Output (Mic)", WS_CHILD | WS_VISIBLE, labelX, S(296), S(140), S(24), hwnd, nullptr, nullptr, nullptr);
        HWND mic = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | CBS_DROPDOWNLIST, fieldX, S(292), micComboW, S(240), hwnd, reinterpret_cast<HMENU>(IDC_MIC_DEVICE), nullptr, nullptr);
        SendMessageW(mic, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), S(24));
        SendMessageW(mic, CB_SETITEMHEIGHT, 0, S(22));
        SendMessageW(mic, CB_SETMINVISIBLE, 12, 0);

        CreateWindowW(L"STATIC", L"Music Volume", WS_CHILD | WS_VISIBLE, labelX, S(388), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        HWND gain = CreateWindowW(TRACKBAR_CLASSW, L"", WS_CHILD | WS_VISIBLE | TBS_AUTOTICKS, fieldX, S(384), gainSliderW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_GAIN), nullptr, nullptr);
        SendMessageW(gain, TBM_SETRANGE, TRUE, MAKELONG(0, kMusicGainSliderMax));
        SendMessageW(gain, TBM_SETTICFREQ, 20, 0);
        SendMessageW(gain, TBM_SETPAGESIZE, 0, 10);
        SendMessageW(gain, TBM_SETLINESIZE, 0, 1);
        CreateWindowW(L"STATIC", L"-15.0 dB", WS_CHILD | WS_VISIBLE, valueTextX, S(388), S(102), S(24), hwnd, reinterpret_cast<HMENU>(IDC_GAIN_VALUE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Max is -2 dB to reduce clipping risk", WS_CHILD | WS_VISIBLE, fieldX, S(412), S(330), S(18), hwnd, reinterpret_cast<HMENU>(IDC_GAIN_HINT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Music Meter", WS_CHILD | WS_VISIBLE, labelX, S(438), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        g_rcMeter = { fieldX, S(436), fieldX + meterW, S(436) + S(20) };
        CreateWindowW(L"STATIC", L"No signal", WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS, statusTextX, S(438), statusTextW, S(22), hwnd, reinterpret_cast<HMENU>(IDC_METER_TEXT), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Send music when mic is silent", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, fieldX, S(466), S(260), S(24), hwnd, reinterpret_cast<HMENU>(IDC_FORCE_TX), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Mute music", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, muteToggleX, S(466), S(120), S(24), hwnd, reinterpret_cast<HMENU>(IDC_MUTE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Music keeps sending even when you are not speaking", WS_CHILD | WS_VISIBLE, fieldX, S(486), S(360), S(18), hwnd, reinterpret_cast<HMENU>(IDC_FORCE_TX_HINT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Mic Meter", WS_CHILD | WS_VISIBLE, labelX, S(518), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        g_rcMicMeter = { fieldX, S(516), fieldX + meterW, S(516) + S(20) };
        CreateWindowW(L"STATIC", L"No signal", WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS, statusTextX, S(518), statusTextW, S(22), hwnd, reinterpret_cast<HMENU>(IDC_MIC_METER_TEXT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Only while connected", WS_CHILD | WS_VISIBLE, fieldX, S(538), S(180), S(18), hwnd, reinterpret_cast<HMENU>(IDC_MIC_METER_HINT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"MicMix-Mute-Hotkey", WS_CHILD | WS_VISIBLE, labelX, S(556), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Set...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, fieldX, S(552), S(160), S(30), hwnd, reinterpret_cast<HMENU>(IDC_MUTE_HOTKEY_SET), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Current: Not set", WS_CHILD | WS_VISIBLE, fieldX + S(172), S(556), S(290), S(24), hwnd, reinterpret_cast<HMENU>(IDC_MUTE_HOTKEY_TEXT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"State: Stopped", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX, statusX, statusY, statusW, statusH, hwnd, reinterpret_cast<HMENU>(IDC_STATUS), nullptr, nullptr);

        ApplyControlTheme(hwnd);
        ApplyFonts(hwnd);
        PopulateCombos(hwnd);
        LoadSettings(hwnd);
        RequestSourceRefresh(hwnd, true);
        UpdateMonitorButton(hwnd);
        SetTimer(hwnd, 1, 80, nullptr);
        UpdateStatus(hwnd);
        return 0;
    }
    case WM_HSCROLL:
        if (reinterpret_cast<HWND>(lParam) == GetDlgItem(hwnd, IDC_GAIN)) {
            UpdateGainLabel(hwnd);
            ApplyLiveSettings(hwnd, false);
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
        return FALSE;
    }
    case kMsgSourceRefreshDone: {
        const uint64_t seq = static_cast<uint64_t>(wParam);
        const uint64_t latest = g_sourceRefreshSeq.load(std::memory_order_acquire);
        if (seq < latest) {
            return 0;
        }
        {
            std::lock_guard<std::mutex> lock(g_enumMutex);
            g_loopbacks = g_pendingLoopbacks;
            g_captureDevices = g_pendingCaptureDevices;
            g_apps = g_pendingApps;
        }
        PopulateCombos(hwnd);
        if (g_sourceRefreshReloadSettings.exchange(false, std::memory_order_acq_rel)) {
            LoadSettings(hwnd);
        }
        HWND refreshBtn = GetDlgItem(hwnd, IDC_RESCAN);
        if (refreshBtn) {
            EnableWindow(refreshBtn, TRUE);
        }
        UpdateStatus(hwnd);
        return 0;
    }
    case WM_TIMER:
        if ((++g_hotkeyRefreshTick % 8) == 0) {
            UpdateMuteHotkeyLabel(hwnd);
        }
        UpdateStatus(hwnd);
        return 0;
    case WM_KEYDOWN:
    case WM_SYSKEYDOWN:
        if (g_waitingForHotkey) {
            const UINT vk = static_cast<UINT>(wParam);
            if (vk == VK_ESCAPE) {
                g_waitingForHotkey = false;
                UpdateMuteHotkeyLabel(hwnd);
                return 0;
            }
            if (vk == VK_DELETE || vk == VK_BACK) {
                g_waitingForHotkey = false;
                g_muteHotkeyModifiers = 0;
                g_muteHotkeyVk = 0;
                ApplyLiveSettings(hwnd, false);
                UpdateMuteHotkeyLabel(hwnd);
                return 0;
            }
            if (IsModifierOnlyKey(vk)) {
                return 0;
            }
            const UINT mods = ReadCurrentModifiers();
            g_waitingForHotkey = false;
            g_muteHotkeyModifiers = mods;
            g_muteHotkeyVk = vk;
            ApplyLiveSettings(hwnd, false);
            UpdateMuteHotkeyLabel(hwnd);
            return 0;
        }
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
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
            ApplyLiveSettings(hwnd, false);
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
                BeginHotkeyCapture(hwnd);
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
        case IDC_MUTE:
        case IDC_AUTOSTART:
            if (HIWORD(wParam) == BN_CLICKED) {
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
        } else if (id == IDC_MONITOR_HINT || id == IDC_MIC_METER_HINT || id == IDC_GAIN_HINT || id == IDC_FORCE_TX_HINT) {
            SetTextColor(hdc, g_theme.muted);
        } else if (id == IDC_STATUS) {
            SetTextColor(hdc, RGB(40, 48, 61));
        } else {
            SetTextColor(hdc, g_theme.text);
        }
        return reinterpret_cast<INT_PTR>(inCard ? g_brushCard : g_brushBg);
    }
    case WM_PAINT: {
        PAINTSTRUCT ps{};
        HDC hdc = BeginPaint(hwnd, &ps);
        FillRect(hdc, &ps.rcPaint, g_brushBg);

        RECT accent{ S(0), S(0), S(kClientWidthPx), S(4) };
        HBRUSH accentBrush = CreateSolidBrush(g_theme.accent);
        FillRect(hdc, &accent, accentBrush);
        DeleteObject(accentBrush);

        DrawCard(hdc, g_rcSource);
        DrawCard(hdc, g_rcMix);
        DrawCard(hdc, g_rcSession);
        DrawLevelMeter(hdc, g_rcMeter, IsMusicMeterActive(g_lastTelemetry), g_meterVisualDb, g_meterHoldDb);
        DrawLevelMeter(hdc, g_rcMicMeter, g_lastTelemetry.micRmsDbfs > -119.0f, g_micMeterVisualDb, g_micMeterHoldDb);

        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, g_theme.muted);
        SelectObject(hdc, g_fontSmall);
        const wchar_t* h1 = L"MICMIX SETTINGS";
        const wchar_t* h2 = L"AUDIO ROUTING";
        const wchar_t* h3 = L"MIX BEHAVIOR";
        TextOutW(hdc, S(36), S(86), h1, lstrlenW(h1));
        TextOutW(hdc, S(36), S(226), h2, lstrlenW(h2));
        TextOutW(hdc, S(36), S(358), h3, lstrlenW(h3));
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;
    case WM_CLOSE:
        if (MicMixApp::Instance().IsMonitorEnabled()) {
            MicMixApp::Instance().SetMonitorEnabled(false);
        }
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
        if (MicMixApp::Instance().IsMonitorEnabled()) {
            MicMixApp::Instance().SetMonitorEnabled(false);
        }
        g_hwnd.store(nullptr, std::memory_order_release);
        g_lastStatusText.clear();
        KillTimer(hwnd, 1);
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

    HINSTANCE hInst = GetPluginModuleHandle();
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
