#include "settings_window.h"

#include <Windows.h>
#include <CommCtrl.h>
#include <uxtheme.h>

#include <atomic>
#include <algorithm>
#include <cstdio>
#include <cwctype>
#include <mutex>
#include <thread>
#include <vector>
#include <string>

#include "micmix_core.h"
#include "resource.h"

#pragma comment(lib, "Comctl32.lib")
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
    IDC_DUCKING = 1009,
    IDC_MUTE = 1010,
    IDC_AUTOSTART = 1011,
    IDC_START = 1012,
    IDC_STOP = 1013,
    IDC_SAVE = 1014,
    IDC_STATUS = 1015,
    IDC_MIC_INFO = 1017,
    IDC_METER_BAR = 1018,
    IDC_METER_TEXT = 1019,
    IDC_MUTE_HOTKEY_SET = 1020,
    IDC_MUTE_HOTKEY_TEXT = 1021,
    IDC_DUCK_AMOUNT = 1022,
    IDC_DUCK_AMOUNT_VALUE = 1023,
    IDC_MIC_DEVICE = 1024,
};

enum class SourceChoiceType {
    Loopback,
    App,
};

struct SourceChoice {
    SourceChoiceType type = SourceChoiceType::Loopback;
    std::string id;
    std::string processName;
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
std::thread g_thread;
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

constexpr INT_PTR kSourceItemHeader = -10;
constexpr INT_PTR kSourceItemDivider = -11;
constexpr INT_PTR kSourceItemPlaceholder = -12;
constexpr float kMusicGainMinDb = -30.0f;
constexpr float kMusicGainMaxDb = -6.0f;
constexpr int kMusicGainStepPerDb = 10;
constexpr int kMusicGainSliderMax = static_cast<int>((kMusicGainMaxDb - kMusicGainMinDb) * static_cast<float>(kMusicGainStepPerDb));

HFONT g_fontBody = nullptr;
HFONT g_fontSmall = nullptr;
HFONT g_fontTitle = nullptr;
HFONT g_fontMono = nullptr;
HBRUSH g_brushBg = nullptr;
HBRUSH g_brushCard = nullptr;

RECT g_rcSource{};
RECT g_rcMix{};
RECT g_rcSession{};
RECT g_rcMeter{};
float g_meterVisualDb = -60.0f;
float g_meterHoldDb = -60.0f;
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
    if (!g_fontBody) g_fontBody = CreateFontW(-S(14), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!g_fontSmall) g_fontSmall = CreateFontW(-S(13), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!g_fontTitle) g_fontTitle = CreateFontW(-S(24), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI Semibold");
    if (!g_fontMono) g_fontMono = CreateFontW(-S(13), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, FIXED_PITCH, L"Consolas");
}

void ReleaseUiResources() {
    if (g_fontBody) { DeleteObject(g_fontBody); g_fontBody = nullptr; }
    if (g_fontSmall) { DeleteObject(g_fontSmall); g_fontSmall = nullptr; }
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

int DuckAmountDbToSlider(float duckAmountDb) {
    const float reductionDb = std::clamp(-duckAmountDb, 0.0f, 30.0f);
    return static_cast<int>(reductionDb * 10.0f);
}

float SliderToDuckAmountDb(int slider) {
    const float reductionDb = std::clamp(static_cast<float>(slider) / 10.0f, 0.0f, 30.0f);
    return -reductionDb;
}

int MeterDbToPixels(float dbfs, int widthPx) {
    const float clamped = std::clamp(dbfs, -60.0f, 0.0f);
    return static_cast<int>(((clamped + 60.0f) / 60.0f) * static_cast<float>(widthPx));
}

void UpdateMusicMeter(HWND hwnd) {
    HWND txt = GetDlgItem(hwnd, IDC_METER_TEXT);
    if (!txt) {
        return;
    }
    const TelemetrySnapshot t = g_lastTelemetry;
    const float shownPeakDb = (t.musicSendPeakDbfs > -119.0f) ? t.musicSendPeakDbfs : t.musicPeakDbfs;

    float targetDb = -60.0f;
    if (t.musicActive) {
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

    if (!t.musicActive) {
        SetWindowTextW(txt, L"No signal");
    } else {
        const wchar_t* grade = L"OK";
        if (shownPeakDb > -8.0f) {
            grade = L"too loud";
        } else if (shownPeakDb > -16.0f) {
            grade = L"reduce volume";
        } else if (shownPeakDb < -35.0f) {
            grade = L"too quiet";
        }
        wchar_t meter[96];
        swprintf_s(meter, L"%.1f dBFS (%s)", shownPeakDb, grade);
        SetWindowTextW(txt, meter);
    }

    InvalidateRect(hwnd, &g_rcMeter, FALSE);
}

void UpdateGainLabel(HWND hwnd) {
    const int slider = static_cast<int>(SendMessageW(GetDlgItem(hwnd, IDC_GAIN), TBM_GETPOS, 0, 0));
    const float db = SliderToGainDb(slider);
    wchar_t buf[64];
    swprintf_s(buf, L"%+.1f dB", db);
    SetWindowTextW(GetDlgItem(hwnd, IDC_GAIN_VALUE), buf);
}

void UpdateDuckAmountLabel(HWND hwnd) {
    const int slider = static_cast<int>(SendMessageW(GetDlgItem(hwnd, IDC_DUCK_AMOUNT), TBM_GETPOS, 0, 0));
    const float reductionDb = std::clamp(static_cast<float>(slider) / 10.0f, 0.0f, 30.0f);
    wchar_t buf[64];
    swprintf_s(buf, L"%.1f dB", reductionDb);
    SetWindowTextW(GetDlgItem(hwnd, IDC_DUCK_AMOUNT_VALUE), buf);
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

void PopulateCombos(HWND hwnd) {
    auto& app = MicMixApp::Instance();
    g_loopbacks = app.GetLoopbackDevices();
    g_captureDevices = app.GetCaptureDevices();
    g_apps = app.GetAppProcesses();
    g_sourceChoices.clear();
    g_lastValidSourceSel = -1;

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
            addSelectableItem(source, label, choice);
        }
    }

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
    s.duckingEnabled = false;
    s.duckingMode = DuckingMode::MicRms;
    s.duckingAmountDb = 0.0f;
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
    std::string line2 = "underrun=" + std::to_string(t.underruns) +
        "  overrun=" + std::to_string(t.overruns) +
        "  clip=" + std::to_string(t.clippedSamples) +
        "  " + micBuf +
        "  " + sendBuf;
    SetStatusText(hwnd, Utf8ToWide(line1 + "\r\n" + line2));
    UpdateMusicMeter(hwnd);
}

void ApplyLiveSettings(HWND hwnd, bool restartSource) {
    if (g_loadingUi) return;
    auto s = CollectSettings(hwnd);
    MicMixApp::Instance().ApplySettings(s, restartSource, true);
    UpdateStatus(hwnd);
}

void ApplyExplorerTheme(HWND hwnd) {
    EnumChildWindows(hwnd, [](HWND child, LPARAM) -> BOOL {
        wchar_t className[32]{};
        GetClassNameW(child, className, static_cast<int>(std::size(className)));
        if (_wcsicmp(className, WC_COMBOBOXW) == 0 ||
            _wcsicmp(className, WC_BUTTONW) == 0) {
            SetWindowTheme(child, L"Explorer", nullptr);
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
    SetControlFont(hwnd, IDC_MIC_INFO, g_fontSmall);
    SetControlFont(hwnd, IDC_METER_TEXT, g_fontSmall);
    SetControlFont(hwnd, IDC_STATUS, g_fontMono);
}

void ComputeLayout() {
    g_rcSource = { S(16), S(82), S(684), S(220) };
    g_rcSession = { S(16), S(232), S(684), S(340) };
    g_rcMix = { S(16), S(352), S(684), S(590) };
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

void DrawMusicMeter(HDC hdc) {
    if (g_rcMeter.right <= g_rcMeter.left || g_rcMeter.bottom <= g_rcMeter.top) {
        return;
    }
    RECT rc = g_rcMeter;
    FillSolidRect(hdc, rc, RGB(235, 240, 246));

    const int width = rc.right - rc.left;
    const int levelPx = std::clamp(MeterDbToPixels(g_meterVisualDb, width), 0, width);
    const int holdPx = std::clamp(MeterDbToPixels(g_meterHoldDb, width), 0, width - 1);
    // Intentionally stricter: show warning zones earlier to encourage lower music levels.
    const int greenEnd = static_cast<int>(width * 0.55f);
    const int yellowEnd = static_cast<int>(width * 0.78f);

    if (g_lastTelemetry.musicActive && levelPx > 0) {
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

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        g_dpi = GetDpiForWindow(hwnd);
        EnsureUiResources();
        ComputeLayout();

        CreateWindowW(L"STATIC", L"MicMix", WS_CHILD | WS_VISIBLE, S(20), S(14), S(220), S(36), hwnd, reinterpret_cast<HMENU>(IDC_TITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Select exactly one source: audio channel or running app", WS_CHILD | WS_VISIBLE, S(22), S(44), S(560), S(20), hwnd, reinterpret_cast<HMENU>(IDC_SUBTITLE), nullptr, nullptr);

        const int labelX = S(36);
        const int fieldX = S(180);
        const int sourceW = S(340);

        CreateWindowW(L"STATIC", L"Audio Source", WS_CHILD | WS_VISIBLE, labelX, S(104), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        HWND source = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | CBS_DROPDOWNLIST, fieldX, S(100), sourceW, S(360), hwnd, reinterpret_cast<HMENU>(IDC_SOURCE), nullptr, nullptr);
        SendMessageW(source, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), S(24));
        SendMessageW(source, CB_SETITEMHEIGHT, 0, S(22));
        SendMessageW(source, CB_SETMINVISIBLE, 18, 0);
        CreateWindowW(L"BUTTON", L"Refresh", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, S(532), S(100), S(130), S(30), hwnd, reinterpret_cast<HMENU>(IDC_RESCAN), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"Microphone", WS_CHILD | WS_VISIBLE, labelX, S(148), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        HWND mic = CreateWindowW(L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | CBS_DROPDOWNLIST, fieldX, S(144), S(482), S(240), hwnd, reinterpret_cast<HMENU>(IDC_MIC_DEVICE), nullptr, nullptr);
        SendMessageW(mic, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), S(24));
        SendMessageW(mic, CB_SETITEMHEIGHT, 0, S(22));
        SendMessageW(mic, CB_SETMINVISIBLE, 12, 0);

        CreateWindowW(L"BUTTON", L"Enable MicMix", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, S(36), S(256), S(180), S(36), hwnd, reinterpret_cast<HMENU>(IDC_START), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Disable MicMix", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, S(228), S(256), S(180), S(36), hwnd, reinterpret_cast<HMENU>(IDC_STOP), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Restart Audio Source", WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON, S(420), S(256), S(242), S(36), hwnd, reinterpret_cast<HMENU>(IDC_SAVE), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"Music Volume", WS_CHILD | WS_VISIBLE, labelX, S(382), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        HWND gain = CreateWindowW(TRACKBAR_CLASSW, L"", WS_CHILD | WS_VISIBLE | TBS_NOTICKS, fieldX, S(378), S(380), S(34), hwnd, reinterpret_cast<HMENU>(IDC_GAIN), nullptr, nullptr);
        SendMessageW(gain, TBM_SETRANGE, TRUE, MAKELONG(0, kMusicGainSliderMax));
        CreateWindowW(L"STATIC", L"-15.0 dB", WS_CHILD | WS_VISIBLE, S(572), S(382), S(90), S(24), hwnd, reinterpret_cast<HMENU>(IDC_GAIN_VALUE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Music Meter", WS_CHILD | WS_VISIBLE, labelX, S(420), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        g_rcMeter = { fieldX, S(418), fieldX + S(380), S(418) + S(20) };
        CreateWindowW(L"STATIC", L"No signal", WS_CHILD | WS_VISIBLE, S(572), S(420), S(120), S(22), hwnd, reinterpret_cast<HMENU>(IDC_METER_TEXT), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"MicMix-Mute-Hotkey", WS_CHILD | WS_VISIBLE, labelX, S(456), S(130), S(24), hwnd, nullptr, nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Set...", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, fieldX, S(452), S(160), S(30), hwnd, reinterpret_cast<HMENU>(IDC_MUTE_HOTKEY_SET), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Current: Not set", WS_CHILD | WS_VISIBLE, fieldX + S(172), S(456), S(290), S(24), hwnd, reinterpret_cast<HMENU>(IDC_MUTE_HOTKEY_TEXT), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Enable on startup", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, fieldX, S(492), S(220), S(24), hwnd, reinterpret_cast<HMENU>(IDC_AUTOSTART), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Send music without speaking", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, S(430), S(492), S(240), S(24), hwnd, reinterpret_cast<HMENU>(IDC_FORCE_TX), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Mute music", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, fieldX, S(520), S(120), S(24), hwnd, reinterpret_cast<HMENU>(IDC_MUTE), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"State: Stopped", WS_CHILD | WS_VISIBLE, S(36), S(598), S(626), S(54), hwnd, reinterpret_cast<HMENU>(IDC_STATUS), nullptr, nullptr);

        ApplyExplorerTheme(hwnd);
        ApplyFonts(hwnd);
        PopulateCombos(hwnd);
        LoadSettings(hwnd);
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
            PopulateCombos(hwnd);
            LoadSettings(hwnd);
            UpdateStatus(hwnd);
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
        } else if (id == IDC_SUBTITLE) {
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

        RECT accent{ S(0), S(0), S(700), S(4) };
        HBRUSH accentBrush = CreateSolidBrush(g_theme.accent);
        FillRect(hdc, &accent, accentBrush);
        DeleteObject(accentBrush);

        DrawCard(hdc, g_rcSource);
        DrawCard(hdc, g_rcMix);
        DrawCard(hdc, g_rcSession);
        DrawMusicMeter(hdc);

        SetBkMode(hdc, TRANSPARENT);
        SetTextColor(hdc, g_theme.muted);
        SelectObject(hdc, g_fontSmall);
        const wchar_t* h1 = L"AUDIO SOURCE";
        const wchar_t* h2 = L"SEND MUSIC";
        const wchar_t* h3 = L"LEVEL & BEHAVIOR";
        TextOutW(hdc, S(30), S(86), h1, lstrlenW(h1));
        TextOutW(hdc, S(30), S(236), h2, lstrlenW(h2));
        TextOutW(hdc, S(30), S(356), h3, lstrlenW(h3));
        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;
    case WM_CLOSE:
        DestroyWindow(hwnd);
        return 0;
    case WM_DESTROY:
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
        S(716),
        S(700),
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
    g_thread = std::thread(WindowThreadMain);
}

void SettingsWindowController::Close() {
    std::lock_guard<std::mutex> lock(g_mutex);
    if (!g_running.load(std::memory_order_acquire)) {
        if (g_thread.joinable()) g_thread.join();
        return;
    }
    HWND hwnd = g_hwnd.load(std::memory_order_acquire);
    if (hwnd) {
        PostMessageW(hwnd, WM_CLOSE, 0, 0);
    }
    if (g_thread.joinable()) {
        g_thread.join();
    }
}
