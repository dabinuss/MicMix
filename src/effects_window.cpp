#include "effects_window.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <CommCtrl.h>
#include <commdlg.h>
#include <shellapi.h>
#include <uxtheme.h>

#include <algorithm>
#include <atomic>
#include <array>
#include <cmath>
#include <cstdio>
#include <exception>
#include <filesystem>
#include <memory>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

#include "micmix_core.h"
#include "micmix_version.h"
#include "resource.h"
#include "ui_shared.h"

#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Shell32.lib")
#pragma comment(lib, "UxTheme.lib")

namespace {

enum ControlId {
    IDC_TITLE = 4200,
    IDC_SUBTITLE = 4201,
    IDC_VERSION = 4202,
    IDC_REPO_LINK = 4203,

    IDC_ENABLE_EFFECTS = 4210,
    IDC_DISABLE_EFFECTS = 4211,
    IDC_MONITOR = 4212,
    IDC_MONITOR_HINT = 4213,

    IDC_MUSIC_LIST = 4220,
    IDC_MUSIC_ADD = 4221,
    IDC_MUSIC_REMOVE = 4222,
    IDC_MUSIC_UP = 4223,
    IDC_MUSIC_DOWN = 4224,
    IDC_MUSIC_BYPASS = 4225,
    IDC_MUSIC_EDITOR = 4226,

    IDC_MIC_LIST = 4230,
    IDC_MIC_ADD = 4231,
    IDC_MIC_REMOVE = 4232,
    IDC_MIC_UP = 4233,
    IDC_MIC_DOWN = 4234,
    IDC_MIC_BYPASS = 4235,
    IDC_MIC_EDITOR = 4236,

    IDC_STATUS = 4240,

    IDC_AUDIO_SECTION_TITLE = 4250,
    IDC_AUDIO_METER_LABEL = 4251,
    IDC_MIC_SECTION_TITLE = 4252,
    IDC_MIC_METER_LABEL = 4253,
};

enum class HeaderBadgeState {
    Active,
    Off,
};

std::mutex g_mutex;
std::thread g_thread;
std::atomic<bool> g_running{false};
std::atomic<HWND> g_hwnd{nullptr};
UiTheme g_theme = DefaultUiTheme();
int g_dpi = 96;
std::wstring g_statusText;
uint64_t g_lastListRefreshTickMs = 0;
std::atomic<bool> g_loading{false};
std::atomic<bool> g_editorOpenPending{false};

constexpr UINT kMsgEditorOpenDone = WM_APP + 42;

struct EditorOpenResult {
    bool ok = false;
    std::wstring message;
};

HFONT g_fontBody = nullptr;
HFONT g_fontSmall = nullptr;
HFONT g_fontHint = nullptr;
HFONT g_fontTiny = nullptr;
HFONT g_fontTitle = nullptr;
HFONT g_fontMono = nullptr;
HBRUSH g_brushBg = nullptr;
HBRUSH g_brushCard = nullptr;

RECT g_rcHeaderBadge{};
RECT g_rcTop{};
RECT g_rcMusic{};
RECT g_rcMic{};
RECT g_rcStatus{};
RECT g_rcMusicMeter{};
RECT g_rcMusicClip{};
RECT g_rcMicMeter{};
RECT g_rcMicClip{};

TelemetrySnapshot g_lastTelemetry{};
float g_musicMeterVisualDb = -60.0f;
float g_musicMeterHoldDb = -60.0f;
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

constexpr int kClientWidthPx = 680;
constexpr int kClientHeightPx = 780;
constexpr int kCardMarginPx = 16;
constexpr int kCardGapPx = 12;
constexpr int kCardInnerPaddingPx = 20;
constexpr UINT kTimerStatus = 1;
constexpr UINT kTimerMs = 250;
constexpr int kListItemHeightPx = 22;
constexpr wchar_t kRepoUrl[] = L"https://github.com/dabinuss/MicMix";

int S(int px) {
    return MulDiv(px, g_dpi, 96);
}

void FillSolidRect(HDC hdc, const RECT& rc, COLORREF color) {
    HBRUSH brush = CreateSolidBrush(color);
    FillRect(hdc, &rc, brush);
    DeleteObject(brush);
}

int MeterDbToPixels(float dbfs, int widthPx) {
    const float clamped = std::clamp(dbfs, -60.0f, 0.0f);
    return static_cast<int>(((clamped + 60.0f) / 60.0f) * static_cast<float>(widthPx));
}

float ClipDangerUnitFromDb(float dbfs) {
    const float clamped = std::clamp(dbfs, -18.0f, 0.0f);
    return (clamped + 18.0f) / 18.0f;
}

int ClipLitSegmentsFromDb(float dbfs) {
    constexpr int kClipSegments = 18;
    const float dangerUnit = ClipDangerUnitFromDb(dbfs);
    return std::clamp(static_cast<int>((dangerUnit * static_cast<float>(kClipSegments)) + 0.5f), 0, kClipSegments);
}

bool IsMusicMeterActive(const TelemetrySnapshot& t) {
    const float shownPeakDb = (t.musicSendPeakDbfs > -119.0f) ? t.musicSendPeakDbfs : t.musicPeakDbfs;
    const bool levelSuggestsSignal = (shownPeakDb > -96.0f) || (t.musicRmsDbfs > -100.0f);
    return t.musicActive || levelSuggestsSignal;
}

bool IsHandCursorControlId(int id) {
    switch (id) {
    case IDC_REPO_LINK:
    case IDC_ENABLE_EFFECTS:
    case IDC_DISABLE_EFFECTS:
    case IDC_MONITOR:
    case IDC_MUSIC_LIST:
    case IDC_MUSIC_ADD:
    case IDC_MUSIC_REMOVE:
    case IDC_MUSIC_UP:
    case IDC_MUSIC_DOWN:
    case IDC_MUSIC_BYPASS:
    case IDC_MUSIC_EDITOR:
    case IDC_MIC_LIST:
    case IDC_MIC_ADD:
    case IDC_MIC_REMOVE:
    case IDC_MIC_UP:
    case IDC_MIC_DOWN:
    case IDC_MIC_BYPASS:
    case IDC_MIC_EDITOR:
        return true;
    default:
        return false;
    }
}

void SetStatusText(HWND hwnd, const std::wstring& text) {
    g_statusText = text;
    HWND ctl = GetDlgItem(hwnd, IDC_STATUS);
    if (ctl) {
        SetWindowTextW(ctl, g_statusText.c_str());
    }
}

HeaderBadgeState ResolveBadgeState() {
    return MicMixApp::Instance().IsEffectsEnabled() ? HeaderBadgeState::Active : HeaderBadgeState::Off;
}

HeaderBadgeVisual GetHeaderBadgeVisual(HeaderBadgeState state) {
    switch (state) {
    case HeaderBadgeState::Active:
        return { L"ACTIVE", RGB(21, 171, 88), RGB(9, 133, 63), RGB(255, 255, 255), RGB(224, 255, 234) };
    default:
        return {};
    }
}

void EnsureUiResources() {
    if (!g_brushBg) {
        g_brushBg = CreateSolidBrush(g_theme.bg);
    }
    if (!g_brushCard) {
        g_brushCard = CreateSolidBrush(g_theme.card);
    }
    if (!g_fontBody) {
        g_fontBody = CreateFontW(-S(13), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    }
    if (!g_fontSmall) {
        g_fontSmall = CreateFontW(-S(12), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    }
    if (!g_fontHint) {
        g_fontHint = CreateFontW(-S(11), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    }
    if (!g_fontTiny) {
        g_fontTiny = CreateFontW(-S(10), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    }
    if (!g_fontTitle) {
        g_fontTitle = CreateFontW(-S(24), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI Semibold");
    }
    if (!g_fontMono) {
        g_fontMono = CreateFontW(-S(13), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, FIXED_PITCH, L"Consolas");
    }
}

void ReleaseUiResources() {
    if (g_fontBody) {
        DeleteObject(g_fontBody);
        g_fontBody = nullptr;
    }
    if (g_fontSmall) {
        DeleteObject(g_fontSmall);
        g_fontSmall = nullptr;
    }
    if (g_fontHint) {
        DeleteObject(g_fontHint);
        g_fontHint = nullptr;
    }
    if (g_fontTiny) {
        DeleteObject(g_fontTiny);
        g_fontTiny = nullptr;
    }
    if (g_fontTitle) {
        DeleteObject(g_fontTitle);
        g_fontTitle = nullptr;
    }
    if (g_fontMono) {
        DeleteObject(g_fontMono);
        g_fontMono = nullptr;
    }
    if (g_brushBg) {
        DeleteObject(g_brushBg);
        g_brushBg = nullptr;
    }
    if (g_brushCard) {
        DeleteObject(g_brushCard);
        g_brushCard = nullptr;
    }
}

void SetControlFont(HWND hwnd, int id, HFONT font) {
    HWND ctl = GetDlgItem(hwnd, id);
    if (ctl && font) {
        SendMessageW(ctl, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }
}

void ComputeLayout(HWND hwnd) {
    const int left = S(kCardMarginPx);
    const int right = S(kClientWidthPx - kCardMarginPx);
    const int cardW = right - left;
    const int topY = S(88);
    const int topH = S(88);
    const int musicY = topY + topH + S(kCardGapPx);
    const int musicH = S(228);
    const int micY = musicY + musicH + S(kCardGapPx);
    const int micH = S(228);
    const int statusY = micY + micH + S(kCardGapPx);
    const int statusH = S(104);
    g_rcTop = { left, topY, left + cardW, topY + topH };
    g_rcMusic = { left, musicY, left + cardW, musicY + musicH };
    g_rcMic = { left, micY, left + cardW, micY + micH };
    g_rcStatus = { left, statusY, left + cardW, statusY + statusH };

    const int meterWidth = S(340);
    const int sectionPad = S(kCardInnerPaddingPx);
    const int musicSectionTop = g_rcMusic.top + sectionPad;
    const int musicSectionLeft = g_rcMusic.left + sectionPad;
    const int musicMeterLeft = musicSectionLeft + S(138);
    g_rcMusicMeter = { musicMeterLeft, musicSectionTop + S(32), musicMeterLeft + meterWidth, musicSectionTop + S(32) + S(28) };
    g_rcMusicClip = { musicMeterLeft, musicSectionTop + S(62), musicMeterLeft + meterWidth, musicSectionTop + S(62) + S(12) };

    const int micSectionTop = g_rcMic.top + sectionPad;
    const int micSectionLeft = g_rcMic.left + sectionPad;
    const int micMeterLeft = micSectionLeft + S(138);
    g_rcMicMeter = { micMeterLeft, micSectionTop + S(32), micMeterLeft + meterWidth, micSectionTop + S(32) + S(28) };
    g_rcMicClip = { micMeterLeft, micSectionTop + S(62), micMeterLeft + meterWidth, micSectionTop + S(62) + S(12) };

    const int contentLeft = S(kCardMarginPx + kCardInnerPaddingPx);
    const int titleX = contentLeft;
    const int titleY = S(14);
    const int titleH = S(36);
    int titleW = S(112);
    int badgeW = S(74);
    HDC dc = GetDC(hwnd);
    if (dc) {
        HGDIOBJ old = SelectObject(dc, g_fontTitle ? g_fontTitle : g_fontBody);
        SIZE titleSize{};
        if (GetTextExtentPoint32W(dc, L"MicMix", 6, &titleSize) != 0) {
            titleW = titleSize.cx + S(2);
        }
        SelectObject(dc, g_fontSmall ? g_fontSmall : g_fontBody);
        SIZE badgeText{};
        if (GetTextExtentPoint32W(dc, L"ACTIVE", 6, &badgeText) != 0) {
            badgeW = std::clamp<int>(badgeText.cx + S(28), S(78), S(108));
        }
        SelectObject(dc, old);
        ReleaseDC(hwnd, dc);
    }
    const int badgeH = S(22);
    const int badgeX = titleX + titleW + S(2);
    const int badgeY = titleY + ((titleH - badgeH) / 2) - S(1);
    g_rcHeaderBadge = { badgeX, badgeY, badgeX + badgeW, badgeY + badgeH };
}

void ApplyFonts(HWND hwnd) {
    EnumChildWindows(hwnd, [](HWND child, LPARAM) -> BOOL {
        SendMessageW(child, WM_SETFONT, reinterpret_cast<WPARAM>(g_fontBody), TRUE);
        return TRUE;
    }, 0);
    SetControlFont(hwnd, IDC_TITLE, g_fontTitle);
    SetControlFont(hwnd, IDC_SUBTITLE, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFont(hwnd, IDC_VERSION, g_fontSmall);
    SetControlFont(hwnd, IDC_REPO_LINK, g_fontSmall);
    SetControlFont(hwnd, IDC_MONITOR_HINT, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFont(hwnd, IDC_AUDIO_SECTION_TITLE, g_fontSmall);
    SetControlFont(hwnd, IDC_MIC_SECTION_TITLE, g_fontSmall);
    SetControlFont(hwnd, IDC_STATUS, g_fontMono ? g_fontMono : g_fontBody);
}

void UpdateMonitorButton(HWND hwnd) {
    HWND btn = GetDlgItem(hwnd, IDC_MONITOR);
    if (!btn) {
        return;
    }
    const bool enabled = MicMixApp::Instance().IsMonitorEnabled();
    SetWindowTextW(btn, enabled ? L"Monitor Mix: On" : L"Monitor Mix: Off");
}

std::wstring BuildEffectRow(const VstEffectSlot& slot) {
    std::wstring row = L"[";
    row += slot.enabled ? L"ON" : L"OFF";
    row += L"] [";
    row += slot.bypass ? L"BYP" : L"ACT";
    row += L"] ";
    if (!slot.lastStatus.empty()) {
        row += L"{";
        row += Utf8ToWide(slot.lastStatus);
        row += L"} ";
    }
    row += Utf8ToWide(slot.name.empty() ? slot.path : slot.name);
    return row;
}

void ReloadEffectList(HWND hwnd, EffectChain chain, int listId) {
    HWND list = GetDlgItem(hwnd, listId);
    if (!list) {
        return;
    }
    const int previousSel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
    const std::vector<VstEffectSlot> slots = MicMixApp::Instance().GetEffects(chain);
    SendMessageW(list, LB_RESETCONTENT, 0, 0);
    for (const auto& slot : slots) {
        const std::wstring row = BuildEffectRow(slot);
        SendMessageW(list, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(row.c_str()));
    }
    if (!slots.empty()) {
        int sel = previousSel;
        if (sel < 0 || sel >= static_cast<int>(slots.size())) {
            sel = static_cast<int>(slots.size() - 1);
        }
        SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(sel), 0);
    }
}

void ReloadAllLists(HWND hwnd) {
    ReloadEffectList(hwnd, EffectChain::Music, IDC_MUSIC_LIST);
    ReloadEffectList(hwnd, EffectChain::Mic, IDC_MIC_LIST);
}

int GetSelectedIndex(HWND hwnd, int listId) {
    return static_cast<int>(SendMessageW(GetDlgItem(hwnd, listId), LB_GETCURSEL, 0, 0));
}

void ApplyActionResult(HWND hwnd, bool ok, const std::string& error, const wchar_t* successText) {
    if (ok) {
        SetStatusText(hwnd, successText ? successText : L"Done.");
        ReloadAllLists(hwnd);
        return;
    }
    SetStatusText(hwnd, Utf8ToWide(error.empty() ? "Operation failed" : error));
}

bool PromptEffectPath(HWND hwnd, std::string& outPath) {
    outPath.clear();
    wchar_t fileBuf[4096]{};
    OPENFILENAMEW ofn{};
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hwnd;
    ofn.lpstrFile = fileBuf;
    ofn.nMaxFile = static_cast<DWORD>(std::size(fileBuf));
    ofn.lpstrTitle = L"Select VST Plugin";
    ofn.lpstrFilter = L"VST3 Plugins (*.vst3)\0*.vst3\0All Files (*.*)\0*.*\0";
    ofn.Flags = OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY;
    if (!GetOpenFileNameW(&ofn)) {
        return false;
    }
    outPath = WideToUtf8(fileBuf);
    return !outPath.empty();
}

void HandleAddEffect(HWND hwnd, EffectChain chain) {
    std::string path;
    if (!PromptEffectPath(hwnd, path)) {
        return;
    }
    VstEffectSlot slot{};
    slot.path = path;
    std::filesystem::path p(Utf8ToWide(path));
    slot.name = WideToUtf8(p.stem().wstring());
    slot.enabled = true;
    slot.bypass = false;
    std::string error;
    const bool ok = MicMixApp::Instance().AddEffect(chain, slot, error);
    ApplyActionResult(hwnd, ok, error, L"Effect added.");
}

void HandleRemoveEffect(HWND hwnd, EffectChain chain, int listId) {
    const int idx = GetSelectedIndex(hwnd, listId);
    if (idx < 0) {
        SetStatusText(hwnd, L"Select an effect first.");
        return;
    }
    std::string error;
    const bool ok = MicMixApp::Instance().RemoveEffect(chain, static_cast<size_t>(idx), error);
    ApplyActionResult(hwnd, ok, error, L"Effect removed.");
}

void HandleMoveEffect(HWND hwnd, EffectChain chain, int listId, int direction) {
    const int idx = GetSelectedIndex(hwnd, listId);
    if (idx < 0) {
        SetStatusText(hwnd, L"Select an effect first.");
        return;
    }
    const int target = idx + direction;
    if (target < 0) {
        return;
    }
    std::string error;
    const bool ok = MicMixApp::Instance().MoveEffect(
        chain, static_cast<size_t>(idx), static_cast<size_t>(target), error);
    ApplyActionResult(hwnd, ok, error, L"Effect order updated.");
    if (ok) {
        SendMessageW(GetDlgItem(hwnd, listId), LB_SETCURSEL, static_cast<WPARAM>(target), 0);
    }
}

void HandleBypassEffect(HWND hwnd, EffectChain chain, int listId) {
    const int idx = GetSelectedIndex(hwnd, listId);
    if (idx < 0) {
        SetStatusText(hwnd, L"Select an effect first.");
        return;
    }
    const std::vector<VstEffectSlot> slots = MicMixApp::Instance().GetEffects(chain);
    if (static_cast<size_t>(idx) >= slots.size()) {
        SetStatusText(hwnd, L"Invalid selection.");
        return;
    }
    const bool nextBypass = !slots[static_cast<size_t>(idx)].bypass;
    std::string error;
    const bool ok = MicMixApp::Instance().SetEffectBypass(chain, static_cast<size_t>(idx), nextBypass, error);
    ApplyActionResult(hwnd, ok, error, nextBypass ? L"Effect bypassed." : L"Effect active.");
}

void HandleOpenEditor(HWND hwnd, EffectChain chain, int listId) {
    const int idx = GetSelectedIndex(hwnd, listId);
    if (idx < 0) {
        SetStatusText(hwnd, L"Select an effect first.");
        return;
    }
    if (g_editorOpenPending.exchange(true, std::memory_order_acq_rel)) {
        SetStatusText(hwnd, L"Editor request already running...");
        return;
    }
    SetStatusText(hwnd, L"Opening editor...");
    std::thread([hwnd, chain, idx]() {
        std::string error;
        const bool ok = MicMixApp::Instance().OpenEffectEditor(chain, static_cast<size_t>(idx), error);
        auto* result = new EditorOpenResult{};
        result->ok = ok;
        if (ok) {
            result->message = L"Editor request sent.";
        } else {
            result->message = Utf8ToWide(error.empty() ? "editor open failed" : error);
        }
        if (!PostMessageW(hwnd, kMsgEditorOpenDone, 0, reinterpret_cast<LPARAM>(result))) {
            delete result;
            g_editorOpenPending.store(false, std::memory_order_release);
        }
    }).detach();
}

void RefreshStatus(HWND hwnd) {
    const bool effectsEnabled = MicMixApp::Instance().IsEffectsEnabled();
    const VstHostStatus host = MicMixApp::Instance().GetVstHostStatus();
    const auto music = MicMixApp::Instance().GetEffects(EffectChain::Music);
    const auto mic = MicMixApp::Instance().GetEffects(EffectChain::Mic);
    std::wstring line1 = L"Effects: ";
    line1 += effectsEnabled ? L"On" : L"Off";
    line1 += L"  |  Host: ";
    line1 += host.running ? L"Running" : L"Stopped";
    if (host.pid != 0) {
        line1 += L" (pid=" + std::to_wstring(host.pid) + L")";
    }
    std::wstring line2 = L"Music plugins=" + std::to_wstring(music.size()) +
                         L"  |  Mic plugins=" + std::to_wstring(mic.size()) +
                         L"  |  Monitor=" + std::wstring(MicMixApp::Instance().IsMonitorEnabled() ? L"On" : L"Off");
    std::wstring line3 = Utf8ToWide(host.message.empty() ? "host_status=idle" : ("host_status=" + host.message));
    SetStatusText(hwnd, line1 + L"\r\n" + line2 + L"\r\n" + line3);
    SetWindowTextW(hwnd, effectsEnabled ? L"MicMix Effects - ACTIVE" : L"MicMix Effects - OFF");
    UpdateMonitorButton(hwnd);
    InvalidateRect(hwnd, &g_rcHeaderBadge, TRUE);
}

void ApplyControlTheme(HWND hwnd) {
    EnumChildWindows(hwnd, [](HWND child, LPARAM) -> BOOL {
        wchar_t className[32]{};
        GetClassNameW(child, className, static_cast<int>(std::size(className)));
        if (_wcsicmp(className, L"ListBox") == 0) {
            SetWindowTheme(child, L"Explorer", nullptr);
        }
        if (_wcsicmp(className, L"Button") == 0) {
            const LONG_PTR style = GetWindowLongPtrW(child, GWL_STYLE);
            const LONG_PTR type = (style & BS_TYPEMASK);
            if (type == BS_PUSHBUTTON || type == BS_DEFPUSHBUTTON) {
                SetWindowLongPtrW(child, GWL_STYLE, style | BS_FLAT);
            }
            SetWindowTheme(child, L"Explorer", nullptr);
        }
        return TRUE;
    }, 0);
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

void DrawEffectListItem(const DRAWITEMSTRUCT* dis) {
    if (!dis || dis->itemID == static_cast<UINT>(-1)) {
        return;
    }
    RECT rc = dis->rcItem;
    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool focused = (dis->itemState & ODS_FOCUS) != 0;
    COLORREF bg = selected ? RGB(218, 234, 255) : RGB(255, 255, 255);
    COLORREF text = g_theme.text;
    HBRUSH bgBrush = CreateSolidBrush(bg);
    FillRect(dis->hDC, &rc, bgBrush);
    DeleteObject(bgBrush);

    std::wstring itemText;
    const LRESULT textLen = SendMessageW(dis->hwndItem, LB_GETTEXTLEN, dis->itemID, 0);
    if (textLen > 0 && textLen != LB_ERR) {
        itemText.resize(static_cast<size_t>(textLen));
        const LRESULT copied = SendMessageW(dis->hwndItem, LB_GETTEXT, dis->itemID, reinterpret_cast<LPARAM>(itemText.data()));
        if (copied == LB_ERR) {
            itemText.clear();
        }
    }
    if (itemText.empty()) {
        itemText = L"(empty)";
    }
    rc.left += S(8);
    rc.right -= S(6);
    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, text);
    SelectObject(dis->hDC, g_fontBody ? g_fontBody : g_fontSmall);
    DrawTextW(dis->hDC, itemText.c_str(), -1, &rc, DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    if (focused) {
        RECT focusRc = dis->rcItem;
        focusRc.left += 1;
        focusRc.right -= 1;
        focusRc.top += 1;
        focusRc.bottom -= 1;
        DrawFocusRect(dis->hDC, &focusRc);
    }
}

void DrawLevelMeter(HDC hdc, const RECT& meterRect, bool active, float visualDb, float holdDb) {
    if (meterRect.right <= meterRect.left || meterRect.bottom <= meterRect.top) {
        return;
    }
    const int targetWidth = static_cast<int>(meterRect.right - meterRect.left);
    const int targetHeight = static_cast<int>(meterRect.bottom - meterRect.top);
    if (targetWidth <= 2 || targetHeight <= 2) {
        return;
    }

    HDC drawDc = hdc;
    HDC memDc = CreateCompatibleDC(hdc);
    HBITMAP memBmp = nullptr;
    HGDIOBJ oldBmp = nullptr;
    bool buffered = false;
    if (memDc) {
        memBmp = CreateCompatibleBitmap(hdc, targetWidth, targetHeight);
        if (memBmp) {
            oldBmp = SelectObject(memDc, memBmp);
            drawDc = memDc;
            buffered = true;
        } else {
            DeleteDC(memDc);
            memDc = nullptr;
        }
    }

    RECT rc = buffered ? RECT{ 0, 0, targetWidth, targetHeight } : meterRect;
    const int left = static_cast<int>(rc.left);
    const int right = static_cast<int>(rc.right);
    const int width = right - left;
    const int height = static_cast<int>(rc.bottom - rc.top);
    if (width <= 2 || height <= 2) {
        if (buffered) {
            SelectObject(memDc, oldBmp);
            DeleteObject(memBmp);
            DeleteDC(memDc);
        }
        return;
    }

    constexpr float kYellowStartDb = -24.0f;
    constexpr float kRedStartDb = -12.0f;
    const int scaleBandH = std::clamp(S(10), 7, std::max(7, height - S(8)));
    RECT scaleRc = rc;
    scaleRc.bottom = std::min(rc.bottom, rc.top + scaleBandH);
    RECT barRc = rc;
    barRc.top = std::min(rc.bottom - 1, scaleRc.bottom);

    FillSolidRect(drawDc, scaleRc, RGB(244, 247, 252));
    FillSolidRect(drawDc, barRc, RGB(235, 240, 246));

    const std::array<int, 7> tickDb = { -60, -36, -24, -18, -12, -6, 0 };
    SetBkMode(drawDc, TRANSPARENT);
    SetTextColor(drawDc, RGB(96, 107, 122));
    HGDIOBJ oldFont = SelectObject(drawDc, g_fontTiny ? g_fontTiny : (g_fontHint ? g_fontHint : g_fontSmall));
    for (size_t i = 0; i < tickDb.size(); ++i) {
        const int db = tickDb[i];
        int x = left + MeterDbToPixels(static_cast<float>(db), width);
        x = std::clamp(x, left, right - 1);

        const COLORREF guideColor = (db == -24)
            ? RGB(198, 176, 104)
            : (db == -12 || db == 0)
                ? RGB(198, 130, 124)
                : RGB(206, 214, 225);
        HPEN guidePen = CreatePen(PS_SOLID, 1, guideColor);
        HGDIOBJ oldGuidePen = SelectObject(drawDc, guidePen);
        MoveToEx(drawDc, x, barRc.top, nullptr);
        LineTo(drawDc, x, barRc.bottom);
        SelectObject(drawDc, oldGuidePen);
        DeleteObject(guidePen);

        RECT tickRc{ x, scaleRc.bottom - S(3), x + 1, scaleRc.bottom };
        FillSolidRect(drawDc, tickRc, RGB(162, 171, 184));

        wchar_t label[8]{};
        swprintf_s(label, L"%d", db);
        SIZE textSize{};
        if (GetTextExtentPoint32W(drawDc, label, lstrlenW(label), &textSize) != 0) {
            int textLeft = x - (textSize.cx / 2);
            if (i == 0) {
                textLeft = left + S(2);
            } else if (i == (tickDb.size() - 1)) {
                textLeft = right - textSize.cx - S(2);
            } else {
                const int labelMinX = left + S(2);
                const int labelMaxX = right - static_cast<int>(textSize.cx) - S(2);
                textLeft = std::clamp(textLeft, labelMinX, labelMaxX);
            }
            TextOutW(drawDc, textLeft, scaleRc.top, label, lstrlenW(label));
        }
    }
    SelectObject(drawDc, oldFont);

    const int barWidth = std::max(1, static_cast<int>(barRc.right - barRc.left));
    const int levelPx = std::clamp(MeterDbToPixels(visualDb, width), 0, width);
    const int holdPx = std::clamp(MeterDbToPixels(holdDb, width), 0, width - 1);
    const int greenEnd = std::clamp(MeterDbToPixels(kYellowStartDb, barWidth), 0, barWidth);
    const int yellowEnd = std::clamp(MeterDbToPixels(kRedStartDb, barWidth), 0, barWidth);

    if (active && levelPx > 0) {
        RECT seg = barRc;
        seg.right = barRc.left + std::min(levelPx, greenEnd);
        if (seg.right > seg.left) FillSolidRect(drawDc, seg, RGB(54, 181, 78));

        if (levelPx > greenEnd) {
            seg.left = barRc.left + greenEnd;
            seg.right = barRc.left + std::min(levelPx, yellowEnd);
            if (seg.right > seg.left) FillSolidRect(drawDc, seg, RGB(232, 191, 58));
        }
        if (levelPx > yellowEnd) {
            seg.left = barRc.left + yellowEnd;
            seg.right = barRc.left + levelPx;
            if (seg.right > seg.left) FillSolidRect(drawDc, seg, RGB(214, 75, 63));
        }

        HPEN holdPen = CreatePen(PS_SOLID, 1, RGB(255, 255, 255));
        HGDIOBJ oldPen = SelectObject(drawDc, holdPen);
        MoveToEx(drawDc, barRc.left + holdPx, barRc.top, nullptr);
        LineTo(drawDc, barRc.left + holdPx, barRc.bottom);
        SelectObject(drawDc, oldPen);
        DeleteObject(holdPen);
    }

    HPEN border = CreatePen(PS_SOLID, 1, RGB(184, 193, 207));
    HGDIOBJ oldPen = SelectObject(drawDc, border);
    HGDIOBJ oldBrush = SelectObject(drawDc, GetStockObject(HOLLOW_BRUSH));
    Rectangle(drawDc, rc.left, rc.top, rc.right, rc.bottom);
    SelectObject(drawDc, oldBrush);
    SelectObject(drawDc, oldPen);
    DeleteObject(border);

    if (buffered) {
        BitBlt(hdc, meterRect.left, meterRect.top, targetWidth, targetHeight, memDc, 0, 0, SRCCOPY);
        SelectObject(memDc, oldBmp);
        DeleteObject(memBmp);
        DeleteDC(memDc);
    }
}

void DrawClipStrip(HDC hdc, const RECT& rcStrip, float dangerUnit, bool clipRecent) {
    if (rcStrip.right <= rcStrip.left || rcStrip.bottom <= rcStrip.top) {
        return;
    }
    const int targetWidth = static_cast<int>(rcStrip.right - rcStrip.left);
    const int targetHeight = static_cast<int>(rcStrip.bottom - rcStrip.top);
    if (targetWidth <= 2 || targetHeight <= 2) {
        return;
    }

    HDC drawDc = hdc;
    HDC memDc = CreateCompatibleDC(hdc);
    HBITMAP memBmp = nullptr;
    HGDIOBJ oldBmp = nullptr;
    bool buffered = false;
    if (memDc) {
        memBmp = CreateCompatibleBitmap(hdc, targetWidth, targetHeight);
        if (memBmp) {
            oldBmp = SelectObject(memDc, memBmp);
            drawDc = memDc;
            buffered = true;
        } else {
            DeleteDC(memDc);
            memDc = nullptr;
        }
    }

    RECT rc = buffered ? RECT{ 0, 0, targetWidth, targetHeight } : rcStrip;
    FillSolidRect(drawDc, rc, RGB(246, 249, 252));

    const int width = rc.right - rc.left;
    constexpr int kSegments = 18;
    const int segGap = std::max(1, S(1));
    const int litSegments = std::clamp(static_cast<int>(std::floor(dangerUnit * static_cast<float>(kSegments) + 0.5f)), 0, kSegments);
    const COLORREF offColor = RGB(218, 225, 235);
    const COLORREF clipColor = clipRecent ? RGB(226, 63, 51) : RGB(196, 72, 64);

    for (int i = 0; i < kSegments; ++i) {
        const int segLeft = rc.left + ((i * width) / kSegments);
        const int segRightFull = rc.left + (((i + 1) * width) / kSegments);
        int segRight = segRightFull;
        if (i < (kSegments - 1)) {
            segRight = std::max(segLeft + 1, segRightFull - segGap);
        }
        RECT seg{ segLeft, rc.top + S(2), segRight, rc.bottom - S(2) };
        float t = static_cast<float>(i) / static_cast<float>(kSegments - 1);
        COLORREF onColor = RGB(
            static_cast<BYTE>(84 + (154.0f * t)),
            static_cast<BYTE>(176 - (126.0f * t)),
            static_cast<BYTE>(86 - (40.0f * t)));
        COLORREF color = (i < litSegments) ? onColor : offColor;
        FillSolidRect(drawDc, seg, color);
    }

    HPEN border = CreatePen(PS_SOLID, 1, clipRecent ? RGB(208, 72, 60) : RGB(188, 197, 210));
    HGDIOBJ oldPen = SelectObject(drawDc, border);
    HGDIOBJ oldBrush = SelectObject(drawDc, GetStockObject(HOLLOW_BRUSH));
    Rectangle(drawDc, rc.left, rc.top, rc.right, rc.bottom);
    SelectObject(drawDc, oldBrush);
    SelectObject(drawDc, oldPen);
    DeleteObject(border);

    SetBkMode(drawDc, TRANSPARENT);
    SetTextColor(drawDc, clipRecent ? RGB(132, 24, 24) : RGB(112, 120, 132));
    SelectObject(drawDc, g_fontTiny ? g_fontTiny : (g_fontHint ? g_fontHint : g_fontSmall));
    RECT labelRc{ rc.left + S(5), rc.top, rc.right - S(52), rc.bottom };
    DrawTextW(drawDc, L"Clip Meter", -1, &labelRc, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);

    RECT inner{ rc.left + 1, rc.top + 1, rc.right - 1, rc.bottom - 1 };
    RECT pulse{};
    const int pulseWidth = S(44);
    if (clipRecent) {
        pulse = { std::max(inner.left, inner.right - pulseWidth), inner.top, inner.right, inner.bottom };
        FillSolidRect(drawDc, pulse, clipColor);
    }

    if (clipRecent) {
        SetTextColor(drawDc, RGB(255, 255, 255));
        SelectObject(drawDc, g_fontTiny ? g_fontTiny : (g_fontHint ? g_fontHint : g_fontSmall));
        DrawTextW(drawDc, L"CLIP", -1, &pulse, DT_SINGLELINE | DT_CENTER | DT_VCENTER);
    }

    if (buffered) {
        BitBlt(hdc, rcStrip.left, rcStrip.top, targetWidth, targetHeight, memDc, 0, 0, SRCCOPY);
        SelectObject(memDc, oldBmp);
        DeleteObject(memBmp);
        DeleteDC(memDc);
    }
}

void UpdateMeterVisuals(HWND hwnd) {
    const TelemetrySnapshot t = MicMixApp::Instance().GetTelemetry();
    g_lastTelemetry = t;

    const float shownMusicPeakDb = (t.musicSendPeakDbfs > -119.0f) ? t.musicSendPeakDbfs : t.musicPeakDbfs;
    const bool musicMeterActive = IsMusicMeterActive(t);
    float musicTargetDb = -60.0f;
    if (musicMeterActive) {
        musicTargetDb = std::clamp(shownMusicPeakDb, -60.0f, 0.0f);
    }
    if (musicTargetDb > g_musicMeterVisualDb) {
        g_musicMeterVisualDb += (musicTargetDb - g_musicMeterVisualDb) * 0.70f;
    } else {
        g_musicMeterVisualDb += (musicTargetDb - g_musicMeterVisualDb) * 0.30f;
    }
    g_musicMeterVisualDb = std::clamp(g_musicMeterVisualDb, -60.0f, 0.0f);
    if (musicTargetDb >= g_musicMeterHoldDb) {
        g_musicMeterHoldDb = musicTargetDb;
    } else {
        g_musicMeterHoldDb = std::max(-60.0f, g_musicMeterHoldDb - 0.6f);
    }

    const int musicMeterWidth = std::max(1, static_cast<int>(g_rcMusicMeter.right - g_rcMusicMeter.left));
    const int musicLevelPx = std::clamp(MeterDbToPixels(g_musicMeterVisualDb, musicMeterWidth), 0, musicMeterWidth);
    const int musicHoldPx = std::clamp(MeterDbToPixels(g_musicMeterHoldDb, musicMeterWidth), 0, musicMeterWidth - 1);
    const bool musicMeterChanged =
        (musicLevelPx != g_lastMusicMeterLevelPx) ||
        (musicHoldPx != g_lastMusicMeterHoldPx) ||
        (musicMeterActive != g_lastMusicMeterActive);
    g_lastMusicMeterLevelPx = musicLevelPx;
    g_lastMusicMeterHoldPx = musicHoldPx;
    g_lastMusicMeterActive = musicMeterActive;
    if (musicMeterChanged) {
        InvalidateRect(hwnd, &g_rcMusicMeter, FALSE);
    }

    const int musicClipLitSegments = ClipLitSegmentsFromDb(shownMusicPeakDb);
    const bool musicClipChanged =
        (musicClipLitSegments != g_lastMusicClipLitSegments) ||
        (t.sourceClipRecent != g_lastMusicClipRecent);
    g_lastMusicClipLitSegments = musicClipLitSegments;
    g_lastMusicClipRecent = t.sourceClipRecent;
    if (musicClipChanged) {
        InvalidateRect(hwnd, &g_rcMusicClip, FALSE);
    }

    float micTargetDb = -60.0f;
    if (t.micRmsDbfs > -119.0f) {
        micTargetDb = std::clamp(t.micRmsDbfs, -60.0f, 0.0f);
    }
    if (micTargetDb > g_micMeterVisualDb) {
        g_micMeterVisualDb += (micTargetDb - g_micMeterVisualDb) * 0.68f;
    } else {
        g_micMeterVisualDb += (micTargetDb - g_micMeterVisualDb) * 0.28f;
    }
    g_micMeterVisualDb = std::clamp(g_micMeterVisualDb, -60.0f, 0.0f);
    if (micTargetDb >= g_micMeterHoldDb) {
        g_micMeterHoldDb = micTargetDb;
    } else {
        g_micMeterHoldDb = std::max(-60.0f, g_micMeterHoldDb - 0.7f);
    }

    const bool micMeterActive = t.micRmsDbfs > -119.0f;
    const int micMeterWidth = std::max(1, static_cast<int>(g_rcMicMeter.right - g_rcMicMeter.left));
    const int micLevelPx = std::clamp(MeterDbToPixels(g_micMeterVisualDb, micMeterWidth), 0, micMeterWidth);
    const int micHoldPx = std::clamp(MeterDbToPixels(g_micMeterHoldDb, micMeterWidth), 0, micMeterWidth - 1);
    const bool micMeterChanged =
        (micLevelPx != g_lastMicMeterLevelPx) ||
        (micHoldPx != g_lastMicMeterHoldPx) ||
        (micMeterActive != g_lastMicMeterActive);
    g_lastMicMeterLevelPx = micLevelPx;
    g_lastMicMeterHoldPx = micHoldPx;
    g_lastMicMeterActive = micMeterActive;
    if (micMeterChanged) {
        InvalidateRect(hwnd, &g_rcMicMeter, FALSE);
    }

    const int micClipLitSegments = ClipLitSegmentsFromDb(t.micPeakDbfs);
    const bool micClipChanged =
        (micClipLitSegments != g_lastMicClipLitSegments) ||
        (t.micClipRecent != g_lastMicClipRecent);
    g_lastMicClipLitSegments = micClipLitSegments;
    g_lastMicClipRecent = t.micClipRecent;
    if (micClipChanged) {
        InvalidateRect(hwnd, &g_rcMicClip, FALSE);
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        g_dpi = GetDpiForWindow(hwnd);
        EnsureUiResources();
        ComputeLayout(hwnd);
        g_lastTelemetry = {};
        g_musicMeterVisualDb = -60.0f;
        g_musicMeterHoldDb = -60.0f;
        g_micMeterVisualDb = -60.0f;
        g_micMeterHoldDb = -60.0f;
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
        g_lastListRefreshTickMs = 0;

        const int contentLeft = S(kCardMarginPx + kCardInnerPaddingPx);
        const int contentRight = S(kClientWidthPx - kCardMarginPx - kCardInnerPaddingPx);
        const int controlGap = S(12);
        const int actionButtonW = S(142);
        const int monitorButtonW = S(122);
        const int headerMetaW = S(300);
        const int headerMetaX = contentRight - headerMetaW;
        const int titleX = contentLeft;
        const int titleY = S(14);
        const int titleH = S(36);
        const int titleW = S(190);

        const std::wstring versionText = Utf8ToWide(std::string("v") + MICMIX_VERSION + "  by dabinuss");

        CreateWindowW(L"STATIC", L"MicMix", WS_CHILD | WS_VISIBLE, titleX, titleY, titleW, titleH, hwnd, reinterpret_cast<HMENU>(IDC_TITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Configure MicMix Effects and route VST chains for music and mic", WS_CHILD | WS_VISIBLE, contentLeft, S(43), headerMetaX - contentLeft - S(8), S(24), hwnd, reinterpret_cast<HMENU>(IDC_SUBTITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", versionText.c_str(), WS_CHILD | WS_VISIBLE | SS_RIGHT, headerMetaX, S(18), headerMetaW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_VERSION), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"github.com/dabinuss/MicMix", WS_CHILD | WS_VISIBLE | SS_RIGHT | SS_NOTIFY, headerMetaX, S(36), headerMetaW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_REPO_LINK), nullptr, nullptr);

        const int topLeft = g_rcTop.left + S(kCardInnerPaddingPx);
        const int topY = g_rcTop.top + S(kCardInnerPaddingPx);
        CreateWindowW(L"BUTTON", L"Enable Effects", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, topLeft, topY, actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_ENABLE_EFFECTS), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Disable Effects", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, topLeft + actionButtonW + controlGap, topY, actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_DISABLE_EFFECTS), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Monitor Mix: Off", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, topLeft + (actionButtonW + controlGap) * 2, topY, monitorButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_MONITOR), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Shared with MicMix monitor path", WS_CHILD | WS_VISIBLE, topLeft + (actionButtonW + controlGap) * 2, topY + S(38), S(220), S(18), hwnd, reinterpret_cast<HMENU>(IDC_MONITOR_HINT), nullptr, nullptr);

        const int sectionLeft = g_rcMusic.left + S(kCardInnerPaddingPx);
        const int sectionTop = g_rcMusic.top + S(kCardInnerPaddingPx);
        const int musicListTop = sectionTop + S(84);
        CreateWindowW(L"STATIC", L"AUDIO SOURCE", WS_CHILD | WS_VISIBLE, sectionLeft, sectionTop, S(220), S(20), hwnd, reinterpret_cast<HMENU>(IDC_AUDIO_SECTION_TITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Audio Meter", WS_CHILD | WS_VISIBLE, sectionLeft, sectionTop + S(34), S(120), S(20), hwnd, reinterpret_cast<HMENU>(IDC_AUDIO_METER_LABEL), nullptr, nullptr);
        CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
            sectionLeft, musicListTop, S(458), S(90), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_LIST), nullptr, nullptr);
        const int musicBtnX = sectionLeft + S(468);
        CreateWindowW(L"BUTTON", L"Add", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX, musicListTop, S(86), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_ADD), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Remove", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX, musicListTop + S(30), S(86), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_REMOVE), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Up", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX, musicListTop + S(60), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_UP), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Down", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX + S(46), musicListTop + S(60), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_DOWN), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Byp", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX, musicListTop + S(90), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_BYPASS), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Open", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX + S(46), musicListTop + S(90), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_EDITOR), nullptr, nullptr);

        const int micLeft = g_rcMic.left + S(kCardInnerPaddingPx);
        const int micTop = g_rcMic.top + S(kCardInnerPaddingPx);
        const int micListTop = micTop + S(84);
        CreateWindowW(L"STATIC", L"MIC SOURCE", WS_CHILD | WS_VISIBLE, micLeft, micTop, S(220), S(20), hwnd, reinterpret_cast<HMENU>(IDC_MIC_SECTION_TITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Mic Meter", WS_CHILD | WS_VISIBLE, micLeft, micTop + S(34), S(120), S(20), hwnd, reinterpret_cast<HMENU>(IDC_MIC_METER_LABEL), nullptr, nullptr);
        CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
            micLeft, micListTop, S(458), S(90), hwnd, reinterpret_cast<HMENU>(IDC_MIC_LIST), nullptr, nullptr);
        const int micBtnX = micLeft + S(468);
        CreateWindowW(L"BUTTON", L"Add", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX, micListTop, S(86), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_ADD), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Remove", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX, micListTop + S(30), S(86), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_REMOVE), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Up", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX, micListTop + S(60), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_UP), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Down", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX + S(46), micListTop + S(60), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_DOWN), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Byp", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX, micListTop + S(90), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_BYPASS), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Open", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX + S(46), micListTop + S(90), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_EDITOR), nullptr, nullptr);

        CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT, g_rcStatus.left + S(kCardInnerPaddingPx), g_rcStatus.top + S(14), (g_rcStatus.right - g_rcStatus.left) - S(kCardInnerPaddingPx * 2), g_rcStatus.bottom - g_rcStatus.top - S(18), hwnd, reinterpret_cast<HMENU>(IDC_STATUS), nullptr, nullptr);

        ApplyControlTheme(hwnd);
        ApplyFonts(hwnd);
        ReloadAllLists(hwnd);
        RefreshStatus(hwnd);
        UpdateMeterVisuals(hwnd);
        SetTimer(hwnd, kTimerStatus, kTimerMs, nullptr);
        return 0;
    }
    case WM_MEASUREITEM: {
        auto* mis = reinterpret_cast<MEASUREITEMSTRUCT*>(lParam);
        if (!mis) {
            return FALSE;
        }
        if (mis->CtlID == IDC_MUSIC_LIST || mis->CtlID == IDC_MIC_LIST) {
            mis->itemHeight = static_cast<UINT>(S(kListItemHeightPx));
            return TRUE;
        }
        break;
    }
    case WM_DRAWITEM: {
        auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
        if (!dis) {
            return FALSE;
        }
        if (dis->CtlID == IDC_MUSIC_LIST || dis->CtlID == IDC_MIC_LIST) {
            DrawEffectListItem(dis);
            return TRUE;
        }
        break;
    }
    case WM_TIMER:
        if (wParam == kTimerStatus) {
            const uint64_t nowMs = GetTickCount64();
            if (nowMs > (g_lastListRefreshTickMs + 1000ULL)) {
                ReloadAllLists(hwnd);
                g_lastListRefreshTickMs = nowMs;
            }
            RefreshStatus(hwnd);
            UpdateMeterVisuals(hwnd);
            return 0;
        }
        break;
    case WM_COMMAND:
        if (HIWORD(wParam) == LBN_DBLCLK) {
            switch (LOWORD(wParam)) {
            case IDC_MUSIC_LIST:
                HandleOpenEditor(hwnd, EffectChain::Music, IDC_MUSIC_LIST);
                return 0;
            case IDC_MIC_LIST:
                HandleOpenEditor(hwnd, EffectChain::Mic, IDC_MIC_LIST);
                return 0;
            default:
                break;
            }
        }
        switch (LOWORD(wParam)) {
        case IDC_REPO_LINK:
            if (HIWORD(wParam) == STN_CLICKED) {
                ShellExecuteW(hwnd, L"open", kRepoUrl, nullptr, nullptr, SW_SHOWNORMAL);
            }
            return 0;
        case IDC_ENABLE_EFFECTS:
            MicMixApp::Instance().SetEffectsEnabled(true, true);
            RefreshStatus(hwnd);
            return 0;
        case IDC_DISABLE_EFFECTS:
            MicMixApp::Instance().SetEffectsEnabled(false, true);
            RefreshStatus(hwnd);
            return 0;
        case IDC_MONITOR:
            MicMixApp::Instance().ToggleMonitor();
            UpdateMonitorButton(hwnd);
            RefreshStatus(hwnd);
            return 0;
        case IDC_MUSIC_ADD:
            HandleAddEffect(hwnd, EffectChain::Music);
            return 0;
        case IDC_MUSIC_REMOVE:
            HandleRemoveEffect(hwnd, EffectChain::Music, IDC_MUSIC_LIST);
            return 0;
        case IDC_MUSIC_UP:
            HandleMoveEffect(hwnd, EffectChain::Music, IDC_MUSIC_LIST, -1);
            return 0;
        case IDC_MUSIC_DOWN:
            HandleMoveEffect(hwnd, EffectChain::Music, IDC_MUSIC_LIST, 1);
            return 0;
        case IDC_MUSIC_BYPASS:
            HandleBypassEffect(hwnd, EffectChain::Music, IDC_MUSIC_LIST);
            return 0;
        case IDC_MUSIC_EDITOR:
            HandleOpenEditor(hwnd, EffectChain::Music, IDC_MUSIC_LIST);
            return 0;
        case IDC_MIC_ADD:
            HandleAddEffect(hwnd, EffectChain::Mic);
            return 0;
        case IDC_MIC_REMOVE:
            HandleRemoveEffect(hwnd, EffectChain::Mic, IDC_MIC_LIST);
            return 0;
        case IDC_MIC_UP:
            HandleMoveEffect(hwnd, EffectChain::Mic, IDC_MIC_LIST, -1);
            return 0;
        case IDC_MIC_DOWN:
            HandleMoveEffect(hwnd, EffectChain::Mic, IDC_MIC_LIST, 1);
            return 0;
        case IDC_MIC_BYPASS:
            HandleBypassEffect(hwnd, EffectChain::Mic, IDC_MIC_LIST);
            return 0;
        case IDC_MIC_EDITOR:
            HandleOpenEditor(hwnd, EffectChain::Mic, IDC_MIC_LIST);
            return 0;
        default:
            break;
        }
        break;
    case kMsgEditorOpenDone: {
        std::unique_ptr<EditorOpenResult> result(reinterpret_cast<EditorOpenResult*>(lParam));
        g_editorOpenPending.store(false, std::memory_order_release);
        if (result) {
            SetStatusText(hwnd, result->message);
            if (result->ok) {
                ReloadAllLists(hwnd);
                RefreshStatus(hwnd);
            }
        }
        return 0;
    }
    case WM_SETCURSOR: {
        const HWND target = reinterpret_cast<HWND>(wParam);
        const UINT hit = LOWORD(lParam);
        if (target && hit == HTCLIENT) {
            const int id = GetDlgCtrlID(target);
            if (IsHandCursorControlId(id)) {
                SetCursor(LoadCursorW(nullptr, IDC_HAND));
                return TRUE;
            }
        }
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
    case WM_CTLCOLORSTATIC: {
        HDC hdc = reinterpret_cast<HDC>(wParam);
        HWND ctl = reinterpret_cast<HWND>(lParam);
        RECT rc{};
        GetWindowRect(ctl, &rc);
        MapWindowPoints(nullptr, hwnd, reinterpret_cast<LPPOINT>(&rc), 2);
        POINT pt{ rc.left, rc.top };
        const bool inCard = PtInRect(&g_rcTop, pt) || PtInRect(&g_rcMusic, pt) || PtInRect(&g_rcMic, pt) || PtInRect(&g_rcStatus, pt);
        SetBkMode(hdc, TRANSPARENT);
        const int id = GetDlgCtrlID(ctl);
        if (id == IDC_TITLE) {
            SetTextColor(hdc, g_theme.text);
        } else if (id == IDC_SUBTITLE || id == IDC_VERSION || id == IDC_MONITOR_HINT ||
                   id == IDC_AUDIO_SECTION_TITLE || id == IDC_MIC_SECTION_TITLE) {
            SetTextColor(hdc, g_theme.muted);
        } else if (id == IDC_REPO_LINK) {
            SetTextColor(hdc, RGB(25, 100, 200));
        } else {
            SetTextColor(hdc, g_theme.text);
        }
        return reinterpret_cast<LRESULT>(inCard ? g_brushCard : g_brushBg);
    }
    case WM_ERASEBKGND:
        return 1;
    case WM_PAINT: {
        PAINTSTRUCT ps{};
        HDC dc = BeginPaint(hwnd, &ps);
        RECT client{};
        GetClientRect(hwnd, &client);
        FillRect(dc, &client, g_brushBg ? g_brushBg : reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
        RECT accent{ S(0), S(0), S(kClientWidthPx), S(4) };
        HBRUSH accentBrush = CreateSolidBrush(g_theme.accent);
        FillRect(dc, &accent, accentBrush);
        DeleteObject(accentBrush);

        DrawFlatCardFrame(dc, g_rcTop, g_brushCard, g_theme.border);
        DrawFlatCardFrame(dc, g_rcMusic, g_brushCard, g_theme.border);
        DrawFlatCardFrame(dc, g_rcMic, g_brushCard, g_theme.border);
        DrawFlatCardFrame(dc, g_rcStatus, g_brushCard, g_theme.border);
        DrawLevelMeter(dc, g_rcMusicMeter, g_lastMusicMeterActive, g_musicMeterVisualDb, g_musicMeterHoldDb);
        DrawClipStrip(dc, g_rcMusicClip, ClipDangerUnitFromDb((g_lastTelemetry.musicSendPeakDbfs > -119.0f) ? g_lastTelemetry.musicSendPeakDbfs : g_lastTelemetry.musicPeakDbfs), g_lastTelemetry.sourceClipRecent);
        DrawLevelMeter(dc, g_rcMicMeter, g_lastMicMeterActive, g_micMeterVisualDb, g_micMeterHoldDb);
        DrawClipStrip(dc, g_rcMicClip, ClipDangerUnitFromDb(g_lastTelemetry.micPeakDbfs), g_lastTelemetry.micClipRecent);
        DrawHeaderBadge(
            dc,
            g_rcHeaderBadge,
            GetHeaderBadgeVisual(ResolveBadgeState()),
            g_fontSmall,
            g_fontBody,
            g_dpi);

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_DESTROY:
        KillTimer(hwnd, kTimerStatus);
        g_editorOpenPending.store(false, std::memory_order_release);
        g_hwnd.store(nullptr, std::memory_order_release);
        PostQuitMessage(0);
        return 0;
    case WM_CLOSE:
        DestroyWindow(hwnd);
        return 0;
    default:
        break;
    }
    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

void WindowThreadMain() {
    HINSTANCE hInst = GetPluginModuleHandle();
    INITCOMMONCONTROLSEX icc{};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_STANDARD_CLASSES | ICC_WIN95_CLASSES | ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icc);

    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInst;
    wc.lpszClassName = L"MicMixEffectsWindow";
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APP_ICON));
    wc.hIconSm = LoadIconW(hInst, MAKEINTRESOURCEW(IDI_APP_ICON_SMALL));
    wc.hbrBackground = nullptr;
    UnregisterClassW(wc.lpszClassName, hInst);
    RegisterClassExW(&wc);

    RECT rc{ 0, 0, S(kClientWidthPx), S(kClientHeightPx) };
    AdjustWindowRectEx(&rc, WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX, FALSE, 0);
    HWND hwnd = CreateWindowExW(
        0,
        wc.lpszClassName,
        L"MicMix Effects",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        CW_USEDEFAULT,
        CW_USEDEFAULT,
        rc.right - rc.left,
        rc.bottom - rc.top,
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

    MSG msg{};
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    ReleaseUiResources();
    g_running.store(false, std::memory_order_release);
}

} // namespace

EffectsWindowController& EffectsWindowController::Instance() {
    static EffectsWindowController instance;
    return instance;
}

void EffectsWindowController::Open() {
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
        LogError(std::string("effects_window thread start failed: ") + ex.what());
    } catch (...) {
        g_running.store(false, std::memory_order_release);
        LogError("effects_window thread start failed: unknown");
    }
}

void EffectsWindowController::Close() {
    std::lock_guard<std::mutex> lock(g_mutex);
    if (!g_running.load(std::memory_order_acquire)) {
        if (g_thread.joinable()) {
            g_thread.join();
        }
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
