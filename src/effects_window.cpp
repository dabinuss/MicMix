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
    IDC_VST_AUTOSTART = 4214,

    IDC_MUSIC_LIST = 4220,
    IDC_MUSIC_ADD = 4221,
    IDC_MUSIC_REMOVE = 4222,
    IDC_MUSIC_UP = 4223,
    IDC_MUSIC_DOWN = 4224,
    IDC_MUSIC_BYPASS = 4225,
    IDC_MUSIC_ENABLE = 4226,
    IDC_MUSIC_EDITOR = 4227,

    IDC_MIC_LIST = 4230,
    IDC_MIC_ADD = 4231,
    IDC_MIC_REMOVE = 4232,
    IDC_MIC_UP = 4233,
    IDC_MIC_DOWN = 4234,
    IDC_MIC_BYPASS = 4235,
    IDC_MIC_ENABLE = 4236,
    IDC_MIC_EDITOR = 4237,

    IDC_STATUS = 4240,

    IDC_AUDIO_SECTION_TITLE = 4250,
    IDC_AUDIO_METER_LABEL = 4251,
    IDC_MIC_SECTION_TITLE = 4252,
    IDC_MIC_METER_LABEL = 4253,
    IDC_AUDIO_METER_TEXT = 4254,
    IDC_AUDIO_CLIP_EVENTS = 4255,
    IDC_MIC_METER_TEXT = 4256,
    IDC_MIC_CLIP_EVENTS = 4257,
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
HeaderBadgeState g_headerBadgeState = HeaderBadgeState::Off;
bool g_headerBadgeStateValid = false;
std::atomic_uint64_t g_statusHoldUntilMs{0};
uint64_t g_lastListRefreshTickMs = 0;
std::atomic<bool> g_loading{false};
std::atomic<bool> g_editorOpenPending{false};
bool g_ownerVstAutostartChecked = false;

constexpr UINT kMsgEditorOpenDone = WM_APP + 42;

struct EditorOpenResult {
    bool ok = false;
    std::wstring message;
};

HFONT g_fontBody = nullptr;
HFONT g_fontSmall = nullptr;
HFONT g_fontHint = nullptr;
HFONT g_fontTiny = nullptr;
HFONT g_fontControlLarge = nullptr;
HFONT g_fontTitle = nullptr;
HFONT g_fontMono = nullptr;
HBRUSH g_brushBg = nullptr;
HBRUSH g_brushCard = nullptr;
UiCommonResources g_commonUi{};

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
std::wstring g_lastMusicMeterText;
std::wstring g_lastMicMeterText;
std::wstring g_lastMusicClipEventsText;
std::wstring g_lastMicClipEventsText;

constexpr int kClientWidthPx = kUiClientWidthPx;
constexpr int kClientHeightPx = kUiClientHeightPx;
constexpr int kCardMarginPx = kUiCardMarginPx;
constexpr int kCardGapPx = kUiCardGapPx;
constexpr int kCardInnerPaddingPx = kUiCardInnerPaddingPx;
constexpr UINT kTimerStatus = 1;
constexpr UINT kTimerMs = 40;
constexpr int kListItemHeightPx = 22;
constexpr wchar_t kRepoUrl[] = L"https://github.com/dabinuss/MicMix";
constexpr wchar_t kEffectsHeaderTitle[] = L"MicMix - Audio Effects";
constexpr wchar_t kEffectsSectionTopTitle[] = L"EFFECTS SETTINGS";
constexpr wchar_t kEffectsSectionMusicTitle[] = L"AUDIO EFFECTS";
constexpr wchar_t kEffectsSectionMicTitle[] = L"MIC EFFECTS";

int S(int px) {
    return MulDiv(px, g_dpi, 96);
}

bool IsHandCursorControlId(int id) {
    switch (id) {
    case IDC_REPO_LINK:
    case IDC_ENABLE_EFFECTS:
    case IDC_DISABLE_EFFECTS:
    case IDC_MONITOR:
    case IDC_VST_AUTOSTART:
    case IDC_MUSIC_LIST:
    case IDC_MUSIC_ADD:
    case IDC_MUSIC_REMOVE:
    case IDC_MUSIC_UP:
    case IDC_MUSIC_DOWN:
    case IDC_MUSIC_ENABLE:
    case IDC_MUSIC_BYPASS:
    case IDC_MUSIC_EDITOR:
    case IDC_MIC_LIST:
    case IDC_MIC_ADD:
    case IDC_MIC_REMOVE:
    case IDC_MIC_UP:
    case IDC_MIC_DOWN:
    case IDC_MIC_ENABLE:
    case IDC_MIC_BYPASS:
    case IDC_MIC_EDITOR:
        return true;
    default:
        return false;
    }
}

void SetStatusText(HWND hwnd, const std::wstring& text, uint64_t holdMs = 0) {
    g_statusText = text;
    if (holdMs > 0) {
        g_statusHoldUntilMs.store(GetTickCount64() + holdMs, std::memory_order_release);
    } else {
        g_statusHoldUntilMs.store(0, std::memory_order_release);
    }
    HWND ctl = GetDlgItem(hwnd, IDC_STATUS);
    if (ctl) {
        SetWindowTextW(ctl, g_statusText.c_str());
    }
}

void SetTransientStatusText(HWND hwnd, const std::wstring& text) {
    constexpr uint64_t kStatusHoldMs = 2400ULL;
    SetStatusText(hwnd, text, kStatusHoldMs);
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
}

void ReleaseUiResources() {
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

void ComputeLayout(HWND hwnd) {
    const int left = S(kCardMarginPx);
    const int right = S(kClientWidthPx - kCardMarginPx);
    const int cardW = right - left;
    const int topY = S(kUiTopCardYpx);
    const int topH = S(128);
    const int musicY = topY + topH + S(kCardGapPx);
    const int musicH = S(214);
    const int micY = musicY + musicH + S(kCardGapPx);
    const int micH = S(214);
    const int statusY = micY + micH + S(kCardGapPx);
    const int statusH = S(84);
    g_rcTop = { left, topY, left + cardW, topY + topH };
    g_rcMusic = { left, musicY, left + cardW, musicY + musicH };
    g_rcMic = { left, micY, left + cardW, micY + micH };
    g_rcStatus = { left, statusY, left + cardW, statusY + statusH };

    const int meterWidth = S(340);
    const int sectionPad = S(kCardInnerPaddingPx);
    const int musicSectionTop = g_rcMusic.top + sectionPad;
    const int musicSectionLeft = g_rcMusic.left + sectionPad;
    const int musicMeterLeft = musicSectionLeft + S(144);
    g_rcMusicMeter = { musicMeterLeft, musicSectionTop + S(32), musicMeterLeft + meterWidth, musicSectionTop + S(32) + S(28) };
    g_rcMusicClip = { musicMeterLeft, musicSectionTop + S(62), musicMeterLeft + meterWidth, musicSectionTop + S(62) + S(12) };

    const int micSectionTop = g_rcMic.top + sectionPad;
    const int micSectionLeft = g_rcMic.left + sectionPad;
    const int micMeterLeft = micSectionLeft + S(144);
    g_rcMicMeter = { micMeterLeft, micSectionTop + S(32), micMeterLeft + meterWidth, micSectionTop + S(32) + S(28) };
    g_rcMicClip = { micMeterLeft, micSectionTop + S(62), micMeterLeft + meterWidth, micSectionTop + S(62) + S(12) };

    const int contentLeft = S(kCardMarginPx + kCardInnerPaddingPx);
    const int titleY = S(10);
    const int titleH = S(42);
    HDC dc = GetDC(hwnd);
    const HeaderLayout headerLayout = ComputeHeaderLayout(
        dc,
        g_fontTitle,
        g_fontSmall,
        g_fontBody,
        g_dpi,
        kEffectsHeaderTitle,
        contentLeft,
        S(kClientWidthPx - kCardMarginPx - kCardInnerPaddingPx),
        titleY,
        titleH,
        300);
    if (dc) {
        ReleaseDC(hwnd, dc);
    }
    g_rcHeaderBadge = headerLayout.badgeRect;
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
    SetControlFontById(hwnd, IDC_MONITOR_HINT, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFontById(hwnd, IDC_AUDIO_METER_TEXT, g_fontSmall);
    SetControlFontById(hwnd, IDC_MIC_METER_TEXT, g_fontSmall);
    SetControlFontById(hwnd, IDC_AUDIO_CLIP_EVENTS, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFontById(hwnd, IDC_MIC_CLIP_EVENTS, g_fontHint ? g_fontHint : g_fontSmall);
    SetControlFontById(hwnd, IDC_VST_AUTOSTART, g_fontControlLarge ? g_fontControlLarge : g_fontBody);
    SetControlFontById(hwnd, IDC_STATUS, g_fontMono ? g_fontMono : g_fontBody);
}

void UpdateMonitorButton(HWND hwnd) {
    const bool enabled = MicMixApp::Instance().IsMonitorEnabled();
    SetMonitorButtonText(hwnd, IDC_MONITOR, enabled);
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
        SetTransientStatusText(hwnd, successText ? successText : L"Done.");
        ReloadAllLists(hwnd);
        return;
    }
    SetTransientStatusText(hwnd, Utf8ToWide(error.empty() ? "Operation failed" : error));
}

void SetStaticTextIfChanged(HWND ctl, std::wstring& cache, const std::wstring& nextText) {
    if (!ctl || cache == nextText) {
        return;
    }
    cache = nextText;
    SetWindowTextW(ctl, cache.c_str());
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
        SetTransientStatusText(hwnd, L"Select an effect first.");
        return;
    }
    std::string error;
    const bool ok = MicMixApp::Instance().RemoveEffect(chain, static_cast<size_t>(idx), error);
    ApplyActionResult(hwnd, ok, error, L"Effect removed.");
}

void HandleMoveEffect(HWND hwnd, EffectChain chain, int listId, int direction) {
    const int idx = GetSelectedIndex(hwnd, listId);
    if (idx < 0) {
        SetTransientStatusText(hwnd, L"Select an effect first.");
        return;
    }
    const std::vector<VstEffectSlot> slots = MicMixApp::Instance().GetEffects(chain);
    if (static_cast<size_t>(idx) >= slots.size()) {
        SetTransientStatusText(hwnd, L"Invalid selection.");
        return;
    }
    const int target = idx + direction;
    if (target < 0 || target >= static_cast<int>(slots.size())) {
        SetTransientStatusText(hwnd, direction < 0 ? L"Already at top." : L"Already at bottom.");
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
        SetTransientStatusText(hwnd, L"Select an effect first.");
        return;
    }
    const std::vector<VstEffectSlot> slots = MicMixApp::Instance().GetEffects(chain);
    if (static_cast<size_t>(idx) >= slots.size()) {
        SetTransientStatusText(hwnd, L"Invalid selection.");
        return;
    }
    const bool nextBypass = !slots[static_cast<size_t>(idx)].bypass;
    std::string error;
    const bool ok = MicMixApp::Instance().SetEffectBypass(chain, static_cast<size_t>(idx), nextBypass, error);
    ApplyActionResult(hwnd, ok, error, nextBypass ? L"Effect bypassed." : L"Effect active.");
}

void HandleToggleEnabledEffect(HWND hwnd, EffectChain chain, int listId) {
    const int idx = GetSelectedIndex(hwnd, listId);
    if (idx < 0) {
        SetTransientStatusText(hwnd, L"Select an effect first.");
        return;
    }
    const std::vector<VstEffectSlot> slots = MicMixApp::Instance().GetEffects(chain);
    if (static_cast<size_t>(idx) >= slots.size()) {
        SetTransientStatusText(hwnd, L"Invalid selection.");
        return;
    }
    const bool nextEnabled = !slots[static_cast<size_t>(idx)].enabled;
    std::string error;
    const bool ok = MicMixApp::Instance().SetEffectEnabled(chain, static_cast<size_t>(idx), nextEnabled, error);
    ApplyActionResult(hwnd, ok, error, nextEnabled ? L"Effect enabled." : L"Effect disabled.");
}

void HandleOpenEditor(HWND hwnd, EffectChain chain, int listId) {
    const int idx = GetSelectedIndex(hwnd, listId);
    if (idx < 0) {
        SetTransientStatusText(hwnd, L"Select an effect first.");
        return;
    }
    if (g_editorOpenPending.exchange(true, std::memory_order_acq_rel)) {
        SetTransientStatusText(hwnd, L"Editor request already running...");
        return;
    }
    SetTransientStatusText(hwnd, L"Opening editor...");
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
    const MicMixSettings settings = MicMixApp::Instance().GetSettings();
    const bool effectsEnabled = settings.vstEffectsEnabled;
    const bool hostAutostart = settings.vstHostAutostart;
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
                         L"  |  Monitor=" + std::wstring(MicMixApp::Instance().IsMonitorEnabled() ? L"On" : L"Off") +
                         L"  |  Host autostart=" + std::wstring(hostAutostart ? L"On" : L"Off");
    std::wstring line3 = Utf8ToWide(host.message.empty() ? "host_status=idle" : ("host_status=" + host.message));
    const uint64_t holdUntil = g_statusHoldUntilMs.load(std::memory_order_acquire);
    const uint64_t nowMs = GetTickCount64();
    if (holdUntil == 0 || nowMs >= holdUntil) {
        SetStatusText(hwnd, line1 + L"\r\n" + line2 + L"\r\n" + line3);
    }
    const HeaderBadgeState nextBadge = effectsEnabled ? HeaderBadgeState::Active : HeaderBadgeState::Off;
    if (!g_headerBadgeStateValid || nextBadge != g_headerBadgeState) {
        g_headerBadgeState = nextBadge;
        g_headerBadgeStateValid = true;
        InvalidateRect(hwnd, &g_rcHeaderBadge, FALSE);
    }
    SetWindowTextW(hwnd, effectsEnabled ? L"MicMix Effects - ACTIVE" : L"MicMix Effects - OFF");
    if (g_ownerVstAutostartChecked != hostAutostart) {
        g_ownerVstAutostartChecked = hostAutostart;
        InvalidateRect(GetDlgItem(hwnd, IDC_VST_AUTOSTART), nullptr, TRUE);
    }
    UpdateMonitorButton(hwnd);
}

void ApplyControlTheme(HWND hwnd) {
    const int unthemeOwnerDrawIds[] = {
        IDC_VST_AUTOSTART,
    };
    ApplyMicMixControlTheme(
        hwnd,
        kUiThemeButtons | kUiThemeListBoxes,
        true,
        true,
        unthemeOwnerDrawIds,
        static_cast<int>(std::size(unthemeOwnerDrawIds)));
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
        std::vector<wchar_t> buffer(static_cast<size_t>(textLen) + 1U, L'\0');
        const LRESULT copied = SendMessageW(dis->hwndItem, LB_GETTEXT, dis->itemID, reinterpret_cast<LPARAM>(buffer.data()));
        if (copied == LB_ERR) {
            itemText.clear();
        } else {
            itemText.assign(buffer.data());
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

void DrawOwnerCheckbox(const DRAWITEMSTRUCT* dis) {
    if (!dis) {
        return;
    }
    DrawOwnerCheckboxShared(
        dis,
        g_ownerVstAutostartChecked,
        g_fontControlLarge,
        g_fontBody,
        g_theme.text,
        g_theme.card,
        g_dpi);
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

void DrawClipStrip(HDC hdc, const RECT& rcStrip, float dangerUnit, bool clipRecent) {
    const MeterFonts fonts{
        g_fontTiny,
        g_fontHint,
        g_fontSmall,
        g_fontBody,
    };
    DrawClipStripShared(hdc, rcStrip, dangerUnit, clipRecent, fonts, g_dpi);
}

void UpdateMeterVisuals(HWND hwnd) {
    const TelemetrySnapshot t = MicMixApp::Instance().GetTelemetry();
    g_lastTelemetry = t;

    const float shownMusicPeakDb = (t.musicSendPeakDbfs > -119.0f) ? t.musicSendPeakDbfs : t.musicPeakDbfs;
    const bool musicMeterActive = IsSignalMeterActive(
        t.musicActive,
        t.musicSendPeakDbfs,
        t.musicPeakDbfs,
        t.musicRmsDbfs);
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
    const int musicLevelPx = std::clamp(::MeterDbToPixels(g_musicMeterVisualDb, musicMeterWidth), 0, musicMeterWidth);
    const int musicHoldPx = std::clamp(::MeterDbToPixels(g_musicMeterHoldDb, musicMeterWidth), 0, musicMeterWidth - 1);
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

    const int musicClipLitSegments = ::ClipLitSegmentsFromDb(shownMusicPeakDb);
    const bool musicClipChanged =
        (musicClipLitSegments != g_lastMusicClipLitSegments) ||
        (t.sourceClipRecent != g_lastMusicClipRecent);
    g_lastMusicClipLitSegments = musicClipLitSegments;
    g_lastMusicClipRecent = t.sourceClipRecent;
    if (musicClipChanged) {
        InvalidateRect(hwnd, &g_rcMusicClip, FALSE);
    }
    if (!musicMeterActive) {
        SetStaticTextIfChanged(GetDlgItem(hwnd, IDC_AUDIO_METER_TEXT), g_lastMusicMeterText, L"No signal");
    } else {
        const wchar_t* grade = L"OK";
        if (shownMusicPeakDb > -8.0f) {
            grade = L"Hot";
        } else if (shownMusicPeakDb > -16.0f) {
            grade = L"Warm";
        } else if (shownMusicPeakDb < -35.0f) {
            grade = L"Quiet";
        }
        wchar_t meterText[96];
        swprintf_s(meterText, L"%.1f dBFS  |  %s", shownMusicPeakDb, grade);
        SetStaticTextIfChanged(GetDlgItem(hwnd, IDC_AUDIO_METER_TEXT), g_lastMusicMeterText, meterText);
    }
    {
        wchar_t clipText[96];
        swprintf_s(clipText, L"CLIP Events: %llu", static_cast<unsigned long long>(t.sourceClipEvents));
        SetStaticTextIfChanged(GetDlgItem(hwnd, IDC_AUDIO_CLIP_EVENTS), g_lastMusicClipEventsText, clipText);
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
    const int micLevelPx = std::clamp(::MeterDbToPixels(g_micMeterVisualDb, micMeterWidth), 0, micMeterWidth);
    const int micHoldPx = std::clamp(::MeterDbToPixels(g_micMeterHoldDb, micMeterWidth), 0, micMeterWidth - 1);
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

    const int micClipLitSegments = ::ClipLitSegmentsFromDb(t.micPeakDbfs);
    const bool micClipChanged =
        (micClipLitSegments != g_lastMicClipLitSegments) ||
        (t.micClipRecent != g_lastMicClipRecent);
    g_lastMicClipLitSegments = micClipLitSegments;
    g_lastMicClipRecent = t.micClipRecent;
    if (micClipChanged) {
        InvalidateRect(hwnd, &g_rcMicClip, FALSE);
    }
    if (!micMeterActive) {
        SetStaticTextIfChanged(GetDlgItem(hwnd, IDC_MIC_METER_TEXT), g_lastMicMeterText, L"No signal");
    } else {
        const wchar_t* grade = L"OK";
        if (t.micPeakDbfs > -8.0f) {
            grade = L"Hot";
        } else if (t.micPeakDbfs > -16.0f) {
            grade = L"Warm";
        } else if (t.micPeakDbfs < -35.0f) {
            grade = L"Quiet";
        }
        wchar_t meterText[96];
        swprintf_s(meterText, L"%.1f dBFS  |  %s", t.micRmsDbfs, grade);
        SetStaticTextIfChanged(GetDlgItem(hwnd, IDC_MIC_METER_TEXT), g_lastMicMeterText, meterText);
    }
    {
        wchar_t clipText[96];
        swprintf_s(clipText, L"CLIP Events: %llu", static_cast<unsigned long long>(t.micClipEvents));
        SetStaticTextIfChanged(GetDlgItem(hwnd, IDC_MIC_CLIP_EVENTS), g_lastMicClipEventsText, clipText);
    }
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        g_dpi = GetDpiForWindow(hwnd);
        EnsureUiResources();
        ComputeLayout(hwnd);
        g_headerBadgeState = HeaderBadgeState::Off;
        g_headerBadgeStateValid = false;
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
        g_lastMusicMeterText.clear();
        g_lastMicMeterText.clear();
        g_lastMusicClipEventsText.clear();
        g_lastMicClipEventsText.clear();
        g_lastListRefreshTickMs = 0;
        g_ownerVstAutostartChecked = MicMixApp::Instance().GetSettings().vstHostAutostart;

        const int contentLeft = S(kCardMarginPx + kCardInnerPaddingPx);
        const int contentRight = S(kClientWidthPx - kCardMarginPx - kCardInnerPaddingPx);
        const int controlGap = S(12);
        const int actionButtonW = S(142);
        const int monitorButtonW = S(122);
        const int autostartCheckboxW = S(kUiTopCardCheckboxWidthPx);
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
            kEffectsHeaderTitle,
            contentLeft,
            contentRight,
            titleY,
            titleH,
            300);
        if (headerMeasureDc) {
            ReleaseDC(hwnd, headerMeasureDc);
        }
        g_rcHeaderBadge = headerLayout.badgeRect;
        const int titleW = headerLayout.titleWidthPx;
        const int headerMetaX = headerLayout.metaX;
        const int headerMetaActualW = headerLayout.metaWidth;

        const std::wstring versionText = Utf8ToWide(std::string("v") + MICMIX_VERSION + "  by dabinuss");

        CreateWindowW(L"STATIC", kEffectsHeaderTitle, WS_CHILD | WS_VISIBLE, titleX, titleY, titleW, titleH, hwnd, reinterpret_cast<HMENU>(IDC_TITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Configure MicMix Effects and route VST chains for music and mic", WS_CHILD | WS_VISIBLE, contentLeft, S(43), headerMetaX - contentLeft - S(8), S(24), hwnd, reinterpret_cast<HMENU>(IDC_SUBTITLE), nullptr, nullptr);
        CreateWindowW(L"STATIC", versionText.c_str(), WS_CHILD | WS_VISIBLE | SS_RIGHT, headerMetaX, S(18), headerMetaActualW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_VERSION), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"github.com/dabinuss/MicMix", WS_CHILD | WS_VISIBLE | SS_RIGHT | SS_NOTIFY, headerMetaX, S(36), headerMetaActualW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_REPO_LINK), nullptr, nullptr);

        const int topLeft = g_rcTop.left + S(kCardInnerPaddingPx);
        const int topY = g_rcTop.top + S(kUiTopCardPrimaryRowOffsetYpx);
        CreateWindowW(L"BUTTON", L"Enable Effects", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, topLeft, topY, actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_ENABLE_EFFECTS), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Disable Effects", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, topLeft + actionButtonW + controlGap, topY, actionButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_DISABLE_EFFECTS), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Monitor Mix: Off", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, topLeft + (actionButtonW + controlGap) * 2, topY, monitorButtonW, S(34), hwnd, reinterpret_cast<HMENU>(IDC_MONITOR), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Shared with MicMix monitor path", WS_CHILD | WS_VISIBLE, topLeft + (actionButtonW + controlGap) * 2, topY + S(kUiTopCardHintOffsetYpx), S(220), S(18), hwnd, reinterpret_cast<HMENU>(IDC_MONITOR_HINT), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"VST Host autostart", WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP, topLeft, g_rcTop.top + S(kUiTopCardCheckboxOffsetYpx), autostartCheckboxW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_VST_AUTOSTART), nullptr, nullptr);

        const int sectionLeft = g_rcMusic.left + S(kCardInnerPaddingPx);
        const int sectionTop = g_rcMusic.top + S(kCardInnerPaddingPx);
        const int musicListTop = sectionTop + S(kUiSectionListTopOffsetYpx);
        const int musicStatusX = g_rcMusicMeter.right + S(4);
        const int musicStatusW = std::max(0, static_cast<int>((g_rcMusic.right - S(kCardInnerPaddingPx)) - musicStatusX));
        CreateWindowW(L"STATIC", L"Audio Meter", WS_CHILD | WS_VISIBLE, sectionLeft, sectionTop + S(kUiSectionMeterLabelOffsetYpx), S(130), S(24), hwnd, reinterpret_cast<HMENU>(IDC_AUDIO_METER_LABEL), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"No signal", WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS, musicStatusX, g_rcMusicMeter.top + S(2), musicStatusW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_AUDIO_METER_TEXT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"CLIP Events: 0", WS_CHILD | WS_VISIBLE, musicStatusX, g_rcMusicClip.top, musicStatusW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_AUDIO_CLIP_EVENTS), nullptr, nullptr);
        CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
            sectionLeft, musicListTop, S(458), S(90), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_LIST), nullptr, nullptr);
        const int musicBtnX = sectionLeft + S(468);
        CreateWindowW(L"BUTTON", L"Add", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX, musicListTop, S(86), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_ADD), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Remove", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX, musicListTop + S(30), S(86), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_REMOVE), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Up", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX, musicListTop + S(60), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_UP), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Down", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX + S(46), musicListTop + S(60), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_DOWN), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"On", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX, musicListTop + S(90), S(26), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_ENABLE), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Byp", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX + S(30), musicListTop + S(90), S(26), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_BYPASS), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Ed", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, musicBtnX + S(60), musicListTop + S(90), S(26), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_EDITOR), nullptr, nullptr);

        const int micLeft = g_rcMic.left + S(kCardInnerPaddingPx);
        const int micTop = g_rcMic.top + S(kCardInnerPaddingPx);
        const int micListTop = micTop + S(kUiSectionListTopOffsetYpx);
        const int micStatusX = g_rcMicMeter.right + S(4);
        const int micStatusW = std::max(0, static_cast<int>((g_rcMic.right - S(kCardInnerPaddingPx)) - micStatusX));
        CreateWindowW(L"STATIC", L"Mic Meter", WS_CHILD | WS_VISIBLE, micLeft, micTop + S(kUiSectionMeterLabelOffsetYpx), S(130), S(24), hwnd, reinterpret_cast<HMENU>(IDC_MIC_METER_LABEL), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"No signal", WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS, micStatusX, g_rcMicMeter.top + S(2), micStatusW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_MIC_METER_TEXT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"CLIP Events: 0", WS_CHILD | WS_VISIBLE, micStatusX, g_rcMicClip.top, micStatusW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_MIC_CLIP_EVENTS), nullptr, nullptr);
        CreateWindowW(L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
            micLeft, micListTop, S(458), S(90), hwnd, reinterpret_cast<HMENU>(IDC_MIC_LIST), nullptr, nullptr);
        const int micBtnX = micLeft + S(468);
        CreateWindowW(L"BUTTON", L"Add", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX, micListTop, S(86), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_ADD), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Remove", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX, micListTop + S(30), S(86), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_REMOVE), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Up", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX, micListTop + S(60), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_UP), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Down", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX + S(46), micListTop + S(60), S(40), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_DOWN), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"On", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX, micListTop + S(90), S(26), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_ENABLE), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Byp", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX + S(30), micListTop + S(90), S(26), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_BYPASS), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Ed", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, micBtnX + S(60), micListTop + S(90), S(26), S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_EDITOR), nullptr, nullptr);

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
        if (dis->CtlID == IDC_VST_AUTOSTART) {
            DrawOwnerCheckbox(dis);
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
        case IDC_VST_AUTOSTART:
            if (HIWORD(wParam) == BN_CLICKED) {
                MicMixSettings settings = MicMixApp::Instance().GetSettings();
                g_ownerVstAutostartChecked = !g_ownerVstAutostartChecked;
                InvalidateRect(GetDlgItem(hwnd, IDC_VST_AUTOSTART), nullptr, TRUE);
                const bool autostartEnabled = g_ownerVstAutostartChecked;
                if (settings.vstHostAutostart != autostartEnabled) {
                    settings.vstHostAutostart = autostartEnabled;
                    MicMixApp::Instance().ApplySettings(settings, false, true);
                }
                RefreshStatus(hwnd);
            }
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
        case IDC_MUSIC_ENABLE:
            HandleToggleEnabledEffect(hwnd, EffectChain::Music, IDC_MUSIC_LIST);
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
        case IDC_MIC_ENABLE:
            HandleToggleEnabledEffect(hwnd, EffectChain::Mic, IDC_MIC_LIST);
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
            SetTransientStatusText(hwnd, result->message);
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
                   id == IDC_AUDIO_CLIP_EVENTS || id == IDC_MIC_CLIP_EVENTS) {
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
        FillRect(paintDc, &bgFillRc, g_brushBg ? g_brushBg : reinterpret_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
        RECT accent{ S(0), S(0), S(kClientWidthPx), S(4) };
        HBRUSH accentBrush = CreateSolidBrush(g_theme.accent);
        FillRect(paintDc, &accent, accentBrush);
        DeleteObject(accentBrush);

        DrawFlatCardFrame(paintDc, g_rcTop, g_brushCard, g_theme.border);
        DrawFlatCardFrame(paintDc, g_rcMusic, g_brushCard, g_theme.border);
        DrawFlatCardFrame(paintDc, g_rcMic, g_brushCard, g_theme.border);
        DrawFlatCardFrame(paintDc, g_rcStatus, g_brushCard, g_theme.border);
        DrawLevelMeter(paintDc, g_rcMusicMeter, g_lastMusicMeterActive, g_musicMeterVisualDb, g_musicMeterHoldDb);
        DrawClipStrip(paintDc, g_rcMusicClip, ::ClipDangerUnitFromDb((g_lastTelemetry.musicSendPeakDbfs > -119.0f) ? g_lastTelemetry.musicSendPeakDbfs : g_lastTelemetry.musicPeakDbfs), g_lastTelemetry.sourceClipRecent);
        DrawLevelMeter(paintDc, g_rcMicMeter, g_lastMicMeterActive, g_micMeterVisualDb, g_micMeterHoldDb);
        DrawClipStrip(paintDc, g_rcMicClip, ::ClipDangerUnitFromDb(g_lastTelemetry.micPeakDbfs), g_lastTelemetry.micClipRecent);
        DrawHeaderBadge(
            paintDc,
            g_rcHeaderBadge,
            GetHeaderBadgeVisual(g_headerBadgeState),
            g_fontSmall,
            g_fontBody,
            g_dpi);
        const SectionHeading sectionHeadings[] = {
            { g_rcTop, kEffectsSectionTopTitle },
            { g_rcMusic, kEffectsSectionMusicTitle },
            { g_rcMic, kEffectsSectionMicTitle },
        };
        DrawSectionHeadings(
            paintDc,
            sectionHeadings,
            static_cast<int>(std::size(sectionHeadings)),
            g_fontSmall,
            g_fontBody,
            g_theme.muted,
            g_dpi);

        if (bufferedPaint && memDc && memBmp) {
            BitBlt(hdc, 0, 0, clientW, clientH, memDc, 0, 0, SRCCOPY);
        }
        if (memDc && oldBmp) {
            SelectObject(memDc, oldBmp);
        }
        if (memBmp) {
            DeleteObject(memBmp);
        }
        if (memDc) {
            DeleteDC(memDc);
        }

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
    HINSTANCE hInst = ResolveModuleHandleFromAddress(reinterpret_cast<const void*>(&WindowThreadMain));
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
