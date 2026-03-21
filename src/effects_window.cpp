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
#include <cctype>
#include <cmath>
#include <cstdio>
#include <exception>
#include <filesystem>
#include <memory>
#include <mutex>
#include <string>
#include <thread>
#include <unordered_map>
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
bool g_windowTitleStateValid = false;
bool g_windowTitleEffectsEnabled = false;
uint64_t g_lastStatusSampleMs = 0;
VstHostStatus g_cachedHostStatus{};
size_t g_cachedMusicEffectsCount = 0;
size_t g_cachedMicEffectsCount = 0;
std::unordered_map<std::string, uint32_t> g_effectDisplayIdByUid;
uint32_t g_nextEffectDisplayId = 1;

enum class InlineAction {
    None = 0,
    ToggleEnabled,
    ToggleBypass,
    Remove,
    Drag,
};

struct EffectRowView {
    std::wstring displayName;
    std::wstring statusText;
    std::wstring idText;
    bool enabled = true;
    bool bypass = false;
};

bool operator==(const EffectRowView& a, const EffectRowView& b) {
    return a.displayName == b.displayName &&
           a.statusText == b.statusText &&
           a.idText == b.idText &&
           a.enabled == b.enabled &&
           a.bypass == b.bypass;
}

struct EffectListUiState {
    std::vector<EffectRowView> rows;
    int hoverItem = -1;
    InlineAction hoverAction = InlineAction::None;
    int pressedItem = -1;
    InlineAction pressedAction = InlineAction::None;
    bool trackingMouse = false;
    bool dragArmed = false;
    bool dragging = false;
    int dragSource = -1;
    int dragInsertIndex = -1;
    POINT dragStart{ 0, 0 };
    uint64_t lastInteractionMs = 0;
};

EffectListUiState g_musicListUi{};
EffectListUiState g_micListUi{};

constexpr int kClientWidthPx = kUiClientWidthPx;
constexpr int kClientHeightPx = kUiClientHeightPx;
constexpr int kCardMarginPx = kUiCardMarginPx;
constexpr int kCardGapPx = kUiCardGapPx;
constexpr int kCardInnerPaddingPx = kUiCardInnerPaddingPx;
constexpr UINT kTimerStatus = 1;
constexpr UINT kTimerMs = 40;
constexpr int kListItemHeightPx = 34;
constexpr int kEffectsTopCardHeightPx = 128;
constexpr int kEffectsMetersCardHeightPx = 146;
constexpr int kEffectsListsCardHeightPx = 282;
constexpr int kEffectsStatusCardHeightPx = 84;
constexpr int kEffectsMeterSectionTopOffsetPx = 10;
constexpr int kEffectsMeterLeftOffsetPx = 144;
constexpr int kEffectsMusicMeterTopOffsetPx = 18;
constexpr int kEffectsMusicClipTopOffsetPx = 48;
constexpr int kEffectsMicMeterTopOffsetPx = 66;
constexpr int kEffectsMicClipTopOffsetPx = 96;
constexpr int kEffectsAudioMeterLabelTopPx = 20;
constexpr int kEffectsMicMeterLabelTopPx = 68;
constexpr int kEffectsListHeightPx = 184;
constexpr int kEffectsListColumnsGapPx = 12;
constexpr int kEffectsListSectionTopOffsetPx = 10;
constexpr int kEffectsListHeaderOffsetPx = 12;
constexpr int kEffectsListTopOffsetPx = 30;
constexpr int kEffectsListButtonTopGapPx = 8;
constexpr int kEffectsListButtonGapPx = 6;
constexpr int kEffectsListButtonAddWidthPx = 122;
constexpr int kEffectsListButtonSmallWidthPx = 82;
constexpr uint64_t kEffectsStatusHeavySampleMs = 250;
constexpr uint64_t kEffectsListInteractionHoldMs = 900;
constexpr wchar_t kRepoUrl[] = L"https://github.com/dabinuss/MicMix";
constexpr wchar_t kEffectsHeaderTitle[] = L"MicMix - Audio Effects";
constexpr wchar_t kEffectsSectionTopTitle[] = L"EFFECTS SETTINGS";
constexpr wchar_t kEffectsSectionMusicTitle[] = L"METERS";
constexpr wchar_t kEffectsSectionMicTitle[] = L"EFFECTS";
constexpr UINT_PTR kMusicListSubclassId = 1001;
constexpr UINT_PTR kMicListSubclassId = 1002;

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
    case IDC_MUSIC_ADD:
    case IDC_MUSIC_UP:
    case IDC_MUSIC_DOWN:
    case IDC_MIC_ADD:
    case IDC_MIC_UP:
    case IDC_MIC_DOWN:
        return true;
    default:
        return false;
    }
}

EffectChain ChainFromListId(int listId) {
    return (listId == IDC_MIC_LIST) ? EffectChain::Mic : EffectChain::Music;
}

EffectListUiState& ListUiStateFromListId(int listId) {
    return (listId == IDC_MIC_LIST) ? g_micListUi : g_musicListUi;
}

void ResetListPointerState(EffectListUiState& state) {
    state.hoverItem = -1;
    state.hoverAction = InlineAction::None;
    state.pressedItem = -1;
    state.pressedAction = InlineAction::None;
    state.trackingMouse = false;
    state.dragArmed = false;
    state.dragging = false;
    state.dragSource = -1;
    state.dragInsertIndex = -1;
    state.dragStart = { 0, 0 };
    state.lastInteractionMs = 0;
}

void MarkListInteraction(EffectListUiState& state) {
    state.lastInteractionMs = GetTickCount64();
}

std::wstring BuildIdText(const VstEffectSlot& slot, size_t index) {
    (void)index;
    if (slot.uid.empty()) {
        return L"#000";
    }
    auto it = g_effectDisplayIdByUid.find(slot.uid);
    if (it == g_effectDisplayIdByUid.end()) {
        const uint32_t next = std::max<uint32_t>(1U, g_nextEffectDisplayId++);
        it = g_effectDisplayIdByUid.emplace(slot.uid, next).first;
    }
    wchar_t idBuf[16]{};
    swprintf_s(idBuf, L"#%03u", static_cast<unsigned int>(it->second));
    return idBuf;
}

std::wstring BuildDisplayName(const VstEffectSlot& slot) {
    if (!slot.name.empty()) {
        return Utf8ToWide(slot.name);
    }
    const std::wstring path = Utf8ToWide(slot.path);
    std::filesystem::path p(path);
    if (!p.stem().wstring().empty()) {
        return p.stem().wstring();
    }
    return path.empty() ? L"(unnamed effect)" : path;
}

std::wstring BuildStatusText(const VstEffectSlot& slot) {
    std::wstring status = slot.enabled ? L"ON" : L"OFF";
    status += L"  |  ";
    status += slot.bypass ? L"BYPASS" : L"ACTIVE";
    if (!slot.lastStatus.empty()) {
        status += L"  |  ";
        status += Utf8ToWide(slot.lastStatus);
    }
    return status;
}

bool AnyListDragging() {
    return g_musicListUi.dragging || g_micListUi.dragging;
}

bool IsListInteracting(HWND hwnd) {
    const uint64_t nowMs = GetTickCount64();
    HWND focus = GetFocus();
    if (focus && (focus == GetDlgItem(hwnd, IDC_MUSIC_LIST) || focus == GetDlgItem(hwnd, IDC_MIC_LIST))) {
        return true;
    }
    if ((nowMs <= (g_musicListUi.lastInteractionMs + kEffectsListInteractionHoldMs)) ||
        (nowMs <= (g_micListUi.lastInteractionMs + kEffectsListInteractionHoldMs))) {
        return true;
    }
    return AnyListDragging();
}

void InvalidateListItem(HWND list, int itemIndex) {
    if (!list || itemIndex < 0) {
        return;
    }
    RECT itemRect{};
    if (SendMessageW(list, LB_GETITEMRECT, static_cast<WPARAM>(itemIndex), reinterpret_cast<LPARAM>(&itemRect)) == LB_ERR) {
        return;
    }
    InvalidateRect(list, &itemRect, FALSE);
}

void SetStatusText(HWND hwnd, const std::wstring& text, uint64_t holdMs = 0) {
    const bool changed = (g_statusText != text);
    g_statusText = text;
    if (holdMs > 0) {
        g_statusHoldUntilMs.store(GetTickCount64() + holdMs, std::memory_order_release);
    } else {
        g_statusHoldUntilMs.store(0, std::memory_order_release);
    }
    HWND ctl = GetDlgItem(hwnd, IDC_STATUS);
    if (ctl && changed) {
        SetWindowTextW(ctl, g_statusText.c_str());
    }
}

void SetTransientStatusText(HWND hwnd, const std::wstring& text) {
    constexpr uint64_t kStatusHoldMs = 2400ULL;
    SetStatusText(hwnd, text, kStatusHoldMs);
}

void UpdateWindowTitleState(HWND hwnd, bool effectsEnabled) {
    if (!hwnd) {
        return;
    }
    if (g_windowTitleStateValid && g_windowTitleEffectsEnabled == effectsEnabled) {
        return;
    }
    g_windowTitleStateValid = true;
    g_windowTitleEffectsEnabled = effectsEnabled;
    SetWindowTextW(hwnd, effectsEnabled ? L"MicMix Effects - ACTIVE" : L"MicMix Effects - OFF");
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
    const int topH = S(kEffectsTopCardHeightPx);
    const int musicY = topY + topH + S(kCardGapPx);
    const int musicH = S(kEffectsMetersCardHeightPx);
    const int micY = musicY + musicH + S(kCardGapPx);
    const int micH = S(kEffectsListsCardHeightPx);
    const int statusY = micY + micH + S(kCardGapPx);
    const int statusH = S(kEffectsStatusCardHeightPx);
    g_rcTop = { left, topY, left + cardW, topY + topH };
    g_rcMusic = { left, musicY, left + cardW, musicY + musicH };
    g_rcMic = { left, micY, left + cardW, micY + micH };
    g_rcStatus = { left, statusY, left + cardW, statusY + statusH };

    const int meterWidth = S(340);
    const int sectionPad = S(kCardInnerPaddingPx);
    const int meterSectionTop = g_rcMusic.top + S(kEffectsMeterSectionTopOffsetPx);
    const int meterSectionLeft = g_rcMusic.left + sectionPad;
    const int meterLeft = meterSectionLeft + S(kEffectsMeterLeftOffsetPx);
    g_rcMusicMeter = { meterLeft, meterSectionTop + S(kEffectsMusicMeterTopOffsetPx), meterLeft + meterWidth, meterSectionTop + S(kEffectsMusicMeterTopOffsetPx) + S(28) };
    g_rcMusicClip = { meterLeft, meterSectionTop + S(kEffectsMusicClipTopOffsetPx), meterLeft + meterWidth, meterSectionTop + S(kEffectsMusicClipTopOffsetPx) + S(12) };
    g_rcMicMeter = { meterLeft, meterSectionTop + S(kEffectsMicMeterTopOffsetPx), meterLeft + meterWidth, meterSectionTop + S(kEffectsMicMeterTopOffsetPx) + S(28) };
    g_rcMicClip = { meterLeft, meterSectionTop + S(kEffectsMicClipTopOffsetPx), meterLeft + meterWidth, meterSectionTop + S(kEffectsMicClipTopOffsetPx) + S(12) };

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

void ReloadEffectList(HWND hwnd, EffectChain chain, int listId) {
    HWND list = GetDlgItem(hwnd, listId);
    if (!list) {
        return;
    }
    EffectListUiState& ui = ListUiStateFromListId(listId);
    const int previousSel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
    const int previousTop = static_cast<int>(SendMessageW(list, LB_GETTOPINDEX, 0, 0));
    const std::vector<VstEffectSlot> slots = MicMixApp::Instance().GetEffects(chain);

    std::vector<EffectRowView> nextRows;
    nextRows.reserve(slots.size());
    for (size_t i = 0; i < slots.size(); ++i) {
        const auto& slot = slots[i];
        EffectRowView row{};
        row.displayName = BuildDisplayName(slot);
        row.statusText = BuildStatusText(slot);
        row.idText = BuildIdText(slot, i);
        row.enabled = slot.enabled;
        row.bypass = slot.bypass;
        nextRows.push_back(std::move(row));
    }

    const int currentCount = static_cast<int>(SendMessageW(list, LB_GETCOUNT, 0, 0));
    const int nextCount = static_cast<int>(nextRows.size());
    const bool countChanged = currentCount != nextCount;
    const bool rowsChanged = ui.rows != nextRows;

    if (countChanged) {
        SendMessageW(list, LB_RESETCONTENT, 0, 0);
        for (int i = 0; i < nextCount; ++i) {
            SendMessageW(list, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(L" "));
        }
    }

    ui.rows = std::move(nextRows);

    if (nextCount > 0) {
        int sel = previousSel;
        if (sel < 0 || sel >= nextCount) {
            sel = nextCount - 1;
        }
        const int currentSel = static_cast<int>(SendMessageW(list, LB_GETCURSEL, 0, 0));
        if (currentSel != sel) {
            SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(sel), 0);
        }
    } else {
        SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(-1), 0);
    }

    if (previousTop >= 0 && nextCount > 0) {
        const int top = std::clamp(previousTop, 0, std::max(0, nextCount - 1));
        SendMessageW(list, LB_SETTOPINDEX, static_cast<WPARAM>(top), 0);
    }

    if (ui.dragSource >= static_cast<int>(slots.size())) {
        ui.dragSource = -1;
    }
    if (ui.dragInsertIndex > static_cast<int>(slots.size())) {
        ui.dragInsertIndex = static_cast<int>(slots.size());
    }
    if (countChanged || rowsChanged) {
        InvalidateRect(list, nullptr, FALSE);
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
        g_lastStatusSampleMs = 0;
        return;
    }
    SetTransientStatusText(hwnd, Utf8ToWide(error.empty() ? "Operation failed" : error));
    g_lastStatusSampleMs = 0;
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

void HandleMoveEffectToIndex(HWND hwnd, EffectChain chain, int listId, int fromIndex, int toIndex) {
    if (fromIndex < 0 || toIndex < 0) {
        return;
    }
    const std::vector<VstEffectSlot> slots = MicMixApp::Instance().GetEffects(chain);
    if (static_cast<size_t>(fromIndex) >= slots.size() || static_cast<size_t>(toIndex) >= slots.size()) {
        SetTransientStatusText(hwnd, L"Invalid move target.");
        return;
    }
    if (fromIndex == toIndex) {
        return;
    }
    std::string error;
    const bool ok = MicMixApp::Instance().MoveEffect(
        chain, static_cast<size_t>(fromIndex), static_cast<size_t>(toIndex), error);
    ApplyActionResult(hwnd, ok, error, L"Effect order updated.");
    if (ok) {
        SendMessageW(GetDlgItem(hwnd, listId), LB_SETCURSEL, static_cast<WPARAM>(toIndex), 0);
    }
}

int OtherListId(int listId) {
    return (listId == IDC_MIC_LIST) ? IDC_MUSIC_LIST : IDC_MIC_LIST;
}

bool ScreenPointToListClient(HWND parent, const POINT& screenPt, int& outListId, POINT& outClientPt) {
    const int candidateIds[] = { IDC_MUSIC_LIST, IDC_MIC_LIST };
    for (int id : candidateIds) {
        HWND list = GetDlgItem(parent, id);
        if (!list) {
            continue;
        }
        RECT rcClient{};
        GetClientRect(list, &rcClient);
        POINT clientPt = screenPt;
        if (!ScreenToClient(list, &clientPt)) {
            continue;
        }
        if (PtInRect(&rcClient, clientPt)) {
            outListId = id;
            outClientPt = clientPt;
            return true;
        }
    }
    return false;
}

void HandleMoveEffectBetweenLists(HWND hwnd, int fromListId, int fromIndex, int toListId, int toInsertIndex) {
    if (fromListId != IDC_MUSIC_LIST && fromListId != IDC_MIC_LIST) {
        return;
    }
    if (toListId != IDC_MUSIC_LIST && toListId != IDC_MIC_LIST) {
        return;
    }
    if (fromIndex < 0) {
        return;
    }
    const EffectChain fromChain = ChainFromListId(fromListId);
    const EffectChain toChain = ChainFromListId(toListId);

    HWND targetList = GetDlgItem(hwnd, toListId);
    const int targetCount = targetList ? static_cast<int>(SendMessageW(targetList, LB_GETCOUNT, 0, 0)) : 0;
    const int insertIndex = std::clamp(toInsertIndex, 0, std::max(0, targetCount));

    std::string error;
    const bool ok = MicMixApp::Instance().MoveEffectBetweenChains(
        fromChain,
        static_cast<size_t>(fromIndex),
        toChain,
        static_cast<size_t>(insertIndex),
        error);
    ApplyActionResult(hwnd, ok, error, L"Effect moved.");
    if (!ok) {
        return;
    }

    HWND reloadedTarget = GetDlgItem(hwnd, toListId);
    if (reloadedTarget) {
        const int reloadedCount = static_cast<int>(SendMessageW(reloadedTarget, LB_GETCOUNT, 0, 0));
        if (reloadedCount > 0) {
            const int sel = std::clamp(insertIndex, 0, reloadedCount - 1);
            SendMessageW(reloadedTarget, LB_SETCURSEL, static_cast<WPARAM>(sel), 0);
            MarkListInteraction(ListUiStateFromListId(toListId));
        }
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

void RefreshStatus(HWND hwnd, bool force) {
    const MicMixSettings settings = MicMixApp::Instance().GetSettings();
    const bool effectsEnabled = settings.vstEffectsEnabled;
    const bool hostAutostart = settings.vstHostAutostart;
    const uint64_t nowMs = GetTickCount64();
    if (force || nowMs >= (g_lastStatusSampleMs + kEffectsStatusHeavySampleMs)) {
        g_cachedHostStatus = MicMixApp::Instance().GetVstHostStatus();
        g_cachedMusicEffectsCount = MicMixApp::Instance().GetEffects(EffectChain::Music).size();
        g_cachedMicEffectsCount = MicMixApp::Instance().GetEffects(EffectChain::Mic).size();
        g_lastStatusSampleMs = nowMs;
    }
    const VstHostStatus& host = g_cachedHostStatus;
    std::wstring line1 = L"Effects: ";
    line1 += effectsEnabled ? L"On" : L"Off";
    line1 += L"  |  Host: ";
    line1 += host.running ? L"Running" : L"Stopped";
    if (host.pid != 0) {
        line1 += L" (pid=" + std::to_wstring(host.pid) + L")";
    }
    std::wstring line2 = L"Music plugins=" + std::to_wstring(g_cachedMusicEffectsCount) +
                         L"  |  Mic plugins=" + std::to_wstring(g_cachedMicEffectsCount) +
                         L"  |  Monitor=" + std::wstring(MicMixApp::Instance().IsMonitorEnabled() ? L"On" : L"Off") +
                         L"  |  Host autostart=" + std::wstring(hostAutostart ? L"On" : L"Off");
    std::wstring line3 = Utf8ToWide(host.message.empty() ? "host_status=idle" : ("host_status=" + host.message));
    const uint64_t holdUntil = g_statusHoldUntilMs.load(std::memory_order_acquire);
    if (holdUntil == 0 || nowMs >= holdUntil) {
        SetStatusText(hwnd, line1 + L"\r\n" + line2 + L"\r\n" + line3);
    }
    const HeaderBadgeState nextBadge = effectsEnabled ? HeaderBadgeState::Active : HeaderBadgeState::Off;
    if (!g_headerBadgeStateValid || nextBadge != g_headerBadgeState) {
        g_headerBadgeState = nextBadge;
        g_headerBadgeStateValid = true;
        InvalidateRect(hwnd, &g_rcHeaderBadge, FALSE);
    }
    UpdateWindowTitleState(hwnd, effectsEnabled);
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

struct EffectRowLayout {
    RECT handleRect{};
    RECT idRect{};
    RECT nameRect{};
    RECT statusRect{};
    RECT enableRect{};
    RECT bypassRect{};
    RECT removeRect{};
};

EffectRowLayout ComputeEffectRowLayout(const RECT& itemRect) {
    EffectRowLayout layout{};
    RECT inner = itemRect;
    InflateRect(&inner, -S(6), -S(4));
    const bool compact = (inner.right - inner.left) <= S(320);
    const int gap = S(compact ? 3 : 6);
    const int buttonH = S(compact ? 17 : 20);
    const int buttonY = static_cast<int>(inner.top) + std::max(0, static_cast<int>(((inner.bottom - inner.top) - buttonH) / 2));

    const int removeW = S(compact ? 22 : 56);
    const int bypassW = S(compact ? 36 : 58);
    const int enableW = S(compact ? 36 : 58);
    const int handleW = S(compact ? 12 : 14);
    const int idW = S(compact ? 30 : 78);
    int right = inner.right;
    layout.removeRect = { right - removeW, buttonY, right, buttonY + buttonH };
    right = layout.removeRect.left - gap;
    layout.bypassRect = { right - bypassW, buttonY, right, buttonY + buttonH };
    right = layout.bypassRect.left - gap;
    layout.enableRect = { right - enableW, buttonY, right, buttonY + buttonH };

    layout.handleRect = { inner.left, inner.top + S(4), inner.left + handleW, inner.bottom - S(4) };
    layout.idRect = { layout.handleRect.right + gap, inner.top + S(3), layout.handleRect.right + gap + idW, inner.bottom - S(3) };
    const int idToNameGap = S(compact ? 1 : 6);
    const int textLeft = layout.idRect.right + idToNameGap;
    const int textRight = std::max(textLeft, static_cast<int>(layout.enableRect.left - gap));
    const int midY = inner.top + ((inner.bottom - inner.top) / 2);
    layout.nameRect = { textLeft, inner.top, textRight, midY };
    layout.statusRect = { textLeft, midY - S(1), textRight, inner.bottom };
    return layout;
}

int HitTestListItem(HWND list, const POINT& clientPt) {
    const DWORD hit = static_cast<DWORD>(SendMessageW(list, LB_ITEMFROMPOINT, 0, MAKELPARAM(clientPt.x, clientPt.y)));
    const int item = static_cast<short>(LOWORD(hit));
    const bool outside = HIWORD(hit) != 0;
    if (outside || item < 0) {
        return -1;
    }
    return item;
}

InlineAction HitTestInlineAction(HWND list, int item, const POINT& clientPt) {
    if (item < 0) {
        return InlineAction::None;
    }
    RECT rowRect{};
    if (SendMessageW(list, LB_GETITEMRECT, static_cast<WPARAM>(item), reinterpret_cast<LPARAM>(&rowRect)) == LB_ERR) {
        return InlineAction::None;
    }
    const EffectRowLayout layout = ComputeEffectRowLayout(rowRect);
    if (PtInRect(&layout.enableRect, clientPt)) {
        return InlineAction::ToggleEnabled;
    }
    if (PtInRect(&layout.bypassRect, clientPt)) {
        return InlineAction::ToggleBypass;
    }
    if (PtInRect(&layout.removeRect, clientPt)) {
        return InlineAction::Remove;
    }
    if (PtInRect(&layout.handleRect, clientPt)) {
        return InlineAction::Drag;
    }
    return InlineAction::None;
}

int ComputeInsertIndexFromPoint(HWND list, const POINT& clientPt) {
    const int count = static_cast<int>(SendMessageW(list, LB_GETCOUNT, 0, 0));
    if (count <= 0) {
        return -1;
    }
    if (clientPt.y <= 0) {
        return 0;
    }
    RECT lastRect{};
    if (SendMessageW(list, LB_GETITEMRECT, static_cast<WPARAM>(count - 1), reinterpret_cast<LPARAM>(&lastRect)) != LB_ERR) {
        if (clientPt.y >= lastRect.bottom) {
            return count;
        }
    }
    const int hitItem = HitTestListItem(list, clientPt);
    if (hitItem < 0) {
        return count;
    }
    RECT hitRect{};
    if (SendMessageW(list, LB_GETITEMRECT, static_cast<WPARAM>(hitItem), reinterpret_cast<LPARAM>(&hitRect)) == LB_ERR) {
        return hitItem;
    }
    const int mid = hitRect.top + ((hitRect.bottom - hitRect.top) / 2);
    return (clientPt.y < mid) ? hitItem : (hitItem + 1);
}

void InvalidateInsertMarker(HWND list, int insertIndex) {
    if (!list || insertIndex < 0) {
        return;
    }
    const int count = static_cast<int>(SendMessageW(list, LB_GETCOUNT, 0, 0));
    if (count <= 0) {
        InvalidateRect(list, nullptr, FALSE);
        return;
    }
    if (insertIndex >= count) {
        InvalidateListItem(list, count - 1);
    } else {
        InvalidateListItem(list, insertIndex);
    }
}

void DrawInlineButton(
    HDC dc,
    const RECT& rc,
    const wchar_t* label,
    bool emphasized,
    bool hover,
    bool pressed,
    HFONT textFont) {
    COLORREF bg = emphasized ? RGB(226, 240, 255) : RGB(242, 245, 249);
    COLORREF border = emphasized ? RGB(129, 165, 214) : RGB(196, 206, 221);
    COLORREF text = emphasized ? RGB(34, 74, 132) : RGB(77, 87, 104);
    if (hover) {
        bg = emphasized ? RGB(213, 232, 255) : RGB(232, 238, 246);
    }
    if (pressed) {
        bg = emphasized ? RGB(193, 220, 252) : RGB(216, 227, 239);
        border = RGB(114, 139, 173);
    }

    HBRUSH fill = CreateSolidBrush(bg);
    HPEN pen = CreatePen(PS_SOLID, 1, border);
    HGDIOBJ oldBrush = SelectObject(dc, fill);
    HGDIOBJ oldPen = SelectObject(dc, pen);
    RoundRect(dc, rc.left, rc.top, rc.right, rc.bottom, S(8), S(8));
    SelectObject(dc, oldPen);
    SelectObject(dc, oldBrush);
    DeleteObject(pen);
    DeleteObject(fill);

    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, text);
    SelectObject(dc, textFont ? textFont : g_fontBody);
    RECT textRect = rc;
    DrawTextW(dc, label, -1, &textRect, DT_SINGLELINE | DT_CENTER | DT_VCENTER | DT_END_ELLIPSIS);
}

void InvokeInlineListAction(HWND parent, int listId, int item, InlineAction action) {
    if (!parent || item < 0) {
        return;
    }
    HWND list = GetDlgItem(parent, listId);
    if (list) {
        SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(item), 0);
    }
    const EffectChain chain = ChainFromListId(listId);
    switch (action) {
    case InlineAction::ToggleEnabled:
        HandleToggleEnabledEffect(parent, chain, listId);
        break;
    case InlineAction::ToggleBypass:
        HandleBypassEffect(parent, chain, listId);
        break;
    case InlineAction::Remove:
        HandleRemoveEffect(parent, chain, listId);
        break;
    default:
        break;
    }
}

LRESULT CALLBACK EffectListSubclassProc(
    HWND list,
    UINT msg,
    WPARAM wParam,
    LPARAM lParam,
    UINT_PTR /*subclassId*/,
    DWORD_PTR refData) {
    const int listId = static_cast<int>(refData);
    EffectListUiState& ui = ListUiStateFromListId(listId);
    HWND parent = GetParent(list);

    switch (msg) {
    case WM_SETFOCUS:
    case WM_MOUSEWHEEL:
    case WM_VSCROLL:
    case WM_KEYDOWN:
        MarkListInteraction(ui);
        break;
    case WM_MOUSELEAVE: {
        const int oldHoverItem = ui.hoverItem;
        ui.trackingMouse = false;
        ui.hoverItem = -1;
        ui.hoverAction = InlineAction::None;
        InvalidateListItem(list, oldHoverItem);
        return 0;
    }
    case WM_MOUSEMOVE: {
        MarkListInteraction(ui);
        POINT pt{
            static_cast<int>(static_cast<short>(LOWORD(lParam))),
            static_cast<int>(static_cast<short>(HIWORD(lParam)))
        };
        if (!ui.trackingMouse) {
            TRACKMOUSEEVENT tme{};
            tme.cbSize = sizeof(tme);
            tme.dwFlags = TME_LEAVE;
            tme.hwndTrack = list;
            TrackMouseEvent(&tme);
            ui.trackingMouse = true;
        }

        if (ui.dragArmed && !ui.dragging) {
            const int dx = std::abs(pt.x - ui.dragStart.x);
            const int dy = std::abs(pt.y - ui.dragStart.y);
            if (dx >= S(3) || dy >= S(3)) {
                ui.dragging = true;
                ui.dragInsertIndex = ComputeInsertIndexFromPoint(list, pt);
            }
        }

        if (ui.dragging) {
            POINT screenPt = pt;
            ClientToScreen(list, &screenPt);
            int dropListId = listId;
            POINT dropClientPt = pt;
            if (parent) {
                int hitListId = listId;
                POINT hitClientPt = pt;
                if (ScreenPointToListClient(parent, screenPt, hitListId, hitClientPt)) {
                    dropListId = hitListId;
                    dropClientPt = hitClientPt;
                }
            }

            if (dropListId == listId) {
                const RECT clientRect = [] (HWND h) {
                    RECT rc{};
                    GetClientRect(h, &rc);
                    return rc;
                }(list);
                if (pt.y < clientRect.top + S(8)) {
                    SendMessageW(list, WM_VSCROLL, SB_LINEUP, 0);
                } else if (pt.y > clientRect.bottom - S(8)) {
                    SendMessageW(list, WM_VSCROLL, SB_LINEDOWN, 0);
                }
                const int nextInsert = ComputeInsertIndexFromPoint(list, pt);
                if (nextInsert != ui.dragInsertIndex) {
                    const int oldInsert = ui.dragInsertIndex;
                    ui.dragInsertIndex = nextInsert;
                    InvalidateInsertMarker(list, oldInsert);
                    InvalidateInsertMarker(list, nextInsert);
                }
            } else {
                if (ui.dragInsertIndex >= 0) {
                    const int oldInsert = ui.dragInsertIndex;
                    ui.dragInsertIndex = -1;
                    InvalidateInsertMarker(list, oldInsert);
                }
                MarkListInteraction(ListUiStateFromListId(dropListId));
            }
            SetCursor(LoadCursorW(nullptr, IDC_SIZEALL));
            return 0;
        }

        const int item = HitTestListItem(list, pt);
        const InlineAction action = HitTestInlineAction(list, item, pt);
        if (item != ui.hoverItem || action != ui.hoverAction) {
            const int oldItem = ui.hoverItem;
            ui.hoverItem = item;
            ui.hoverAction = action;
            InvalidateListItem(list, oldItem);
            InvalidateListItem(list, item);
        }
        if (action != InlineAction::None) {
            SetCursor(LoadCursorW(nullptr, IDC_HAND));
            return 0;
        }
        break;
    }
    case WM_LBUTTONDOWN: {
        MarkListInteraction(ui);
        POINT pt{
            static_cast<int>(static_cast<short>(LOWORD(lParam))),
            static_cast<int>(static_cast<short>(HIWORD(lParam)))
        };
        const int item = HitTestListItem(list, pt);
        const InlineAction action = HitTestInlineAction(list, item, pt);
        if (item >= 0) {
            const int oldPressedItem = ui.pressedItem;
            SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(item), 0);
            if (action == InlineAction::ToggleEnabled ||
                action == InlineAction::ToggleBypass ||
                action == InlineAction::Remove) {
                ui.pressedItem = item;
                ui.pressedAction = action;
            } else {
                ui.dragArmed = true;
                ui.dragSource = item;
                ui.dragStart = pt;
                ui.dragInsertIndex = item;
            }
            SetCapture(list);
            InvalidateListItem(list, oldPressedItem);
            InvalidateListItem(list, item);
            return 0;
        }
        break;
    }
    case WM_LBUTTONUP: {
        MarkListInteraction(ui);
        POINT pt{
            static_cast<int>(static_cast<short>(LOWORD(lParam))),
            static_cast<int>(static_cast<short>(HIWORD(lParam)))
        };
        const int pressedItem = ui.pressedItem;
        const InlineAction pressedAction = ui.pressedAction;
        const bool wasDragging = ui.dragging;
        const int dragSource = ui.dragSource;
        const int dragInsert = ui.dragInsertIndex;

        ui.pressedItem = -1;
        ui.pressedAction = InlineAction::None;
        ui.dragArmed = false;
        ui.dragging = false;
        ui.dragSource = -1;
        ui.dragInsertIndex = -1;
        if (GetCapture() == list) {
            ReleaseCapture();
        }

        if (pressedItem >= 0 && pressedAction != InlineAction::None) {
            const int hitItem = HitTestListItem(list, pt);
            const InlineAction hitAction = HitTestInlineAction(list, hitItem, pt);
            if (pressedItem == hitItem && pressedAction == hitAction) {
                InvokeInlineListAction(parent, listId, pressedItem, pressedAction);
            }
            InvalidateListItem(list, pressedItem);
            return 0;
        }

        if (wasDragging && dragSource >= 0) {
            int dropListId = listId;
            POINT dropClientPt = pt;
            POINT screenPt = pt;
            ClientToScreen(list, &screenPt);
            if (parent) {
                int hitListId = listId;
                POINT hitClientPt = pt;
                if (ScreenPointToListClient(parent, screenPt, hitListId, hitClientPt)) {
                    dropListId = hitListId;
                    dropClientPt = hitClientPt;
                }
            }

            if (dropListId == listId) {
                const int count = static_cast<int>(SendMessageW(list, LB_GETCOUNT, 0, 0));
                if (count > 0 && dragInsert >= 0) {
                    int target = dragInsert;
                    if (target > dragSource) {
                        target -= 1;
                    }
                    target = std::clamp(target, 0, count - 1);
                    if (target != dragSource) {
                        HandleMoveEffectToIndex(parent, ChainFromListId(listId), listId, dragSource, target);
                    }
                }
            } else if (parent) {
                HWND dropList = GetDlgItem(parent, dropListId);
                int dropInsertIndex = 0;
                if (dropList) {
                    dropInsertIndex = ComputeInsertIndexFromPoint(dropList, dropClientPt);
                    if (dropInsertIndex < 0) {
                        dropInsertIndex = static_cast<int>(SendMessageW(dropList, LB_GETCOUNT, 0, 0));
                    }
                }
                HandleMoveEffectBetweenLists(parent, listId, dragSource, dropListId, dropInsertIndex);
            }
            InvalidateInsertMarker(list, dragInsert);
            InvalidateListItem(list, dragSource);
            return 0;
        }
        break;
    }
    case WM_CAPTURECHANGED:
        MarkListInteraction(ui);
        InvalidateListItem(list, ui.pressedItem);
        InvalidateInsertMarker(list, ui.dragInsertIndex);
        ui.pressedItem = -1;
        ui.pressedAction = InlineAction::None;
        ui.dragArmed = false;
        ui.dragging = false;
        ui.dragSource = -1;
        ui.dragInsertIndex = -1;
        break;
    case WM_LBUTTONDBLCLK: {
        POINT pt{
            static_cast<int>(static_cast<short>(LOWORD(lParam))),
            static_cast<int>(static_cast<short>(HIWORD(lParam)))
        };
        const int item = HitTestListItem(list, pt);
        const InlineAction action = HitTestInlineAction(list, item, pt);
        if (action != InlineAction::None) {
            return 0;
        }
        if (item >= 0 && parent) {
            SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(item), 0);
            HandleOpenEditor(parent, ChainFromListId(listId), listId);
            return 0;
        }
        break;
    }
    default:
        break;
    }
    return DefSubclassProc(list, msg, wParam, lParam);
}

void DrawEffectListItem(const DRAWITEMSTRUCT* dis) {
    if (!dis || dis->itemID == static_cast<UINT>(-1)) {
        return;
    }
    const int listId = GetDlgCtrlID(dis->hwndItem);
    EffectListUiState& ui = ListUiStateFromListId(listId);
    if (dis->itemID >= ui.rows.size()) {
        return;
    }
    const EffectRowView& row = ui.rows[dis->itemID];
    const bool compactRow = (dis->rcItem.right - dis->rcItem.left) <= S(320);
    const bool selected = (dis->itemState & ODS_SELECTED) != 0;
    const bool focused = (dis->itemState & ODS_FOCUS) != 0;
    const bool draggingThis = ui.dragging && (static_cast<int>(dis->itemID) == ui.dragSource);
    const bool hoveredItem = (ui.hoverItem == static_cast<int>(dis->itemID));

    RECT rc = dis->rcItem;
    COLORREF bg = RGB(252, 253, 255);
    if (selected) {
        bg = RGB(226, 238, 255);
    } else if (hoveredItem) {
        bg = RGB(245, 249, 255);
    }
    if (draggingThis) {
        bg = RGB(236, 244, 255);
    }
    HBRUSH bgBrush = CreateSolidBrush(bg);
    FillRect(dis->hDC, &rc, bgBrush);
    DeleteObject(bgBrush);

    const EffectRowLayout layout = ComputeEffectRowLayout(rc);

    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, RGB(92, 102, 118));
    SelectObject(dis->hDC, g_fontTiny ? g_fontTiny : g_fontSmall);
    RECT idTextRect = layout.idRect;
    DrawTextW(dis->hDC, row.idText.c_str(), -1, &idTextRect, DT_SINGLELINE | DT_LEFT | DT_VCENTER | DT_END_ELLIPSIS);

    const int handleW = std::max(1, static_cast<int>(layout.handleRect.right - layout.handleRect.left));
    const bool compactHandle = handleW <= S(12);
    const int dotW = S(compactHandle ? 3 : 4);
    const int dotH = S(compactHandle ? 3 : 4);
    const int dotGap = S(compactHandle ? 2 : 3);
    const int handleCenterY = layout.handleRect.top + ((layout.handleRect.bottom - layout.handleRect.top) / 2);
    const int firstDotY = handleCenterY - dotH - (dotGap / 2);
    for (int col = 0; col < 2; ++col) {
        for (int rowDot = 0; rowDot < 3; ++rowDot) {
            RECT dot{
                layout.handleRect.left + col * (dotW + dotGap),
                firstDotY + rowDot * (dotH + dotGap),
                layout.handleRect.left + col * (dotW + dotGap) + dotW,
                firstDotY + rowDot * (dotH + dotGap) + dotH
            };
            HBRUSH dotBrush = CreateSolidBrush(RGB(157, 171, 194));
            FillRect(dis->hDC, &dot, dotBrush);
            DeleteObject(dotBrush);
        }
    }

    SetTextColor(dis->hDC, g_theme.text);
    SelectObject(dis->hDC, g_fontBody ? g_fontBody : g_fontSmall);
    RECT nameRect = layout.nameRect;
    DrawTextW(dis->hDC, row.displayName.c_str(), -1, &nameRect, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
    SetTextColor(dis->hDC, g_theme.muted);
    SelectObject(dis->hDC, g_fontTiny ? g_fontTiny : g_fontSmall);
    std::wstring statusText = row.statusText;
    if (compactRow) {
        statusText = row.enabled ? L"ON" : L"OFF";
        statusText += row.bypass ? L"  |  BYP ON" : L"  |  BYP OFF";
    }
    RECT statusRect = layout.statusRect;
    DrawTextW(dis->hDC, statusText.c_str(), -1, &statusRect, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);

    const bool hoverEnable = (ui.hoverItem == static_cast<int>(dis->itemID) && ui.hoverAction == InlineAction::ToggleEnabled);
    const bool hoverBypass = (ui.hoverItem == static_cast<int>(dis->itemID) && ui.hoverAction == InlineAction::ToggleBypass);
    const bool hoverRemove = (ui.hoverItem == static_cast<int>(dis->itemID) && ui.hoverAction == InlineAction::Remove);
    const bool pressedEnable = (ui.pressedItem == static_cast<int>(dis->itemID) && ui.pressedAction == InlineAction::ToggleEnabled);
    const bool pressedBypass = (ui.pressedItem == static_cast<int>(dis->itemID) && ui.pressedAction == InlineAction::ToggleBypass);
    const bool pressedRemove = (ui.pressedItem == static_cast<int>(dis->itemID) && ui.pressedAction == InlineAction::Remove);
    const wchar_t* enableLabel = row.enabled ? L"On" : L"Off";
    const wchar_t* bypassLabel = L"Byp";
    const wchar_t* removeLabel = compactRow ? L"X" : L"Remove";
    HFONT buttonFont = compactRow ? (g_fontTiny ? g_fontTiny : g_fontSmall) : g_fontSmall;

    DrawInlineButton(dis->hDC, layout.enableRect, enableLabel, row.enabled, hoverEnable, pressedEnable, buttonFont);
    DrawInlineButton(dis->hDC, layout.bypassRect, bypassLabel, row.bypass, hoverBypass, pressedBypass, buttonFont);
    DrawInlineButton(dis->hDC, layout.removeRect, removeLabel, false, hoverRemove, pressedRemove, buttonFont);

    const int count = static_cast<int>(ui.rows.size());
    if (ui.dragging && ui.dragInsertIndex >= 0) {
        HPEN linePen = CreatePen(PS_SOLID, std::max(1, S(2)), g_theme.accent);
        HGDIOBJ oldLinePen = SelectObject(dis->hDC, linePen);
        if (ui.dragInsertIndex == static_cast<int>(dis->itemID)) {
            MoveToEx(dis->hDC, rc.left + S(2), rc.top + 1, nullptr);
            LineTo(dis->hDC, rc.right - S(2), rc.top + 1);
        } else if (ui.dragInsertIndex == count && static_cast<int>(dis->itemID) == count - 1) {
            MoveToEx(dis->hDC, rc.left + S(2), rc.bottom - 1, nullptr);
            LineTo(dis->hDC, rc.right - S(2), rc.bottom - 1);
        }
        SelectObject(dis->hDC, oldLinePen);
        DeleteObject(linePen);
    }

    HPEN dividerPen = CreatePen(PS_SOLID, 1, RGB(225, 231, 240));
    HGDIOBJ oldDivider = SelectObject(dis->hDC, dividerPen);
    MoveToEx(dis->hDC, rc.left + S(2), rc.bottom - 1, nullptr);
    LineTo(dis->hDC, rc.right - S(2), rc.bottom - 1);
    SelectObject(dis->hDC, oldDivider);
    DeleteObject(dividerPen);

    if (focused) {
        RECT focusRc = dis->rcItem;
        focusRc.left += S(2);
        focusRc.right -= S(2);
        focusRc.top += S(2);
        focusRc.bottom -= S(2);
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
        g_windowTitleStateValid = false;
        g_windowTitleEffectsEnabled = false;
        g_lastStatusSampleMs = 0;
        g_cachedHostStatus = {};
        g_cachedMusicEffectsCount = 0;
        g_cachedMicEffectsCount = 0;
        g_effectDisplayIdByUid.clear();
        g_nextEffectDisplayId = 1;
        ResetListPointerState(g_musicListUi);
        ResetListPointerState(g_micListUi);
        g_musicListUi.rows.clear();
        g_micListUi.rows.clear();
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

        const int meterLeft = g_rcMusic.left + S(kCardInnerPaddingPx);
        const int meterTop = g_rcMusic.top + S(kEffectsMeterSectionTopOffsetPx);
        const int meterStatusX = g_rcMusicMeter.right + S(4);
        const int meterStatusW = std::max(0, static_cast<int>((g_rcMusic.right - S(kCardInnerPaddingPx)) - meterStatusX));
        CreateWindowW(L"STATIC", L"Audio Meter", WS_CHILD | WS_VISIBLE, meterLeft, meterTop + S(kEffectsAudioMeterLabelTopPx), S(130), S(24), hwnd, reinterpret_cast<HMENU>(IDC_AUDIO_METER_LABEL), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"No signal", WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS, meterStatusX, g_rcMusicMeter.top + S(2), meterStatusW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_AUDIO_METER_TEXT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"CLIP Events: 0", WS_CHILD | WS_VISIBLE, meterStatusX, g_rcMusicClip.top, meterStatusW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_AUDIO_CLIP_EVENTS), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"Mic Meter", WS_CHILD | WS_VISIBLE, meterLeft, meterTop + S(kEffectsMicMeterLabelTopPx), S(130), S(24), hwnd, reinterpret_cast<HMENU>(IDC_MIC_METER_LABEL), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"No signal", WS_CHILD | WS_VISIBLE | SS_ENDELLIPSIS, meterStatusX, g_rcMicMeter.top + S(2), meterStatusW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_MIC_METER_TEXT), nullptr, nullptr);
        CreateWindowW(L"STATIC", L"CLIP Events: 0", WS_CHILD | WS_VISIBLE, meterStatusX, g_rcMicClip.top, meterStatusW, S(18), hwnd, reinterpret_cast<HMENU>(IDC_MIC_CLIP_EVENTS), nullptr, nullptr);

        const int effectsLeft = g_rcMic.left + S(kCardInnerPaddingPx);
        const int effectsRight = g_rcMic.right - S(kCardInnerPaddingPx);
        const int effectsTop = g_rcMic.top + S(kEffectsListSectionTopOffsetPx);
        const int listH = S(kEffectsListHeightPx);
        const int colGap = S(kEffectsListColumnsGapPx);
        const int colW = std::max(S(220), ((effectsRight - effectsLeft) - colGap) / 2);
        const int leftColX = effectsLeft;
        const int rightColX = leftColX + colW + colGap;
        const int effectsLabelY = effectsTop + S(kEffectsListHeaderOffsetPx);
        const int listTop = effectsTop + S(kEffectsListTopOffsetPx);
        const int buttonsTop = listTop + listH + S(kEffectsListButtonTopGapPx);
        const int buttonGap = S(kEffectsListButtonGapPx);
        const int addBtnW = S(kEffectsListButtonAddWidthPx);
        const int smallBtnW = S(kEffectsListButtonSmallWidthPx);

        CreateWindowW(L"STATIC", L"Audio Effects", WS_CHILD | WS_VISIBLE, leftColX, effectsLabelY, S(160), S(20), hwnd, nullptr, nullptr, nullptr);
        HWND musicList = CreateWindowW(
            L"LISTBOX",
            L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
            leftColX,
            listTop,
            colW,
            listH,
            hwnd,
            reinterpret_cast<HMENU>(IDC_MUSIC_LIST),
            nullptr,
            nullptr);
        CreateWindowW(L"BUTTON", L"Add", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, leftColX, buttonsTop, addBtnW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_ADD), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Up", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, leftColX + addBtnW + buttonGap, buttonsTop, smallBtnW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_UP), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Down", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, leftColX + addBtnW + buttonGap + smallBtnW + buttonGap, buttonsTop, smallBtnW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_MUSIC_DOWN), nullptr, nullptr);
        if (musicList) {
            SetWindowSubclass(musicList, EffectListSubclassProc, kMusicListSubclassId, static_cast<DWORD_PTR>(IDC_MUSIC_LIST));
        }

        CreateWindowW(L"STATIC", L"Mic Effects", WS_CHILD | WS_VISIBLE, rightColX, effectsLabelY, S(160), S(20), hwnd, nullptr, nullptr, nullptr);
        HWND micList = CreateWindowW(
            L"LISTBOX",
            L"",
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS,
            rightColX,
            listTop,
            colW,
            listH,
            hwnd,
            reinterpret_cast<HMENU>(IDC_MIC_LIST),
            nullptr,
            nullptr);
        CreateWindowW(L"BUTTON", L"Add", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, rightColX, buttonsTop, addBtnW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_ADD), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Up", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, rightColX + addBtnW + buttonGap, buttonsTop, smallBtnW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_UP), nullptr, nullptr);
        CreateWindowW(L"BUTTON", L"Down", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, rightColX + addBtnW + buttonGap + smallBtnW + buttonGap, buttonsTop, smallBtnW, S(28), hwnd, reinterpret_cast<HMENU>(IDC_MIC_DOWN), nullptr, nullptr);
        if (micList) {
            SetWindowSubclass(micList, EffectListSubclassProc, kMicListSubclassId, static_cast<DWORD_PTR>(IDC_MIC_LIST));
        }

        CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT, g_rcStatus.left + S(kCardInnerPaddingPx), g_rcStatus.top + S(14), (g_rcStatus.right - g_rcStatus.left) - S(kCardInnerPaddingPx * 2), g_rcStatus.bottom - g_rcStatus.top - S(18), hwnd, reinterpret_cast<HMENU>(IDC_STATUS), nullptr, nullptr);

        ApplyControlTheme(hwnd);
        ApplyFonts(hwnd);
        ReloadAllLists(hwnd);
        RefreshStatus(hwnd, true);
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
            if (!IsListInteracting(hwnd) && nowMs > (g_lastListRefreshTickMs + 1000ULL)) {
                ReloadAllLists(hwnd);
                g_lastListRefreshTickMs = nowMs;
            }
            RefreshStatus(hwnd, false);
            UpdateMeterVisuals(hwnd);
            return 0;
        }
        break;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case IDC_REPO_LINK:
            if (HIWORD(wParam) == STN_CLICKED) {
                ShellExecuteW(hwnd, L"open", kRepoUrl, nullptr, nullptr, SW_SHOWNORMAL);
            }
            return 0;
        case IDC_ENABLE_EFFECTS:
            MicMixApp::Instance().SetEffectsEnabled(true, true);
            RefreshStatus(hwnd, true);
            return 0;
        case IDC_DISABLE_EFFECTS:
            MicMixApp::Instance().SetEffectsEnabled(false, true);
            RefreshStatus(hwnd, true);
            return 0;
        case IDC_MONITOR:
            MicMixApp::Instance().ToggleMonitor();
            UpdateMonitorButton(hwnd);
            RefreshStatus(hwnd, true);
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
                RefreshStatus(hwnd, true);
            }
            return 0;
        case IDC_MUSIC_ADD:
            HandleAddEffect(hwnd, EffectChain::Music);
            return 0;
        case IDC_MUSIC_UP:
            HandleMoveEffect(hwnd, EffectChain::Music, IDC_MUSIC_LIST, -1);
            return 0;
        case IDC_MUSIC_DOWN:
            HandleMoveEffect(hwnd, EffectChain::Music, IDC_MUSIC_LIST, 1);
            return 0;
        case IDC_MIC_ADD:
            HandleAddEffect(hwnd, EffectChain::Mic);
            return 0;
        case IDC_MIC_UP:
            HandleMoveEffect(hwnd, EffectChain::Mic, IDC_MIC_LIST, -1);
            return 0;
        case IDC_MIC_DOWN:
            HandleMoveEffect(hwnd, EffectChain::Mic, IDC_MIC_LIST, 1);
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
                RefreshStatus(hwnd, true);
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
        if (HWND musicList = GetDlgItem(hwnd, IDC_MUSIC_LIST)) {
            RemoveWindowSubclass(musicList, EffectListSubclassProc, kMusicListSubclassId);
        }
        if (HWND micList = GetDlgItem(hwnd, IDC_MIC_LIST)) {
            RemoveWindowSubclass(micList, EffectListSubclassProc, kMicListSubclassId);
        }
        ResetListPointerState(g_musicListUi);
        ResetListPointerState(g_micListUi);
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
