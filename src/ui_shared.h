#pragma once

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>

#include <string>

struct UiTheme {
    COLORREF bg = RGB(241, 244, 248);
    COLORREF card = RGB(255, 255, 255);
    COLORREF border = RGB(212, 218, 229);
    COLORREF text = RGB(30, 33, 41);
    COLORREF muted = RGB(96, 105, 119);
    COLORREF accent = RGB(33, 112, 218);
};

struct HeaderBadgeVisual {
    const wchar_t* label = L"OFF";
    COLORREF bg = RGB(134, 150, 173);
    COLORREF border = RGB(102, 118, 141);
    COLORREF text = RGB(255, 255, 255);
    COLORREF dot = RGB(246, 249, 255);
};

inline constexpr int kUiClientWidthPx = 680;
inline constexpr int kUiClientHeightPx = 780;
inline constexpr int kUiCardMarginPx = 16;
inline constexpr int kUiCardGapPx = 12;
inline constexpr int kUiCardInnerPaddingPx = 20;
inline constexpr int kUiTopCardYpx = 82;
inline constexpr int kUiTopCardPrimaryRowOffsetYpx = 26;
inline constexpr int kUiTopCardHintOffsetYpx = 38;
inline constexpr int kUiTopCardCheckboxOffsetYpx = 88;
inline constexpr int kUiTopCardCheckboxWidthPx = 320;
inline constexpr int kUiSectionMeterLabelOffsetYpx = 34;
inline constexpr int kUiSectionListTopOffsetYpx = 84;

struct HeaderLayout {
    int titleWidthPx = 0;
    RECT badgeRect{};
    int metaX = 0;
    int metaWidth = 0;
};

struct SectionHeading {
    RECT cardRect{};
    const wchar_t* label = L"";
};

struct MeterFonts {
    HFONT tinyFont = nullptr;
    HFONT hintFont = nullptr;
    HFONT smallFont = nullptr;
    HFONT bodyFont = nullptr;
};

struct UiCommonResources {
    HFONT bodyFont = nullptr;
    HFONT smallFont = nullptr;
    HFONT hintFont = nullptr;
    HFONT tinyFont = nullptr;
    HFONT controlLargeFont = nullptr;
    HFONT titleFont = nullptr;
    HFONT monoFont = nullptr;
    HBRUSH bgBrush = nullptr;
    HBRUSH cardBrush = nullptr;
};

enum UiControlThemeFlags : unsigned int {
    kUiThemeButtons = 1u << 0,
    kUiThemeComboBoxes = 1u << 1,
    kUiThemeTrackBars = 1u << 2,
    kUiThemeListBoxes = 1u << 3,
};

UiTheme DefaultUiTheme();

void EnsureCommonUiResources(UiCommonResources& resources, const UiTheme& theme, int dpi);
void ReleaseCommonUiResources(UiCommonResources& resources);

HINSTANCE ResolveModuleHandleFromAddress(const void* anchorAddress);
void SetControlFontById(HWND hwnd, int id, HFONT font);
bool GetCheckboxChecked(HWND hwnd, int id);
void SetCheckboxChecked(HWND hwnd, int id, bool checked);
void SetMonitorButtonText(HWND hwnd, int id, bool enabled);
bool IsSignalMeterActive(
    bool activeFlag,
    float sendPeakDbfs,
    float peakDbfs,
    float rmsDbfs,
    float inactiveSentinelDbfs = -119.0f,
    float peakActiveThresholdDbfs = -96.0f,
    float rmsActiveThresholdDbfs = -100.0f);

int ScaleByDpi(int px, int dpi);

std::wstring Utf8ToWide(const std::string& text);
std::string WideToUtf8(const std::wstring& text);

void DrawFlatCardFrame(HDC dc, const RECT& rc, HBRUSH fillBrush, COLORREF borderColor);
void DrawHeaderBadge(
    HDC dc,
    const RECT& rc,
    const HeaderBadgeVisual& visual,
    HFONT textFont,
    HFONT fallbackFont,
    int dpi);

HeaderLayout ComputeHeaderLayout(
    HDC dc,
    HFONT titleFont,
    HFONT badgeFont,
    HFONT fallbackFont,
    int dpi,
    const wchar_t* titleText,
    int contentLeft,
    int contentRight,
    int titleY,
    int titleHeight,
    int metaPreferredWidthPx = 300);

void DrawSectionHeadings(
    HDC dc,
    const SectionHeading* headings,
    int headingCount,
    HFONT headingFont,
    HFONT fallbackFont,
    COLORREF textColor,
    int dpi,
    int insetXPx = kUiCardInnerPaddingPx,
    int insetYPx = 4);

void DrawOwnerCheckboxShared(
    const DRAWITEMSTRUCT* dis,
    bool checked,
    HFONT textFont,
    HFONT fallbackFont,
    COLORREF textColor,
    COLORREF cardColor,
    int dpi);

int MeterDbToPixels(float dbfs, int widthPx);
float ClipDangerUnitFromDb(float dbfs);
int ClipLitSegmentsFromDb(float dbfs, int segments = 18);

void DrawLevelMeterShared(
    HDC hdc,
    const RECT& meterRect,
    bool active,
    float visualDb,
    float holdDb,
    const MeterFonts& fonts,
    int dpi);

void DrawClipStripShared(
    HDC hdc,
    const RECT& rcStrip,
    float dangerUnit,
    bool clipRecent,
    const MeterFonts& fonts,
    int dpi);

void ApplyMicMixControlTheme(
    HWND hwnd,
    unsigned int flags,
    bool flatButtons = true,
    bool unthemeNativeCheckboxes = true,
    const int* unthemeButtonIds = nullptr,
    int unthemeButtonIdCount = 0);
