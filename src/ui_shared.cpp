#include "ui_shared.h"

#include <algorithm>
#include <array>
#include <cmath>

#include <CommCtrl.h>
#include <uxtheme.h>

UiTheme DefaultUiTheme() {
    return UiTheme{};
}

void EnsureCommonUiResources(UiCommonResources& resources, const UiTheme& theme, int dpi) {
    if (!resources.bgBrush) resources.bgBrush = CreateSolidBrush(theme.bg);
    if (!resources.cardBrush) resources.cardBrush = CreateSolidBrush(theme.card);
    if (!resources.bodyFont) resources.bodyFont = CreateFontW(-ScaleByDpi(13, dpi), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!resources.smallFont) resources.smallFont = CreateFontW(-ScaleByDpi(12, dpi), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!resources.hintFont) resources.hintFont = CreateFontW(-ScaleByDpi(11, dpi), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!resources.tinyFont) resources.tinyFont = CreateFontW(-ScaleByDpi(10, dpi), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!resources.controlLargeFont) resources.controlLargeFont = CreateFontW(-ScaleByDpi(14, dpi), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI");
    if (!resources.titleFont) resources.titleFont = CreateFontW(-ScaleByDpi(24, dpi), 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH, L"Segoe UI Semibold");
    if (!resources.monoFont) resources.monoFont = CreateFontW(-ScaleByDpi(13, dpi), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_OUTLINE_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, FIXED_PITCH, L"Consolas");
}

void ReleaseCommonUiResources(UiCommonResources& resources) {
    if (resources.bodyFont) { DeleteObject(resources.bodyFont); resources.bodyFont = nullptr; }
    if (resources.smallFont) { DeleteObject(resources.smallFont); resources.smallFont = nullptr; }
    if (resources.hintFont) { DeleteObject(resources.hintFont); resources.hintFont = nullptr; }
    if (resources.tinyFont) { DeleteObject(resources.tinyFont); resources.tinyFont = nullptr; }
    if (resources.controlLargeFont) { DeleteObject(resources.controlLargeFont); resources.controlLargeFont = nullptr; }
    if (resources.titleFont) { DeleteObject(resources.titleFont); resources.titleFont = nullptr; }
    if (resources.monoFont) { DeleteObject(resources.monoFont); resources.monoFont = nullptr; }
    if (resources.bgBrush) { DeleteObject(resources.bgBrush); resources.bgBrush = nullptr; }
    if (resources.cardBrush) { DeleteObject(resources.cardBrush); resources.cardBrush = nullptr; }
}

HINSTANCE ResolveModuleHandleFromAddress(const void* anchorAddress) {
    HMODULE mod = nullptr;
    if (anchorAddress &&
        GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
            reinterpret_cast<LPCWSTR>(anchorAddress),
            &mod) != 0 &&
        mod != nullptr) {
        return mod;
    }
    return GetModuleHandleW(nullptr);
}

void SetControlFontById(HWND hwnd, int id, HFONT font) {
    HWND ctl = GetDlgItem(hwnd, id);
    if (ctl && font) {
        SendMessageW(ctl, WM_SETFONT, reinterpret_cast<WPARAM>(font), TRUE);
    }
}

bool GetCheckboxChecked(HWND hwnd, int id) {
    const HWND ctl = GetDlgItem(hwnd, id);
    if (!ctl) {
        return false;
    }
    return SendMessageW(ctl, BM_GETCHECK, 0, 0) == BST_CHECKED;
}

void SetCheckboxChecked(HWND hwnd, int id, bool checked) {
    const HWND ctl = GetDlgItem(hwnd, id);
    if (!ctl) {
        return;
    }
    SendMessageW(ctl, BM_SETCHECK, checked ? BST_CHECKED : BST_UNCHECKED, 0);
}

void SetMonitorButtonText(HWND hwnd, int id, bool enabled) {
    const HWND ctl = GetDlgItem(hwnd, id);
    if (!ctl) {
        return;
    }
    SetWindowTextW(ctl, enabled ? L"Monitor Mix: On" : L"Monitor Mix: Off");
}

bool IsSignalMeterActive(
    bool activeFlag,
    float sendPeakDbfs,
    float peakDbfs,
    float rmsDbfs,
    float inactiveSentinelDbfs,
    float peakActiveThresholdDbfs,
    float rmsActiveThresholdDbfs) {
    const float shownPeakDb = (sendPeakDbfs > inactiveSentinelDbfs) ? sendPeakDbfs : peakDbfs;
    const bool levelSuggestsSignal = (shownPeakDb > peakActiveThresholdDbfs) || (rmsDbfs > rmsActiveThresholdDbfs);
    return activeFlag || levelSuggestsSignal;
}

int ScaleByDpi(int px, int dpi) {
    return MulDiv(px, dpi, 96);
}

std::wstring Utf8ToWide(const std::string& text) {
    if (text.empty()) {
        return {};
    }
    const int len = MultiByteToWideChar(CP_UTF8, 0, text.c_str(), -1, nullptr, 0);
    if (len <= 0) {
        return {};
    }
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
    const int len = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, nullptr, 0, nullptr, nullptr);
    if (len <= 0) {
        return {};
    }
    std::string out(static_cast<size_t>(len), '\0');
    WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, out.data(), len, nullptr, nullptr);
    if (!out.empty() && out.back() == '\0') {
        out.pop_back();
    }
    return out;
}

void DrawFlatCardFrame(HDC dc, const RECT& rc, HBRUSH fillBrush, COLORREF borderColor) {
    FillRect(dc, &rc, fillBrush);
    HPEN pen = CreatePen(PS_SOLID, 1, borderColor);
    HGDIOBJ oldPen = SelectObject(dc, pen);
    HGDIOBJ oldBrush = SelectObject(dc, GetStockObject(HOLLOW_BRUSH));
    Rectangle(dc, rc.left, rc.top, rc.right, rc.bottom);
    SelectObject(dc, oldBrush);
    SelectObject(dc, oldPen);
    DeleteObject(pen);
}

void DrawHeaderBadge(
    HDC dc,
    const RECT& rc,
    const HeaderBadgeVisual& visual,
    HFONT textFont,
    HFONT fallbackFont,
    int dpi) {
    if (rc.right <= rc.left || rc.bottom <= rc.top) {
        return;
    }

    HPEN pen = CreatePen(PS_SOLID, 1, visual.border);
    HBRUSH brush = CreateSolidBrush(visual.bg);
    HGDIOBJ oldPen = SelectObject(dc, pen);
    HGDIOBJ oldBrush = SelectObject(dc, brush);
    RoundRect(
        dc,
        rc.left,
        rc.top,
        rc.right,
        rc.bottom,
        ScaleByDpi(10, dpi),
        ScaleByDpi(10, dpi));
    SelectObject(dc, oldBrush);
    SelectObject(dc, oldPen);
    DeleteObject(brush);
    DeleteObject(pen);

    const int dotSize = ScaleByDpi(6, dpi);
    const int dotInsetX = ScaleByDpi(8, dpi);
    const int dotY = rc.top + ((rc.bottom - rc.top - dotSize) / 2);
    RECT dotRc{
        rc.left + dotInsetX,
        dotY,
        rc.left + dotInsetX + dotSize,
        dotY + dotSize
    };

    HPEN dotPen = CreatePen(PS_SOLID, 1, visual.dot);
    HBRUSH dotBrush = CreateSolidBrush(visual.dot);
    HGDIOBJ oldDotPen = SelectObject(dc, dotPen);
    HGDIOBJ oldDotBrush = SelectObject(dc, dotBrush);
    Ellipse(dc, dotRc.left, dotRc.top, dotRc.right, dotRc.bottom);
    SelectObject(dc, oldDotBrush);
    SelectObject(dc, oldDotPen);
    DeleteObject(dotBrush);
    DeleteObject(dotPen);

    RECT textRc = rc;
    textRc.left += ScaleByDpi(12, dpi);
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, visual.text);
    SelectObject(dc, textFont ? textFont : fallbackFont);
    DrawTextW(dc, visual.label, -1, &textRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
}

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
    int metaPreferredWidthPx) {
    HeaderLayout out{};
    const wchar_t* safeTitle = (titleText && titleText[0] != L'\0') ? titleText : L"MicMix";

    int titleWidthPx = ScaleByDpi(112, dpi);
    int badgeWidthPx = ScaleByDpi(74, dpi);
    if (dc) {
        HGDIOBJ old = SelectObject(dc, titleFont ? titleFont : fallbackFont);
        SIZE titleSize{};
        if (GetTextExtentPoint32W(dc, safeTitle, lstrlenW(safeTitle), &titleSize) != 0) {
            titleWidthPx = titleSize.cx + ScaleByDpi(2, dpi);
        }
        SelectObject(dc, badgeFont ? badgeFont : fallbackFont);
        SIZE badgeTextSize{};
        if (GetTextExtentPoint32W(dc, L"ACTIVE", 6, &badgeTextSize) != 0) {
            badgeWidthPx = std::clamp<int>(
                static_cast<int>(badgeTextSize.cx) + ScaleByDpi(28, dpi),
                ScaleByDpi(78, dpi),
                ScaleByDpi(108, dpi));
        }
        SelectObject(dc, old);
    }

    const int titlePadRight = ScaleByDpi(2, dpi);
    const int badgeHeightPx = ScaleByDpi(22, dpi);
    const int badgeX = contentLeft + titleWidthPx + titlePadRight;
    const int badgeY = titleY + ((titleHeight - badgeHeightPx) / 2) - ScaleByDpi(1, dpi);
    out.titleWidthPx = titleWidthPx;
    out.badgeRect = { badgeX, badgeY, badgeX + badgeWidthPx, badgeY + badgeHeightPx };

    const int metaPreferredWidth = ScaleByDpi(metaPreferredWidthPx, dpi);
    const int metaMinX = out.badgeRect.right + ScaleByDpi(10, dpi);
    out.metaX = std::max(contentRight - metaPreferredWidth, metaMinX);
    out.metaWidth = std::max(ScaleByDpi(140, dpi), contentRight - out.metaX);
    return out;
}

void DrawSectionHeadings(
    HDC dc,
    const SectionHeading* headings,
    int headingCount,
    HFONT headingFont,
    HFONT fallbackFont,
    COLORREF textColor,
    int dpi,
    int insetXPx,
    int insetYPx) {
    if (!dc || !headings || headingCount <= 0) {
        return;
    }
    SetBkMode(dc, TRANSPARENT);
    SetTextColor(dc, textColor);
    HGDIOBJ oldFont = SelectObject(dc, headingFont ? headingFont : fallbackFont);
    for (int i = 0; i < headingCount; ++i) {
        const SectionHeading& heading = headings[i];
        if (!heading.label || heading.label[0] == L'\0') {
            continue;
        }
        const int x = heading.cardRect.left + ScaleByDpi(insetXPx, dpi);
        const int y = heading.cardRect.top + ScaleByDpi(insetYPx, dpi);
        TextOutW(dc, x, y, heading.label, lstrlenW(heading.label));
    }
    SelectObject(dc, oldFont);
}

namespace {
void FillSolidRect(HDC hdc, const RECT& rc, COLORREF color);
}

void DrawOwnerCheckboxShared(
    const DRAWITEMSTRUCT* dis,
    bool checked,
    HFONT textFont,
    HFONT fallbackFont,
    COLORREF textColor,
    COLORREF cardColor,
    int dpi) {
    if (!dis) {
        return;
    }
    RECT rc = dis->rcItem;
    FillSolidRect(dis->hDC, rc, cardColor);

    const bool disabled = (dis->itemState & ODS_DISABLED) != 0;
    const int boxSize = ScaleByDpi(15, dpi);
    RECT box{
        rc.left + ScaleByDpi(2, dpi),
        rc.top + ((rc.bottom - rc.top - boxSize) / 2),
        rc.left + ScaleByDpi(2, dpi) + boxSize,
        rc.top + ((rc.bottom - rc.top - boxSize) / 2) + boxSize
    };

    FillSolidRect(dis->hDC, box, RGB(252, 253, 255));
    const COLORREF borderColor = disabled ? GetSysColor(COLOR_GRAYTEXT) : GetSysColor(COLOR_BTNSHADOW);
    HPEN border = CreatePen(PS_SOLID, 1, borderColor);
    HGDIOBJ oldPen = SelectObject(dis->hDC, border);
    HGDIOBJ oldBrush = SelectObject(dis->hDC, GetStockObject(HOLLOW_BRUSH));
    Rectangle(dis->hDC, box.left, box.top, box.right, box.bottom);
    SelectObject(dis->hDC, oldBrush);
    SelectObject(dis->hDC, oldPen);
    DeleteObject(border);

    if (checked) {
        HPEN markPen = CreatePen(PS_SOLID, std::max(1, ScaleByDpi(2, dpi)), disabled ? RGB(130, 136, 148) : RGB(54, 68, 92));
        HGDIOBJ oldMarkPen = SelectObject(dis->hDC, markPen);
        const int pad = ScaleByDpi(3, dpi);
        MoveToEx(dis->hDC, box.left + pad, box.top + pad, nullptr);
        LineTo(dis->hDC, box.right - pad - 1, box.bottom - pad - 1);
        MoveToEx(dis->hDC, box.right - pad - 1, box.top + pad, nullptr);
        LineTo(dis->hDC, box.left + pad, box.bottom - pad - 1);
        SelectObject(dis->hDC, oldMarkPen);
        DeleteObject(markPen);
    }

    wchar_t text[256]{};
    GetWindowTextW(dis->hwndItem, text, static_cast<int>(std::size(text)));
    RECT textRc = rc;
    textRc.left = box.right + ScaleByDpi(8, dpi);
    SetBkMode(dis->hDC, TRANSPARENT);
    SetTextColor(dis->hDC, disabled ? RGB(144, 150, 162) : textColor);
    SelectObject(dis->hDC, textFont ? textFont : fallbackFont);
    DrawTextW(dis->hDC, text, -1, &textRc, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
}

int MeterDbToPixels(float dbfs, int widthPx) {
    const float clamped = std::clamp(dbfs, -60.0f, 0.0f);
    return static_cast<int>(((clamped + 60.0f) / 60.0f) * static_cast<float>(widthPx));
}

float ClipDangerUnitFromDb(float dbfs) {
    const float clamped = std::clamp(dbfs, -18.0f, 0.0f);
    return (clamped + 18.0f) / 18.0f;
}

int ClipLitSegmentsFromDb(float dbfs, int segments) {
    if (segments <= 0) {
        return 0;
    }
    const float dangerUnit = ClipDangerUnitFromDb(dbfs);
    return std::clamp(static_cast<int>((dangerUnit * static_cast<float>(segments)) + 0.5f), 0, segments);
}

namespace {

void FillSolidRect(HDC hdc, const RECT& rc, COLORREF color) {
    HBRUSH brush = CreateSolidBrush(color);
    FillRect(hdc, &rc, brush);
    DeleteObject(brush);
}

} // namespace

void DrawLevelMeterShared(
    HDC hdc,
    const RECT& meterRect,
    bool active,
    float visualDb,
    float holdDb,
    const MeterFonts& fonts,
    int dpi) {
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
    const int scaleBandH = std::clamp(
        ScaleByDpi(10, dpi),
        7,
        std::max(7, height - ScaleByDpi(8, dpi)));
    RECT scaleRc = rc;
    scaleRc.bottom = std::min(rc.bottom, rc.top + scaleBandH);
    RECT barRc = rc;
    barRc.top = std::min(rc.bottom - 1, scaleRc.bottom);

    FillSolidRect(drawDc, scaleRc, RGB(244, 247, 252));
    FillSolidRect(drawDc, barRc, RGB(235, 240, 246));

    const std::array<int, 7> tickDb = { -60, -36, -24, -18, -12, -6, 0 };
    SetBkMode(drawDc, TRANSPARENT);
    SetTextColor(drawDc, RGB(96, 107, 122));
    HGDIOBJ oldFont = SelectObject(
        drawDc,
        fonts.tinyFont ? fonts.tinyFont : (fonts.hintFont ? fonts.hintFont : (fonts.smallFont ? fonts.smallFont : fonts.bodyFont)));
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

        RECT tickRc{ x, scaleRc.bottom - ScaleByDpi(3, dpi), x + 1, scaleRc.bottom };
        FillSolidRect(drawDc, tickRc, RGB(162, 171, 184));

        wchar_t label[8]{};
        swprintf_s(label, L"%d", db);
        SIZE textSize{};
        if (GetTextExtentPoint32W(drawDc, label, lstrlenW(label), &textSize) != 0) {
            int textLeft = x - (textSize.cx / 2);
            if (i == 0) {
                textLeft = left + ScaleByDpi(2, dpi);
            } else if (i == (tickDb.size() - 1)) {
                textLeft = right - textSize.cx - ScaleByDpi(2, dpi);
            } else {
                const int labelMinX = left + ScaleByDpi(2, dpi);
                const int labelMaxX = right - static_cast<int>(textSize.cx) - ScaleByDpi(2, dpi);
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
        BitBlt(
            hdc,
            meterRect.left,
            meterRect.top,
            targetWidth,
            targetHeight,
            memDc,
            0,
            0,
            SRCCOPY);
        SelectObject(memDc, oldBmp);
        DeleteObject(memBmp);
        DeleteDC(memDc);
    }
}

void DrawClipStripShared(
    HDC hdc,
    const RECT& rcStrip,
    float dangerUnit,
    bool clipRecent,
    const MeterFonts& fonts,
    int dpi) {
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
    const int width = static_cast<int>(rc.right - rc.left);
    const int height = static_cast<int>(rc.bottom - rc.top);
    if (width <= 2 || height <= 2) {
        if (buffered) {
            SelectObject(memDc, oldBmp);
            DeleteObject(memBmp);
            DeleteDC(memDc);
        }
        return;
    }

    FillSolidRect(drawDc, rc, RGB(246, 249, 252));

    constexpr int kSegments = 18;
    const int segGap = std::max(1, ScaleByDpi(1, dpi));
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
        RECT seg{ segLeft, rc.top + ScaleByDpi(2, dpi), segRight, rc.bottom - ScaleByDpi(2, dpi) };
        const float t = static_cast<float>(i) / static_cast<float>(kSegments - 1);
        const COLORREF onColor = RGB(
            static_cast<BYTE>(84 + (154.0f * t)),
            static_cast<BYTE>(176 - (126.0f * t)),
            static_cast<BYTE>(86 - (40.0f * t)));
        const COLORREF color = (i < litSegments) ? onColor : offColor;
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
    HGDIOBJ oldLabelFont = SelectObject(
        drawDc, fonts.tinyFont ? fonts.tinyFont : (fonts.hintFont ? fonts.hintFont : (fonts.smallFont ? fonts.smallFont : fonts.bodyFont)));
    RECT labelRc{ rc.left + ScaleByDpi(5, dpi), rc.top, rc.right - ScaleByDpi(52, dpi), rc.bottom };
    DrawTextW(drawDc, L"Clip Meter", -1, &labelRc, DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS);
    SelectObject(drawDc, oldLabelFont);

    RECT inner{ rc.left + 1, rc.top + 1, rc.right - 1, rc.bottom - 1 };
    RECT pulse{};
    const int pulseWidth = ScaleByDpi(44, dpi);
    if (clipRecent) {
        pulse = { std::max(inner.left, inner.right - pulseWidth), inner.top, inner.right, inner.bottom };
        FillSolidRect(drawDc, pulse, clipColor);
    }

    if (clipRecent) {
        SetTextColor(drawDc, RGB(255, 255, 255));
        HGDIOBJ oldPulseFont = SelectObject(
            drawDc, fonts.tinyFont ? fonts.tinyFont : (fonts.hintFont ? fonts.hintFont : (fonts.smallFont ? fonts.smallFont : fonts.bodyFont)));
        DrawTextW(drawDc, L"CLIP", -1, &pulse, DT_SINGLELINE | DT_CENTER | DT_VCENTER);
        SelectObject(drawDc, oldPulseFont);
    }

    if (buffered) {
        BitBlt(
            hdc,
            rcStrip.left,
            rcStrip.top,
            targetWidth,
            targetHeight,
            memDc,
            0,
            0,
            SRCCOPY);
        SelectObject(memDc, oldBmp);
        DeleteObject(memBmp);
        DeleteDC(memDc);
    }
}

void ApplyMicMixControlTheme(
    HWND hwnd,
    unsigned int flags,
    bool flatButtons,
    bool unthemeNativeCheckboxes,
    const int* unthemeButtonIds,
    int unthemeButtonIdCount) {
    struct ThemeOpts {
        unsigned int flags = 0;
        bool flatButtons = true;
        bool unthemeNativeCheckboxes = true;
        const int* unthemeButtonIds = nullptr;
        int unthemeButtonIdCount = 0;
    } opts{ flags, flatButtons, unthemeNativeCheckboxes, unthemeButtonIds, unthemeButtonIdCount };

    EnumChildWindows(hwnd, [](HWND child, LPARAM lp) -> BOOL {
        const ThemeOpts& opts = *reinterpret_cast<const ThemeOpts*>(lp);
        wchar_t className[32]{};
        GetClassNameW(child, className, static_cast<int>(std::size(className)));

        if ((opts.flags & kUiThemeComboBoxes) != 0 && _wcsicmp(className, WC_COMBOBOXW) == 0) {
            SetWindowTheme(child, L"Explorer", nullptr);
        }
        if ((opts.flags & kUiThemeTrackBars) != 0 && _wcsicmp(className, TRACKBAR_CLASSW) == 0) {
            SetWindowTheme(child, L"", L"");
        }
        if ((opts.flags & kUiThemeListBoxes) != 0 && _wcsicmp(className, L"ListBox") == 0) {
            SetWindowTheme(child, L"Explorer", nullptr);
        }
        if ((opts.flags & kUiThemeButtons) != 0 && _wcsicmp(className, WC_BUTTONW) == 0) {
            const LONG_PTR style = GetWindowLongPtrW(child, GWL_STYLE);
            const LONG_PTR type = (style & BS_TYPEMASK);
            if (opts.flatButtons &&
                (type == BS_PUSHBUTTON || type == BS_DEFPUSHBUTTON || type == BS_CHECKBOX || type == BS_AUTOCHECKBOX)) {
                SetWindowLongPtrW(child, GWL_STYLE, style | BS_FLAT);
            }

            bool untheme = false;
            if (opts.unthemeNativeCheckboxes && (type == BS_CHECKBOX || type == BS_AUTOCHECKBOX)) {
                untheme = true;
            }
            if (!untheme && opts.unthemeButtonIds && opts.unthemeButtonIdCount > 0) {
                const int id = static_cast<int>(GetDlgCtrlID(child));
                for (int i = 0; i < opts.unthemeButtonIdCount; ++i) {
                    if (opts.unthemeButtonIds[i] == id) {
                        untheme = true;
                        break;
                    }
                }
            }
            if (untheme) {
                SetWindowTheme(child, L"", L"");
            } else {
                SetWindowTheme(child, L"Explorer", nullptr);
            }
        }
        return TRUE;
    }, reinterpret_cast<LPARAM>(&opts));
}
