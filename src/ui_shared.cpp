#include "ui_shared.h"

#include <algorithm>

UiTheme DefaultUiTheme() {
    return UiTheme{};
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
