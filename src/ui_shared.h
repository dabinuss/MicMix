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

UiTheme DefaultUiTheme();

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
