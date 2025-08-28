// darktheme.h
#pragma once
#include <windows.h>
#include <Uxtheme.h>

#pragma comment(lib, "uxtheme.lib")

enum PreferredAppMode { Default, AllowDark, ForceDark, ForceLight, Max };
using fnSetPreferredAppMode = PreferredAppMode(WINAPI*)(PreferredAppMode appMode);
using fnAllowDarkModeForWindow = bool (WINAPI*)(HWND hWnd, bool allow);
using fnFlushMenuThemes = void (WINAPI*)();

fnSetPreferredAppMode SetPreferredAppMode = nullptr;
fnAllowDarkModeForWindow AllowDarkModeForWindow = nullptr;
fnFlushMenuThemes FlushMenuThemes = nullptr;

inline void InitDarkMode() {
    HMODULE hUxtheme = LoadLibraryW(L"uxtheme.dll");
    if (hUxtheme) {
        SetPreferredAppMode = (fnSetPreferredAppMode)GetProcAddress(hUxtheme, MAKEINTRESOURCEA(135));
        AllowDarkModeForWindow = (fnAllowDarkModeForWindow)GetProcAddress(hUxtheme, MAKEINTRESOURCEA(133));
        FlushMenuThemes = (fnFlushMenuThemes)GetProcAddress(hUxtheme, MAKEINTRESOURCEA(136));
        if (SetPreferredAppMode) { SetPreferredAppMode(AllowDark); }
        if (FlushMenuThemes) { FlushMenuThemes(); }
    }
}