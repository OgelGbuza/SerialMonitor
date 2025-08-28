// darktheme.h
#pragma once
#include <windows.h>
#include <Uxtheme.h>

#pragma comment(lib, "uxtheme.lib")

// These are undocumented functions from uxtheme.dll, used to enable dark mode
// We must link them dynamically at runtime.

enum PreferredAppMode {
    Default,
    AllowDark,
    ForceDark,
    ForceLight,
    Max
};

using fnSetPreferredAppMode = PreferredAppMode(WINAPI*)(PreferredAppMode appMode);
using fnAllowDarkModeForWindow = bool (WINAPI*)(HWND hWnd, bool allow);
using fnFlushMenuThemes = void (WINAPI*)();

fnSetPreferredAppMode SetPreferredAppMode = nullptr;
fnAllowDarkModeForWindow AllowDarkModeForWindow = nullptr;
fnFlushMenuThemes FlushMenuThemes = nullptr;

// Call this once at the beginning of your application
void InitDarkMode() {
    HMODULE hUxtheme = LoadLibraryW(L"uxtheme.dll");
    if (hUxtheme) {
        SetPreferredAppMode = (fnSetPreferredAppMode)GetProcAddress(hUxtheme, MAKEINTRESOURCEA(135));
        AllowDarkModeForWindow = (fnAllowDarkModeForWindow)GetProcAddress(hUxtheme, MAKEINTRESOURCEA(133));
        FlushMenuThemes = (fnFlushMenuThemes)GetProcAddress(hUxtheme, MAKEINTRESOURCEA(136));

        if (SetPreferredAppMode && AllowDarkModeForWindow && FlushMenuThemes) {
            SetPreferredAppMode(AllowDark);
            FlushMenuThemes();
        }
    }
}