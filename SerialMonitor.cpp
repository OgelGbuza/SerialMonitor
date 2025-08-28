#include "framework.h"
#include "SerialMonitor.h"
#include "darktheme.h" // >> DARK THEME: Include our new header
#include <windows.h>
#include <string>
#include <vector>
#include <CommCtrl.h> 
#include <ShlObj.h>   

#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Shell32.lib")

#define MAX_LOADSTRING 100

#define WM_SERIAL_DATA_RECEIVED (WM_APP + 1)
#define WM_UPDATE_STATUS        (WM_APP + 2)

// Global Variables
HINSTANCE hInst;
WCHAR szTitle[MAX_LOADSTRING];
WCHAR szWindowClass[MAX_LOADSTRING];

HWND hPortCombo, hBaudCombo, hStartButton, hStopButton, hOutputEdit, hRefreshButton;
HWND hLogDirEdit, hBrowseButton, hStatusLabel, hCancelButton;

HANDLE hThread = NULL;
volatile bool bShouldBeMonitoring = false;

// >> DARK THEME: Brushes for our dark theme colors
HBRUSH g_hbrBackground = CreateSolidBrush(RGB(32, 32, 32));
HBRUSH g_hbrEditBackground = CreateSolidBrush(RGB(43, 43, 43));

// Forward declarations
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
void                CreateControls(HWND hWnd);
void                StartMonitoring(HWND hWnd);
void                StopMonitoring();
DWORD WINAPI        ConnectionManagerThread(LPVOID lpParam);
void                AppendTextToEdit(HWND hEdit, const wchar_t* text);
void                SaveSettings();
void                LoadSettings();
void                PopulatePorts();

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow)
{
    InitDarkMode(); // >> DARK THEME: Initialize dark mode support at startup

    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_SERIALMONITOR, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    if (!InitInstance(hInstance, nCmdShow)) return FALSE;

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return (int)msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex = {};
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_SERIALMONITOR));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = g_hbrBackground; // >> DARK THEME: Use our dark background brush
    wcex.lpszClassName = szWindowClass;
    wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));
    return RegisterClassExW(&wcex);
}

BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    hInst = hInstance;
    HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, 0, 640, 520, nullptr, nullptr, hInstance, nullptr);

    if (!hWnd) return FALSE;

    // >> DARK THEME: Tell this specific window to use the dark title bar
    if (AllowDarkModeForWindow) {
        AllowDarkModeForWindow(hWnd, true);
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
        // >> DARK THEME: Handle messages to color the controls
    case WM_CTLCOLORSTATIC:
    {
        HDC hdcStatic = (HDC)wParam;
        SetTextColor(hdcStatic, RGB(255, 255, 255));
        SetBkColor(hdcStatic, RGB(32, 32, 32));
        return (INT_PTR)g_hbrBackground;
    }
    case WM_CTLCOLOREDIT:
    {
        HDC hdcEdit = (HDC)wParam;
        SetTextColor(hdcEdit, RGB(220, 220, 220));
        SetBkColor(hdcEdit, RGB(43, 43, 43));
        return (INT_PTR)g_hbrEditBackground;
    }
    case WM_CREATE:
        CreateControls(hWnd);
        break;

        // >> NEW: Handle window resizing
    case WM_SIZE:
    {
        // Get the new width and height of the window's client area
        int newWidth = LOWORD(lParam);
        int newHeight = HIWORD(lParam);

        // Calculate the new size and position for the output edit box
        // The top is fixed at 105. It spans the new width, minus 20px for margins.
        // Its height is the new window height minus the top and bottom panels.
        int editWidth = newWidth - 20;
        int editHeight = newHeight - 105 - 45; // 105 for top panel, 45 for bottom status area
        MoveWindow(hOutputEdit, 10, 105, editWidth, editHeight, TRUE);

        // Reposition the status label and cancel button to the new bottom
        MoveWindow(hStatusLabel, 10, newHeight - 35, 450, 20, TRUE);
        MoveWindow(hCancelButton, newWidth - 150, newHeight - 40, 140, 25, TRUE);

        break;
    }
    case WM_COMMAND:
    {
        switch (LOWORD(wParam))
        {
        case IDC_START_BUTTON:   StartMonitoring(hWnd); break;
        case IDC_STOP_BUTTON:    StopMonitoring(); break;
        case IDC_CANCEL_BUTTON:  StopMonitoring(); break;
        case IDC_REFRESH_BUTTON: PopulatePorts(); break;
        case IDC_BROWSE_BUTTON:
        {
            BROWSEINFOW bi = { 0 };
            bi.lpszTitle = L"Select a folder to save logs";
            bi.hwndOwner = hWnd;
            bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
            LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
            if (pidl != NULL) {
                wchar_t path[MAX_PATH];
                if (SHGetPathFromIDListW(pidl, path)) {
                    SetWindowTextW(hLogDirEdit, path);
                }
                CoTaskMemFree(pidl);
            }
            break;
        }
        }
        return 0;
    }
    case WM_SERIAL_DATA_RECEIVED:
    {
        wchar_t* buffer = (wchar_t*)wParam;
        AppendTextToEdit(hOutputEdit, buffer);
        delete[] buffer;
        break;
    }
    case WM_UPDATE_STATUS:
    {
        wchar_t* status = (wchar_t*)wParam;
        SetWindowTextW(hStatusLabel, status);
        delete[] status;
        break;
    }
    case WM_CLOSE:
        SaveSettings();
        StopMonitoring();
        DestroyWindow(hWnd);
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

void CreateControls(HWND hWnd)
{
    // Row 1: Port selection
    CreateWindowW(L"STATIC", L"Port:", WS_CHILD | WS_VISIBLE, 10, 15, 80, 20, hWnd, NULL, hInst, NULL);
    hPortCombo = CreateWindowW(WC_COMBOBOXW, L"", CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL, 100, 10, 180, 150, hWnd, (HMENU)IDC_PORT_COMBO, hInst, NULL);
    hRefreshButton = CreateWindowW(L"BUTTON", L"Refresh", WS_CHILD | WS_VISIBLE, 290, 10, 80, 25, hWnd, (HMENU)IDC_REFRESH_BUTTON, hInst, NULL);

    // Row 2: Baud rate
    CreateWindowW(L"STATIC", L"Baud Rate:", WS_CHILD | WS_VISIBLE, 10, 45, 80, 20, hWnd, NULL, hInst, NULL);
    hBaudCombo = CreateWindowW(WC_COMBOBOXW, L"", CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL, 100, 40, 180, 200, hWnd, (HMENU)IDC_BAUD_COMBO, hInst, NULL);

    // Row 3: Log directory
    CreateWindowW(L"STATIC", L"Log Folder:", WS_CHILD | WS_VISIBLE, 10, 75, 80, 20, hWnd, NULL, hInst, NULL);
    hLogDirEdit = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 100, 70, 410, 25, hWnd, (HMENU)IDC_LOGDIR_EDIT, hInst, NULL);
    hBrowseButton = CreateWindowW(L"BUTTON", L"...", WS_CHILD | WS_VISIBLE, 520, 70, 30, 25, hWnd, (HMENU)IDC_BROWSE_BUTTON, hInst, NULL);

    // Main control buttons
    hStartButton = CreateWindowW(L"BUTTON", L"Start", WS_CHILD | WS_VISIBLE, 400, 10, 110, 25, hWnd, (HMENU)IDC_START_BUTTON, hInst, NULL);
    hStopButton = CreateWindowW(L"BUTTON", L"Stop", WS_CHILD | WS_VISIBLE, 400, 40, 110, 25, hWnd, (HMENU)IDC_STOP_BUTTON, hInst, NULL);

    // Output area
    hOutputEdit = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | ES_MULTILINE | ES_AUTOVSCROLL | ES_READONLY, 10, 105, 605, 330, hWnd, (HMENU)IDC_OUTPUT_EDIT, hInst, NULL);

    // Status bar area
    hStatusLabel = CreateWindowW(L"STATIC", L"Ready.", WS_CHILD | WS_VISIBLE, 10, 445, 450, 20, hWnd, (HMENU)IDC_STATUS_LABEL, hInst, NULL);
    hCancelButton = CreateWindowW(L"BUTTON", L"Cancel Reconnect", WS_CHILD, 470, 440, 140, 25, hWnd, (HMENU)IDC_CANCEL_BUTTON, hInst, NULL);

    // >> DARK THEME: Tell the standard controls to adopt the modern, dark-theme-aware style
    SetWindowTheme(hPortCombo, L"Explorer", NULL);
    SetWindowTheme(hBaudCombo, L"Explorer", NULL);
    SetWindowTheme(hRefreshButton, L"Explorer", NULL);
    SetWindowTheme(hStartButton, L"Explorer", NULL);
    SetWindowTheme(hStopButton, L"Explorer", NULL);
    SetWindowTheme(hBrowseButton, L"Explorer", NULL);
    SetWindowTheme(hLogDirEdit, L"Explorer", NULL);
    SetWindowTheme(hCancelButton, L"Explorer", NULL);

    // Set initial states
    EnableWindow(hStopButton, FALSE);
    std::vector<std::string> bauds = { "9600", "57600", "115200", "250000", "921600" };
    for (const auto& r : bauds) SendMessageA(hBaudCombo, CB_ADDSTRING, 0, (LPARAM)r.c_str());

    PopulatePorts();
    LoadSettings();
}

void PopulatePorts()
{
    SendMessageW(hPortCombo, CB_RESETCONTENT, 0, 0); // Clear existing items

    wchar_t targetPath[255];
    wchar_t comName[32];

    for (int i = 1; i < 256; i++) {
        wsprintfW(comName, L"COM%d", i);
        if (QueryDosDeviceW(comName, targetPath, 255) != 0) {
            SendMessageW(hPortCombo, CB_ADDSTRING, 0, (LPARAM)comName);
        }
    }
    SendMessageW(hPortCombo, CB_SETCURSEL, 0, 0); // Select the first available port
}

void StartMonitoring(HWND hWnd)
{
    if (bShouldBeMonitoring) return;

    EnableWindow(hStartButton, FALSE);
    EnableWindow(hStopButton, TRUE);
    ShowWindow(hCancelButton, SW_HIDE);

    bShouldBeMonitoring = true;
    hThread = CreateThread(NULL, 0, ConnectionManagerThread, hWnd, 0, NULL);
}

void StopMonitoring()
{
    if (bShouldBeMonitoring) {
        bShouldBeMonitoring = false;
        if (hThread != NULL) {
            WaitForSingleObject(hThread, 2000);
            CloseHandle(hThread);
            hThread = NULL;
        }
        SetWindowTextW(hStatusLabel, L"Stopped.");
    }
    EnableWindow(hStartButton, TRUE);
    EnableWindow(hStopButton, FALSE);
    ShowWindow(hCancelButton, SW_HIDE);
}

DWORD WINAPI ConnectionManagerThread(LPVOID lpParam)
{
    HWND hWnd = (HWND)lpParam;
    HANDLE hSerial = INVALID_HANDLE_VALUE;
    HANDLE hLogFile = INVALID_HANDLE_VALUE;

    wchar_t portW[32], baudW[16], logDirW[MAX_PATH];
    GetWindowTextW(hPortCombo, portW, 32);
    GetWindowTextW(hBaudCombo, baudW, 16);
    GetWindowTextW(hLogDirEdit, logDirW, MAX_PATH);
    std::wstring fullPortName = L"\\\\.\\" + std::wstring(portW);

    std::string dataBuffer;
    std::wstring statusBuffer;
    DWORD lastUpdateTime = GetTickCount();

    while (bShouldBeMonitoring) {
        hSerial = CreateFileW(fullPortName.c_str(), GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

        if (hSerial == INVALID_HANDLE_VALUE) {
            statusBuffer = L"Failed to open " + std::wstring(portW) + L". Retrying in 5s...";
            ShowWindow(hCancelButton, SW_SHOW);
            EnableWindow(hStopButton, FALSE);

            for (int i = 0; i < 50 && bShouldBeMonitoring; ++i) Sleep(100);
            continue;
        }

        DCB dcb = { 0 };
        dcb.DCBlength = sizeof(dcb);
        if (GetCommState(hSerial, &dcb)) {
            dcb.BaudRate = _wtoi(baudW);
            dcb.ByteSize = 8;
            dcb.Parity = NOPARITY;
            dcb.StopBits = ONESTOPBIT;
            SetCommState(hSerial, &dcb);
        }

        COMMTIMEOUTS timeouts = { 0 };
        timeouts.ReadIntervalTimeout = 100;
        SetCommTimeouts(hSerial, &timeouts);

        statusBuffer = L"✅ Connected to " + std::wstring(portW);
        ShowWindow(hCancelButton, SW_HIDE);
        EnableWindow(hStopButton, TRUE);

        SYSTEMTIME st;
        GetLocalTime(&st);
        wchar_t logFileName[MAX_PATH];
        wsprintfW(logFileName, L"%s\\log_%s_%04d-%02d-%02d_%02d-%02d-%02d.txt", logDirW, portW, st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
        hLogFile = CreateFileW(logFileName, FILE_APPEND_DATA, FILE_SHARE_READ, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

        char readBuf[512];
        DWORD bytesRead;
        while (bShouldBeMonitoring) {
            if (ReadFile(hSerial, readBuf, sizeof(readBuf), &bytesRead, NULL)) {
                if (bytesRead > 0) {
                    dataBuffer.append(readBuf, bytesRead);
                    if (hLogFile != INVALID_HANDLE_VALUE) {
                        DWORD bytesWritten;
                        WriteFile(hLogFile, readBuf, bytesRead, &bytesWritten, NULL);
                    }
                }
            }
            else {
                CloseHandle(hSerial);
                hSerial = INVALID_HANDLE_VALUE;
                if (hLogFile != INVALID_HANDLE_VALUE) {
                    CloseHandle(hLogFile);
                    hLogFile = INVALID_HANDLE_VALUE;
                }
                break;
            }

            if (GetTickCount() - lastUpdateTime > 100) {
                if (!statusBuffer.empty()) {
                    wchar_t* statusMsg = new wchar_t[statusBuffer.length() + 1];
                    wcscpy_s(statusMsg, statusBuffer.length() + 1, statusBuffer.c_str());
                    PostMessageW(hWnd, WM_UPDATE_STATUS, (WPARAM)statusMsg, 0);
                    statusBuffer.clear();
                }

                if (!dataBuffer.empty()) {
                    wchar_t* wideBuf = new wchar_t[dataBuffer.length() + 1];
                    MultiByteToWideChar(CP_UTF8, 0, dataBuffer.c_str(), -1, wideBuf, dataBuffer.length() + 1);
                    PostMessageW(hWnd, WM_SERIAL_DATA_RECEIVED, (WPARAM)wideBuf, 0);
                    dataBuffer.clear();

                    if (hLogFile != INVALID_HANDLE_VALUE) {
                        FlushFileBuffers(hLogFile);
                    }
                }
                lastUpdateTime = GetTickCount();
            }
        }
    }

    if (hSerial != INVALID_HANDLE_VALUE) CloseHandle(hSerial);
    if (hLogFile != INVALID_HANDLE_VALUE) CloseHandle(hLogFile);
    return 0;
}

void SaveSettings()
{
    HKEY hKey;
    RegCreateKeyExW(HKEY_CURRENT_USER, L"Software\\CppSerialMonitor", 0, NULL, 0, KEY_WRITE, NULL, &hKey, NULL);

    wchar_t buffer[MAX_PATH];
    GetWindowTextW(hPortCombo, buffer, 32);
    RegSetValueExW(hKey, L"LastPort", 0, REG_SZ, (BYTE*)buffer, (wcslen(buffer) + 1) * sizeof(wchar_t));

    GetWindowTextW(hBaudCombo, buffer, 16);
    RegSetValueExW(hKey, L"LastBaud", 0, REG_SZ, (BYTE*)buffer, (wcslen(buffer) + 1) * sizeof(wchar_t));

    GetWindowTextW(hLogDirEdit, buffer, MAX_PATH);
    RegSetValueExW(hKey, L"LastLogDir", 0, REG_SZ, (BYTE*)buffer, (wcslen(buffer) + 1) * sizeof(wchar_t));

    RegCloseKey(hKey);
}

void LoadSettings()
{
    HKEY hKey;
    if (RegOpenKeyExW(HKEY_CURRENT_USER, L"Software\\CppSerialMonitor", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        wchar_t buffer[MAX_PATH];
        DWORD bufferSize = sizeof(buffer);
        if (RegQueryValueExW(hKey, L"LastPort", NULL, NULL, (LPBYTE)buffer, &bufferSize) == ERROR_SUCCESS) {
            SendMessageW(hPortCombo, CB_SELECTSTRING, -1, (LPARAM)buffer);
        }

        bufferSize = sizeof(buffer);
        if (RegQueryValueExW(hKey, L"LastBaud", NULL, NULL, (LPBYTE)buffer, &bufferSize) == ERROR_SUCCESS) {
            SendMessageW(hBaudCombo, CB_SELECTSTRING, -1, (LPARAM)buffer);
        }
        else {
            SendMessageW(hBaudCombo, CB_SELECTSTRING, -1, (LPARAM)L"250000");
        }

        bufferSize = sizeof(buffer);
        if (RegQueryValueExW(hKey, L"LastLogDir", NULL, NULL, (LPBYTE)buffer, &bufferSize) == ERROR_SUCCESS) {
            SetWindowTextW(hLogDirEdit, buffer);
        }
        else {
            SHGetFolderPathW(NULL, CSIDL_MYDOCUMENTS, NULL, 0, buffer);
            SetWindowTextW(hLogDirEdit, buffer);
        }

        RegCloseKey(hKey);
    }
}

void AppendTextToEdit(HWND hEdit, const wchar_t* newText)
{
    const int MAX_LINES = 5000;

    int lineCount = SendMessage(hEdit, EM_GETLINECOUNT, 0, 0);
    if (lineCount > MAX_LINES) {
        int firstLineLen = SendMessage(hEdit, EM_LINELENGTH, 0, 0);
        SendMessage(hEdit, EM_SETSEL, 0, firstLineLen + 2);
        SendMessage(hEdit, EM_REPLACESEL, 0, (LPARAM)L"");
    }

    int len = GetWindowTextLengthW(hEdit);
    SendMessageW(hEdit, EM_SETSEL, (WPARAM)len, (LPARAM)len);
    SendMessageW(hEdit, EM_REPLACESEL, 0, (LPARAM)newText);
}