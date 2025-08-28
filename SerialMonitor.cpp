#include "framework.h"
#include "SerialMonitor.h"
#include "darktheme.h" 
#include <windows.h>
#include <string>
#include <vector>
#include <CommCtrl.h> 
#include <ShlObj.h>   

#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Shell32.lib")

#define MAX_LOADSTRING 100

// Message Definitions
#define WM_SERIAL_DATA_RECEIVED (WM_APP + 1)
#define WM_UPDATE_STATUS        (WM_APP + 2)
#define WM_GUI_STATE_CONNECTING (WM_APP + 3)
#define WM_GUI_STATE_CONNECTED  (WM_APP + 4)
#define WM_CONNECTION_LOST      (WM_APP + 5)

// Timer ID
#define IDT_RECONNECT_TIMER 1

// Struct to pass timestamped log data
struct LogEntry {
    std::wstring timestamp;
    std::wstring message;
};

// Global Variables
HINSTANCE hInst;
WCHAR szTitle[MAX_LOADSTRING];
WCHAR szWindowClass[MAX_LOADSTRING];
HWND hPortCombo, hBaudCombo, hStartButton, hStopButton, hOutputListView, hRefreshButton;
HWND hLogDirEdit, hBrowseButton, hStatusLabel, hCancelButton;
HWND hClearButton;
HANDLE hThread = NULL;
volatile bool bShouldBeMonitoring = false;
HBRUSH g_hbrBackground = CreateSolidBrush(RGB(32, 32, 32));
HBRUSH g_hbrEditBackground = CreateSolidBrush(RGB(43, 43, 43));

// Forward Declarations
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
void                CreateControls(HWND hWnd);
void                StartMonitoring(HWND hWnd);
void                StopMonitoring();
DWORD WINAPI        SerialThread(LPVOID lpParam);
void                AddLogEntry(const LogEntry* entry);
void                SaveSettings();
void                LoadSettings();
void                PopulatePorts();

int APIENTRY wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nCmdShow)
{
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icex);

    InitDarkMode();
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_SERIALMONITOR, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);
    if (!InitInstance(hInstance, nCmdShow)) return FALSE;
    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
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
    wcex.hbrBackground = g_hbrBackground;
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
    if (AllowDarkModeForWindow) AllowDarkModeForWindow(hWnd, true);
    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);
    return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CTLCOLORSTATIC: {
        HDC hdcStatic = (HDC)wParam;
        SetTextColor(hdcStatic, RGB(255, 255, 255));
        SetBkColor(hdcStatic, RGB(32, 32, 32));
        return (INT_PTR)g_hbrBackground;
    }
    case WM_CTLCOLOREDIT: {
        HDC hdcEdit = (HDC)wParam;
        SetTextColor(hdcEdit, RGB(220, 220, 220));
        SetBkColor(hdcEdit, RGB(43, 43, 43));
        return (INT_PTR)g_hbrEditBackground;
    }
    case WM_GUI_STATE_CONNECTING: {
        EnableWindow(hStartButton, FALSE);
        EnableWindow(hStopButton, FALSE);
        ShowWindow(hCancelButton, SW_SHOW);
        break;
    }
    case WM_GUI_STATE_CONNECTED: {
        EnableWindow(hStartButton, FALSE);
        EnableWindow(hStopButton, TRUE);
        ShowWindow(hCancelButton, SW_HIDE);
        break;
    }
    case WM_SERIAL_DATA_RECEIVED: {
        LogEntry* entry = (LogEntry*)wParam;
        AddLogEntry(entry);
        delete entry;
        break;
    }
    case WM_UPDATE_STATUS: {
        wchar_t* status = (wchar_t*)wParam;
        SetWindowTextW(hStatusLabel, status);
        delete[] status;
        break;
    }
    case WM_CONNECTION_LOST: {
        // >> FIX: Clean up the handle of the thread that just ended.
        if (hThread != NULL) {
            CloseHandle(hThread);
            hThread = NULL;
        }

        SetWindowTextW(hStatusLabel, L"Connection lost. Retrying in 5s...");
        SetTimer(hWnd, IDT_RECONNECT_TIMER, 5000, NULL);
        EnableWindow(hStartButton, FALSE);
        EnableWindow(hStopButton, FALSE);
        ShowWindow(hCancelButton, SW_SHOW);
        break;
    }
    case WM_TIMER: {
        if (wParam == IDT_RECONNECT_TIMER) {
            KillTimer(hWnd, IDT_RECONNECT_TIMER);
            StartMonitoring(hWnd);
        }
        break;
    }
    case WM_CREATE: CreateControls(hWnd); break;
    case WM_SIZE: {
        int newWidth = LOWORD(lParam);
        int newHeight = HIWORD(lParam);
        MoveWindow(hOutputListView, 10, 105, newWidth - 20, newHeight - 150, TRUE);
        ListView_SetColumnWidth(hOutputListView, 1, newWidth - 125);
        MoveWindow(hStatusLabel, 10, newHeight - 35, newWidth - 170, 20, TRUE);
        MoveWindow(hCancelButton, newWidth - 150, newHeight - 40, 140, 25, TRUE);
        break;
    }
    case WM_COMMAND: {
        switch (LOWORD(wParam)) {
        case IDC_START_BUTTON:   StartMonitoring(hWnd); break;
        case IDC_STOP_BUTTON:    StopMonitoring(); break;
        case IDC_CANCEL_BUTTON:  StopMonitoring(); break;
        case IDC_REFRESH_BUTTON: PopulatePorts(); break;
        case IDC_CLEAR_BUTTON:
            ListView_DeleteAllItems(hOutputListView);
            break;
        case IDC_BROWSE_BUTTON: {
            BROWSEINFOW bi = { 0 };
            bi.lpszTitle = L"Select a folder to save logs";
            bi.hwndOwner = hWnd;
            bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;
            LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
            if (pidl != NULL) {
                wchar_t path[MAX_PATH];
                if (SHGetPathFromIDListW(pidl, path)) SetWindowTextW(hLogDirEdit, path);
                CoTaskMemFree(pidl);
            }
            break;
        }
        default: return DefWindowProc(hWnd, message, wParam, lParam);
        }
        break;
    }
    case WM_CLOSE:
        SaveSettings();
        StopMonitoring();
        DestroyWindow(hWnd);
        break;
    case WM_DESTROY: PostQuitMessage(0); break;
    default: return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}


void CreateControls(HWND hWnd)
{
    CreateWindowW(L"STATIC", L"Port:", WS_CHILD | WS_VISIBLE, 10, 15, 80, 20, hWnd, NULL, hInst, NULL);
    hPortCombo = CreateWindowW(WC_COMBOBOXW, L"", CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL, 100, 10, 180, 150, hWnd, (HMENU)IDC_PORT_COMBO, hInst, NULL);
    hRefreshButton = CreateWindowW(L"BUTTON", L"Refresh", WS_CHILD | WS_VISIBLE, 290, 10, 80, 25, hWnd, (HMENU)IDC_REFRESH_BUTTON, hInst, NULL);
    CreateWindowW(L"STATIC", L"Baud Rate:", WS_CHILD | WS_VISIBLE, 10, 45, 80, 20, hWnd, NULL, hInst, NULL);
    hBaudCombo = CreateWindowW(WC_COMBOBOXW, L"", CBS_DROPDOWNLIST | WS_CHILD | WS_VISIBLE | WS_VSCROLL, 100, 40, 180, 200, hWnd, (HMENU)IDC_BAUD_COMBO, hInst, NULL);
    CreateWindowW(L"STATIC", L"Log Folder:", WS_CHILD | WS_VISIBLE, 10, 75, 80, 20, hWnd, NULL, hInst, NULL);
    hLogDirEdit = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL, 100, 70, 410, 25, hWnd, (HMENU)IDC_LOGDIR_EDIT, hInst, NULL);
    hBrowseButton = CreateWindowW(L"BUTTON", L"...", WS_CHILD | WS_VISIBLE, 520, 70, 30, 25, hWnd, (HMENU)IDC_BROWSE_BUTTON, hInst, NULL);
    hStartButton = CreateWindowW(L"BUTTON", L"Start", WS_CHILD | WS_VISIBLE, 400, 10, 110, 25, hWnd, (HMENU)IDC_START_BUTTON, hInst, NULL);
    hStopButton = CreateWindowW(L"BUTTON", L"Stop", WS_CHILD | WS_VISIBLE, 400, 40, 110, 25, hWnd, (HMENU)IDC_STOP_BUTTON, hInst, NULL);
    hClearButton = CreateWindowW(L"BUTTON", L"Clear Output", WS_CHILD | WS_VISIBLE, 520, 10, 95, 55, hWnd, (HMENU)IDC_CLEAR_BUTTON, hInst, NULL);

    hOutputListView = CreateWindowExW(0, WC_LISTVIEWW, L"", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT, 10, 105, 605, 330, hWnd, (HMENU)IDC_OUTPUT_EDIT, hInst, NULL);
    LVCOLUMNW lvc = { 0 };
    lvc.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_SUBITEM;
    lvc.cx = 100;
    lvc.pszText = (LPWSTR)L"Time";
    ListView_InsertColumn(hOutputListView, 0, &lvc);
    lvc.cx = 500;
    lvc.pszText = (LPWSTR)L"Message";
    ListView_InsertColumn(hOutputListView, 1, &lvc);
    ListView_SetExtendedListViewStyle(hOutputListView, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER);

    hStatusLabel = CreateWindowW(L"STATIC", L"Ready.", WS_CHILD | WS_VISIBLE, 10, 445, 450, 20, hWnd, (HMENU)IDC_STATUS_LABEL, hInst, NULL);
    hCancelButton = CreateWindowW(L"BUTTON", L"Cancel Reconnect", WS_CHILD, 470, 440, 140, 25, hWnd, (HMENU)IDC_CANCEL_BUTTON, hInst, NULL);

    SetWindowTheme(hPortCombo, L"Explorer", NULL);
    SetWindowTheme(hBaudCombo, L"Explorer", NULL);
    SetWindowTheme(hRefreshButton, L"Explorer", NULL);
    SetWindowTheme(hStartButton, L"Explorer", NULL);
    SetWindowTheme(hStopButton, L"Explorer", NULL);
    SetWindowTheme(hBrowseButton, L"Explorer", NULL);
    SetWindowTheme(hLogDirEdit, L"Explorer", NULL);
    SetWindowTheme(hCancelButton, L"Explorer", NULL);
    SetWindowTheme(hOutputListView, L"Explorer", NULL);
    SetWindowTheme(hClearButton, L"Explorer", NULL);

    EnableWindow(hStopButton, FALSE);
    ShowWindow(hCancelButton, SW_HIDE);
    std::vector<std::string> bauds = { "9600", "57600", "115200", "250000", "333333", "444444","555555", "888888", "921600" };
    for (const auto& r : bauds) SendMessageA(hBaudCombo, CB_ADDSTRING, 0, (LPARAM)r.c_str());

    PopulatePorts();
    LoadSettings();
}

void PopulatePorts()
{
    SendMessageW(hPortCombo, CB_RESETCONTENT, 0, 0);
    wchar_t targetPath[255];
    wchar_t comName[32];
    for (int i = 1; i < 256; i++) {
        wsprintfW(comName, L"COM%d", i);
        if (QueryDosDeviceW(comName, targetPath, 255) != 0) {
            SendMessageW(hPortCombo, CB_ADDSTRING, 0, (LPARAM)comName);
        }
    }
    SendMessageW(hPortCombo, CB_SETCURSEL, 0, 0);
}

// Replace your existing StartMonitoring function with this one
void StartMonitoring(HWND hWnd)
{
    if (hThread != NULL) return;

    KillTimer(hWnd, IDT_RECONNECT_TIMER);

    bShouldBeMonitoring = true;
    PostMessage(hWnd, WM_GUI_STATE_CONNECTING, 0, 0);

    hThread = CreateThread(NULL, 0, SerialThread, hWnd, 0, NULL);
}

void StopMonitoring()
{
    HWND hWnd = GetParent(hStartButton);
    KillTimer(hWnd, IDT_RECONNECT_TIMER);
    if (bShouldBeMonitoring) {
        bShouldBeMonitoring = false;
        if (hThread != NULL) {
            WaitForSingleObject(hThread, 2000);
            CloseHandle(hThread);
            hThread = NULL;
        }
    }
    SetWindowTextW(hStatusLabel, L"Stopped.");
    EnableWindow(hStartButton, TRUE);
    EnableWindow(hStopButton, FALSE);
    ShowWindow(hCancelButton, SW_HIDE);
}
// Replace your existing SerialThread function with this one
DWORD WINAPI SerialThread(LPVOID lpParam)
{
    HWND hWnd = (HWND)lpParam;
    HANDLE hSerial = INVALID_HANDLE_VALUE;
    HANDLE hLogFile = INVALID_HANDLE_VALUE;

    wchar_t portW[32], baudW[16], logDirW[MAX_PATH];
    GetWindowTextW(hPortCombo, portW, 32);
    GetWindowTextW(hBaudCombo, baudW, 16);
    GetWindowTextW(hLogDirEdit, logDirW, MAX_PATH);
    std::wstring fullPortName = L"\\\\.\\" + std::wstring(portW);

    wchar_t status[128];
    wsprintfW(status, L"Connecting to %s...", portW);
    wchar_t* statusMsg = new wchar_t[128];
    wcscpy_s(statusMsg, 128, status);
    PostMessageW(hWnd, WM_UPDATE_STATUS, (WPARAM)statusMsg, 0);

    hSerial = CreateFileW(fullPortName.c_str(), GENERIC_READ | GENERIC_WRITE, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

    if (hSerial == INVALID_HANDLE_VALUE) {
        PostMessage(hWnd, WM_CONNECTION_LOST, 0, 0);
        return 1;
    }

    DCB dcb = { 0 };
    dcb.DCBlength = sizeof(dcb);
    if (GetCommState(hSerial, &dcb)) {
        dcb.BaudRate = _wtoi(baudW);
        dcb.ByteSize = 8;
        dcb.Parity = NOPARITY;
        dcb.StopBits = ONESTOPBIT;
        dcb.fDtrControl = DTR_CONTROL_ENABLE;
        SetCommState(hSerial, &dcb);
    }
    COMMTIMEOUTS timeouts = { 0 };
    timeouts.ReadIntervalTimeout = 100;
    SetCommTimeouts(hSerial, &timeouts);
    EscapeCommFunction(hSerial, CLRDTR); Sleep(100);
    EscapeCommFunction(hSerial, SETDTR); Sleep(500);

    PostMessage(hWnd, WM_GUI_STATE_CONNECTED, 0, 0);
    wsprintfW(status, L"✅ Connected to %s", portW);
    statusMsg = new wchar_t[128];
    wcscpy_s(statusMsg, 128, status);
    PostMessageW(hWnd, WM_UPDATE_STATUS, (WPARAM)statusMsg, 0);

    SYSTEMTIME st;
    GetLocalTime(&st);
    wchar_t logFileName[MAX_PATH];
    wsprintfW(logFileName, L"%s\\log_%s_%04d-%02d-%02d_%02d-%02d-%02d.txt", logDirW, portW, st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    hLogFile = CreateFileW(logFileName, FILE_APPEND_DATA, FILE_SHARE_READ, NULL, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

    std::string dataBuffer;
    ULONGLONG lastUpdateTime = GetTickCount64();
    char readBuf[512];
    DWORD bytesRead;

    while (bShouldBeMonitoring) {
        if (ReadFile(hSerial, readBuf, sizeof(readBuf), &bytesRead, NULL)) {
            if (bytesRead > 0) {
                dataBuffer.append(readBuf, bytesRead);
                if (hLogFile != INVALID_HANDLE_VALUE) WriteFile(hLogFile, readBuf, bytesRead, &bytesRead, NULL);
            }
        }
        else {
            PostMessage(hWnd, WM_CONNECTION_LOST, 0, 0);
            break;
        }

        if (GetTickCount64() - lastUpdateTime > 100) {

            // >> THE FIX: Process the buffer line by line
            size_t newline_pos;
            while ((newline_pos = dataBuffer.find('\n')) != std::string::npos)
            {
                // Extract one full message (line) from the buffer
                std::string message = dataBuffer.substr(0, newline_pos + 1);
                // Remove that message from the buffer
                dataBuffer.erase(0, newline_pos + 1);

                // Create a LogEntry for this single message
                LogEntry* entry = new LogEntry();

                // Generate timestamp
                SYSTEMTIME st_now;
                GetLocalTime(&st_now);
                wchar_t timeBuf[32];
                wsprintfW(timeBuf, L"%02d:%02d:%02d.%03d", st_now.wHour, st_now.wMinute, st_now.wSecond, st_now.wMilliseconds);
                entry->timestamp = timeBuf;

                // Convert the single message to wide string (the newline normalization is now in AppendTextToEdit)
                wchar_t* wideBuf = new wchar_t[message.length() + 1];
                MultiByteToWideChar(CP_UTF8, 0, message.c_str(), -1, wideBuf, message.length() + 1);
                entry->message = wideBuf;
                delete[] wideBuf;

                // Post the message for this single line
                PostMessageW(hWnd, WM_SERIAL_DATA_RECEIVED, (WPARAM)entry, 0);
            }

            if (hLogFile != INVALID_HANDLE_VALUE) FlushFileBuffers(hLogFile);
            lastUpdateTime = GetTickCount64();
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
    RegSetValueExW(hKey, L"LastPort", 0, REG_SZ, (BYTE*)buffer, static_cast<DWORD>((wcslen(buffer) + 1) * sizeof(wchar_t)));
    GetWindowTextW(hBaudCombo, buffer, 16);
    RegSetValueExW(hKey, L"LastBaud", 0, REG_SZ, (BYTE*)buffer, static_cast<DWORD>((wcslen(buffer) + 1) * sizeof(wchar_t)));
    GetWindowTextW(hLogDirEdit, buffer, MAX_PATH);
    RegSetValueExW(hKey, L"LastLogDir", 0, REG_SZ, (BYTE*)buffer, static_cast<DWORD>((wcslen(buffer) + 1) * sizeof(wchar_t)));
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

void AddLogEntry(const LogEntry* entry)
{
    const int MAX_ITEMS = 5000;
    int itemCount = ListView_GetItemCount(hOutputListView);
    if (itemCount > MAX_ITEMS) {
        ListView_DeleteItem(hOutputListView, 0);
    }

    itemCount = ListView_GetItemCount(hOutputListView);
    LVITEMW lvi = { 0 };
    lvi.mask = LVIF_TEXT;
    lvi.iItem = itemCount;
    lvi.iSubItem = 0;
    lvi.pszText = (LPWSTR)entry->timestamp.c_str();
    int newIndex = ListView_InsertItem(hOutputListView, &lvi);

    ListView_SetItemText(hOutputListView, newIndex, 1, (LPWSTR)entry->message.c_str());
    ListView_EnsureVisible(hOutputListView, newIndex, FALSE);
}