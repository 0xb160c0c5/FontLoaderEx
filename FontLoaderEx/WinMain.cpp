#if !defined(UNICODE) || !defined(_UNICODE)
#error Unicode character set required
#endif // UNICODE && _UNICODE

#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "ComCtl32.lib")
#pragma comment(lib, "Shlwapi.lib")

#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <shlwapi.h>
#include <tlhelp32.h>
#include <process.h>
#include <sddl.h>
#include <versionhelpers.h>
#include <cwchar>
#include <cwctype>
#include <string>
#include <iomanip>
#include <sstream>
#include <list>
#include <vector>
#include <algorithm>
#include <memory>
#include "FontResource.h"
#include "Globals.h"
#include "Splitter.h"
#include "resource.h"

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);

const WCHAR szAppName[]{ L"FontLoaderEx" };

std::list<FontResource> FontList{};

HWND hWndMain{};
HMENU hMenuContextListViewFontList{};
HMENU hMenuContextTray{};
HICON hIconApplication{};

bool bDragDropHasFonts{ false };

// Create an unique string by scope
enum class Scope { Machine, User, Session, WindowStation, Desktop };
std::wstring GetUniqueName(LPCWSTR lpszString, Scope scope);

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
	// Check Windows version
	if (!IsWindows7OrGreater())
	{
		MessageBox(NULL, L"Windows 7 or higher required.", szAppName, MB_ICONWARNING);

		return 0;
	}

	// Prevent multiple instances of FontLoaderEx in the same session
	Scope scope{ Scope::Session };
	std::wstring strMutexName{ GetUniqueName(L"FontLoaderEx-656A8394-5AB8-4061-8882-2FE2E7940C2E", scope) };
	HANDLE hMutexSingleton{ CreateMutex(NULL, FALSE, strMutexName.c_str()) };
	if (!hMutexSingleton)
	{
		MessageBox(NULL, L"Failed to create singleton mutex.", szAppName, MB_ICONERROR);

		return -1;
	}
	else
	{
		if (GetLastError() == ERROR_ALREADY_EXISTS || GetLastError() == ERROR_ACCESS_DENIED)
		{
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			ssMessage << L"An instance of " << szAppName << L" is already running ";
			switch (scope)
			{
			case Scope::Machine:
				{
					ssMessage << L"on the same machine.";
				}
				break;
			case Scope::User:
				{
					ssMessage << L"by the same user.";
				}
				break;
			case Scope::Session:
				{
					ssMessage << L"in the same session.";
				}
				break;
			case Scope::WindowStation:
				{
					ssMessage << L"in the same window station.";
				}
				break;
			case Scope::Desktop:
				{
					ssMessage << L"on the same desktop.";
				}
				break;
			default:
				break;
			}
			strMessage = ssMessage.str();
			MessageBox(NULL, strMessage.c_str(), szAppName, MB_ICONWARNING);

			return 0;
		}
	}

	// Register global AddFont() and RemoveFont() procedure
	FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

	// Process drag-drop font files onto the application icon stage I
	int argc{};
	LPWSTR* argv = CommandLineToArgvW(GetCommandLine(), &argc);
	if (argc > 1)
	{
		bDragDropHasFonts = true;

		for (int i = 1; i < argc; i++)
		{
			if (PathMatchSpec(argv[i], L"*.ttf") || PathMatchSpec(argv[i], L"*.ttc") || PathMatchSpec(argv[i], L"*.otf"))
			{
				FontList.push_back(argv[i]);
			}
		}
	}
	LocalFree(argv);

	// Initialize common controls and user controls
	InitCommonControls();
	InitSplitter();

	// Get HICON to application
	hIconApplication = LoadIcon(NULL, IDI_APPLICATION);

	// Create window
	WNDCLASS wc{ 0, WndProc, 0, 0, hInstance, hIconApplication, LoadCursor(NULL, IDC_ARROW), GetSysColorBrush(COLOR_WINDOW), NULL, szAppName };
	if (!RegisterClass(&wc))
	{
		return -1;
	}
	if (!(hWndMain = CreateWindow(szAppName, szAppName, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 700, 700, NULL, NULL, hInstance, NULL)))
	{
		return -1;
	}
	ShowWindow(hWndMain, nShowCmd);
	UpdateWindow(hWndMain);

	// Get HMENU to context menus
	hMenuContextListViewFontList = GetSubMenu(LoadMenu(hInstance, MAKEINTRESOURCE(IDR_CONTEXTMENU1)), 0);
	hMenuContextTray = GetSubMenu(LoadMenu(hInstance, MAKEINTRESOURCE(IDR_CONTEXTMENU1)), 1);

	MSG Message{};
	BOOL bRet{};
	while ((bRet = GetMessage(&Message, NULL, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			CloseHandle(hMutexSingleton);

			return (int)GetLastError();
		}
		else
		{
			if (!IsDialogMessage(hWndMain, &Message))
			{
				TranslateMessage(&Message);
				DispatchMessage(&Message);
			}
		}
	}

	CloseHandle(hMutexSingleton);

	return (int)Message.wParam;
}

LRESULT CALLBACK EditTimeoutSubclassProc(HWND hWndEditTimeout, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
LRESULT CALLBACK ListViewFontListSubclassProc(HWND hWndListViewFontList, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
LRESULT CALLBACK EditMessageSubclassProc(HWND hWndEditMessage, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT Message, WPARAM wParam, LPARAM IParam);

void* lpRemoteAddFontProcAddr{};
void* lpRemoteRemoveFontProcAddr{};

HANDLE hThreadWatch{};
HANDLE hThreadMessage{};

HANDLE hProcessCurrentDuplicated{};
HANDLE hProcessTargetDuplicated{};

HANDLE hEventParentProcessRunning{};
HANDLE hEventMessageThreadNotReady{};
HANDLE hEventMessageThreadReady{};
HANDLE hEventTerminateWatchThread{};
HANDLE hEventProxyProcessReady{};
HANDLE hEventProxyProcessDebugPrivilegeEnablingFinished{};
HANDLE hEventProxyProcessHWNDRevieved{};
HANDLE hEventProxyDllInjectionFinished{};
HANDLE hEventProxyDllPullingFinished{};

PROXYPROCESSDEBUGPRIVILEGEENABLING ProxyDebugPrivilegeEnablingResult{};
PROXYDLLINJECTION ProxyDllInjectionResult{};
PROXYDLLPULL ProxyDllPullingResult{};

ProcessInfo TargetProcessInfo{}, ProxyProcessInfo{};

int MessageBoxCentered(HWND hWnd, LPCTSTR lpText, LPCTSTR lpCaption, UINT uType);

bool EnableDebugPrivilege();
bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout);
bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout);

enum class ID : WORD { ButtonOpen = 20, ButtonClose, ButtonLoad, ButtonUnload, ButtonBroadcastWM_FONTCHANGE, StaticTimeout, EditTimeout, ButtonSelectProcess, ButtonMinimizeToTray, ListViewFontList, Splitter, EditMessage, StatusBarFontInfo };

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	static DWORD dwTimeout{};

#ifdef _WIN64
	const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll64.dll" };
	const WCHAR szInjectionDllNameByProxy[]{ L"FontLoaderExInjectionDll.dll" };
#else
	const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll.dll" };
	const WCHAR szInjectionDllNameByProxy[]{ L"FontLoaderExInjectionDll64.dll" };
#endif // _WIN64

	switch ((USERMESSAGE)Message)
	{
		// Close worker thread termination notification
		// wParam = Whether font list was modified : bool
		// LOWORD(lParam) = Whether fonts unloading is interrupted by proxy/target termination: bool
		// HIWORD(lParam) = Whether are all fonts are unloaded : bool
	case USERMESSAGE::CLOSEWORKERTHREADTERMINATED:
		{
			// If unloading is interrupted
			if (LOWORD(lParam))
			{
				// Re-enable and re-disable controls
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), TRUE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::EditTimeout), TRUE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), TRUE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), TRUE);
				EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);

				// Update StatusBarFontInfo
				std::wstringstream ssFontInfo{};
				std::wstring strFontInfo{};
				std::size_t nLoadedFonts{};
				for (const auto& i : FontList)
				{
					if (i.IsLoaded())
					{
						nLoadedFonts++;
					}
				}
				ssFontInfo << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
				strFontInfo = ssFontInfo.str();
				SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), strFontInfo.c_str());

				// Update syatem tray icon tip
				if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
				{
					NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
					wcscpy_s(nid.szTip, strFontInfo.c_str());
					Shell_NotifyIcon(NIM_MODIFY, &nid);
				}
			}
			// If unloading is not interrupted
			else
			{
				// If unloading successful, do cleanup
				if (HIWORD(lParam))
				{
					PostMessage(hWnd, WM_CLOSE, 0, 0);
				}
				else
				{
					// Else, prompt user whether inisit to exit
					switch (MessageBoxCentered(hWnd, L"Some fonts are not successfully unloaded.\r\n\r\nDo you still want to exit?", szAppName, MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON1 | MB_APPLMODAL))
					{
					case IDYES:
						{
							// Do cleanup regardless of loaded fonts
							PostMessage(hWnd, WM_CLOSE, 0, 0);
						}
						break;
					case IDNO:
						{
							// Re-enable controls
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), TRUE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), TRUE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), TRUE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), TRUE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), TRUE);
							if (!TargetProcessInfo.hProcess)
							{
								EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
							}

							// Update StatusBarFontInfo
							std::wstringstream ssFontInfo{};
							std::wstring strFontInfo{};
							std::size_t nLoadedFonts{};
							for (const auto& i : FontList)
							{
								if (i.IsLoaded())
								{
									nLoadedFonts++;
								}
							}
							ssFontInfo << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
							strFontInfo = ssFontInfo.str();
							SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), strFontInfo.c_str());

							// Update syatem tray icon tip
							if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
							{
								NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
								wcscpy_s(nid.szTip, strFontInfo.c_str());
								Shell_NotifyIcon(NIM_MODIFY, &nid);
							}
						}
						break;
					default:
						break;
					}
				}
			}

			// Broadcast WM_FONTCHANGE if ButtonBroadcastWM_FONTCHANGE is checked
			HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{};
			if (wParam)
			{
				if ((Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE)) == BST_CHECKED) && (!TargetProcessInfo.hProcess))
				{
					FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
				}
			}
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, L"\r\n");
		}
		break;
		// Button worker thread termination notification
		// wParam = Whether some fonts are loaded/unloaded : bool
	case USERMESSAGE::BUTTONCLOSEWORKERTHREADTERMINATED:
	case USERMESSAGE::DRAGDROPWORKERTHREADTERMINATED:
	case USERMESSAGE::BUTTONLOADWORKERTHREADTERMINATED:
	case USERMESSAGE::BUTTONUNLOADWORKERTHREADTERMINATED:
		{
			// Re-enable and re-disable controls
			EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), TRUE);
			EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), TRUE);
			EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
			if (FontList.empty())
			{
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
			}
			else
			{
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), TRUE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), TRUE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), TRUE);
			}
			if (!TargetProcessInfo.hProcess)
			{
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
			}
			bool bIsSomeFontsLoaded{};
			for (const auto& i : FontList)
			{
				if (i.IsLoaded())
				{
					bIsSomeFontsLoaded = true;

					break;
				}
			}
			if (!bIsSomeFontsLoaded)
			{
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), TRUE);
			}

			// Broadcast WM_FONTCHANGE if ButtonBroadcastWM_FONTCHANGE is checked
			HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{};
			if (wParam)
			{
				if ((Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE)) == BST_CHECKED) && (!TargetProcessInfo.hProcess))
				{
					FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
				}
			}
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, L"\r\n");

			// Update StatusBarFontInfo
			std::wstringstream ssFontInfo{};
			std::wstring strFontInfo{};
			std::size_t nLoadedFonts{};
			for (const auto& i : FontList)
			{
				if (i.IsLoaded())
				{
					nLoadedFonts++;
				}
			}
			ssFontInfo << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
			strFontInfo = ssFontInfo.str();
			SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), strFontInfo.c_str());

			// Update syatem tray icon tip
			if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
			{
				NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
				wcscpy_s(nid.szTip, strFontInfo.c_str());
				Shell_NotifyIcon(NIM_MODIFY, &nid);
			}
		}
		break;
		// Watch thread terminated notofication
		// LOWORD(wParam) = Which thread(TargetProcessWatchThreadProc/ProxyAndTargetProcessWatchThreadProc) : enum WATCHTHREADTERMINATED
		// HIWORD(wParam) = What terminated(Proxy/Target) : enum TERMINATION
		// lParam = Whether worker thread is still running : bool
	case USERMESSAGE::WATCHTHREADTERMINATED:
		{
			// Disable controls
			if (!lParam)
			{
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), FALSE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), FALSE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), FALSE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), FALSE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), FALSE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), FALSE);
				EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
			}

			// Clear FontList and ListViewFontList
			FontResource::RegisterAddRemoveFontProc(NullAddFontProc, NullRemoveFontProc);
			FontList.clear();
			ListView_DeleteAllItems(GetDlgItem(hWnd, (int)ID::ListViewFontList));

			HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{};
			switch ((TERMINATION)HIWORD(wParam))
			{
				// If proxy process terminates, just print message
			case TERMINATION::PROXY:
				{
					ssMessage << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") terminated.\r\n\r\n";
					strMessage = ssMessage.str();
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
				}
				break;
				// If target process termiantes, print messages and terminate proxy process
			case TERMINATION::TARGET:
				{
					ssMessage << L"Target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L") terminated.\r\n\r\n";
					strMessage = ssMessage.str();
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
					ssMessage.str(L"");

					COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
					FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
					WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
					ssMessage << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") successfully terminated.\r\n\r\n";
					strMessage = ssMessage.str();
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
				}
				break;
			default:
				break;
			}

			if ((TERMINATION)LOWORD(wParam) == TERMINATION::PROXY)
			{
				// Terminate message thread
				SendMessage(hWndMessage, WM_CLOSE, 0, 0);
				WaitForSingleObject(hThreadMessage, INFINITE);
			}

			// Register global AddFont() and RemoveFont() procedures
			FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

			// Revert the caption of ButtonSelectProcess to default
			Button_SetText(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), L"Click to select process");

			// Close HANDLE to proxy process and target process, duplicated handles and synchronization objects
			CloseHandle(TargetProcessInfo.hProcess);
			TargetProcessInfo.hProcess = NULL;
			CloseHandle(hProcessCurrentDuplicated);
			CloseHandle(hProcessTargetDuplicated);
			if ((TERMINATION)LOWORD(wParam) == TERMINATION::PROXY)
			{
				CloseHandle(ProxyProcessInfo.hProcess);
				ProxyProcessInfo.hProcess = NULL;
				CloseHandle(hEventProxyAddFontFinished);
				CloseHandle(hEventProxyRemoveFontFinished);
			}

			// Enable controls
			if (!lParam)
			{
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), TRUE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), TRUE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), TRUE);
				EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
			}
			EnableWindow(GetDlgItem(hWnd, (int)ID::EditTimeout), TRUE);
			EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);

			// Update StatusBarFontInfo
			SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), L"0 font(s) opened, 0 font(s) loaded.");

			// Update syatem tray icon tip
			if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
			{
				NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP, 0, NULL, L"0 font(s) opened, 0 font(s) loaded." };
				Shell_NotifyIcon(NIM_MODIFY, &nid);
			}
		}
		break;
		// Font list changed notofication
		// wParam = Font change event : enum FONTLISTCHANGED
		// lParam = iItem in ListViewFontList and font name : struct FONTLISTCHANGEDSTRUCT*
	case USERMESSAGE::FONTLISTCHANGED:
		{
			// Modify ListViewFontList and print messages to EditMessage
			HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
			HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{};
			LVITEM lvi{ LVIF_TEXT, ((FONTLISTCHANGEDSTRUCT*)lParam)->iItem };
			switch ((FONTLISTCHANGED)wParam)
			{
			case FONTLISTCHANGED::OPENED:
				{
					lvi.iSubItem = 0;
					lvi.pszText = (LPWSTR)((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName;
					ListView_InsertItem(hWndListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = (LPWSTR)L"Not loaded";
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_SetItemState(hWndListViewFontList, lvi.iItem, LVIS_SELECTED, LVIS_SELECTED);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << ((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName << L" opened\r\n";
				}
				break;
			case FONTLISTCHANGED::OPENED_LOADED:
				{
					lvi.iSubItem = 0;
					lvi.pszText = (LPWSTR)((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName;
					ListView_InsertItem(hWndListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = (LPWSTR)L"Loaded";
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_SetItemState(hWndListViewFontList, lvi.iItem, LVIS_SELECTED, LVIS_SELECTED);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << ((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName << L" opened and successfully loaded\r\n";
				}
				break;
			case FONTLISTCHANGED::OPENED_NOTLOADED:
				{
					lvi.iSubItem = 0;
					lvi.pszText = (LPWSTR)((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName;
					ListView_InsertItem(hWndListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = (LPWSTR)L"Load failed";
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_SetItemState(hWndListViewFontList, lvi.iItem, LVIS_SELECTED, LVIS_SELECTED);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << L"Opened but failed to load " << ((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName << L"\r\n";
				}
				break;
			case FONTLISTCHANGED::LOADED:
				{
					lvi.iSubItem = 1;
					lvi.pszText = (LPWSTR)L"Loaded";
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << ((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName << L" successfully loaded\r\n";
				}
				break;
			case FONTLISTCHANGED::NOTLOADED:
				{
					lvi.iSubItem = 1;
					lvi.pszText = (LPWSTR)L"Load failed";
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << L"Failed to load " << ((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName << L"\r\n";
				}
				break;
			case FONTLISTCHANGED::UNLOADED:
				{
					lvi.iSubItem = 1;
					lvi.pszText = (LPWSTR)L"Unloaded";
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << ((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName << L" successfully unloaded\r\n";
				}
				break;
			case FONTLISTCHANGED::NOTUNLOADED:
				{
					lvi.iSubItem = 1;
					lvi.pszText = (LPWSTR)L"Unload failed";
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << L"Failed to unload " << ((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName << L"\r\n";
				}
				break;
			case FONTLISTCHANGED::UNLOADED_CLOSED:
				{
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);
					ListView_DeleteItem(hWndListViewFontList, lvi.iItem);

					ssMessage << ((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName << L" successfully unloaded and closed\r\n";
				}
				break;
			case FONTLISTCHANGED::CLOSED:
				{
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);
					ListView_DeleteItem(hWndListViewFontList, lvi.iItem);

					ssMessage << ((FONTLISTCHANGEDSTRUCT*)lParam)->lpszFontName << L" closed\r\n";
				}
				break;
			default:
				break;
			}
			strMessage = ssMessage.str();
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
		}
		break;
		// Child window size/position changed notidfication
		// wParam = Child window ID : enum ID
	case USERMESSAGE::CHILDWINDOWPOSCHANGED:
		{
			switch ((ID)wParam)
			{
			case ID::ListViewFontList:
				{
					// Adjust column width
					HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)wParam) };

					RECT rcListViewFontListClient{};
					GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
					ListView_SetColumnWidth(hWndListViewFontList, 0, rcListViewFontListClient.right - rcListViewFontListClient.left - ListView_GetColumnWidth(hWndListViewFontList, 1));
				}
				break;
			case ID::EditMessage:
				{
					// Scroll caret into view
					Edit_ScrollCaret(GetDlgItem(hWnd, (int)wParam));
				}
				break;
			default:
				break;
			}
		}
		break;
		// System tray notification
	case USERMESSAGE::TRAYNOTIFYICON:
		{
			switch (LOWORD(lParam))
			{
			case NIN_SELECT:
			case NIN_KEYSELECT:
				{
					if (IsWindowVisible(hWnd))
					{
						SetForegroundWindow(hWnd);
					}
					else
					{
						ShowWindow(hWnd, SW_SHOW);
					}
				}
				break;
			case WM_CONTEXTMENU:
				{
					// Windows bug, see "KB135788 - PRB: Menus for Notification Icons Do Not Work Correctly"
					SetForegroundWindow(hWnd);

					UINT uFlags{};
					if (GetSystemMetrics(SM_MENUDROPALIGNMENT))
					{
						uFlags |= TPM_RIGHTALIGN;
					}
					else
					{
						uFlags |= TPM_LEFTALIGN;
					}
					TrackPopupMenu(hMenuContextTray, uFlags | TPM_BOTTOMALIGN | TPM_RIGHTBUTTON, GET_X_LPARAM(wParam), GET_Y_LPARAM(wParam), 0, hWnd, NULL);

					PostMessage(hWnd, WM_NULL, 0, 0);
				}
				break;
			default:
				break;
			}
		}
		break;
	default:
		break;
	}

	static HFONT hFontMain{};

	static LONG cyEditMessageTextMargin{};

	static UINT_PTR SizingEdge{};
	static RECT rcStatusBarFontInfoOld{};
	static int PreviousShowCmd{};

	switch (Message)
	{
	case WM_CREATE:
		{
			/*
			┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
			┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃
			┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨
			┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃
			┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃
			┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃
			┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨
			┃ Font Name                                                         │ State        ┃
			┠───────────────────────────────────────────────────────────────────┼──────────────┨
			┃                                                                   ┆              ┃
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
			┃                                                                   ┆              ┃
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
			┃                                                                   ┆              ┃
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
			┃                                                                   ┆              ┃
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
			┠───────────────────────────────────────────────────────────────────┴──────────────┨
			┠──────────────────────────────────────────────────────────────────────────────┬───┨
			┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃
			┃                                                                              ├───┨
			┃ How to load fonts to Windows:                                                │▓▓▓┃
			┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃
			┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃
			┃  view, then click "Load" button.                                             │▓▓▓┃
			┃                                                                              ├───┨
			┃ How to unload fonts from Windows:                                            │   ┃
			┃ Select all fonts then click "Unload" or "Close" button or the X at the       │   ┃
			┃ upper-right cornor.                                                          │   ┃
			┃                                                                              │   ┃
			┃ How to load fonts to process:                                                │   ┃
			┃ 1.Click "Click to select process", select a process.                         │   ┃
			┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨
			┃ view, then click "Load" button.                                              │ ↓ ┃
			┠──────────────────────────────────────────────────────────────────────────────┴───┨
			┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃
			┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
			*/

			// Get initial window state
			WINDOWPLACEMENT wp{ sizeof(WINDOWPLACEMENT) };
			GetWindowPlacement(hWnd, &wp);
			PreviousShowCmd = wp.showCmd;

			NONCLIENTMETRICS ncm{ sizeof(NONCLIENTMETRICS) };
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
			hFontMain = CreateFontIndirect(&ncm.lfMessageFont);

			RECT rcMainClient{};
			GetClientRect(hWnd, &rcMainClient);

			// Initialize ButtonOpen
			HWND hWndButtonOpen{ CreateWindow(WC_BUTTON, L"&Open", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, rcMainClient.left, rcMainClient.top, 50, 50, hWnd, (HMENU)ID::ButtonOpen, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonOpen, hFontMain, TRUE);

			// Initialize ButtonClose
			RECT rcButtonOpen{};
			GetWindowRect(hWndButtonOpen, &rcButtonOpen);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);

			HWND hWndButtonClose{ CreateWindow(WC_BUTTON, L"&Close", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, rcButtonOpen.right, rcMainClient.top, 50, 50, hWnd, (HMENU)ID::ButtonClose, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonClose, hFontMain, TRUE);

			// Initialize ButtonLoad
			RECT rcButtonClose{};
			GetWindowRect(hWndButtonClose, &rcButtonClose);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonClose);
			HWND hWndButtonLoad{ CreateWindow(WC_BUTTON, L"&Load", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, rcButtonClose.right, rcMainClient.top, 50, 50, hWnd, (HMENU)ID::ButtonLoad, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonLoad, hFontMain, TRUE);

			// Initialize ButtonUnload
			RECT rcButtonLoad{};
			GetWindowRect(hWndButtonLoad, &rcButtonLoad);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonLoad);
			HWND hWndButtonUnload{ CreateWindow(WC_BUTTON, L"&Unload", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, rcButtonLoad.right, rcMainClient.top, 50, 50, hWnd, (HMENU)ID::ButtonUnload, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonUnload, hFontMain, TRUE);

			// Initialize ButtonBroadcastWM_FONTCHANGE
			RECT rcButtonUnload{};
			GetWindowRect(hWndButtonUnload, &rcButtonUnload);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonUnload);
			HWND hWndButtonBroadcastWM_FONTCHANGE{ CreateWindow(WC_BUTTON, L"&Broadcast WM_FONTCHANGE", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, rcButtonUnload.right, rcMainClient.top, 250, 21, hWnd, (HMENU)ID::ButtonBroadcastWM_FONTCHANGE, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonBroadcastWM_FONTCHANGE, hFontMain, TRUE);

			// Initialize EditTimeout and its label
			RECT rcButtonBroadcastWM_FONTCHANGE{};
			GetWindowRect(hWndButtonBroadcastWM_FONTCHANGE, &rcButtonBroadcastWM_FONTCHANGE);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonBroadcastWM_FONTCHANGE);
			HWND hWndStaticTimeout{ CreateWindow(WC_STATIC, L"&Timeout:", WS_CHILD | WS_VISIBLE | SS_LEFT , rcButtonBroadcastWM_FONTCHANGE.right + 20, rcMainClient.top + 1, 50, 19, hWnd, (HMENU)ID::StaticTimeout, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndStaticTimeout, hFontMain, TRUE);

			RECT rcStaticTimeout{};
			GetWindowRect(hWndStaticTimeout, &rcStaticTimeout);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcStaticTimeout);
			HWND hWndEditTimeout{ CreateWindow(WC_EDIT, L"5000", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_LEFT | ES_NUMBER | ES_AUTOHSCROLL | ES_NOHIDESEL, rcStaticTimeout.right, rcMainClient.top, 80, 21, hWnd, (HMENU)ID::EditTimeout, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndEditTimeout, hFontMain, TRUE);

			Edit_LimitText(hWndEditTimeout, 10);

			SetWindowSubclass(hWndEditTimeout, EditTimeoutSubclassProc, 0, NULL);

			dwTimeout = 5000;

			// Initialize ButtonSelectProcess
			HWND hWndButtonSelectProcess{ CreateWindow(WC_BUTTON, L"&Select process", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, rcButtonUnload.right, rcButtonUnload.bottom - 21, 250, 21, hWnd, (HMENU)ID::ButtonSelectProcess, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonSelectProcess, hFontMain, TRUE);

			// Initialize ButtonMinimizeToTray
			HWND hWndButtonMinimizeToTray{ CreateWindow(WC_BUTTON, L"&Minimize to tray", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, rcButtonBroadcastWM_FONTCHANGE.right + 20, rcButtonUnload.bottom - 21, 130, 21, hWnd, (HMENU)ID::ButtonMinimizeToTray, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonMinimizeToTray, hFontMain, TRUE);

			// Initialize StatusBar
			HWND hWndStatusBarFontInfo{ CreateWindow(STATUSCLASSNAME, L"0 font(s) opened, 0 font(s) loaded.", WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0, hWnd, (HMENU)ID::StatusBarFontInfo, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonMinimizeToTray, hFontMain, TRUE);

			// Initialize Splitter
			RECT rcStatusBarFontInfo{};
			GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

			HWND hWndSplitter{ CreateWindow(UC_SPLITTER, NULL, WS_CHILD | WS_VISIBLE, rcMainClient.left, rcButtonOpen.bottom + ((rcMainClient.bottom - rcMainClient.top) - (rcButtonOpen.bottom - rcButtonOpen.top) - (rcStatusBarFontInfo.bottom - rcStatusBarFontInfo.top)) / 2 - 2, rcMainClient.right - rcMainClient.left, 5, hWnd, (HMENU)ID::Splitter, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };

			// Initialize ListViewFontList
			RECT rcSplitter{};
			GetWindowRect(hWndSplitter, &rcSplitter);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);

			HWND hWndListViewFontList{ CreateWindow(WC_LISTVIEW, L"FontList", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_NOSORTHEADER, rcMainClient.left, rcButtonOpen.bottom, rcMainClient.right - rcMainClient.left, rcSplitter.top - rcButtonOpen.bottom, hWnd, (HMENU)ID::ListViewFontList, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			ListView_SetExtendedListViewStyle(hWndListViewFontList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
			SetWindowFont(hWndListViewFontList, hFontMain, TRUE);

			RECT rcListViewFontListClient{};
			GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
			LVCOLUMN lvc1{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rcListViewFontListClient.right - rcListViewFontListClient.left) * 4 / 5 , (LPWSTR)L"Font Name" };
			LVCOLUMN lvc2{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rcListViewFontListClient.right - rcListViewFontListClient.left) * 1 / 5 , (LPWSTR)L"State" };
			ListView_InsertColumn(hWndListViewFontList, 0, &lvc1);
			ListView_InsertColumn(hWndListViewFontList, 1, &lvc2);

			DragAcceptFiles(hWndListViewFontList, TRUE);

			SetWindowSubclass(hWndListViewFontList, ListViewFontListSubclassProc, 0, NULL);

			CHANGEFILTERSTRUCT cfs{ sizeof(CHANGEFILTERSTRUCT) };
			ChangeWindowMessageFilterEx(hWndListViewFontList, WM_DROPFILES, MSGFLT_ALLOW, &cfs);
			ChangeWindowMessageFilterEx(hWndListViewFontList, WM_COPYDATA, MSGFLT_ALLOW, &cfs);
			ChangeWindowMessageFilterEx(hWndListViewFontList, 0x0049, MSGFLT_ALLOW, &cfs);	// 0x0049 == WM_COPYGLOBALDATA

			// Initialize EditMessage
			HWND hWndEditMessage{ CreateWindow(WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | ES_READONLY | ES_LEFT | ES_MULTILINE | ES_NOHIDESEL, rcMainClient.left, rcSplitter.bottom, rcMainClient.right - rcMainClient.left, rcStatusBarFontInfo.top - rcSplitter.bottom, hWnd, (HMENU)ID::EditMessage, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndEditMessage, hFontMain, TRUE);
			Edit_LimitText(hWndEditMessage, 0);
			Edit_SetText(hWndEditMessage,
				LR"(Temporarily load fonts to Windows or specific process.)""\r\n"
				"\r\n"
				R"(How to load fonts to Windows:)""\r\n"
				R"(1.Drag-drop font files onto the icon of this application.)""\r\n"
				R"(2.Click "Open" button to select fonts or drag-drop font files onto the list view, then click "Load" button.)""\r\n"
				"\r\n"
				R"(How to unload fonts from Windows:)""\r\n"
				R"(Select all fonts then click "Unload" or "Close" button or the X at the upper-right cornor.)""\r\n"
				"\r\n"
				R"(How to load fonts to process:)""\r\n"
				R"(1.Click "Click to select process", select a process.)""\r\n"
				R"(2.Click "Open" button to select fonts or drag-drop font files onto the list view, then click "Load" button.)""\r\n"
				"\r\n"
				R"(How to unload fonts from process:)""\r\n"
				R"(Select all fonts then click "Unload" or "Close" button or the X at the upper-right cornor or terminate selected process.)""\r\n"
				"\r\n"
				R"(UI description:)""\r\n"
				R"("Open": Add fonts to the list view.)""\r\n"
				R"("Close": Remove selected fonts from Windows or target process and the list view.)""\r\n"
				R"("Load": Add selected fonts to Windows or target process.)""\r\n"
				R"("Unload": Remove selected fonts from Windows or target process.)""\r\n"
				R"("Broadcast WM_FONTCHANGE": If checked, broadcast WM_FONTCHANGE message to all top windows when loading or unloading fonts.)""\r\n"
				R"("Select process": Select a process to only load fonts to selected process.)""\r\n"
				R"("Timeout": The time in milliseconds FontLoaderEx waits before reporting failure while injecting dll into target process via proxy process, the default value is 5000. Type 0, 4294967295 or clear content to wait infinitely.)""\r\n"
				R"("Minimize to tray": If checked, click minimize or close button will minimize the window to system tray.)""\r\n"
				R"("Font Name": Names of the fonts added to the list view.)""\r\n"
				R"("State": State of the font. There are five states, "Not loaded", "Loaded", "Load failed", "Unloaded" and "Unload failed".)""\r\n"
				"\r\n"
			);

			SetWindowSubclass(hWndEditMessage, EditMessageSubclassProc, 0, NULL);

			// Get vertical margin of the formatting rectangle in EditMessage
			RECT rcEditMessageClient{}, rcEditMessageFormatting{};
			GetClientRect(hWndEditMessage, &rcEditMessageClient);
			Edit_GetRect(hWndEditMessage, &rcEditMessageFormatting);
			cyEditMessageTextMargin = (rcEditMessageClient.bottom - rcEditMessageClient.top) - (rcEditMessageFormatting.bottom - rcEditMessageFormatting.top);
		}
		break;
	case WM_ACTIVATE:
		{
			// Process drag-drop font files onto the application icon stage II
			if (bDragDropHasFonts)
			{
				// Disable controls
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), FALSE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), FALSE);
				EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), FALSE);
				EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

				// Update StatusBarFontInfo
				std::wstringstream ssFontInfo{};
				std::wstring strFontInfo{};
				std::size_t nLoadedFonts{};
				for (const auto& i : FontList)
				{
					if (i.IsLoaded())
					{
						nLoadedFonts++;
					}
				}
				ssFontInfo << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
				strFontInfo = ssFontInfo.str();
				SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), strFontInfo.c_str());

				// Update syatem tray icon tip
				if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
				{
					NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
					wcscpy_s(nid.szTip, strFontInfo.c_str());
					Shell_NotifyIcon(NIM_MODIFY, &nid);
				}

				_beginthread(DragDropWorkerThreadProc, 0, nullptr);

				bDragDropHasFonts = false;
			}

			SetFocus(GetDlgItem(hWnd, (int)ID::ButtonOpen));
		}
		break;
	case WM_COMMAND:
		{
			switch ((ID)LOWORD(wParam))
			{
				// "Open" Button
			case ID::ButtonOpen:
				{
					switch (HIWORD(wParam))
					{
						// Open “Open” dialog and add fonts
					case BN_CLICKED:
						{
							bool bIsFontListEmptyBefore{ FontList.empty() };

							HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
							HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };

							const DWORD cchBuffer{ 1024 };
							OPENFILENAME ofn{ sizeof(ofn), hWnd, NULL, L"Font Files(*.ttf;*.ttc;*.otf)\0*.ttf;*.ttc;*.otf\0", NULL, 0, 0, new WCHAR[cchBuffer]{}, cchBuffer * sizeof(WCHAR), NULL, 0, NULL, L"Select fonts", OFN_EXPLORER | OFN_ENABLESIZING | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_ENABLEHOOK, 0, 0, NULL, (LPARAM)&ofn,
								[](HWND hWndOpenDialogChild, UINT Message, WPARAM wParam, LPARAM lParam) -> UINT_PTR
								{
									UINT_PTR ret{};

									static HWND hWndOpenDialog{};

									static LPOPENFILENAME lpofn{};

									switch (Message)
									{
									case WM_INITDIALOG:
										{
											// Get the pointer to original ofn
											lpofn = (LPOPENFILENAME)lParam;

											// Get HWND to open dialog
											hWndOpenDialog = GetParent(hWndOpenDialogChild);

											// Change default font
											SetWindowFont(GetDlgItem(hWndOpenDialog, chx1), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, cmb1), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, stc2), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, cmb2), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, stc4), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, cmb13), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, edt1), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, stc3), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, lst1), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, stc1), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, IDOK), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, IDCANCEL), hFontMain, TRUE);
											SetWindowFont(GetDlgItem(hWndOpenDialog, pshHelp), hFontMain, TRUE);

											ret = (UINT_PTR)FALSE;
										}
										break;
									case WM_NOTIFY:
										{
											switch (((LPOFNOTIFY)lParam)->hdr.code)
											{
												// Adjust buffer size
											case CDN_SELCHANGE:
												{
													DWORD cchLength{ (DWORD)CommDlg_OpenSave_GetSpec(hWndOpenDialog, NULL, 0) + MAX_PATH };
													if (lpofn->nMaxFile < cchLength)
													{
														delete[] lpofn->lpstrFile;
														lpofn->lpstrFile = new WCHAR[cchLength]{};
														lpofn->nMaxFile = cchLength * sizeof(WCHAR);
													}

													ret = (UINT_PTR)TRUE;
												}
												break;
											default:
												break;
											}
										}
										break;
									default:
										break;
									}

									return ret;
								},
								NULL, NULL, 0, 0 };
							if (GetOpenFileName(&ofn))
							{
								FONTLISTCHANGEDSTRUCT flcs{ ListView_GetItemCount(hWndListViewFontList) };
								if (PathIsDirectory(ofn.lpstrFile))
								{
									LPWSTR lpszFileName{ ofn.lpstrFile + ofn.nFileOffset };
									do
									{
										WCHAR lpszPath[MAX_PATH]{};
										PathCombine(lpszPath, ofn.lpstrFile, lpszFileName);
										lpszFileName += std::wcslen(lpszFileName) + 1;
										FontList.push_back(lpszPath);

										flcs.lpszFontName = lpszPath;
										SendMessage(hWnd, (UINT)USERMESSAGE::FONTLISTCHANGED, (WPARAM)FONTLISTCHANGED::OPENED, (LPARAM)&flcs);

										flcs.iItem++;
									} while (*lpszFileName);
								}
								else
								{
									FontList.push_back(ofn.lpstrFile);

									flcs.lpszFontName = ofn.lpstrFile;
									SendMessage(hWnd, (UINT)USERMESSAGE::FONTLISTCHANGED, (WPARAM)FONTLISTCHANGED::OPENED, (LPARAM)&flcs);
								}
								int cchMessageLength{ Edit_GetTextLength(hWndEditMessage) };
								Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
								Edit_ReplaceSel(hWndEditMessage, L"\r\n");

								if (bIsFontListEmptyBefore)
								{
									EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), TRUE);
									EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), TRUE);
									EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), TRUE);
									EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_ENABLED);
									EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_ENABLED);
									EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_ENABLED);
									EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_ENABLED);
								}
							}
							delete[] ofn.lpstrFile;

							// Update StatusBarFontInfo
							std::wstringstream ssFontInfo{};
							std::wstring strFontInfo{};
							std::size_t nLoadedFonts{};
							for (const auto& i : FontList)
							{
								if (i.IsLoaded())
								{
									nLoadedFonts++;
								}
							}
							ssFontInfo << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
							strFontInfo = ssFontInfo.str();
							SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), strFontInfo.c_str());

							// Update syatem tray icon tip
							if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
							{
								NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
								wcscpy_s(nid.szTip, strFontInfo.c_str());
								Shell_NotifyIcon(NIM_MODIFY, &nid);
							}
						}
						break;
					default:
						break;
					}
				}
				break;
				// "Close" button
			case ID::ButtonClose:
				{
					switch (HIWORD(wParam))
					{
						// Unload and close selected fonts
						// Won't close those failed to unload
					case BN_CLICKED:
						{
							// Disable controls
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), FALSE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							// Update StatusBarFontInfo
							SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), L"Unloading and closing fonts...");

							// Start worker thread
							_beginthread(ButtonCloseWorkerThreadProc, 0, (void*)ID::ListViewFontList);
						}
						break;
					default:
						break;
					}
				}
				break;
				// "Load" button
			case ID::ButtonLoad:
				{
					switch (HIWORD(wParam))
					{
						// Load selected fonts
					case BN_CLICKED:
						{
							// Disable controls
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), FALSE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							// Update StatusBarFontInfo
							SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), L"Loading fonts...");

							// Start worker thread
							_beginthread(ButtonLoadWorkerThreadProc, 0, (void*)ID::ListViewFontList);
						}
					default:
						break;
					}
				}
				break;
				// "Unload" button
			case ID::ButtonUnload:
				{
					switch (HIWORD(wParam))
					{
						// Unload selected fonts
					case BN_CLICKED:
						{
							// Disable controls
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), FALSE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							// Update StatusBarFontInfo
							SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), L"Unloading fonts...");

							// Start worker thread
							_beginthread(ButtonUnloadWorkerThreadProc, 0, (void*)ID::ListViewFontList);
						}
					default:
						break;
					}
				}
				break;
				// "Timeout" edit
			case ID::EditTimeout:
				{
					switch (HIWORD(wParam))
					{
					case EN_UPDATE:
						{
							// Show balloon tip if number is out of range
							HWND hWndEditTimeout{ GetDlgItem(hWnd, (int)ID::EditTimeout) };

							std::wstringstream ssTipEdit{};
							std::wstring strTipEdit{};

							BOOL bIsConverted{};
							DWORD dwTimeoutTemp{ (DWORD)GetDlgItemInt(hWnd, (int)ID::EditTimeout, &bIsConverted, FALSE) };
							if (!bIsConverted)
							{
								if (Edit_GetTextLength(hWndEditTimeout) == 0)
								{
									ssTipEdit << L"Empty text will be treated as infinite.";
									strTipEdit = ssTipEdit.str();
									EDITBALLOONTIP ebt{ sizeof(EDITBALLOONTIP), L"Infinite", strTipEdit.c_str(), TTI_INFO };
									Edit_ShowBalloonTip(hWndEditTimeout, &ebt);

									dwTimeout = INFINITE;
								}
								else
								{
									DWORD dwCaretIndex{ Edit_GetCaretIndex(hWndEditTimeout) };
									SetDlgItemInt(hWnd, (int)ID::EditTimeout, (UINT)dwTimeout, FALSE);
									Edit_SetCaretIndex(hWndEditTimeout, dwCaretIndex);

									EDITBALLOONTIP ebt{ sizeof(EDITBALLOONTIP), L"Out of range", L"Valid timeout value is 0 ~ 4294967295.", TTI_ERROR };
									Edit_ShowBalloonTip(hWndEditTimeout, &ebt);
								}
								MessageBeep(0xFFFFFFFF);
							}
							else
							{
								if (dwTimeoutTemp == 0 || dwTimeoutTemp == 4294967295)
								{
									ssTipEdit << dwTimeoutTemp << L" will be treated as infinite.";
									strTipEdit = ssTipEdit.str();
									EDITBALLOONTIP ebt{ sizeof(EDITBALLOONTIP), L"Infinite", strTipEdit.c_str(), TTI_INFO };
									Edit_ShowBalloonTip(hWndEditTimeout, &ebt);
									MessageBeep(0xFFFFFFFF);

									dwTimeout = INFINITE;
								}
								else
								{
									dwTimeout = dwTimeoutTemp;
								}
							}
						}
						break;
					default:
						break;
					}
				}
				break;
				// "Select process" button
			case ID::ButtonSelectProcess:
				{
					switch (HIWORD(wParam))
					{
						// Select a process to load fonts to
					case BN_CLICKED:
						{
							HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
							HWND hWndButtonSelectProcess{ GetDlgItem(hWnd, (int)ID::ButtonSelectProcess) };

							static bool bIsSeDebugPrivilegeEnabled{ false };

							std::wstringstream ssMessage{};
							std::wstring strMessage{};
							int cchMessageLength{};

							// Enable SeDebugPrivilege
							if (!bIsSeDebugPrivilegeEnabled)
							{
								if (!EnableDebugPrivilege())
								{
									ssMessage << L"Failed to enable " << SE_DEBUG_NAME << L" for " << szAppName << L".";
									strMessage = ssMessage.str();
									MessageBoxCentered(NULL, strMessage.c_str(), szAppName, MB_ICONERROR);

									break;
								}
								bIsSeDebugPrivilegeEnabled = true;
							}

							// Select process
							ProcessInfo* pi{ (ProcessInfo*)DialogBox(NULL, MAKEINTRESOURCE(IDD_DIALOG1), hWnd, DialogProc) };

							// If p != nullptr, select it
							if (pi)
							{
								ProcessInfo SelectedProcessInfo = *pi;
								delete pi;

								// Clear selected process
								// If loaded via proxy
								if (ProxyProcessInfo.hProcess)
								{
									// Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
									COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::PULLDLL, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
									WaitForSingleObject(hEventProxyDllPullingFinished, INFINITE);
									CloseHandle(hEventProxyDllPullingFinished);
									switch (ProxyDllPullingResult)
									{
									case PROXYDLLPULL::SUCCESSFUL:
										{
											ssMessage << szInjectionDllNameByProxy << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
											ssMessage.str(L"");
										}
										goto continue_B9A25A68;
									case PROXYDLLPULL::FAILED:
										{
											ssMessage << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
										}
										goto break_B9A25A68;
									default:
										break;
									}
								break_B9A25A68:
									break;
								continue_B9A25A68:

									// Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hThreadWatch, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hThreadWatch);

									// Terminate message thread
									PostMessage(hWndMessage, WM_CLOSE, 0, 0);
									WaitForSingleObject(hThreadMessage, INFINITE);
									DWORD dwMessageThreadExitCode{};
									GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
									if (dwMessageThreadExitCode)
									{
										std::wstringstream ssMessage{};
										std::wstring strMessage{};
										ssMessage << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
										strMessage = ssMessage.str();
										MessageBoxCentered(NULL, strMessage.c_str(), szAppName, MB_ICONERROR);
									}
									CloseHandle(hThreadMessage);

									// Terminate proxy process
									COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
									WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
									ssMessage << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") successfully terminated.\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
									CloseHandle(ProxyProcessInfo.hProcess);
									ProxyProcessInfo.hProcess = NULL;

									// Close HANDLE to target process, duplicated handles and synchronization objects
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
									CloseHandle(hProcessCurrentDuplicated);
									CloseHandle(hProcessTargetDuplicated);
									CloseHandle(hEventProxyAddFontFinished);
									CloseHandle(hEventProxyRemoveFontFinished);
								}
								// Else DIY
								else if (TargetProcessInfo.hProcess)
								{
									// Unload FontLoaderExInjectionDll(64).dll from target process
									if (PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										ssMessage << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
										ssMessage.str(L"");
									}
									else
									{
										ssMessage << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

										break;
									}

									// Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hThreadWatch, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hThreadWatch);

									// Close HANDLE to target process
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
								}

								// Get HANDLE to target process
								SelectedProcessInfo.hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, SelectedProcessInfo.dwProcessID);
								if (!SelectedProcessInfo.hProcess)
								{
									ssMessage << L"Failed to open process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

									break;
								}

								// Determine current and target process machine architecture
								BOOL bIsCurrentProcessWOW64{}, bIsTargetProcessWOW64{}, bIsTargetProcessAArch64{};
								// Get the function pointer to IsWow64Process2() for compatibility
								typedef BOOL(__stdcall *pfnIsWow64Process2)(HANDLE hProcess, USHORT *pProcessMachine, USHORT *pNativeMachine);
								pfnIsWow64Process2 IsWow64Process2_{ (pfnIsWow64Process2)GetProcAddress(GetModuleHandle(L"Kernel32.dll"), "IsWow64Process2") };
								// For Win10 1709+, IsWow64Process2 exists
								if (IsWow64Process2_)
								{
									USHORT usCurrentProcessMachineArchitecture{}, usTargetProcessMachineArchitecture{}, usNativeMachineArchitecture{};
									IsWow64Process2_(GetCurrentProcess(), &usCurrentProcessMachineArchitecture, &usNativeMachineArchitecture);
									IsWow64Process2_(SelectedProcessInfo.hProcess, &usTargetProcessMachineArchitecture, &usNativeMachineArchitecture);
									if ((usNativeMachineArchitecture == IMAGE_FILE_MACHINE_ARM64) && (usTargetProcessMachineArchitecture == IMAGE_FILE_MACHINE_UNKNOWN))
									{
										bIsTargetProcessAArch64 = TRUE;
									}
									if ((usNativeMachineArchitecture == IMAGE_FILE_MACHINE_AMD64) && (usCurrentProcessMachineArchitecture == IMAGE_FILE_MACHINE_I386))
									{
										bIsCurrentProcessWOW64 = TRUE;
									}
									if ((usNativeMachineArchitecture == IMAGE_FILE_MACHINE_AMD64) && (usTargetProcessMachineArchitecture == IMAGE_FILE_MACHINE_I386))
									{
										bIsTargetProcessWOW64 = TRUE;
									}
								}
								// Else call bIsWow64Process() instead
								else
								{
									IsWow64Process(GetCurrentProcess(), &bIsCurrentProcessWOW64);
									IsWow64Process(SelectedProcessInfo.hProcess, &bIsTargetProcessWOW64);
								}

								// If target process is AArch64 architecture, pop up message box
								if (bIsTargetProcessAArch64)
								{
									MessageBoxCentered(hWnd, L"Process of AArch64 architecture not supported.", szAppName, MB_ICONERROR);

									break;
								}

								// If process architectures are different(one is WOW64 and another isn't), launch FontLoaderExProxy.exe to inject dll
								if (bIsCurrentProcessWOW64 != bIsTargetProcessWOW64)
								{
									// Create synchronization objects
									SECURITY_ATTRIBUTES sa{ sizeof(SECURITY_ATTRIBUTES), NULL, TRUE };
									hEventParentProcessRunning = CreateEvent(NULL, TRUE, FALSE, L"FontLoaderEx_EventParentProcessRunning_B980D8A4-C487-4306-9D17-3BA6A2CCA4A4");
									hEventMessageThreadNotReady = CreateEvent(NULL, TRUE, FALSE, NULL);
									hEventMessageThreadReady = CreateEvent(&sa, TRUE, FALSE, NULL);
									hEventProxyProcessReady = CreateEvent(&sa, TRUE, FALSE, NULL);
									hEventProxyProcessDebugPrivilegeEnablingFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
									hEventProxyProcessHWNDRevieved = CreateEvent(NULL, TRUE, FALSE, NULL);
									hEventProxyDllInjectionFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
									hEventProxyDllPullingFinished = CreateEvent(NULL, TRUE, FALSE, NULL);

									// Start message thread
									hThreadMessage = (HANDLE)_beginthreadex(nullptr, 0, MessageThreadProc, nullptr, 0, nullptr);
									HANDLE handles[]{ hEventMessageThreadNotReady, hEventMessageThreadReady };
									switch (WaitForMultipleObjects(2, handles, FALSE, INFINITE))
									{
									case WAIT_OBJECT_0:
										{
											MessageBoxCentered(NULL, L"Failed to create message-only window.", szAppName, MB_ICONERROR);

											WaitForSingleObject(hThreadMessage, INFINITE);
											CloseHandle(hThreadMessage);

											CloseHandle(hEventParentProcessRunning);
											CloseHandle(hEventMessageThreadNotReady);
											CloseHandle(hEventMessageThreadReady);
											CloseHandle(hEventProxyProcessReady);
											CloseHandle(hEventProxyProcessDebugPrivilegeEnablingFinished);
											CloseHandle(hEventProxyProcessHWNDRevieved);
											CloseHandle(hEventProxyDllInjectionFinished);
											CloseHandle(hEventProxyDllPullingFinished);
										}
										goto break_721EFBC1;
									case WAIT_OBJECT_0 + 1:
										{
											CloseHandle(hEventMessageThreadNotReady);
										}
										goto continue_721EFBC1;
									default:
										break;
									}
								break_721EFBC1:
									break;
								continue_721EFBC1:

									// Launch proxy process, send HANDLE to current process and target process, HWND to message window, HANDLE to synchronization objects and timeout as arguments to proxy process
									const WCHAR szProxyAppName[]{ L"FontLoaderExProxy.exe" };

									DuplicateHandle(GetCurrentProcess(), GetCurrentProcess(), GetCurrentProcess(), &hProcessCurrentDuplicated, 0, TRUE, DUPLICATE_SAME_ACCESS);
									DuplicateHandle(GetCurrentProcess(), SelectedProcessInfo.hProcess, GetCurrentProcess(), &hProcessTargetDuplicated, 0, TRUE, DUPLICATE_SAME_ACCESS);
									std::wstringstream ssParams{};
									ssParams << (UINT_PTR)hProcessCurrentDuplicated << L" " << (UINT_PTR)hProcessTargetDuplicated << L" " << (UINT_PTR)hWndMessage << L" " << (UINT_PTR)hEventMessageThreadReady << L" " << (UINT_PTR)hEventProxyProcessReady << L" " << dwTimeout;
									std::size_t cchParamLength{ ssParams.str().length() };
									std::unique_ptr<WCHAR[]> lpszParams{ new WCHAR[cchParamLength + 1]{} };
									wcsncpy_s(lpszParams.get(), cchParamLength + 1, ssParams.str().c_str(), cchParamLength);
									std::wstringstream ssProxyPath{};
									STARTUPINFO si{ sizeof(STARTUPINFO) };
									PROCESS_INFORMATION piProxyProcess{};
#ifdef _DEBUG
#ifdef _WIN64
									ssProxyPath << LR"(..\Debug\)" << szProxyAppName;
#else
									ssProxyPath << LR"(..\x64\Debug\)" << szProxyAppName;
#endif // _WIN64
#else
									ssProxyPath << szProxyAppName;
#endif // _DEBUG
									if (!CreateProcess(ssProxyPath.str().c_str(), lpszParams.get(), NULL, NULL, TRUE, 0, NULL, NULL, &si, &piProxyProcess))
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hProcessCurrentDuplicated);
										CloseHandle(hProcessTargetDuplicated);

										ssMessage << L"Failed to launch " << szProxyAppName << L"." << L"\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

										// Terminate message thread
										PostMessage(hWndMessage, WM_CLOSE, 0, 0);
										WaitForSingleObject(hThreadMessage, INFINITE);
										DWORD dwMessageThreadExitCode{};
										GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
										if (dwMessageThreadExitCode)
										{
											std::wstringstream ssMessage{};
											std::wstring strMessage{};
											ssMessage << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
											strMessage = ssMessage.str();
											MessageBoxCentered(NULL, strMessage.c_str(), szAppName, MB_ICONERROR);
										}
										CloseHandle(hThreadMessage);

										CloseHandle(hEventParentProcessRunning);
										CloseHandle(hEventMessageThreadReady);
										CloseHandle(hEventProxyProcessReady);
										CloseHandle(hEventProxyProcessDebugPrivilegeEnablingFinished);
										CloseHandle(hEventProxyProcessHWNDRevieved);
										CloseHandle(hEventProxyDllInjectionFinished);
										CloseHandle(hEventProxyDllPullingFinished);

										break;
									}
									CloseHandle(piProxyProcess.hThread);

									// Wait for proxy process to ready
									WaitForSingleObject(hEventProxyProcessReady, INFINITE);
									CloseHandle(hEventProxyProcessReady);
									CloseHandle(hEventMessageThreadReady);
									CloseHandle(hEventParentProcessRunning);

									ssMessage << szProxyAppName << L"(" << piProxyProcess.dwProcessId << L") succesfully launched.\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
									ssMessage.str(L"");

									// Wait for message-only window to recieve HWND to proxy process
									WaitForSingleObject(hEventProxyProcessHWNDRevieved, INFINITE);
									CloseHandle(hEventProxyProcessHWNDRevieved);

									// Wait for proxy process to enable SeDebugPrivilege
									WaitForSingleObject(hEventProxyProcessDebugPrivilegeEnablingFinished, INFINITE);
									CloseHandle(hEventProxyProcessDebugPrivilegeEnablingFinished);
									switch (ProxyDebugPrivilegeEnablingResult)
									{
									case PROXYPROCESSDEBUGPRIVILEGEENABLING::SUCCESSFUL:
										goto continue_90567013;
									case PROXYPROCESSDEBUGPRIVILEGEENABLING::FAILED:
										{
											ssMessage << L"Failed to enable " << SE_DEBUG_NAME << L" for " << szProxyAppName << L".";
											strMessage = ssMessage.str();
											MessageBoxCentered(NULL, strMessage.c_str(), szAppName, MB_ICONERROR);
											ssMessage.str(L"");

											// Terminate proxy process
											COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
											FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
											WaitForSingleObject(piProxyProcess.hProcess, INFINITE);

											ssMessage << szProxyAppName << L"(" << piProxyProcess.dwProcessId << L") successfully terminated.\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

											CloseHandle(piProxyProcess.hThread);
											CloseHandle(piProxyProcess.hProcess);
											piProxyProcess.hProcess = NULL;

											// Terminate message thread
											PostMessage(hWndMessage, WM_CLOSE, 0, 0);
											WaitForSingleObject(hThreadMessage, INFINITE);
											DWORD dwMessageThreadExitCode{};
											GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
											if (dwMessageThreadExitCode)
											{
												std::wstringstream ssMessage{};
												std::wstring strMessage{};
												ssMessage << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
												strMessage = ssMessage.str();
												MessageBoxCentered(NULL, strMessage.c_str(), szAppName, MB_ICONERROR);
											}
											CloseHandle(hThreadMessage);

											// Close HANDLE to selected process and duplicated handles
											CloseHandle(SelectedProcessInfo.hProcess);
											CloseHandle(hProcessCurrentDuplicated);
											CloseHandle(hProcessTargetDuplicated);
											CloseHandle(hEventProxyDllPullingFinished);
										}
										goto break_90567013;
									default:
										break;
									}
								break_90567013:
									break;
								continue_90567013:

									// Begin dll injection
									COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::INJECTDLL, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);

									// Wait for proxy process to inject dll into target process
									WaitForSingleObject(hEventProxyDllInjectionFinished, INFINITE);
									CloseHandle(hEventProxyDllInjectionFinished);
									switch (ProxyDllInjectionResult)
									{
									case PROXYDLLINJECTION::SUCCESSFUL:
										{
											ssMessage << szInjectionDllNameByProxy << L" successfully injected into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

											// Register proxy AddFont() and RemoveFont() procedure and create synchronization objects
											FontResource::RegisterAddRemoveFontProc(ProxyAddFontProc, ProxyRemoveFontProc);
											hEventProxyAddFontFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
											hEventProxyRemoveFontFinished = CreateEvent(NULL, TRUE, FALSE, NULL);

											// Disable EditTimeout and ButtonBroadcastWM_FONTCHANGE
											EnableWindow(GetDlgItem(hWnd, (int)ID::EditTimeout), FALSE);
											EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);

											// Change the caption of ButtonSelectProcess
											std::wstringstream Caption{};
											Caption << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")";
											Button_SetText(hWndButtonSelectProcess, (LPCWSTR)Caption.str().c_str());

											// Set TargetProcessInfo and ProxyProcessInfo
											TargetProcessInfo = SelectedProcessInfo;
											ProxyProcessInfo = { piProxyProcess.hProcess, szProxyAppName, piProxyProcess.dwProcessId };

											// Create synchronization object and start watch thread
											hEventTerminateWatchThread = CreateEvent(NULL, TRUE, FALSE, NULL);
											hThreadWatch = (HANDLE)_beginthreadex(nullptr, 0, ProxyAndTargetProcessWatchThreadProc, nullptr, 0, nullptr);
										}
										break;
									case PROXYDLLINJECTION::FAILED:
										{
											ssMessage << L"Failed to inject " << szInjectionDllNameByProxy << L" into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::FAILEDTOENUMERATEMODULES:
										{
											ssMessage << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::GDI32NOTLOADED:
										{
											ssMessage << L"gdi32.dll not loaded by target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::MODULENOTFOUND:
										{
											ssMessage << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									continue_DBEA36FE:
										{
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
											ssMessage.str(L"");

											// Terminate proxy process
											COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
											FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
											WaitForSingleObject(piProxyProcess.hProcess, INFINITE);

											ssMessage << szProxyAppName << L"(" << piProxyProcess.dwProcessId << L") successfully terminated.\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

											CloseHandle(piProxyProcess.hProcess);
											piProxyProcess.hProcess = NULL;

											// Terminate message thread
											PostMessage(hWndMessage, WM_CLOSE, 0, 0);
											WaitForSingleObject(hThreadMessage, INFINITE);
											DWORD dwMessageThreadExitCode{};
											GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
											if (dwMessageThreadExitCode)
											{
												std::wstringstream ssMessage{};
												std::wstring strMessage{};
												ssMessage << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
												strMessage = ssMessage.str();
												MessageBoxCentered(NULL, strMessage.c_str(), szAppName, MB_ICONERROR);
											}
											CloseHandle(hThreadMessage);

											// Close HANDLE to selected process and duplicated handles
											CloseHandle(SelectedProcessInfo.hProcess);
											CloseHandle(hProcessCurrentDuplicated);
											CloseHandle(hProcessTargetDuplicated);
											CloseHandle(hEventProxyDllPullingFinished);
										}
										break;
									default:
										break;
									}
								}
								// Else DIY
								else
								{
									// Check whether target process loads gdi32.dll as AddFontResourceEx() and RemoveFontResourceEx() are in it
									HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, SelectedProcessInfo.dwProcessID) };
									MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
									bool bIsGDI32Loaded{ false };
									if (!Module32First(hModuleSnapshot, &me32))
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot);

										ssMessage << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
										break;
									}
									do
									{
										if (!lstrcmpi(me32.szModule, L"gdi32.dll"))
										{
											bIsGDI32Loaded = true;

											break;
										}
									} while (Module32Next(hModuleSnapshot, &me32));
									if (!bIsGDI32Loaded)
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot);

										ssMessage << L"gdi32.dll not loaded by target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

										break;
									}
									CloseHandle(hModuleSnapshot);

									// Inject FontLoaderExInjectionDll(64).dll into target process
									if (!InjectModule(SelectedProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										CloseHandle(SelectedProcessInfo.hProcess);

										ssMessage << L"Failed to inject " << szInjectionDllName << L" into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

										break;
									}
									ssMessage << szInjectionDllName << L" successfully injected into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
									ssMessage.str(L"");

									// Get base address of FontLoaderExInjectionDll(64).dll in target process
									HANDLE hModuleSnapshot2{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, SelectedProcessInfo.dwProcessID) };
									MODULEENTRY32 me322{ sizeof(MODULEENTRY32) };
									BYTE* lpModBaseAddr{};
									if (!Module32First(hModuleSnapshot2, &me322))
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot2);

										ssMessage << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

										break;
									}
									do
									{
										if (!lstrcmpi(me322.szModule, szInjectionDllName))
										{
											lpModBaseAddr = me322.modBaseAddr;

											break;
										}
									} while (Module32Next(hModuleSnapshot2, &me322));
									if (!lpModBaseAddr)
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot2);

										ssMessage << szInjectionDllName << " not found in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

										break;
									}
									CloseHandle(hModuleSnapshot2);

									// Calculate the addresses of AddFont() and RemoveFont() in target process
									HMODULE hModInjectionDll{ LoadLibrary(szInjectionDllName) };
									void* lpLocalAddFontProcAddr{ GetProcAddress(hModInjectionDll, "AddFont") };
									void* lpLocalRemoveFontProcAddr{ GetProcAddress(hModInjectionDll, "RemoveFont") };
									FreeLibrary(hModInjectionDll);
									lpRemoteAddFontProcAddr = (void*)((UINT_PTR)lpModBaseAddr + ((UINT_PTR)lpLocalAddFontProcAddr - (UINT_PTR)hModInjectionDll));
									lpRemoteRemoveFontProcAddr = (void*)((UINT_PTR)lpModBaseAddr + ((UINT_PTR)lpLocalRemoveFontProcAddr - (UINT_PTR)hModInjectionDll));

									// Register remote AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(RemoteAddFontProc, RemoteRemoveFontProc);

									// Disable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWnd, (int)ID::EditTimeout), FALSE);
									EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);

									// Change the caption of ButtonSelectProcess
									std::wstringstream Caption{};
									Caption << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")";
									Button_SetText(hWndButtonSelectProcess, (LPCWSTR)Caption.str().c_str());

									// Set TargetProcessInfo
									TargetProcessInfo = SelectedProcessInfo;

									// Create synchronization object and start watch thread
									hEventTerminateWatchThread = CreateEvent(NULL, TRUE, FALSE, NULL);
									hThreadWatch = (HANDLE)_beginthreadex(nullptr, 0, TargetProcessWatchThreadProc, nullptr, 0, nullptr);
								}
							}
							// If p == nullptr, clear selected process
							else
							{
								// If loaded via proxy
								if (ProxyProcessInfo.hProcess)
								{
									// Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
									COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::PULLDLL, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
									WaitForSingleObject(hEventProxyDllPullingFinished, INFINITE);
									CloseHandle(hEventProxyDllPullingFinished);
									switch (ProxyDllPullingResult)
									{
									case PROXYDLLPULL::SUCCESSFUL:
										{
											ssMessage << szInjectionDllNameByProxy << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
											ssMessage.str(L"");
										}
										goto continue_0F70B465;
									case PROXYDLLPULL::FAILED:
										{
											ssMessage << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
										}
										goto break_0F70B465;
									default:
										break;
									}
								break_0F70B465:
									break;
								continue_0F70B465:

									// Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hThreadWatch, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hThreadWatch);

									// Terminate message thread
									PostMessage(hWndMessage, WM_CLOSE, 0, 0);
									WaitForSingleObject(hThreadMessage, INFINITE);
									DWORD dwMessageThreadExitCode{};
									GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
									if (dwMessageThreadExitCode)
									{
										std::wstringstream ssMessage{};
										std::wstring strMessage{};
										ssMessage << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
										strMessage = ssMessage.str();
										MessageBoxCentered(NULL, strMessage.c_str(), szAppName, MB_ICONERROR);
									}
									CloseHandle(hThreadMessage);

									// Terminate proxy process
									COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
									WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);

									ssMessage << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") successfully terminated.\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

									CloseHandle(ProxyProcessInfo.hProcess);
									ProxyProcessInfo.hProcess = NULL;

									// Close HANDLE to target process, duplicated handles and synchronization objects
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
									CloseHandle(hProcessCurrentDuplicated);
									CloseHandle(hProcessTargetDuplicated);
									CloseHandle(hEventProxyAddFontFinished);
									CloseHandle(hEventProxyRemoveFontFinished);

									// Register global AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

									// Enable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWnd, (int)ID::EditTimeout), TRUE);
									EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);

									// Revert the caption of ButtonSelectProcess to default
									Button_SetText(hWndButtonSelectProcess, L"Select process");
								}
								// Else DIY
								else if (TargetProcessInfo.hProcess)
								{
									// Unload FontLoaderExInjectionDll(64).dll from target process
									if (PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										ssMessage << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
									}
									else
									{
										ssMessage << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

										break;
									}

									// Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hThreadWatch, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hThreadWatch);

									// Close HANDLE to target process
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;

									// Register global AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

									// Enable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWnd, (int)ID::EditTimeout), TRUE);
									EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);

									// Revert the caption of ButtonSelectProcess to default
									Button_SetText(hWndButtonSelectProcess, L"Select process");
								}
							}
						}
						break;
					default:
						break;
					}
				}
				break;
				// "Minimize to trat" button
			case ID::ButtonMinimizeToTray:
				{
					switch (HIWORD(wParam))
					{
					case BN_CLICKED:
						{
							// If unchecked, remove the icon from system tray
							if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_UNCHECKED)
							{
								NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWnd, 0 };
								Shell_NotifyIcon(NIM_DELETE, &nid);
							}
						}
						break;
					default:
						break;
					}
				}
				break;
			default:
				break;
			}
			switch (LOWORD(wParam))
			{
				// "Load" menu item in hMenuContextListViewFontList
			case ID_MENU_LOAD:
				{
					// Simulate clicking "Load" button
					SendMessage(GetDlgItem(hWnd, (int)ID::ButtonLoad), BM_CLICK, 0, 0);
				}
				break;
				// "Unload" menu item in hMenuContextListViewFontList
			case ID_MENU_UNLOAD:
				{
					// Simulate clicking "Unload" button
					SendMessage(GetDlgItem(hWnd, (int)ID::ButtonUnload), BM_CLICK, 0, 0);
				}
				break;
				// "Close" menu item in hMenuContextListViewFontList
			case ID_MENU_CLOSE:
				{
					// Simulate clicking "Close" button
					SendMessage(GetDlgItem(hWnd, (int)ID::ButtonClose), BM_CLICK, 0, 0);
				}
				break;
				// "Select All" menu item in hMenuContextListViewFontList
			case ID_MENU_SELECTALL:
				{
					// Select all items in ListViewFontList
					ListView_SetItemState(GetDlgItem(hWnd, (int)ID::ListViewFontList), -1, LVIS_SELECTED, LVIS_SELECTED);
				}
				break;
				// "Show window" menu item in hMenuContextTray
			case ID_TRAYMENU_SHOWWINDOW:
				{
					// Show main window
					if (IsWindowVisible(hWnd))
					{
						SetForegroundWindow(hWnd);
					}
					else
					{
						ShowWindow(hWnd, SW_SHOW);
					}
				}
				break;
				// "Exit" menu item in hMenuContextTrat
			case ID_TRAYMENU_EXIT:
				{
					// Show main window
					if (IsWindowVisible(hWnd))
					{
						SetForegroundWindow(hWnd);
					}
					else
					{
						ShowWindow(hWnd, SW_SHOW);
					}

					// Simulate clicking close button
					Button_SetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray), BST_UNCHECKED);
					PostMessage(hWnd, WM_SYSCOMMAND, (WPARAM)SC_CLOSE, 0);
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_SYSCOMMAND:
		{
			switch (wParam & 0xFFF0)
			{
			case SC_MINIMIZE:
				{
					// If ButtonMinimizeToTray is checked, minimize to system tray 
					if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
					{
						NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWnd, 0, NIF_MESSAGE | NIF_ICON | NIF_TIP | NIF_SHOWTIP, (UINT)USERMESSAGE::TRAYNOTIFYICON, hIconApplication };
						std::wstringstream ssTip{};
						std::wstring strTip{};
						std::size_t nLoadedFonts{};
						for (const auto& i : FontList)
						{
							if (i.IsLoaded())
							{
								nLoadedFonts++;
							}
						}
						ssTip << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
						strTip = ssTip.str();
						wcscpy_s(nid.szTip, strTip.c_str());
						Shell_NotifyIcon(NIM_ADD, &nid);
						nid.uFlags = NIF_INFO;
						nid.uVersion = NOTIFYICON_VERSION_4;
						Shell_NotifyIcon(NIM_SETVERSION, &nid);

						ShowWindow(hWnd, SW_HIDE);
					}
					// Else do as usual
					else
					{
						ret = DefWindowProc(hWnd, Message, wParam, lParam);
					}
				}
				break;
			case SC_CLOSE:
				{
					// If ButtonMinimizeToTray is checked, minimize to system tray
					if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
					{
						NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWnd, 0, NIF_MESSAGE | NIF_ICON | NIF_TIP | NIF_SHOWTIP, (UINT)USERMESSAGE::TRAYNOTIFYICON, hIconApplication };
						std::wstringstream ssTip{};
						std::wstring strTip{};
						std::size_t nLoadedFonts{};
						for (const auto& i : FontList)
						{
							if (i.IsLoaded())
							{
								nLoadedFonts++;
							}
						}
						ssTip << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
						strTip = ssTip.str();
						wcscpy_s(nid.szTip, strTip.c_str());
						Shell_NotifyIcon(NIM_ADD, &nid);
						nid.uFlags = NIF_INFO;
						nid.uVersion = NOTIFYICON_VERSION_4;
						Shell_NotifyIcon(NIM_SETVERSION, &nid);

						ShowWindow(hWnd, SW_HIDE);
					}
					// Else do as usual
					else
					{
						// If font list is not empty, unload all fonts first
						if (!FontList.empty())
						{
							// Disable controls
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonSelectProcess), FALSE);
							EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), FALSE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							// Update StatusBarFontInfo
							SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), L"Unloading and closing fonts...");

							// Start worker thread
							_beginthread(CloseWorkerThreadProc, 0, nullptr);
						}
						// Else do as usual
						else
						{
							ret = DefWindowProc(hWnd, Message, wParam, lParam);
						}
					}
				}
				break;
			default:
				{
					ret = DefWindowProc(hWnd, Message, wParam, lParam);
				}
				break;
			}
		}
		break;
	case WM_NOTIFY:
		{
			switch ((ID)wParam)
			{
			case ID::ListViewFontList:
				{
					switch (((LPNMHDR)lParam)->code)
					{
						// Set empty text
					case LVN_GETEMPTYMARKUP:
						{
							((NMLVEMPTYMARKUP*)lParam)->dwFlags = 0;
							wcscpy_s(((NMLVEMPTYMARKUP*)lParam)->szMarkup, LR"(Click "Open" or drag-drop font files here to add fonts.)");

							ret = TRUE;
						}
						break;
					default:
						break;
					}
				}
				break;
			case ID::Splitter:
				{
					static LONG cyCursorOffset{};
					static POINT ptSpliitterRange{};

					switch ((SPLITTERNOTIFICATION)((LPNMHDR)lParam)->code)
					{
						// Begin dragging Splitter
					case SPLITTERNOTIFICATION::DRAGBEGIN:
						{
							// Caculate cursor offset to Splitter top in y orientation
							HWND hWndSplitter{ ((LPSPLITTERSTRUCT)lParam)->nmhdr.hwndFrom };
							RECT rcSplitter{}, rcSplitterClient{};
							GetWindowRect(hWndSplitter, &rcSplitter);
							GetClientRect(hWndSplitter, &rcSplitterClient);
							cyCursorOffset = ((LPSPLITTERSTRUCT)lParam)->ptCursor.y - rcSplitterClient.top;
							MapWindowRect(hWndSplitter, HWND_DESKTOP, &rcSplitterClient);
							cyCursorOffset += rcSplitterClient.top - rcSplitter.top;

							// Confine cursor to a specific rectangle
							/*
							┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
							┃                                                                                                                               ┃
							┃                                                                                                                               ┃
							┃	┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓                                        ┃← Desktop
							┃	┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃                                        ┃
							┃	┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨                                        ┃
							┃	┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃                                        ┃
							┃	┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃                                        ┃
							┃	┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃                                        ┃
							┃	┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────╂────────                                ┃
							┃	┃ Font Name                                                         │ State        ┃        ↑                               ┃
							┃	┠───────────────────────────────────────────────────────────────────┼──────────────┨        │ cyListViewFontListMin         ┃
							┃	┃                                                                   ┆              ┃        ↓                               ┃
							┠┄┄┄╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	┃                                                                   ┆              ┃                          ↑             ┃
							┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                ptSpliitterRange.top    ┃
							┃	┃                                                                   ┆              ┃                                        ┃
							┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                                        ┃
							┃	┃                                                                   ┆              ┃                                        ┃
							┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                                        ┃
							┃	┠───────────────────────────────────────────────────────────────────┴──────────────┨                                        ┃
							┃	┠──────────────────────────────────────────────────────────────────────────────┬───┨                                        ┃
							┃	┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃                                        ┃
							┃	┃                                                                              ├───┨                                        ┃
							┃	┃ How to load fonts to Windows:                                                │▓▓▓┃                                        ┃
							┃	┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃                                        ┃
							┃	┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃                                        ┃
							┃	┃  view, then click "Load" button.                                             │▓▓▓┃                                        ┃
							┃	┃                                                                              ├───┨                                        ┃
							┃	┃ How to unload fonts from Windows:                                            │   ┃                                        ┃
							┃	┃ Select all fonts then click "Unload" or "Close" button or the X at the       │   ┃                                        ┃
							┃	┃ upper-right cornor.                                                          │   ┃                                        ┃
							┃	┃                                                                              │   ┃                                        ┃
							┃	┃ How to load fonts to process:                                                │   ┃                                        ┃
							┃	┃ 1.Click "Click to select process", select a process.                         │   ┃              ptSpliitterRange.bottom   ┃
							┃	┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │   ┃                          ↓             ┃
							┠┄┄┄╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼───╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	┃ view, then click "Load" button.                                              │ ↓ ┃        } cyEditMessageMin              ┃
							┃	┠──────────────────────────────────────────────────────────────────────────────┴───┨────────                                ┃
							┃	┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃        } cyStatusBarFontInfo           ┃
							┃	┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛────────                                ┃
							┃                                                                                                                               ┃
							┃                                                                                                                               ┃
							┃                                                                                                                               ┃
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
							*/

							// Calculate the minimal height of ListViewFontList
							HWND hWndListViewFontList{ GetDlgItem(hWnd,(int)ID::ListViewFontList) };
							RECT rcListViewFontList{}, rcListViewFontListClient{};
							GetWindowRect(hWndListViewFontList, &rcListViewFontList);
							GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
							bool bIsInserted{ false };
							if (ListView_GetItemCount(hWndListViewFontList) == 0)
							{
								bIsInserted = true;

								LVITEM lvi{};
								ListView_InsertItem(hWndListViewFontList, &lvi);
							}
							RECT rcListViewFontListItem{};
							ListView_GetItemRect(hWndListViewFontList, 0, &rcListViewFontListItem, LVIR_BOUNDS);
							if (bIsInserted)
							{
								ListView_DeleteAllItems(hWndListViewFontList);
							}
							LONG cyListViewFontListMin{ (rcListViewFontListItem.bottom - rcListViewFontListClient.top) + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

							// Calculate the minimal height of EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
							HDC hDCEditMessage{ GetDC(hWndEditMessage) };
							SelectFont(hDCEditMessage, GetWindowFont(hWndEditMessage));
							TEXTMETRIC tm{};
							GetTextMetrics(hDCEditMessage, &tm);
							ReleaseDC(hWndEditMessage, hDCEditMessage);
							RECT rcEditMessage{}, rcEditMessageClient{};
							GetWindowRect(hWndEditMessage, &rcEditMessage);
							GetClientRect(hWndEditMessage, &rcEditMessageClient);
							LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + cyEditMessageTextMargin };

							// Calculate the Splitter range
							RECT rcMainClient{}, rcButtonOpen{}, rcStatusBarFontInfo{};
							GetClientRect(hWnd, &rcMainClient);
							GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
							GetWindowRect(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), &rcStatusBarFontInfo);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);
							ptSpliitterRange = { rcMainClient.top + (rcButtonOpen.bottom - rcButtonOpen.top) + cyListViewFontListMin, rcStatusBarFontInfo.top - (rcSplitter.bottom - rcSplitter.top) - cyEditMessageMin };
						}
						break;
						// Dragging Splitter
					case SPLITTERNOTIFICATION::DRAGGING:
						{
							// Convert cursor coordinates from Splitter to main window
							POINT ptCursor{ ((LPSPLITTERSTRUCT)lParam)->ptCursor };
							MapWindowPoints(((LPSPLITTERSTRUCT)lParam)->nmhdr.hwndFrom, hWnd, &ptCursor, 1);

							// Move Splitter
							HWND hWndSplitter{ ((LPSPLITTERSTRUCT)lParam)->nmhdr.hwndFrom };
							RECT rcSplitter{};
							GetWindowRect(hWndSplitter, &rcSplitter);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
							int yPosSplitter{ ptCursor.y - cyCursorOffset };
							if (yPosSplitter < ptSpliitterRange.x)
							{
								yPosSplitter = ptSpliitterRange.x;
							}
							else if (yPosSplitter > ptSpliitterRange.y)
							{
								yPosSplitter = ptSpliitterRange.y;
							}
							MoveWindow(hWndSplitter, rcSplitter.left, yPosSplitter, rcSplitter.right - rcSplitter.left, rcSplitter.bottom - rcSplitter.top, TRUE);

							// Resize ListViewFontList
							HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
							RECT rcListViewFontList{};
							GetWindowRect(hWndListViewFontList, &rcListViewFontList);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
							MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, rcListViewFontList.right - rcListViewFontList.left, yPosSplitter - rcListViewFontList.top, TRUE);

							// Resize EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
							RECT rcEditMessage{}, rcStatusBarFontInfo{};
							GetWindowRect(hWndEditMessage, &rcEditMessage);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);
							GetWindowRect(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), &rcStatusBarFontInfo);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);
							MoveWindow(hWndEditMessage, rcEditMessage.left, yPosSplitter + (rcSplitter.bottom - rcSplitter.top), rcEditMessage.right - rcEditMessage.left, rcStatusBarFontInfo.top - rcSplitter.bottom, TRUE);
						}
						break;
						// End dragging Splitter
					case SPLITTERNOTIFICATION::DRAGEND:
						{
							// Redraw ListViewFontList, Splitter and EditMessage
							InvalidateRect(GetDlgItem(hWnd, (int)ID::ListViewFontList), NULL, FALSE);
							InvalidateRect(GetDlgItem(hWnd, (int)ID::EditMessage), NULL, FALSE);
						}
						break;
					default:
						break;
					}
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_CONTEXTMENU:
		{
			// Show context menu in ListViewFontList
			HWND hWndListViewFontList{ GetDlgItem(hWnd,(int)ID::ListViewFontList) };

			if ((HWND)wParam == hWndListViewFontList)
			{
				POINT ptContextMenu{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
				if (ptContextMenu.x == -1 && ptContextMenu.y == -1)
				{
					int iSelectionMark{ ListView_GetSelectionMark(hWndListViewFontList) };
					if (iSelectionMark == -1)
					{
						RECT rcListViewFontListClient{}, rcHeaderListViewFontList{};
						GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
						MapWindowRect(hWndListViewFontList, HWND_DESKTOP, &rcListViewFontListClient);
						GetWindowRect(ListView_GetHeader(hWndListViewFontList), &rcHeaderListViewFontList);
						ptContextMenu = { rcListViewFontListClient.left, rcHeaderListViewFontList.bottom };
					}
					else
					{
						ListView_EnsureVisible(hWndListViewFontList, iSelectionMark, FALSE);
						ListView_GetItemPosition(hWndListViewFontList, iSelectionMark, &ptContextMenu);
						ClientToScreen(hWndListViewFontList, &ptContextMenu);
					}
				}
				UINT uFlags{};
				if (GetSystemMetrics(SM_MENUDROPALIGNMENT))
				{
					uFlags |= TPM_RIGHTALIGN;
				}
				else
				{
					uFlags |= TPM_LEFTALIGN;
				}
				TrackPopupMenu(hMenuContextListViewFontList, uFlags | TPM_TOPALIGN | TPM_RIGHTBUTTON, ptContextMenu.x, ptContextMenu.y, 0, hWnd, NULL);
			}
			else
			{
				ret = DefWindowProc(hWnd, Message, wParam, lParam);
			}
		}
		break;
	case WM_CTLCOLORSTATIC:
		{
			// Change the background color of ButtonBroadcastWM_FONTCHANGE, StaticTimeout, ButtonMinimizeToTray and EditMessage to default window background color
			// From https://social.msdn.microsoft.com/Forums/vstudio/en-US/7b6d1815-87e3-4f47-b5d5-fd4caa0e0a89/why-is-wmctlcolorstatic-sent-for-a-button-instead-of-wmctlcolorbtn?forum=vclanguage
			// "WM_CTLCOLORSTATIC is sent by any control that displays text which would be displayed using the default dialog/window background color. 
			// This includes check boxes, radio buttons, group boxes, static text, read-only or disabled edit controls, and disabled combo boxes (all styles)."
			if (((HWND)lParam == GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE)) || ((HWND)lParam == GetDlgItem(hWnd, (int)ID::StaticTimeout)) || ((HWND)lParam == GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) || (HWND)lParam == GetDlgItem(hWnd, (int)ID::EditMessage))
			{
				ret = (LRESULT)GetSysColorBrush(COLOR_WINDOW);;
			}
			else
			{
				ret = (LRESULT)DefWindowProc(hWnd, Message, wParam, lParam);
			}
		}
		break;
	case WM_WINDOWPOSCHANGING:
		{
			// Get StatusBarFontList window rectangle before main window changes size
			GetWindowRect(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), &rcStatusBarFontInfoOld);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfoOld);

			ret = DefWindowProc(hWnd, Message, wParam, lParam);
		}
		break;
	case WM_SIZING:
		{
			// Get the sizing edge
			SizingEdge = wParam;
		}
		break;
	case WM_SIZE:
		{
			// Resize controls
			switch (PreviousShowCmd)
			{
				// If previous window state is restored
			case SW_SHOWDEFAULT:
			case SW_SHOWNORMAL:
			case SW_RESTORE:
				{
					switch (wParam)
					{
						// Resize controls in respective ways by SizingEdge
					case SIZE_RESTORED:
						{
							switch (SizingEdge)
							{
								// Resize ListViewFontList and keep the height of EditMessage as far as possible
							case WMSZ_TOP:
							case WMSZ_TOPLEFT:
							case WMSZ_TOPRIGHT:
								{
									/*
																															  ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
																															  ┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃
																															  ┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨
									┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓      ┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃
									┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃      ┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃
									┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨      ┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃
									┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃      ┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨
									┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃      ┃ Font Name                                                         │ State        ┃
									┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃      ┠───────────────────────────────────────────────────────────────────┼──────────────┨
									┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨      ┃                                                                   ┆              ┃
									┃ Font Name                                                         │ State        ┃      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┠───────────────────────────────────────────────────────────────────┼──────────────┨      ┃                                                                   ┆              ┃
									┃                                                                   ┆              ┃      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┃                                                                   ┆              ┃
									┃                                                                   ┆              ┃      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┃                                                                   ┆              ┃
									┃                                                                   ┆              ┃      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┃                                                                   ┆              ┃
									┃                                                                   ┆              ┃      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┃                                                                   ┆              ┃
									┠───────────────────────────────────────────────────────────────────┴──────────────┨  =>  ┠───────────────────────────────────────────────────────────────────┴──────────────┨
									┠──────────────────────────────────────────────────────────────────────────────┬───┨      ┠──────────────────────────────────────────────────────────────────────────────┬───┨
									┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃      ┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃
									┃                                                                              ├───┨      ┃                                                                              ├───┨
									┃ How to load fonts to Windows:                                                │▓▓▓┃      ┃ How to load fonts to Windows:                                                │▓▓▓┃
									┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃      ┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃
									┃  view, then click "Load" button.                                             │▓▓▓┃      ┃  view, then click "Load" button.                                             │▓▓▓┃
									┃                                                                              ├───┨      ┃                                                                              ├───┨
									┃ How to unload fonts from Windows:                                            │   ┃      ┃ How to unload fonts from Windows:                                            │   ┃
									┃ Select all fonts then click "Unload" or "Close" button or the X at the       │   ┃      ┃ Select all fonts then click "Unload" or "Close" button or the X at the       │   ┃
									┃ upper-right cornor.                                                          │   ┃      ┃ upper-right cornor.                                                          │   ┃
									┃                                                                              │   ┃      ┃                                                                              │   ┃
									┃ How to load fonts to process:                                                │   ┃      ┃ How to load fonts to process:                                                │   ┃
									┃ 1.Click "Click to select process", select a process.                         │   ┃      ┃ 1.Click "Click to select process", select a process.                         │   ┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨
									┃ view, then click "Load" button.                                              │ ↓ ┃      ┃ view, then click "Load" button.                                              │ ↓ ┃
									┠──────────────────────────────────────────────────────────────────────────────┴───┨      ┠──────────────────────────────────────────────────────────────────────────────┴───┨
									┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃      ┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛      ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
									*/

									// Resize StatusBarFontInfo
									HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, (int)ID::StatusBarFontInfo) };
									FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);
									RECT rcStatusBarFontInfo{};
									GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

									// Calculate the minimal height of ListViewFontList
									HWND hWndListViewFontList{ GetDlgItem(hWnd,(int)ID::ListViewFontList) };
									RECT rcListViewFontList{}, rcListViewFontListClient{};
									GetWindowRect(hWndListViewFontList, &rcListViewFontList);
									GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
									bool bIsInserted{ false };
									if (ListView_GetItemCount(hWndListViewFontList) == 0)
									{
										bIsInserted = true;

										LVITEM lvi{};
										ListView_InsertItem(hWndListViewFontList, &lvi);
									}
									RECT rcListViewFontListItem{};
									ListView_GetItemRect(hWndListViewFontList, 0, &rcListViewFontListItem, LVIR_BOUNDS);
									if (bIsInserted)
									{
										ListView_DeleteAllItems(hWndListViewFontList);
									}
									LONG cyListViewFontListMin{ (rcListViewFontListItem.bottom - rcListViewFontListClient.top) + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

									// Resize ListViewFontList
									RECT rcButtonOpen{}, rcSplitter{}, rcEditMessage{};
									GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
									GetWindowRect(GetDlgItem(hWnd, (int)ID::Splitter), &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									GetWindowRect(GetDlgItem(hWnd, (int)ID::EditMessage), &rcEditMessage);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);

									bool bIsListViewFontListMinimized{ false };
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									if (rcStatusBarFontInfo.top - rcButtonOpen.bottom - (rcSplitter.bottom - rcSplitter.top) - (rcEditMessage.bottom - rcEditMessage.top) < cyListViewFontListMin)
									{
										bIsListViewFontListMinimized = true;

										MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), cyListViewFontListMin, FALSE);
									}
									else
									{
										MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcStatusBarFontInfo.top - rcButtonOpen.bottom - (rcSplitter.bottom - rcSplitter.top) - (rcEditMessage.bottom - rcEditMessage.top), FALSE);
									}

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
									if (bIsListViewFontListMinimized)
									{
										MoveWindow(hWndSplitter, rcSplitter.left, rcButtonOpen.bottom + cyListViewFontListMin, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, FALSE);
									}
									else
									{
										MoveWindow(hWndSplitter, rcSplitter.left, rcStatusBarFontInfo.top - (rcSplitter.bottom - rcSplitter.top) - (rcEditMessage.bottom - rcEditMessage.top), LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, FALSE);
									}

									// Resize EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
									if (bIsListViewFontListMinimized)
									{
										MoveWindow(hWndEditMessage, rcEditMessage.left, rcButtonOpen.bottom + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcStatusBarFontInfo.top - rcButtonOpen.bottom - cyListViewFontListMin - (rcSplitter.bottom - rcSplitter.top), FALSE);
									}
									else
									{
										MoveWindow(hWndEditMessage, rcEditMessage.left, rcStatusBarFontInfo.top - (rcEditMessage.bottom - rcEditMessage.top), LOWORD(lParam), rcEditMessage.bottom - rcEditMessage.top, FALSE);
									}
								}
								break;
								// Resize EditMessage and keep the height of ListViewFontList as far as possible
							case WMSZ_BOTTOM:
							case WMSZ_BOTTOMLEFT:
							case WMSZ_BOTTOMRIGHT:
								{
									/*
									┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓      ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
									┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃      ┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃
									┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨      ┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨
									┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃      ┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃
									┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃      ┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃
									┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃      ┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃
									┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨      ┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨
									┃ Font Name                                                         │ State        ┃      ┃ Font Name                                                         │ State        ┃
									┠───────────────────────────────────────────────────────────────────┼──────────────┨      ┠───────────────────────────────────────────────────────────────────┼──────────────┨
									┃                                                                   ┆              ┃      ┃                                                                   ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                   ┆              ┃      ┃                                                                   ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                   ┆              ┃      ┃                                                                   ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                   ┆              ┃      ┃                                                                   ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┠───────────────────────────────────────────────────────────────────┴──────────────┨  =>  ┠───────────────────────────────────────────────────────────────────┴──────────────┨
									┠──────────────────────────────────────────────────────────────────────────────┬───┨      ┠──────────────────────────────────────────────────────────────────────────────┬───┨
									┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃      ┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃
									┃                                                                              ├───┨      ┃                                                                              ├───┨
									┃ How to load fonts to Windows:                                                │▓▓▓┃      ┃ How to load fonts to Windows:                                                │▓▓▓┃
									┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃      ┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃
									┃  view, then click "Load" button.                                             │▓▓▓┃      ┃  view, then click "Load" button.                                             │▓▓▓┃
									┃                                                                              ├───┨      ┃                                                                              │▓▓▓┃
									┃ How to unload fonts from Windows:                                            │   ┃      ┃ How to unload fonts from Windows:                                            │▓▓▓┃
									┃ Select all fonts then click "Unload" or "Close" button or the X at the       │   ┃      ┃ Select all fonts then click "Unload" or "Close" button or the X at the       ├───┨
									┃ upper-right cornor.                                                          │   ┃      ┃ upper-right cornor.                                                          │   ┃
									┃                                                                              │   ┃      ┃                                                                              │   ┃
									┃ How to load fonts to process:                                                │   ┃      ┃ How to load fonts to process:                                                │   ┃
									┃ 1.Click "Click to select process", select a process.                         │   ┃      ┃ 1.Click "Click to select process", select a process.                         │   ┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │   ┃
									┃ view, then click "Load" button.                                              │ ↓ ┃      ┃ view, then click "Load" button.                                              │   ┃
									┠──────────────────────────────────────────────────────────────────────────────┴───┨      ┃ How to unload fonts from process:                                            │   ┃
									┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃	  ┃ Select all fonts then click "Unload" or "Close" button or the X at the       ├───┨
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛	  ┃  upper-right cornor or terminate selected process.                           │ ↓ ┃
																															  ┠──────────────────────────────────────────────────────────────────────────────┴───┨
																															  ┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃
																															  ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
									*/

									// Resize StatusBarFontInfo
									HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, (int)ID::StatusBarFontInfo) };
									FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);
									RECT rcStatusBarFontInfo{};
									GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

									// Calculate the minimal height of EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
									HDC hDCEditMessage{ GetDC(hWndEditMessage) };
									SelectFont(hDCEditMessage, (HFONT)GetWindowFont(hWndEditMessage));
									TEXTMETRIC tm{};
									GetTextMetrics(hDCEditMessage, &tm);
									ReleaseDC(hWndEditMessage, hDCEditMessage);
									RECT rcEditMessage{}, rcEditMessageClient{};
									GetWindowRect(hWndEditMessage, &rcEditMessage);
									GetClientRect(hWndEditMessage, &rcEditMessageClient);
									LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + cyEditMessageTextMargin };

									// Resize EditMessage
									bool bIsEditMessageMinimized{ false };
									MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);
									if (rcStatusBarFontInfo.top - rcEditMessage.top < cyEditMessageMin)
									{
										bIsEditMessageMinimized = true;

										MoveWindow(hWndEditMessage, rcEditMessage.left, rcStatusBarFontInfo.top - cyEditMessageMin, LOWORD(lParam), cyEditMessageMin, FALSE);
									}
									else
									{
										MoveWindow(hWndEditMessage, rcEditMessage.left, rcEditMessage.top, LOWORD(lParam), rcStatusBarFontInfo.top - rcEditMessage.top, FALSE);
									}

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
									RECT rcSplitter{};
									GetWindowRect(hWndSplitter, &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									if (bIsEditMessageMinimized)
									{
										MoveWindow(hWndSplitter, rcSplitter.left, rcStatusBarFontInfo.top - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, FALSE);
									}
									else
									{
										MoveWindow(hWndSplitter, rcSplitter.left, rcSplitter.top, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, FALSE);
									}

									// Resize ListViewFontList
									HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
									RECT rcButtonOpen{}, rcListViewFontList{};
									GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
									GetWindowRect(hWndListViewFontList, &rcListViewFontList);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									if (bIsEditMessageMinimized)
									{
										MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcStatusBarFontInfo.top - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top) - rcButtonOpen.bottom, FALSE);
									}
									else
									{
										MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcListViewFontList.bottom - rcListViewFontList.top, FALSE);
									}
								}
								break;
								// Just modify width
							case WMSZ_LEFT:
							case WMSZ_RIGHT:
								{
									/*
									┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓      ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
									┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃      ┃FontLoaderEx                                                                        ┃_  ┃ □ ┃ x ┃
									┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨      ┠────────┬────────┬────────┬────────┬───────────────────────────────────┬─────────┬──┸───┸───┸───┨
									┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃      ┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │              ┃
									┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃      ┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘              ┃
									┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃      ┃        │        │        │        │     Select Process     │  □ Minimize to tray               ┃
									┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨      ┠────────┴────────┴────────┴────────┴────────────────────────┴────────────────────┬──────────────┨
									┃ Font Name                                                         │ State        ┃      ┃ Font Name                                                                       │ State        ┃
									┠───────────────────────────────────────────────────────────────────┼──────────────┨      ┠─────────────────────────────────────────────────────────────────────────────────┼──────────────┨
									┃                                                                   ┆              ┃      ┃                                                                                 ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                   ┆              ┃      ┃                                                                                 ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                   ┆              ┃      ┃                                                                                 ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                   ┆              ┃      ┃                                                                                 ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┠───────────────────────────────────────────────────────────────────┴──────────────┨  =>  ┠─────────────────────────────────────────────────────────────────────────────────┴──────────────┨
									┠──────────────────────────────────────────────────────────────────────────────┬───┨      ┠────────────────────────────────────────────────────────────────────────────────────────────┬───┨
									┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃      ┃ Temporarily load fonts to Windows or specific process                                      │ ↑ ┃
									┃                                                                              ├───┨      ┃                                                                                            ├───┨
									┃ How to load fonts to Windows:                                                │▓▓▓┃      ┃ How to load fonts to Windows:                                                              │▓▓▓┃
									┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃      ┃ 1.Drag-drop font files onto the icon of this application.                                  │▓▓▓┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list view, then     │▓▓▓┃
									┃  view, then click "Load" button.                                             │▓▓▓┃      ┃  click "Load" button.                                                                      │▓▓▓┃
									┃                                                                              ├───┨      ┃                                                                                            ├───┨
									┃ How to unload fonts from Windows:                                            │   ┃      ┃ How to unload fonts from Windows:                                                          │   ┃
									┃ Select all fonts then click "Unload" or "Close" button or the X at the       │   ┃      ┃ Select all fonts then click "Unload" or "Close" button or the X at the upper-right cornor. │   ┃
									┃ upper-right cornor.                                                          │   ┃      ┃                                                                                            │   ┃
									┃                                                                              │   ┃      ┃ How to load fonts to process:                                                              │   ┃
									┃ How to load fonts to process:                                                │   ┃      ┃ 1.Click "Click to select process", select a process.                                       │   ┃
									┃ 1.Click "Click to select process", select a process.                         │   ┃      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list view, then     │   ┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨      ┃  click "Load" button.                                                                      ├───┨
									┃ view, then click "Load" button.                                              │ ↓ ┃      ┃                                                                                            │ ↓ ┃
									┠──────────────────────────────────────────────────────────────────────────────┴───┨      ┠────────────────────────────────────────────────────────────────────────────────────────────┴───┨
									┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃      ┃ 0 font(s) opened, 0 font(s) loaded.                                                            ┃
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛      ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
									*/

									// Resize StatusBarFontInfo
									HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, (int)ID::StatusBarFontInfo) };
									FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);

									// Resize ListViewFontList
									HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
									RECT rcListViewFontList{};
									GetWindowRect(hWndListViewFontList, &rcListViewFontList);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcListViewFontList.bottom - rcListViewFontList.top, FALSE);

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
									RECT rcSplitter{};
									GetWindowRect(hWndSplitter, &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									MoveWindow(hWndSplitter, rcSplitter.left, rcSplitter.top, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, FALSE);

									// Resize EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
									RECT rcEditMessage{};
									GetWindowRect(hWndEditMessage, &rcEditMessage);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);
									MoveWindow(hWndEditMessage, rcEditMessage.left, rcEditMessage.top, LOWORD(lParam), rcEditMessage.bottom - rcEditMessage.top, FALSE);
								}
								break;
							default:
								break;
							}

							PreviousShowCmd = SW_RESTORE;
						}
						break;
						// Proportionally scale the position of splitter
					case SIZE_MAXIMIZED:
						{
							/*
							┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
							┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃
							┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨
							┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃
							┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃
							┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃
							┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨
							┃ Font Name                                                         │ State        ┃
							┠───────────────────────────────────────────────────────────────────┼──────────────┨
							┃                                                                   ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                   ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                   ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                   ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┠───────────────────────────────────────────────────────────────────┴──────────────┨
							┠──────────────────────────────────────────────────────────────────────────────┬───┨
							┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃
							┃                                                                              ├───┨
							┃ How to load fonts to Windows:                                                │▓▓▓┃
							┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃
							┃  view, then click "Load" button.                                             │▓▓▓┃
							┃                                                                              ├───┨
							┃ How to unload fonts from Windows:                                            │   ┃
							┃ Select all fonts then click "Unload" or "Close" button or the X at the       │   ┃
							┃ upper-right cornor.                                                          │   ┃
							┃                                                                              │   ┃
							┃ How to load fonts to process:                                                │   ┃
							┃ 1.Click "Click to select process", select a process.                         │   ┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨
							┃ view, then click "Load" button.                                              │ ↓ ┃
							┠──────────────────────────────────────────────────────────────────────────────┴───┨
							┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

							┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
							┃FontLoaderEx                                                                                                       ┃_  ┃ □ ┃ x ┃
							┠────────┬────────┬────────┬────────┬───────────────────────────────────┬─────────┬─────────────────────────────────┸───┸───┸───┨
							┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │                                             ┃
							┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘                                             ┃
							┃        │        │        │        │     Select Process     │  □ Minimize to tray                                              ┃
							┠────────┴────────┴────────┴────────┴────────────────────────┴───────────────────────────────────────────────────┬──────────────┨
							┃ Font Name                                                                                                      │ State        ┃
							┠────────────────────────────────────────────────────────────────────────────────────────────────────────────────┼──────────────┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠────────────────────────────────────────────────────────────────────────────────────────────────────────────────┴──────────────┨
							┠───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┬───┨
							┃ Temporarily load fonts to Windows or specific process                                                                     │ ↑ ┃
							┃                                                                                                                           ├───┨
							┃ How to load fonts to Windows:                                                                                             │▓▓▓┃
							┃ 1.Drag-drop font files onto the icon of this application.                                                                 │▓▓▓┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list view, then click "Load" button.               │▓▓▓┃
							┃	                                                                                                                        │▓▓▓┃
							┃ How to unload fonts from Windows:                                                                                         │▓▓▓┃
							┃ Select all fonts then click "Unload" or "Close" button or the X at the upper-right cornor.                                │▓▓▓┃
							┃                                                                                                                           │▓▓▓┃
							┃ How to load fonts to process:                                                                                             ├───┨
							┃ 1.Click "Click to select process", select a process.                                                                      │   ┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list view, then click "Load" button.               │   ┃
							┃                                                                                                                           │   ┃
							┃ How to unload fonts from process:	                                                                                        │   ┃
							┃ Select all fonts then click "Unload" or "Close" button or the X at the upper-right cornor or terminate selected process.	│   ┃
							┃                                                                                                                           │   ┃
							┃ UI description:                                                                                                           │   ┃
							┃ "Open": Add fonts to the list view.                                                                                       │   ┃
							┃ "Close": Remove selected fonts from Windows or target process and the list view.                                          │   ┃
							┃ "Load": Add selected fonts to Windows or target process.                                                                  │   ┃
							┃ "Unload": Remove selected fonts from Windows or target process.	                                                        ├───┨
							┃ "Broadcast WM_FONTCHANGE": If checked, broadcast WM_FONTCHANGE message to all top windows when loading or unloading fonts.│ ↓ ┃
							┠───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┴───┨
							┃ 0 font(s) opened, 0 font(s) loaded.                                                                                           ┃
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
							*/

							// Resize StatusBarFontInfo
							HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, (int)ID::StatusBarFontInfo) };
							FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);
							RECT rcStatusBarFontInfo{};
							GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

							// Resize Splitter
							HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
							RECT rcButtonOpen{}, rcListViewFontList{}, rcSplitter{}, rcEditMessage{};
							GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
							GetWindowRect(GetDlgItem(hWnd, (int)ID::ListViewFontList), &rcListViewFontList);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
							GetWindowRect(hWndSplitter, &rcSplitter);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
							GetWindowRect(GetDlgItem(hWnd, (int)ID::EditMessage), &rcEditMessage);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);

							LONG ySplitterTopNew{ (((rcSplitter.top - rcButtonOpen.bottom) + (rcSplitter.bottom - rcButtonOpen.bottom)) * (rcStatusBarFontInfo.top - rcButtonOpen.bottom)) / ((rcStatusBarFontInfoOld.top - rcButtonOpen.bottom) * 2) - ((rcSplitter.bottom - rcSplitter.top) / 2) + rcButtonOpen.bottom };
							MoveWindow(hWndSplitter, rcSplitter.left, ySplitterTopNew, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, FALSE);

							// Resize ListViewFontList
							HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
							MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), ySplitterTopNew - rcButtonOpen.bottom, FALSE);

							// Resize EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
							MoveWindow(hWndEditMessage, rcEditMessage.left, ySplitterTopNew + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcStatusBarFontInfo.top - (ySplitterTopNew + (rcSplitter.bottom - rcSplitter.top)), FALSE);

							PreviousShowCmd = SW_MAXIMIZE;
						}
						break;
						// Do nothing
					case SIZE_MINIMIZED:
						{
							PreviousShowCmd = SW_MINIMIZE;
						}
						break;
					default:
						break;
					}
				}
				break;
				// If previous window state is maximized
			case SW_MAXIMIZE:
				{
					switch (wParam)
					{
						// Proportionally scale the position of splitter
					case SIZE_RESTORED:
						{
							/*
							┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
							┃FontLoaderEx                                                                                                       ┃_  ┃ □ ┃ x ┃
							┠────────┬────────┬────────┬────────┬───────────────────────────────────┬─────────┬─────────────────────────────────┸───┸───┸───┨
							┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │                                             ┃
							┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘                                             ┃
							┃        │        │        │        │     Select Process     │  □ Minimize to tray                                              ┃
							┠────────┴────────┴────────┴────────┴────────────────────────┴───────────────────────────────────────────────────┬──────────────┨
							┃ Font Name                                                                                                      │ State        ┃
							┠────────────────────────────────────────────────────────────────────────────────────────────────────────────────┼──────────────┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	                                                                                                             │              ┃
							┠────────────────────────────────────────────────────────────────────────────────────────────────────────────────┴──────────────┨  =>
							┠───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┬───┨
							┃ Temporarily load fonts to Windows or specific process                                                                     │ ↑ ┃
							┃                                                                                                                           ├───┨
							┃ How to load fonts to Windows:                                                                                             │▓▓▓┃
							┃ 1.Drag-drop font files onto the icon of this application.                                                                 │▓▓▓┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list view, then click "Load" button.               │▓▓▓┃
							┃	                                                                                                                        │▓▓▓┃
							┃ How to unload fonts from Windows:                                                                                         │▓▓▓┃
							┃ Select all fonts then click "Unload" or "Close" button or the X at the upper-right cornor.                                │▓▓▓┃
							┃                                                                                                                           │▓▓▓┃
							┃ How to load fonts to process:                                                                                             ├───┨
							┃ 1.Click "Click to select process", select a process.                                                                      │   ┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list view, then click "Load" button.               │   ┃
							┃                                                                                                                           │   ┃
							┃ How to unload fonts from process:	                                                                                        │   ┃
							┃ Select all fonts then click "Unload" or "Close" button or the X at the upper-right cornor or terminate selected process.	│   ┃
							┃                                                                                                                           │   ┃
							┃ UI description:                                                                                                           │   ┃
							┃ "Open": Add fonts to the list view.                                                                                       │   ┃
							┃ "Close": Remove selected fonts from Windows or target process and the list view.                                          │   ┃
							┃ "Load": Add selected fonts to Windows or target process.                                                                  │   ┃
							┃ "Unload": Remove selected fonts from Windows or target process.	                                                        ├───┨
							┃ "Broadcast WM_FONTCHANGE": If checked, broadcast WM_FONTCHANGE message to all top windows when loading or unloading fonts.│ ↓ ┃
							┠───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┴───┨
							┃ 0 font(s) opened, 0 font(s) loaded.                                                                                           ┃
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

							┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
							┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃
							┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨
							┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃
							┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃
							┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃
							┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨
							┃ Font Name                                                         │ State        ┃
							┠───────────────────────────────────────────────────────────────────┼──────────────┨
							┃                                                                   ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                   ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                   ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                   ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┠───────────────────────────────────────────────────────────────────┴──────────────┨
							┠──────────────────────────────────────────────────────────────────────────────┬───┨
							┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃
							┃                                                                              ├───┨
							┃ How to load fonts to Windows:                                                │▓▓▓┃
							┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃
							┃  view, then click "Load" button.                                             │▓▓▓┃
							┃                                                                              ├───┨
							┃ How to unload fonts from Windows:                                            │   ┃
							┃ Select all fonts then click "Unload" or "Close" button or the X at the       │   ┃
							┃ upper-right cornor.                                                          │   ┃
							┃                                                                              │   ┃
							┃ How to load fonts to process:                                                │   ┃
							┃ 1.Click "Click to select process", select a process.                         │   ┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨
							┃ view, then click "Load" button.                                              │ ↓ ┃
							┠──────────────────────────────────────────────────────────────────────────────┴───┨
							┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
							*/

							// Resize StatusBarFontInfo
							HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, (int)ID::StatusBarFontInfo) };
							FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);
							RECT rcStatusBarFontInfo{};
							GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

							// Calculate the minimal height of ListViewFontList
							HWND hWndListViewFontList{ GetDlgItem(hWnd,(int)ID::ListViewFontList) };
							RECT rcListViewFontList{}, rcListViewFontListClient{};
							GetWindowRect(hWndListViewFontList, &rcListViewFontList);
							GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
							bool bIsInserted{ false };
							if (ListView_GetItemCount(hWndListViewFontList) == 0)
							{
								bIsInserted = true;

								LVITEM lvi{};
								ListView_InsertItem(hWndListViewFontList, &lvi);
							}
							RECT rcListViewFontListItem{};
							ListView_GetItemRect(hWndListViewFontList, 0, &rcListViewFontListItem, LVIR_BOUNDS);
							if (bIsInserted)
							{
								ListView_DeleteAllItems(hWndListViewFontList);
							}
							LONG cyListViewFontListMin{ (rcListViewFontListItem.bottom - rcListViewFontListClient.top) + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

							// Calculate the minimal height of EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
							HDC hDCEditMessage{ GetDC(hWndEditMessage) };
							SelectFont(hDCEditMessage, (HFONT)GetWindowFont(hWndEditMessage));
							TEXTMETRIC tm{};
							GetTextMetrics(hDCEditMessage, &tm);
							ReleaseDC(hWndEditMessage, hDCEditMessage);
							RECT rcEditMessage{}, rcEditMessageClient{};
							GetWindowRect(hWndEditMessage, &rcEditMessage);
							GetClientRect(hWndEditMessage, &rcEditMessageClient);
							LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + cyEditMessageTextMargin };

							// Calculate new position of splitter
							HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
							RECT rcButtonOpen{}, rcSplitter{};
							GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
							GetWindowRect(hWndSplitter, &rcSplitter);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
							LONG ySplitterTopNew{ (((rcSplitter.top - rcButtonOpen.bottom) + (rcSplitter.bottom - rcButtonOpen.bottom)) * (rcStatusBarFontInfo.top - rcButtonOpen.bottom)) / ((rcStatusBarFontInfoOld.top - rcButtonOpen.bottom) * 2) - ((rcSplitter.bottom - rcSplitter.top) / 2) + rcButtonOpen.bottom };

							MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);
							LONG cyListViewFontListNew{ ySplitterTopNew - rcButtonOpen.bottom };
							LONG cyEditMessageNew{ rcStatusBarFontInfo.top - ySplitterTopNew - (rcSplitter.bottom - rcSplitter.top) };
							// If cyListViewFontListNew < cyListViewFontListMin, keep the minimal height of ListViewFontList
							if (cyListViewFontListNew < cyListViewFontListMin)
							{
								// Resize ListViewFontList
								HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
								MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), cyListViewFontListMin, FALSE);

								// Resize Splitter
								MoveWindow(hWndSplitter, rcSplitter.left, rcButtonOpen.bottom + cyListViewFontListMin, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, FALSE);

								// Resize EditMessage
								HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
								MoveWindow(hWndEditMessage, rcEditMessage.left, rcButtonOpen.bottom + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcStatusBarFontInfo.top - rcButtonOpen.bottom - cyListViewFontListMin - (rcSplitter.bottom - rcSplitter.top), FALSE);
							}
							// If cyEditMessageNew < cyEditMessageMin, keep the minimal height of EditMessage
							else if (cyEditMessageNew < cyEditMessageMin)
							{
								// Resize EditMessage
								HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
								MoveWindow(hWndEditMessage, rcEditMessage.left, rcStatusBarFontInfo.top - cyEditMessageMin, LOWORD(lParam), cyEditMessageMin, FALSE);

								// Resize Splitter
								MoveWindow(hWndSplitter, rcSplitter.left, rcStatusBarFontInfo.top - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, FALSE);

								// Resize ListViewFontList
								HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
								MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcStatusBarFontInfo.top - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top) - rcButtonOpen.bottom, FALSE);
							}
							// Else resize as usual
							else
							{
								// Resize Splitter
								MoveWindow(hWndSplitter, rcSplitter.left, ySplitterTopNew, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, FALSE);

								// Resize ListViewFontList
								HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
								MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), cyListViewFontListNew, FALSE);

								// Resize EditMessage
								HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
								MoveWindow(hWndEditMessage, rcEditMessage.left, ySplitterTopNew + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), cyEditMessageNew, FALSE);
							}

							PreviousShowCmd = SW_RESTORE;
						}
						break;
						// Do nothing
					case SIZE_MINIMIZED:
						{
							PreviousShowCmd = SW_MINIMIZE;
						}
						break;
					default:
						break;
					}
				}
				break;
				// If previous window state is minimized
			case SW_MINIMIZE:
			case SW_FORCEMINIMIZE:
			case SW_SHOWMINIMIZED:
				{
					// Do nothing
					switch (wParam)
					{
					case SIZE_RESTORED:
						{
							PreviousShowCmd = SW_RESTORE;
						}
						break;
					case SIZE_MAXIMIZED:
						{
							PreviousShowCmd = SW_MAXIMIZE;
						}
						break;
					default:
						break;
					}
				}
				break;
			default:
				break;
			}

			// Redraw controls
			InvalidateRect(hWnd, NULL, FALSE);
		}
		break;
	case WM_GETMINMAXINFO:
		{
			// Limit minimal window size
			/*
			┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
			┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃
			┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸───╂───────
			┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     ┃       ↑
			┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┨       │ cyButtonOpen
			┃        │        │        │        │     Select Process     │  □ Minimize to tray┃       ↓
			┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────╂───────
			┃ Font Name                                                        │ State        ┃       ↑
			┠──────────────────────────────────────────────────────────────────┼──────────────┨       │ cyListViewFontListMin
			┃                                                                  ┆              ┃       ↓
			┠──────────────────────────────────────────────────────────────────┴──────────────╂───────
			┠─────────────────────────────────────────────────────────────────────────────┬───╂───────
			┃ Temporarily load fonts to Windows or specific process                       ├───┨       } cyEditMessagemin
			┠─────────────────────────────────────────────────────────────────────────────┴───┨───────
			┃ 0 font(s) opened, 0 font(s) loaded.                                             ┃       } cyStatusBarFontList
			┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛───────
			│                               rcEditTimeout.right                               │
			│←───────────────────────────────────────────────────────────────────────────────→│
			│                                                                                 │
			*/

			// Get ButtonOpen, Splitter, StatusBarFontInfo and EditTimeout window rectangle
			RECT rcButtonOpen{}, rcSplitter{}, rcStatusBarFontInfo{}, rcEditTimeout{};
			GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
			GetWindowRect(GetDlgItem(hWnd, (int)ID::Splitter), &rcSplitter);
			GetWindowRect(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), &rcStatusBarFontInfo);
			GetWindowRect(GetDlgItem(hWnd, (int)ID::EditTimeout), &rcEditTimeout);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcEditTimeout);

			// Calculate the minimal height of ListViewFontList
			HWND hWndListViewFontList{ GetDlgItem(hWnd,(int)ID::ListViewFontList) };
			RECT rcListViewFontList{}, rcListViewFontListClient{};
			GetWindowRect(hWndListViewFontList, &rcListViewFontList);
			GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
			bool bIsInserted{ false };
			if (ListView_GetItemCount(hWndListViewFontList) == 0)
			{
				bIsInserted = true;

				LVITEM lvi{};
				ListView_InsertItem(hWndListViewFontList, &lvi);
			}
			RECT rcListViewFontListItem{};
			ListView_GetItemRect(hWndListViewFontList, 0, &rcListViewFontListItem, LVIR_BOUNDS);
			if (bIsInserted)
			{
				ListView_DeleteAllItems(hWndListViewFontList);
			}
			LONG cyListViewFontListMin{ (rcListViewFontListItem.bottom - rcListViewFontListClient.top) + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

			// Calculate the minimal height of Editmessage
			HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
			HDC hDCEditMessage{ GetDC(hWndEditMessage) };
			SelectFont(hDCEditMessage, (HFONT)GetWindowFont(hWndEditMessage));
			TEXTMETRIC tm{};
			GetTextMetrics(hDCEditMessage, &tm);
			ReleaseDC(hWndEditMessage, hDCEditMessage);
			RECT rcEditMessage{}, rcEditMessageClient{};
			GetWindowRect(hWndEditMessage, &rcEditMessage);
			GetClientRect(hWndEditMessage, &rcEditMessageClient);
			LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + cyEditMessageTextMargin };

			// Calculate the minimal window size
			RECT rcMainMin{ 0, 0, rcEditTimeout.right, (rcButtonOpen.bottom - rcButtonOpen.top) + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top) + cyEditMessageMin + (rcStatusBarFontInfo.bottom - rcStatusBarFontInfo.top) };
			AdjustWindowRect(&rcMainMin, (DWORD)GetWindowLongPtr(hWnd, GWL_STYLE), FALSE);
			((LPMINMAXINFO)lParam)->ptMinTrackSize = { rcMainMin.right - rcMainMin.left, rcMainMin.bottom - rcMainMin.top };
		}
		break;
	case WM_CLOSE:
		{
			// Do cleanup
			HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };

			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{};

			// If loaded via proxy
			if (ProxyProcessInfo.hProcess)
			{
				// Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
				COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::PULLDLL, 0, NULL };
				FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
				WaitForSingleObject(hEventProxyDllPullingFinished, INFINITE);
				CloseHandle(hEventProxyDllPullingFinished);
				switch (ProxyDllPullingResult)
				{
				case PROXYDLLPULL::SUCCESSFUL:
					goto continue_69504405;
				case PROXYDLLPULL::FAILED:
					{
						// Print message
						ssMessage << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
						strMessage = ssMessage.str();
						cchMessageLength = Edit_GetTextLength(hWndEditMessage);
						Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
						Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

						// Re-enable controls
						EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), TRUE);
						EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), TRUE);
						EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
						EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), TRUE);
						EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), TRUE);
						EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), TRUE);

						// Update StatusBarFontInfo
						std::wstringstream ssFontInfo{};
						std::wstring strFontInfo{};
						std::size_t nLoadedFonts{};
						for (const auto& i : FontList)
						{
							if (i.IsLoaded())
							{
								nLoadedFonts++;
							}
						}
						ssFontInfo << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
						strFontInfo = ssFontInfo.str();
						SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), strFontInfo.c_str());

						// Update syatem tray icon tip
						if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
						{
							NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
							wcscpy_s(nid.szTip, strFontInfo.c_str());
							Shell_NotifyIcon(NIM_MODIFY, &nid);
						}
					}
					goto break_69504405;
				default:
					break;
				}
			break_69504405:
				break;
			continue_69504405:

				// Terminate watch thread
				SetEvent(hEventTerminateWatchThread);
				WaitForSingleObject(hThreadWatch, INFINITE);
				CloseHandle(hEventTerminateWatchThread);
				CloseHandle(hThreadWatch);

				// Terminate message thread
				PostMessage(hWndMessage, WM_CLOSE, 0, 0);
				WaitForSingleObject(hThreadMessage, INFINITE);
				DWORD dwMessageThreadExitCode{};
				GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
				if (dwMessageThreadExitCode)
				{
					std::wstringstream ssMessage{};
					std::wstring strMessage{};
					ssMessage << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
					strMessage = ssMessage.str();
					MessageBoxCentered(NULL, strMessage.c_str(), szAppName, MB_ICONERROR);
				}
				CloseHandle(hThreadMessage);

				// Terminate proxy process
				COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
				FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
				WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
				CloseHandle(ProxyProcessInfo.hProcess);
				ProxyProcessInfo.hProcess = NULL;

				// Close HANDLE to target process and duplicated handles
				CloseHandle(TargetProcessInfo.hProcess);
				TargetProcessInfo.hProcess = NULL;
				CloseHandle(hProcessCurrentDuplicated);
				CloseHandle(hProcessTargetDuplicated);
			}
			// Else DIY
			else if (TargetProcessInfo.hProcess)
			{
				// Unload FontLoaderExInjectionDll(64).dll from target process
				if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
				{
					// Print message
					ssMessage << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
					strMessage = ssMessage.str();
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

					// Re-enable controls
					EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonOpen), TRUE);
					EnableWindow(GetDlgItem(hWnd, (int)ID::ListViewFontList), TRUE);
					EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
					EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonClose), TRUE);
					EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonLoad), TRUE);
					EnableWindow(GetDlgItem(hWnd, (int)ID::ButtonUnload), TRUE);

					// Update StatusBarFontInfo
					std::wstringstream ssFontInfo{};
					std::wstring strFontInfo{};
					std::size_t nLoadedFonts{};
					for (const auto& i : FontList)
					{
						if (i.IsLoaded())
						{
							nLoadedFonts++;
						}
					}
					ssFontInfo << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
					strFontInfo = ssFontInfo.str();
					SetWindowText(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), strFontInfo.c_str());

					// Update syatem tray icon tip
					if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
					{
						NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
						wcscpy_s(nid.szTip, strFontInfo.c_str());
						Shell_NotifyIcon(NIM_MODIFY, &nid);
					}
					break;
				}

				// Terminate watch thread
				SetEvent(hEventTerminateWatchThread);
				WaitForSingleObject(hThreadWatch, INFINITE);
				CloseHandle(hEventTerminateWatchThread);
				CloseHandle(hThreadWatch);

				// Close HANDLE to target process
				CloseHandle(TargetProcessInfo.hProcess);
				TargetProcessInfo.hProcess = NULL;
			}

			// Remove the icon from system tray
			if (Button_GetCheck(GetDlgItem(hWnd, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
			{
				NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWnd, 0 };
				Shell_NotifyIcon(NIM_DELETE, &nid);
			}

			ret = DefWindowProc(hWnd, Message, wParam, lParam);
		}
		break;
	case WM_DESTROY:
		{
			DeleteFont(hFontMain);

			PostQuitMessage(0);
		}
		break;
	default:
		{
			ret = DefWindowProc(hWnd, Message, wParam, lParam);
		}
		break;
	}

	return ret;
}

LRESULT CALLBACK EditTimeoutSubclassProc(HWND hWndEditTimeout, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData)
{
	LRESULT ret{};

	switch (Message)
	{
	case WM_PASTE:
		{
			// Play beep sound when pasting invalid text
			if (IsClipboardFormatAvailable(CF_UNICODETEXT))
			{
				if (OpenClipboard(NULL))
				{
					HANDLE hClipboardData{ GetClipboardData(CF_UNICODETEXT) };
					LPCWSTR lpszText{ (LPCWSTR)GlobalLock(hClipboardData) };

					bool bIsNumeric{ true };
					std::wstring_view wstr_v{ lpszText };
					for (const auto& i : wstr_v)
					{
						if (!std::iswdigit(i))
						{
							bIsNumeric = false;

							break;
						}
					}

					GlobalUnlock(hClipboardData);
					CloseClipboard();

					ret = DefSubclassProc(hWndEditTimeout, Message, wParam, lParam);

					if (!bIsNumeric)
					{
						MessageBeep(0xFFFFFFFF);
					}
				}
				else
				{
					ret = DefSubclassProc(hWndEditTimeout, Message, wParam, lParam);
				}
			}
			else
			{
				ret = DefSubclassProc(hWndEditTimeout, Message, wParam, lParam);

				MessageBeep(0xFFFFFFFF);
			}
		}
		break;
	default:
		{
			ret = DefSubclassProc(hWndEditTimeout, Message, wParam, lParam);
		}
		break;
	}

	return ret;
}

LRESULT CALLBACK ListViewFontListSubclassProc(HWND hWndListViewFontList, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData)
{
	LRESULT ret{};

	switch (Message)
	{
	case WM_KEYDOWN:
		{
			// Select all items in ListViewFontList when pressing CTRL+A
			if (wParam == 0x41)	// Virtual key code of 'A' key
			{
				if (GetKeyState(VK_CONTROL))
				{
					ListView_SetItemState(hWndListViewFontList, -1, LVIS_SELECTED, LVIS_SELECTED);
				}
			}

			ret = DefSubclassProc(hWndListViewFontList, Message, wParam, lParam);
		}
		break;
	case WM_DROPFILES:
		{
			// Process drag-drop and open fonts
			HWND hWndParent{ GetParent(hWndListViewFontList) };
			HWND hWndEditMessage{ GetDlgItem(hWndParent, (int)ID::EditMessage) };

			bool bIsFontListEmptyBefore{ FontList.empty() };

			UINT nFileCount{ DragQueryFile((HDROP)wParam, 0xFFFFFFFF, NULL, 0) };
			FONTLISTCHANGEDSTRUCT flcs{ ListView_GetItemCount(hWndListViewFontList) };
			for (UINT i = 0; i < nFileCount; i++)
			{
				WCHAR szFileName[MAX_PATH]{};
				DragQueryFile((HDROP)wParam, i, szFileName, MAX_PATH);
				if (PathMatchSpec(szFileName, L"*.ttf") || PathMatchSpec(szFileName, L"*.ttc") || PathMatchSpec(szFileName, L"*.otf"))
				{
					FontList.push_back(szFileName);

					flcs.lpszFontName = szFileName;
					SendMessage(GetParent(hWndListViewFontList), (UINT)USERMESSAGE::FONTLISTCHANGED, (WPARAM)FONTLISTCHANGED::OPENED, (LPARAM)&flcs);

					flcs.iItem++;
				}
			}
			int cchMessageLength{ Edit_GetTextLength(hWndEditMessage) };
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, L"\r\n");

			DragFinish((HDROP)wParam);

			if (bIsFontListEmptyBefore && (!FontList.empty()))
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_ENABLED);
			}

			// Update StatusBarFontInfo
			std::wstringstream ssFontInfo{};
			std::wstring strFontInfo{};
			std::size_t nLoadedFonts{};
			for (const auto& i : FontList)
			{
				if (i.IsLoaded())
				{
					nLoadedFonts++;
				}
			}
			ssFontInfo << FontList.size() << L" font(s) opened, " << nLoadedFonts << L" font(s) loaded.";
			strFontInfo = ssFontInfo.str();
			SetWindowText(GetDlgItem(hWndParent, (int)ID::StatusBarFontInfo), strFontInfo.c_str());

			// Update syatem tray icon tip
			if (Button_GetCheck(GetDlgItem(hWndParent, (int)ID::ButtonMinimizeToTray)) == BST_CHECKED)
			{
				NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
				wcscpy_s(nid.szTip, strFontInfo.c_str());
				Shell_NotifyIcon(NIM_MODIFY, &nid);
			}
		}
		break;
	case WM_WINDOWPOSCHANGED:
		{
			// Post USERMESSAGE::CHILDWINDOWPOSCHANGED to parent window
			ret = DefSubclassProc(hWndListViewFontList, Message, wParam, lParam);

			PostMessage(GetParent(hWndListViewFontList), (UINT)USERMESSAGE::CHILDWINDOWPOSCHANGED, (WPARAM)GetDlgCtrlID(hWndListViewFontList), NULL);
		}
		break;
	default:
		{
			ret = DefSubclassProc(hWndListViewFontList, Message, wParam, lParam);
		}
		break;
	}

	return ret;
}

LRESULT CALLBACK EditMessageSubclassProc(HWND hWndEditMessage, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData)
{
	LRESULT ret{};

	switch (Message)
	{
	case WM_KEYDOWN:
		{
			// Select all text in EditMessage when pressing CTRL+A
			if (wParam == 0x41)	// Virtual key code of 'A' key
			{
				if (GetKeyState(VK_CONTROL))
				{
					Edit_SetSel(hWndEditMessage, 0, Edit_GetTextLength(hWndEditMessage));
				}
			}

			ret = DefSubclassProc(hWndEditMessage, Message, wParam, lParam);
		}
		break;
	case WM_CONTEXTMENU:
		{
			// Delete "Undo", "Cut", "Paste" and "Clear" from context menu
			static HWND hWndOwner{};
			static POINT ptCursor{};
			hWndOwner = (HWND)wParam;
			ptCursor = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };

			HWINEVENTHOOK hWinEventHook{ SetWinEventHook(EVENT_SYSTEM_MENUPOPUPSTART, EVENT_SYSTEM_MENUPOPUPSTART, NULL,
				[](HWINEVENTHOOK hWinEventHook, DWORD Event, HWND hWnd, LONG idObject, LONG idChild, DWORD idEventThread, DWORD dwmsEventTime)
				{
					if (idObject == OBJID_CLIENT && idChild == CHILDID_SELF)
					{
						// Test whether context menu is triggered by mouse or by keyboard
						if (ptCursor.x == -1 && ptCursor.y == -1)
						{
							LRESULT lRet{ SendMessage(hWndOwner, EM_POSFROMCHAR, Edit_GetCaretIndex(hWndOwner), NULL) };
							ptCursor = { LOWORD(lRet), HIWORD(lRet) };
							ClientToScreen(hWndOwner, &ptCursor);
						}

						// Test whether context menu pops up in the client area of EditMessage
						RECT rcEditMessageClient{};
						GetClientRect(hWndOwner, &rcEditMessageClient);
						MapWindowRect(hWndOwner, HWND_DESKTOP, &rcEditMessageClient);
						if (PtInRect(&rcEditMessageClient, ptCursor))
						{
							HMENU hMenuContextEdit{ (HMENU)SendMessage(hWnd, MN_GETHMENU, 0, 0) };

							// Menu item identifiers in Edit control context menu is the same as corresponding Windows messages
							DeleteMenu(hMenuContextEdit, WM_UNDO, MF_BYCOMMAND);	// Undo
							DeleteMenu(hMenuContextEdit, WM_CUT, MF_BYCOMMAND);		// Cut
							DeleteMenu(hMenuContextEdit, WM_PASTE, MF_BYCOMMAND);	// Paste
							DeleteMenu(hMenuContextEdit, WM_CLEAR, MF_BYCOMMAND);	// Clear
							DeleteMenu(hMenuContextEdit, 0, MF_BYPOSITION);			// Seperator

							// Adjust context menu position
							RECT rcContextMenuEdit{};
							GetWindowRect(hWnd, &rcContextMenuEdit);
							MONITORINFO mi{ sizeof(MONITORINFO) };
							GetMonitorInfo(MonitorFromPoint(ptCursor, MONITOR_DEFAULTTONEAREST), &mi);
							SIZE sizeContextMenu{ rcContextMenuEdit.right - rcContextMenuEdit.left, rcContextMenuEdit.bottom - rcContextMenuEdit.top };
							UINT uFlags{ TPM_WORKAREA };
							if (ptCursor.y > mi.rcWork.bottom - sizeContextMenu.cy)
							{
								uFlags |= TPM_BOTTOMALIGN;
							}
							else
							{
								uFlags |= TPM_TOPALIGN;
							}
							if (GetSystemMetrics(SM_MENUDROPALIGNMENT))
							{
								uFlags |= TPM_RIGHTALIGN;
							}
							else
							{
								uFlags |= TPM_LEFTALIGN;
							}
							CalculatePopupWindowPosition(&ptCursor, &sizeContextMenu, uFlags, NULL, &rcContextMenuEdit);
							MoveWindow(hWnd, rcContextMenuEdit.left, rcContextMenuEdit.top, sizeContextMenu.cx, sizeContextMenu.cy, FALSE);
						}
					}
				},
				GetCurrentProcessId(), GetCurrentThreadId(), WINEVENT_OUTOFCONTEXT) };

			ret = DefSubclassProc(hWndEditMessage, Message, wParam, lParam);

			UnhookWinEvent(hWinEventHook);
		}
		break;
	case WM_WINDOWPOSCHANGED:
		{
			// Post USERMESSAGE::CHILDWINDOWPOSCHANGED to parent window
			ret = DefSubclassProc(hWndEditMessage, Message, wParam, lParam);

			PostMessage(GetParent(hWndEditMessage), (UINT)USERMESSAGE::CHILDWINDOWPOSCHANGED, (WPARAM)GetDlgCtrlID(hWndEditMessage), NULL);
		}
		break;
	default:
		{
			ret = DefSubclassProc(hWndEditMessage, Message, wParam, lParam);
		}
		break;
	}

	return ret;
}

INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT Message, WPARAM wParam, LPARAM lParam)
{
	INT_PTR ret{};

	static HFONT hFontDialog{};

	static std::vector<ProcessInfo> ProcessList{};

	static bool bOrderByProcessAscending{ true };
	static bool bOrderByPIDAscending{ true };

	switch (Message)
	{
	case WM_INITDIALOG:
		{
			bOrderByProcessAscending = true;
			bOrderByPIDAscending = true;

			NONCLIENTMETRICS ncm{ sizeof(NONCLIENTMETRICS) };
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
			hFontDialog = CreateFontIndirect(&ncm.lfMessageFont);

			// Initialize ListViewProcessList
			HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };
			SetWindowLongPtr(hWndListViewProcessList, GWL_STYLE, GetWindowLongPtr(hWndListViewProcessList, GWL_STYLE) | LVS_REPORT | LVS_SINGLESEL);
			ListView_SetExtendedListViewStyle(hWndListViewProcessList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
			SetWindowFont(hWndListViewProcessList, hFontDialog, TRUE);

			LVCOLUMN lvc1{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, 1 , (LPWSTR)L"Process" };
			LVCOLUMN lvc2{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, 1 , (LPWSTR)L"PID" };
			ListView_InsertColumn(hWndListViewProcessList, 0, &lvc1);
			ListView_InsertColumn(hWndListViewProcessList, 1, &lvc2);

			// Initialize ButtonOK
			HWND hWndButtonOK{ GetDlgItem(hWndDialog, IDOK) };
			SetWindowFont(hWndButtonOK, hFontDialog, TRUE);

			// Initialize ButtonCancel
			HWND hWndButtonCancel{ GetDlgItem(hWndDialog, IDCANCEL) };
			SetWindowFont(hWndButtonCancel, hFontDialog, TRUE);

			// Fill ProcessList
			ProcessList.clear();
			PROCESSENTRY32 pe32{ sizeof(PROCESSENTRY32) };
			HANDLE hProcessSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) };
			LVITEM lvi{ LVIF_TEXT, 0 };
			if (Process32First(hProcessSnapshot, &pe32))
			{
				do
				{
					lvi.iSubItem = 0;
					lvi.pszText = pe32.szExeFile;
					ListView_InsertItem(hWndListViewProcessList, &lvi);
					lvi.iSubItem = 1;
					std::wstring str{ std::to_wstring(pe32.th32ProcessID) };
					lvi.pszText = (LPWSTR)str.c_str();
					ListView_SetItem(hWndListViewProcessList, &lvi);
					lvi.iItem++;

					ProcessList.push_back({ NULL, pe32.szExeFile, pe32.th32ProcessID });
				} while (Process32Next(hProcessSnapshot, &pe32));
			}

			// Set the header width in ListViewProcessList
			RECT rcListViewProcessListClient{};
			GetClientRect(hWndListViewProcessList, &rcListViewProcessListClient);
			ListView_SetColumnWidth(hWndListViewProcessList, 0, (rcListViewProcessListClient.right - rcListViewProcessListClient.left) * 4 / 5);
			ListView_SetColumnWidth(hWndListViewProcessList, 1, (rcListViewProcessListClient.right - rcListViewProcessListClient.left) * 1 / 5);

			ret = (INT_PTR)TRUE;
		}
		break;
	case WM_COMMAND:
		{
			switch (LOWORD(wParam))
			{
				// Return selected process
			case IDOK:
				{
					HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };

					int iSelected{ ListView_GetSelectionMark(hWndListViewProcessList) };
					if (iSelected == -1)
					{
						EndDialog(hWndDialog, NULL);
					}
					else
					{
						ProcessInfo* TargetProcessInfo = new ProcessInfo{ ProcessList[iSelected] };

						EndDialog(hWndDialog, (INT_PTR)TargetProcessInfo);
					}

					ret = (INT_PTR)TRUE;
				}
				break;
				// Return null
			case IDCANCEL:
				{
					EndDialog(hWndDialog, NULL);

					ret = (INT_PTR)TRUE;
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_NOTIFY:
		{
			switch (wParam)
			{
			case IDC_LIST1:
				{
					switch (((LPNMLISTVIEW)lParam)->hdr.code)
					{
						// Sort items
					case LVN_COLUMNCLICK:
						{
							HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };
							HWND hWndHeaderListViewProcessList{ ListView_GetHeader(hWndListViewProcessList) };

							HDITEM hdi{ HDI_FORMAT };

							switch (((LPNMLISTVIEW)lParam)->iSubItem)
							{
							case 0:
								{
									// Sort items by Process
									bOrderByProcessAscending ? std::sort(ProcessList.begin(), ProcessList.end(),
										[](const ProcessInfo& value1, const ProcessInfo& value2) -> bool
										{
											int i{ lstrcmpi(value1.strProcessName.c_str(), value2.strProcessName.c_str()) };
											if (i < 0)
											{
												return true;
											}
											else
											{
												return false;
											}
										}) : std::sort(ProcessList.begin(), ProcessList.end(),
											[](const ProcessInfo& value1, const ProcessInfo& value2) -> bool
											{
												int i{ lstrcmpi(value2.strProcessName.c_str(), value1.strProcessName.c_str()) };
												if (i < 0)
												{
													return true;
												}
												else
												{
													return false;
												}
											});

										// Add an arrow to the header in ListViewProcessList
										Header_GetItem(hWndHeaderListViewProcessList, 1, &hdi);
										hdi.fmt = hdi.fmt & (~(HDF_SORTDOWN | HDF_SORTUP));
										Header_SetItem(hWndHeaderListViewProcessList, 1, &hdi);
										Header_GetItem(hWndHeaderListViewProcessList, 0, &hdi);
										bOrderByProcessAscending ? hdi.fmt = (hdi.fmt & (~HDF_SORTDOWN)) | HDF_SORTUP : hdi.fmt = (hdi.fmt & (~HDF_SORTUP)) | HDF_SORTDOWN;
										Header_SetItem(hWndHeaderListViewProcessList, 0, &hdi);

										bOrderByProcessAscending = !bOrderByProcessAscending;
								}
								break;
							case 1:
								{
									// Sort items by PID
									bOrderByPIDAscending ? std::sort(ProcessList.begin(), ProcessList.end(),
										[](const ProcessInfo& value1, const ProcessInfo& value2) -> bool
										{
											return value1.dwProcessID < value2.dwProcessID;
										}) : std::sort(ProcessList.begin(), ProcessList.end(),
											[](const ProcessInfo& value1, const ProcessInfo& value2) -> bool
											{
												return value1.dwProcessID > value2.dwProcessID;
											});

										// Add an arrow to the header in ListViewProcessList
										Header_GetItem(hWndHeaderListViewProcessList, 0, &hdi);
										hdi.fmt = hdi.fmt & (~(HDF_SORTDOWN | HDF_SORTUP));
										Header_SetItem(hWndHeaderListViewProcessList, 0, &hdi);
										Header_GetItem(hWndHeaderListViewProcessList, 1, &hdi);
										bOrderByPIDAscending ? hdi.fmt = (hdi.fmt & (~HDF_SORTDOWN)) | HDF_SORTUP : hdi.fmt = (hdi.fmt & (~HDF_SORTUP)) | HDF_SORTDOWN;
										Header_SetItem(hWndHeaderListViewProcessList, 1, &hdi);

										bOrderByPIDAscending = !bOrderByPIDAscending;
								}
								break;
							default:
								break;
							}

							// Reset the contents of ListViewProcessList
							LVITEM lvi{ LVIF_TEXT, 0 };
							for (auto&& i : ProcessList)
							{
								lvi.iSubItem = 0;
								lvi.pszText = (LPWSTR)i.strProcessName.c_str();
								ListView_SetItem(hWndListViewProcessList, &lvi);
								lvi.iSubItem = 1;
								std::wstring strPID{ std::to_wstring(i.dwProcessID) };
								lvi.pszText = (LPWSTR)strPID.c_str();
								ListView_SetItem(hWndListViewProcessList, &lvi);
								lvi.iItem++;
							}

							ret = (INT_PTR)TRUE;
						}
						break;
						// Select double-clicked item
					case NM_DBLCLK:
						{
							// Simulate clicking "OK" button
							SendMessage(GetDlgItem(hWndDialog, IDOK), BM_CLICK, 0, 0);

							ret = (INT_PTR)TRUE;
						}
						break;
					default:
						break;
					}
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_DESTROY:
		{
			DeleteFont(hFontDialog);

			ret = (INT_PTR)FALSE;
		}
		break;
	default:
		break;
	}

	return ret;
}

int MessageBoxCentered(HWND hWnd, LPCTSTR lpText, LPCTSTR lpCaption, UINT uType)
{
	// Center message box at its parent window
	static HHOOK hHookCBT{};
	hHookCBT = SetWindowsHookEx(WH_CBT,
		[](int nCode, WPARAM wParam, LPARAM lParam) -> LRESULT
		{
			if (nCode == HCBT_CREATEWND)
			{
				if (((LPCBT_CREATEWND)lParam)->lpcs->lpszClass == (LPWSTR)(ATOM)32770)	// #32770 = dialog box class
				{
					RECT rcParent{};
					GetWindowRect(((LPCBT_CREATEWND)lParam)->lpcs->hwndParent, &rcParent);
					((LPCBT_CREATEWND)lParam)->lpcs->x = rcParent.left + ((rcParent.right - rcParent.left) - ((LPCBT_CREATEWND)lParam)->lpcs->cx) / 2;
					((LPCBT_CREATEWND)lParam)->lpcs->y = rcParent.top + ((rcParent.bottom - rcParent.top) - ((LPCBT_CREATEWND)lParam)->lpcs->cy) / 2;
				}
			}

			return CallNextHookEx(hHookCBT, nCode, wParam, lParam);
		},
		0, GetCurrentThreadId());

	int iRet{ MessageBox(hWnd, lpText, lpCaption, uType) };

	UnhookWindowsHookEx(hHookCBT);

	return iRet;
}

bool EnableDebugPrivilege()
{
	// Enable SeDebugPrivilege
	bool bRet{};

	do
	{
		HANDLE hToken{};
		LUID luid{};
		if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken))
		{
			bRet = false;

			break;
		}

		if (!LookupPrivilegeValue(NULL, SE_DEBUG_NAME, &luid))
		{
			CloseHandle(hToken);

			bRet = false;

			break;
		}

		TOKEN_PRIVILEGES tp{ 1 , {luid, SE_PRIVILEGE_ENABLED} };
		if (!AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(tp), 0, 0))
		{
			CloseHandle(hToken);

			bRet = false;

			break;
		}
		CloseHandle(hToken);

		bRet = true;
	} while (false);

	return bRet;
}

bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout)
{
	// Inject dll into target process
	bool bRet{};

	do
	{
		// Make full path to module
		WCHAR szDllPath[MAX_PATH]{};
		GetModuleFileName(NULL, szDllPath, MAX_PATH);
		PathRemoveFileSpec(szDllPath);
		PathAppend(szDllPath, szModuleName);

		// Call LoadLibraryW with module full path to inject dll into hProcess
		bRet = CallRemoteProc(hProcess, GetProcAddress(GetModuleHandle(L"Kernel32"), "LoadLibraryW"), (void*)szDllPath, (std::wcslen(szDllPath) + 1) * sizeof(WCHAR), dwTimeout);
	} while (false);

	return bRet;
}

bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout)
{
	// Unload dll from target process
	bool bRet{};

	do
	{
		// Find HMODULE of szModuleName in target process
		HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetProcessId(hProcess)) };
		MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
		HMODULE hModInjectionDll{};
		if (!Module32First(hModuleSnapshot, &me32))
		{
			CloseHandle(hModuleSnapshot);

			bRet = false;

			break;
		}
		do
		{
			if (!lstrcmpi(me32.szModule, szModuleName))
			{
				hModInjectionDll = me32.hModule;

				break;
			}
		} while (Module32Next(hModuleSnapshot, &me32));
		if (!hModInjectionDll)
		{
			CloseHandle(hModuleSnapshot);

			bRet = false;

			break;
		}
		CloseHandle(hModuleSnapshot);

		// Call FreeLibrary with HMODULE to unload dll from hProcess
		bRet = CallRemoteProc(hProcess, GetProcAddress(GetModuleHandle(L"Kernel32"), "FreeLibrary"), (void*)hModInjectionDll, 0, dwTimeout);
	} while (false);

	return bRet;
}

DWORD CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, std::size_t cbParamSize, DWORD dwTimeout)
{
	DWORD dwRet{};

	do
	{
		LPVOID lpRemoteBuffer{};
		// If cbParamSize == 0, directly copy lpParameter to lpRemoteBuffer
		if (cbParamSize == 0)
		{
			lpRemoteBuffer = lpParameter;
		}
		// Else do as usual
		else
		{
			// Allocate buffer in target process
			lpRemoteBuffer = VirtualAllocEx(hProcess, NULL, cbParamSize, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
			if (!lpRemoteBuffer)
			{
				dwRet = 0;

				break;
			}

			// Write parameter to remote buffer
			if (!WriteProcessMemory(hProcess, lpRemoteBuffer, lpParameter, cbParamSize, NULL))
			{
				VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

				dwRet = 0;

				break;
			}
		}

		// Create remote thread to call function
		HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, (LPTHREAD_START_ROUTINE)lpRemoteProcAddr, lpRemoteBuffer, 0, NULL) };
		if (!hRemoteThread)
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			dwRet = 0;

			break;
		}

		// Wait for remote thread to terminate with timeout
		if (WaitForSingleObject(hRemoteThread, dwTimeout) == WAIT_TIMEOUT)
		{
			CloseHandle(hRemoteThread);
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			dwRet = 0;

			break;
		}
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

		// Get exit code of remote thread
		DWORD dwRemoteThreadExitCode{};
		if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
		{
			CloseHandle(hRemoteThread);

			dwRet = 0;

			break;
		}
		CloseHandle(hRemoteThread);

		dwRet = dwRemoteThreadExitCode;
	} while (false);

	return dwRet;
}

std::wstring GetUniqueName(LPCWSTR lpszString, Scope scope)
{
	// Create an unique string by scope
	std::wstringstream ssRet{};

	// On the same computer
	if (scope == Scope::Machine)
	{
		ssRet << LR"(Global\)" << lpszString;
	}
	else
	{
		do
		{
			// By the same user
			ssRet << LR"(Local\)" << lpszString << L"--";

			HANDLE hTokenProcess{};
			DWORD dwLength{};
			OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hTokenProcess);
			GetTokenInformation(hTokenProcess, TokenUser, NULL, 0, &dwLength);
			std::unique_ptr<BYTE[]> lpBuffer{ new BYTE[dwLength]{} };
			GetTokenInformation(hTokenProcess, TokenUser, lpBuffer.get(), dwLength, &dwLength);
			LPWSTR lpszSID{};
			ConvertSidToStringSid(((PTOKEN_USER)lpBuffer.get())->User.Sid, &lpszSID);
			ssRet << lpszSID;
			LocalFree(lpszSID);

			if (scope == Scope::User)
			{
				CloseHandle(hTokenProcess);

				break;
			}

			// In the same session
			ssRet << L"--";

			DWORD dwSessionID{};
			ProcessIdToSessionId(GetCurrentProcessId(), &dwSessionID);
			ssRet << dwSessionID;

			if (scope == Scope::Session)
			{
				CloseHandle(hTokenProcess);

				break;
			}

			// In the same window station
			ssRet << L"--";

			DWORD dwLength2{};
			HWINSTA hWinStaProcess{ GetProcessWindowStation() };
			GetUserObjectInformation(hWinStaProcess, UOI_NAME, NULL, 0, &dwLength2);
			std::unique_ptr<BYTE[]> lpBuffer2{ new BYTE[dwLength2]{} };
			GetUserObjectInformation(hWinStaProcess, UOI_NAME, lpBuffer2.get(), dwLength2, &dwLength2);
			ssRet << (LPCWSTR)lpBuffer2.get();

			if (scope == Scope::WindowStation)
			{
				break;
			}

			// On the same desktop
			ssRet << L"--";

			DWORD dwLength3{};
			HDESK hDeskProcess{ GetThreadDesktop(GetCurrentThreadId()) };
			GetUserObjectInformation(hDeskProcess, UOI_NAME, NULL, 0, &dwLength3);
			std::unique_ptr<BYTE[]> lpBuffer3{ new BYTE[dwLength3] };
			GetUserObjectInformation(hDeskProcess, UOI_NAME, lpBuffer3.get(), dwLength3, &dwLength3);
			ssRet << (LPCWSTR)lpBuffer3.get();

			if (scope == Scope::Desktop)
			{
				break;
			}
		} while (false);
	}

	return ssRet.str();
}