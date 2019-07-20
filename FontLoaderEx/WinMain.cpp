#if !defined(UNICODE) || !defined(_UNICODE)
#error Unicode character set required
#endif // UNICODE && _UNICODE

#ifdef _DEBUG
#define DBGPRINTWNDPOSINFO
#endif // _DEBUG

#pragma comment(linker, "\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
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
#include <cassert>
#include "FontResource.h"
#include "Globals.h"
#include "Splitter.h"
#include "resource.h"

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);

constexpr WCHAR szAppName[]{ L"FontLoaderEx" };

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
				{
					assert(0 && "invalid scope");
				}
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
	int iRet{};
	BOOL bRet{};
	do
	{
		switch (bRet = GetMessage(&Message, NULL, 0, 0))
		{
		case -1:
			{
				CloseHandle(hMutexSingleton);

				iRet = static_cast<int>(GetLastError());
			}
			break;
		case 0:
			{
				CloseHandle(hMutexSingleton);

				iRet = static_cast<int>(Message.wParam);
			}
			break;
		default:
			{
				if (!IsDialogMessage(hWndMain, &Message))
				{
					TranslateMessage(&Message);
					DispatchMessage(&Message);
				}
			}
			break;
		}
	} while (bRet);

	return iRet;
}

LRESULT CALLBACK EditTimeoutSubclassProc(HWND hWndEditTimeout, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
LRESULT CALLBACK ListViewFontListSubclassProc(HWND hWndListViewFontList, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
LRESULT CALLBACK EditMessageSubclassProc(HWND hWndEditMessage, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
#ifdef DBGPRINTWNDPOSINFO
LRESULT CALLBACK SplitterSubclassProc(HWND hWndSplitter, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
LRESULT CALLBACK ProgressBarFontSubclassProc(HWND hWndProgressBarFont, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
#endif // SHOWPOSINFO
INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT Message, WPARAM wParam, LPARAM IParam);

void* lpRemoteAddFontProcAddr{};
void* lpRemoteRemoveFontProcAddr{};

HANDLE hThreadCloseWorkerThreadProc{};
HANDLE hThreadButtonCloseWorkerThreadProc{};
HANDLE hThreadButtonLoadWorkerThreadProc{};
HANDLE hThreadButtonUnloadWorkerThreadProc{};
HANDLE hThreadWatch{};
HANDLE hThreadMessage{};

HANDLE hEventWorkerThreadReadyToTerminate{};

HWND hWndProxy{};

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
HANDLE hEventProxyAddFontFinished{};
HANDLE hEventProxyRemoveFontFinished{};

PROXYPROCESSDEBUGPRIVILEGEENABLING ProxyDebugPrivilegeEnablingResult{};
PROXYDLLINJECTION ProxyDllInjectionResult{};
PROXYDLLPULL ProxyDllPullingResult{};

ProcessInfo TargetProcessInfo{}, ProxyProcessInfo{};

int MessageBoxCentered(HWND hWnd, LPCTSTR lpText, LPCTSTR lpCaption, UINT uType);

bool EnableDebugPrivilege();
bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout);
bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout);

enum class ID : WORD { ButtonOpen = 0x20, ButtonClose, ButtonLoad, ButtonUnload, ButtonBroadcastWM_FONTCHANGE, StaticTimeout, EditTimeout, ButtonSelectProcess, ButtonMinimizeToTray, ListViewFontList, Splitter, EditMessage, StatusBarFontInfo, ProgressBarFont };
enum class MENUITEMID : UINT { SC_RESETSIZE = 0xA000 };

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	static DWORD dwTimeout{};

#ifdef _WIN64
	constexpr WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll64.dll" };
	constexpr WCHAR szInjectionDllNameByProxy[]{ L"FontLoaderExInjectionDll.dll" };
#else
	constexpr WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll.dll" };
	constexpr WCHAR szInjectionDllNameByProxy[]{ L"FontLoaderExInjectionDll64.dll" };
#endif // _WIN64

	switch (static_cast<USERMESSAGE>(Message))
	{
		// Close worker thread termination notification
		// wParam = Whether font list was modified : bool
		// LOWORD(lParam) = Whether fonts unloading is interrupted by proxy/target termination: bool
		// HIWORD(lParam) = Whether are all fonts are unloaded : bool
	case USERMESSAGE::CLOSEWORKERTHREADTERMINATED:
		{
			// Wait for close worker thread to terminate
			CloseHandle(hThreadCloseWorkerThreadProc);
			hThreadCloseWorkerThreadProc = NULL;

			// Close the handle to synchronization object
			CloseHandle(hEventWorkerThreadReadyToTerminate);

			// If unloading is interrupted
			if (LOWORD(lParam))
			{
				// Re-enable and re-disable controls
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), TRUE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), TRUE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::EditTimeout)), TRUE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), TRUE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), TRUE);
				EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
				if (!TargetProcessInfo.hProcess)
				{
					EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), TRUE);
				}

				// Update StatusBarFontInfo
				HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
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
				SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(strFontInfo.c_str()));

				// Set ProgressBarFont
				HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
				SendMessage(hWndProgressBarFont, PBM_SETPOS, 0, 0);
				ShowWindow(hWndProgressBarFont, SW_HIDE);

				int aiStatusBarFontInfoParts[]{ -1 };
				SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));

				// Update syatem tray icon tip
				if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
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
					switch (MessageBoxCentered(hWnd, L"Some fonts are not successfully unloaded\r\n\r\nDo you still want to exit?", szAppName, MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON1 | MB_APPLMODAL))
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
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), TRUE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), TRUE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), TRUE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), TRUE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), TRUE);
							if (!TargetProcessInfo.hProcess)
							{
								EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), TRUE);
							}

							// Update StatusBarFontInfo
							HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
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
							SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(strFontInfo.c_str()));

							// Set ProgressBarFont
							HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
							SendMessage(hWndProgressBarFont, PBM_SETPOS, 0, 0);
							ShowWindow(hWndProgressBarFont, SW_HIDE);

							int aiStatusBarFontInfoParts[]{ -1 };
							SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));

							// Update syatem tray icon tip
							if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
							{
								NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
								wcscpy_s(nid.szTip, strFontInfo.c_str());
								Shell_NotifyIcon(NIM_MODIFY, &nid);
							}
						}
						break;
					default:
						{
							assert(0 && "invalid option");
						}
						break;
					}
				}
			}

			// Broadcast WM_FONTCHANGE if ButtonBroadcastWM_FONTCHANGE is checked
			HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{};
			if (wParam)
			{
				if ((Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE))) == BST_CHECKED) && (!TargetProcessInfo.hProcess))
				{
					FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows\r\n\r\n");
				}
			}
		}
		break;
		// Button worker thread termination notification
		// wParam = Whether some fonts are loaded/unloaded : bool
		// lParam = Whether fonts unloading is interrupted by proxy/target termination: bool
	case USERMESSAGE::DRAGDROPWORKERTHREADTERMINATED:
	case USERMESSAGE::BUTTONCLOSEWORKERTHREADTERMINATED:
	case USERMESSAGE::BUTTONLOADWORKERTHREADTERMINATED:
	case USERMESSAGE::BUTTONUNLOADWORKERTHREADTERMINATED:
		{
			// Wait for worker thread to terminate
			// Because only one worker thread runs at a time, so use bitwise-or to get the handle to running worker thread
			HANDLE hThreadWorker{ reinterpret_cast<HANDLE>(reinterpret_cast<UINT_PTR>(hThreadButtonCloseWorkerThreadProc) | reinterpret_cast<UINT_PTR>(hThreadButtonLoadWorkerThreadProc) | reinterpret_cast<UINT_PTR>(hThreadButtonUnloadWorkerThreadProc)) };
			WaitForSingleObject(hThreadWorker, INFINITE);
			CloseHandle(hThreadWorker);
			hThreadButtonCloseWorkerThreadProc = NULL;
			hThreadButtonLoadWorkerThreadProc = NULL;
			hThreadButtonUnloadWorkerThreadProc = NULL;

			// Close the handle to synchronization object
			CloseHandle(hEventWorkerThreadReadyToTerminate);

			// Re-enable and re-disable controls
			EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), TRUE);
			EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), TRUE);
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
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), TRUE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), TRUE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), TRUE);
			}
			if (!TargetProcessInfo.hProcess)
			{
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), TRUE);
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
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), TRUE);
			}

			// Broadcast WM_FONTCHANGE if ButtonBroadcastWM_FONTCHANGE is checked
			HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{};
			if (wParam)
			{
				if ((Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE))) == BST_CHECKED) && (!TargetProcessInfo.hProcess))
				{
					FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows\r\n\r\n");
				}
			}

			// Update StatusBarFontInfo
			HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
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
			SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(strFontInfo.c_str()));

			// Set ProgressBarFont
			HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
			SendMessage(hWndProgressBarFont, PBM_SETPOS, 0, 0);
			ShowWindow(hWndProgressBarFont, SW_HIDE);

			int aiStatusBarFontInfoParts[]{ -1 };
			SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));

			// Update syatem tray icon tip
			if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
			{
				NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
				wcscpy_s(nid.szTip, strFontInfo.c_str());
				Shell_NotifyIcon(NIM_MODIFY, &nid);
			}

			// If worker thread is not interrupted, print an extra CR LF
			if (!lParam)
			{
				cchMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");
			}
		}
		break;
		// Watch thread about to termiate notification
		// wParam = What terminated(Proxy/Target) : enum TERMINATION
		// lParam = Whether worker thread is still running : bool
	case USERMESSAGE::WATCHTHREADTERMINATING:
		{
			// Disable controls
			if (!lParam)
			{
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), FALSE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), FALSE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), FALSE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), FALSE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), FALSE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), FALSE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), FALSE);
				EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
			}

			// Clear ListViewFontList
			ListView_DeleteAllItems(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)));

			HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{ Edit_GetTextLength(hWndEditMessage) };
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, L"\r\n");
			switch (static_cast<TERMINATION>(wParam))
			{
				// If target process terminates, just print message
			case TERMINATION::DIRECT:
				{
					ssMessage << L"Target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L") terminated\r\n\r\n";
					strMessage = ssMessage.str();
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
				}
				break;
				// If proxy process terminates, just print message
			case TERMINATION::PROXY:
				{
					ssMessage << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") terminated\r\n\r\n";
					strMessage = ssMessage.str();
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
				}
				break;
				// If target process termiantes and proxy process is launched, print messages and terminate proxy process
			case TERMINATION::TARGET:
				{
					ssMessage << L"Target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L") terminated\r\n\r\n";
					strMessage = ssMessage.str();
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
					ssMessage.str(L"");

					COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::TERMINATE), 0, NULL };
					FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
					WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
					ssMessage << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") successfully terminated\r\n\r\n";
					strMessage = ssMessage.str();
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
				}
				break;
			default:
				{
					assert(0 && "invalid termination type");
				}
				break;
			}

			// Revert the caption of ButtonSelectProcess to default
			Button_SetText(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), L"Select process");
		}
		break;
		// Watch thread terminated notofication
		// wParam = The exit code of message thread
		// lParam = Whether worker thread is still running : bool
	case USERMESSAGE::WATCHTHREADTERMINATED:
		{
			// Check whether message thread exited normally
			if (wParam)
			{
				HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };

				std::wstringstream ssMessage{};
				std::wstring strMessage{};
				int cchMessageLength{};

				ssMessage << L"Message thread exited abnormally with code " << wParam << L"\r\n\r\n";
				strMessage = ssMessage.str();
				Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
				Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
			}

			// Update StatusBarFontInfo
			SendMessage(GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)), SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(L"0 font(s) opened, 0 font(s) loaded."));

			// Update syatem tray icon tip
			if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
			{
				NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP, 0, NULL, L"0 font(s) opened, 0 font(s) loaded." };
				Shell_NotifyIcon(NIM_MODIFY, &nid);
			}

			// Enable controls
			if (!lParam)
			{
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), TRUE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), TRUE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), TRUE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), TRUE);
				EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
			}
			EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::EditTimeout)), TRUE);
			EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), TRUE);

			// If worker thread is still running, wait for worker thread to terminate and close the handle to watch thread
			// Because only one worker thread runs at a time, so use bitwise-or to get the handle to running worker thread
			if (lParam)
			{
				WaitForSingleObject(hThreadWatch, INFINITE);
				WaitForSingleObject(ULongToHandle(HandleToULong(hThreadCloseWorkerThreadProc) | HandleToULong(hThreadButtonCloseWorkerThreadProc) | HandleToULong(hThreadButtonLoadWorkerThreadProc) | HandleToULong(hThreadButtonUnloadWorkerThreadProc)), INFINITE);
				CloseHandle(hThreadWatch);
			}
			else
			{
				WaitForSingleObject(hThreadWatch, INFINITE);
				CloseHandle(hThreadWatch);
			}
		}
		break;
		// Font list changed notofication
		// wParam = Font change event : enum FONTLISTCHANGED
		// lParam = iItem in ListViewFontList and font name : struct FONTLISTCHANGEDSTRUCT*
	case USERMESSAGE::FONTLISTCHANGED:
		{
			// Modify ListViewFontList and print messages to EditMessage
			HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
			HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
			HWND hWndProgressBarFont{ GetDlgItem(GetDlgItem(hWnd, (int)ID::StatusBarFontInfo), static_cast<int>(ID::ProgressBarFont)) };

			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{};
			LVITEM lvi{ LVIF_TEXT };
			switch (static_cast<FONTLISTCHANGED>(wParam))
			{
			case FONTLISTCHANGED::OPENED:
				{
					lvi.iItem = reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->iItem;
					lvi.iSubItem = 0;
					lvi.pszText = const_cast<LPWSTR>(reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName);
					ListView_InsertItem(hWndListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = const_cast<LPWSTR>(L"Not loaded");
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_SetItemState(hWndListViewFontList, lvi.iItem, LVIS_SELECTED, LVIS_SELECTED);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName << L" opened\r\n";
				}
				break;
			case FONTLISTCHANGED::OPENED_LOADED:
				{
					lvi.iItem = reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->iItem;
					lvi.iSubItem = 0;
					lvi.pszText = const_cast<LPWSTR>(reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName);
					ListView_InsertItem(hWndListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = const_cast<LPWSTR>(L"Loaded");
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_SetItemState(hWndListViewFontList, lvi.iItem, LVIS_SELECTED, LVIS_SELECTED);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName << L" opened and successfully loaded\r\n";

					SendMessage(hWndProgressBarFont, PBM_STEPIT, 0, 0);
				}
				break;
			case FONTLISTCHANGED::OPENED_NOTLOADED:
				{
					lvi.iItem = reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->iItem;
					lvi.iSubItem = 0;
					lvi.pszText = const_cast<LPWSTR>(reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName);
					ListView_InsertItem(hWndListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = const_cast<LPWSTR>(L"Load failed");
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_SetItemState(hWndListViewFontList, lvi.iItem, LVIS_SELECTED, LVIS_SELECTED);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << L"Opened but failed to load " << reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName << L"\r\n";

					SendMessage(hWndProgressBarFont, PBM_STEPIT, 0, 0);
				}
				break;
			case FONTLISTCHANGED::LOADED:
				{
					lvi.iItem = reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->iItem;
					lvi.iSubItem = 1;
					lvi.pszText = const_cast<LPWSTR>(L"Loaded");
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName << L" successfully loaded\r\n";

					SendMessage(hWndProgressBarFont, PBM_STEPIT, 0, 0);
				}
				break;
			case FONTLISTCHANGED::NOTLOADED:
				{
					lvi.iItem = reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->iItem;
					lvi.iSubItem = 1;
					lvi.pszText = const_cast<LPWSTR>(L"Load failed");
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << L"Failed to load " << reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName << L"\r\n";

					SendMessage(hWndProgressBarFont, PBM_STEPIT, 0, 0);
				}
				break;
			case FONTLISTCHANGED::UNLOADED:
				{
					lvi.iItem = reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->iItem;
					lvi.iSubItem = 1;
					lvi.pszText = const_cast<LPWSTR>(L"Unloaded");
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName << L" successfully unloaded\r\n";

					SendMessage(hWndProgressBarFont, PBM_STEPIT, 0, 0);
				}
				break;
			case FONTLISTCHANGED::NOTUNLOADED:
				{
					lvi.iItem = reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->iItem;
					lvi.iSubItem = 1;
					lvi.pszText = const_cast<LPWSTR>(L"Unload failed");
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);

					ssMessage << L"Failed to unload " << reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName << L"\r\n";

					SendMessage(hWndProgressBarFont, PBM_STEPIT, 0, 0);
				}
				break;
			case FONTLISTCHANGED::UNLOADED_CLOSED:
				{
					lvi.iItem = reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->iItem;
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);
					ListView_DeleteItem(hWndListViewFontList, lvi.iItem);

					ssMessage << reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName << L" successfully unloaded and closed\r\n";

					SendMessage(hWndProgressBarFont, PBM_STEPIT, 0, 0);
				}
				break;
			case FONTLISTCHANGED::CLOSED:
				{
					lvi.iItem = reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->iItem;
					ListView_EnsureVisible(hWndListViewFontList, lvi.iItem, FALSE);
					ListView_DeleteItem(hWndListViewFontList, lvi.iItem);

					ssMessage << reinterpret_cast<FONTLISTCHANGEDSTRUCT*>(lParam)->lpszFontName << L" closed\r\n";

					SendMessage(hWndProgressBarFont, PBM_STEPIT, 0, 0);
				}
				break;
			case FONTLISTCHANGED::UNTOUCHED:
				{
					SendMessage(hWndProgressBarFont, PBM_STEPIT, 0, 0);
				}
				break;
			default:
				{
					assert(0 && "invalid font list change event");
				}
				break;
			}
			strMessage = ssMessage.str();
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
		}
		break;
		// Child window size/position changed notidfication
		// wParam = Handle to child window : HWND
		// lParam = WINDOWPOSCHANGE::flags in WM_WINDOWPOSCHANGED : UINT
	case USERMESSAGE::CHILDWINDOWPOSCHANGED:
		{
			switch (static_cast<ID>(GetDlgCtrlID(reinterpret_cast<HWND>(wParam))))
			{
			case ID::ListViewFontList:
				{
					// Adjust column width
					if ((lParam & SWP_DRAWFRAME) || (!(lParam & SWP_NOSIZE)))
					{
						RECT rcListViewFontListClient{};
						GetClientRect(reinterpret_cast<HWND>(wParam), &rcListViewFontListClient);
						ListView_SetColumnWidth(reinterpret_cast<HWND>(wParam), 0, rcListViewFontListClient.right - rcListViewFontListClient.left - ListView_GetColumnWidth(reinterpret_cast<HWND>(wParam), 1));
					}

					// Ensure the item with selection mark or last selected item visible
					if (!(lParam & SWP_NOSIZE))
					{
						int iSelectionMark{ ListView_GetSelectionMark(reinterpret_cast<HWND>(wParam)) };
						if (iSelectionMark == -1)
						{
							int iItemCount{ ListView_GetItemCount(reinterpret_cast<HWND>(wParam)) };
							for (int i = iItemCount - 1; i >= 0; i--)
							{
								if (ListView_GetItemState(reinterpret_cast<HWND>(wParam), LVIS_SELECTED, LVIS_SELECTED) & LVIS_SELECTED)
								{
									ListView_EnsureVisible(reinterpret_cast<HWND>(wParam), i, FALSE);

									break;
								}
							}
						}
						else
						{
							ListView_EnsureVisible(reinterpret_cast<HWND>(wParam), iSelectionMark, FALSE);
						}
					}
				}
				break;
			case ID::EditMessage:
				{
					// Scroll caret into view
					if (!(lParam & SWP_NOSIZE))
					{
						Edit_ScrollCaret(reinterpret_cast<HWND>(wParam));
					}
				}
				break;
			default:
				{
					assert(0 && "invalid child control HWND");
				}
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
					BOOL bRetTrackPopupMenu{ TrackPopupMenu(hMenuContextTray, uFlags | TPM_BOTTOMALIGN | TPM_RIGHTBUTTON, GET_X_LPARAM(wParam), GET_Y_LPARAM(wParam), 0, hWnd, NULL) };
					assert(bRetTrackPopupMenu);

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

	static HPEN hPenSplitter{};

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
			┃ 1.Click "Select process", select a process.                                  │   ┃
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
			HWND hWndButtonOpen{ CreateWindow(WC_BUTTON, L"&Open", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, rcMainClient.left, rcMainClient.top, 50, 50, hWnd, reinterpret_cast<HMENU>(ID::ButtonOpen), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndButtonOpen);
			SetWindowFont(hWndButtonOpen, hFontMain, TRUE);

			// Initialize ButtonClose
			RECT rcButtonOpen{};
			GetWindowRect(hWndButtonOpen, &rcButtonOpen);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
			HWND hWndButtonClose{ CreateWindow(WC_BUTTON, L"&Close", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, rcButtonOpen.right, rcMainClient.top, 50, 50, hWnd, reinterpret_cast<HMENU>(ID::ButtonClose), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndButtonClose);
			SetWindowFont(hWndButtonClose, hFontMain, TRUE);

			// Initialize ButtonLoad
			RECT rcButtonClose{};
			GetWindowRect(hWndButtonClose, &rcButtonClose);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonClose);
			HWND hWndButtonLoad{ CreateWindow(WC_BUTTON, L"&Load", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, rcButtonClose.right, rcMainClient.top, 50, 50, hWnd, reinterpret_cast<HMENU>(ID::ButtonLoad), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndButtonLoad);
			SetWindowFont(hWndButtonLoad, hFontMain, TRUE);

			// Initialize ButtonUnload
			RECT rcButtonLoad{};
			GetWindowRect(hWndButtonLoad, &rcButtonLoad);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonLoad);
			HWND hWndButtonUnload{ CreateWindow(WC_BUTTON, L"&Unload", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, rcButtonLoad.right, rcMainClient.top, 50, 50, hWnd, reinterpret_cast<HMENU>(ID::ButtonUnload), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndButtonUnload);
			SetWindowFont(hWndButtonUnload, hFontMain, TRUE);

			// Initialize ButtonBroadcastWM_FONTCHANGE
			RECT rcButtonUnload{};
			GetWindowRect(hWndButtonUnload, &rcButtonUnload);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonUnload);
			HWND hWndButtonBroadcastWM_FONTCHANGE{ CreateWindow(WC_BUTTON, L"&Broadcast WM_FONTCHANGE", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, rcButtonUnload.right, rcMainClient.top, 250, 21, hWnd, reinterpret_cast<HMENU>(ID::ButtonBroadcastWM_FONTCHANGE), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndButtonBroadcastWM_FONTCHANGE);
			SetWindowFont(hWndButtonBroadcastWM_FONTCHANGE, hFontMain, TRUE);

			// Initialize EditTimeout and its label
			RECT rcButtonBroadcastWM_FONTCHANGE{};
			GetWindowRect(hWndButtonBroadcastWM_FONTCHANGE, &rcButtonBroadcastWM_FONTCHANGE);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonBroadcastWM_FONTCHANGE);
			HWND hWndStaticTimeout{ CreateWindow(WC_STATIC, L"&Timeout:", WS_CHILD | WS_VISIBLE | SS_LEFT, rcButtonBroadcastWM_FONTCHANGE.right + 20, rcMainClient.top + 1, 50, 19, hWnd, reinterpret_cast<HMENU>(ID::StaticTimeout), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndStaticTimeout);
			SetWindowFont(hWndStaticTimeout, hFontMain, TRUE);

			RECT rcStaticTimeout{};
			GetWindowRect(hWndStaticTimeout, &rcStaticTimeout);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcStaticTimeout);
			HWND hWndEditTimeout{ CreateWindow(WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_LEFT | ES_NUMBER | ES_AUTOHSCROLL | ES_NOHIDESEL, rcStaticTimeout.right, rcMainClient.top, 80, 21, hWnd, reinterpret_cast<HMENU>(ID::EditTimeout), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndEditTimeout);
			SetWindowFont(hWndEditTimeout, hFontMain, TRUE);

			Edit_SetText(hWndEditTimeout, L"5000");
			Edit_LimitText(hWndEditTimeout, 10);

			SetWindowSubclass(hWndEditTimeout, EditTimeoutSubclassProc, 0, 0);

			// Initialize ButtonSelectProcess
			HWND hWndButtonSelectProcess{ CreateWindow(WC_BUTTON, L"&Select process", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, rcButtonUnload.right, rcButtonUnload.bottom - 21, 250, 21, hWnd, reinterpret_cast<HMENU>(ID::ButtonSelectProcess), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndButtonSelectProcess);
			SetWindowFont(hWndButtonSelectProcess, hFontMain, TRUE);

			// Initialize ButtonMinimizeToTray
			HWND hWndButtonMinimizeToTray{ CreateWindow(WC_BUTTON, L"&Minimize to tray", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, rcButtonBroadcastWM_FONTCHANGE.right + 20, rcButtonUnload.bottom - 21, 130, 21, hWnd, reinterpret_cast<HMENU>(ID::ButtonMinimizeToTray), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndButtonMinimizeToTray);
			SetWindowFont(hWndButtonMinimizeToTray, hFontMain, TRUE);

			// Initialize StatusBar
			HWND hWndStatusBarFontInfo{ CreateWindow(STATUSCLASSNAME, NULL, WS_CHILD | WS_VISIBLE | SBARS_SIZEGRIP, 0, 0, 0, 0, hWnd, reinterpret_cast<HMENU>(ID::StatusBarFontInfo), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndStatusBarFontInfo);
			SetWindowFont(hWndButtonMinimizeToTray, hFontMain, TRUE);

			SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(L"0 font(s) opened, 0 font(s) loaded."));

			// Initialize ProgreeBarFont
			HWND hWndProgressBarFont{ CreateWindow(PROGRESS_CLASS, NULL, WS_CHILD, 0, 0, 0, 0, hWndStatusBarFontInfo, reinterpret_cast<HMENU>(ID::ProgressBarFont), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndProgressBarFont);

			SendMessage(hWndProgressBarFont, PBM_SETRANGE, 0, MAKELPARAM(0, 1));
			SendMessage(hWndProgressBarFont, PBM_SETSTEP, 1, 0);
			SendMessage(hWndProgressBarFont, PBM_SETSTATE, PBST_NORMAL, 0);

#ifdef DBGPRINTWNDPOSINFO
			SetWindowSubclass(hWndProgressBarFont, ProgressBarFontSubclassProc, 0, 0);
#endif // SHOWPOSINFO

			// Initialize Splitter
			RECT rcStatusBarFontInfo{};
			GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);
			HWND hWndSplitter{ CreateWindow(UC_SPLITTER, NULL, WS_CHILD | WS_VISIBLE | SPS_HORZ | SPS_PARENTWIDTH | SPS_AUTODRAG | SPS_NONOTIFY, rcMainClient.left, (rcButtonOpen.bottom + rcStatusBarFontInfo.top) / 2 - 2, 0, 5, hWnd, reinterpret_cast<HMENU>(ID::Splitter), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndSplitter);
			//HWND hWndSplitter{ NULL };

#ifdef DBGPRINTWNDPOSINFO
			SetWindowSubclass(hWndSplitter, SplitterSubclassProc, 0, 0);
#endif // SHOWPOSINFO

			SendMessage(hWndSplitter, SPM_SETMARGIN, 3, 0);

			RECT rcSplitterClient{};
			GetClientRect(hWndSplitter, &rcSplitterClient);
			hPenSplitter = CreatePen(PS_SOLID, (rcSplitterClient.bottom - rcSplitterClient.top) / 5, static_cast<COLORREF>(GetSysColor(COLOR_GRAYTEXT)));

			// Initialize ListViewFontList
			RECT rcSplitter{};
			GetWindowRect(hWndSplitter, &rcSplitter);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
			HWND hWndListViewFontList{ CreateWindow(WC_LISTVIEW, NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | WS_CLIPSIBLINGS | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_NOSORTHEADER, rcMainClient.left, rcButtonOpen.bottom, rcMainClient.right - rcMainClient.left, rcSplitter.top - rcButtonOpen.bottom, hWnd, reinterpret_cast<HMENU>(ID::ListViewFontList), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndListViewFontList);
			ListView_SetExtendedListViewStyle(hWndListViewFontList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_INFOTIP | LVS_EX_LABELTIP);
			SetWindowFont(hWndListViewFontList, hFontMain, TRUE);

			RECT rcListViewFontListClient{};
			GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
			LVCOLUMN lvc1{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rcListViewFontListClient.right - rcListViewFontListClient.left) * 4 / 5 , const_cast<LPWSTR>(L"Font Name") };
			LVCOLUMN lvc2{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rcListViewFontListClient.right - rcListViewFontListClient.left) * 1 / 5 , const_cast<LPWSTR>(L"State") };
			ListView_InsertColumn(hWndListViewFontList, 0, &lvc1);
			ListView_InsertColumn(hWndListViewFontList, 1, &lvc2);

			SetWindowSubclass(hWndListViewFontList, ListViewFontListSubclassProc, 0, 0);

			DragAcceptFiles(hWndListViewFontList, TRUE);

			CHANGEFILTERSTRUCT cfs{ sizeof(CHANGEFILTERSTRUCT) };
			ChangeWindowMessageFilterEx(hWndListViewFontList, WM_DROPFILES, MSGFLT_ALLOW, &cfs);
			assert(cfs.ExtStatus == MSGFLTINFO_NONE);
			ChangeWindowMessageFilterEx(hWndListViewFontList, WM_COPYDATA, MSGFLT_ALLOW, &cfs);
			assert(cfs.ExtStatus == MSGFLTINFO_NONE);
			ChangeWindowMessageFilterEx(hWndListViewFontList, 0x0049, MSGFLT_ALLOW, &cfs);	// 0x0049 == WM_COPYGLOBALDATA
			assert(cfs.ExtStatus == MSGFLTINFO_NONE);

			// Initialize EditMessage
			HWND hWndEditMessage{ CreateWindow(WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_CLIPSIBLINGS | ES_READONLY | ES_LEFT | ES_MULTILINE | ES_NOHIDESEL, rcMainClient.left, rcSplitter.bottom, rcMainClient.right - rcMainClient.left, rcStatusBarFontInfo.top - rcSplitter.bottom, hWnd, reinterpret_cast<HMENU>(ID::EditMessage), reinterpret_cast<LPCREATESTRUCT>(lParam)->hInstance, NULL) };
			assert(hWndEditMessage);
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
				R"(1.Click "Select process", select a process.)""\r\n"
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
				R"("Minimize to tray": If checked, click the minimize or close button in the upper-right cornor will minimize the window to system tray.)""\r\n"
				R"("Font Name": Names of the fonts added to the list view.)""\r\n"
				R"("State": State of the font. There are five states, "Not loaded", "Loaded", "Load failed", "Unloaded" and "Unload failed".)""\r\n"
				"\r\n"
			);

			SetWindowSubclass(hWndEditMessage, EditMessageSubclassProc, 0, NULL);

			// Link ListViewFontList and EditMessage to Splitter
			SendMessage(hWndSplitter, SPM_SETLINKEDCTL, MAKEWPARAM(1, SLC_TOP), reinterpret_cast<LPARAM>(&hWndListViewFontList));
			SendMessage(hWndSplitter, SPM_SETLINKEDCTL, MAKEWPARAM(1, SLC_BOTTOM), reinterpret_cast<LPARAM>(&hWndEditMessage));

			// Get vertical margin of the formatting rectangle in EditMessage
			RECT rcEditMessageClient{}, rcEditMessageFormatting{};
			GetClientRect(hWndEditMessage, &rcEditMessageClient);
			Edit_GetRect(hWndEditMessage, &rcEditMessageFormatting);
			cyEditMessageTextMargin = (rcEditMessageClient.bottom - rcEditMessageClient.top) - (rcEditMessageFormatting.bottom - rcEditMessageFormatting.top);

			// Add "Reset window size" to system menu
			HMENU hMenuSystem{ GetSystemMenu(hWnd, FALSE) };
			InsertMenu(hMenuSystem, 5, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
			InsertMenu(hMenuSystem, 6, MF_BYPOSITION | MF_STRING, static_cast<UINT_PTR>(MENUITEMID::SC_RESETSIZE), L"&Reset window size");

			ret = 0;
		}
		break;
	case WM_ACTIVATE:
		{
			// Process drag-drop font files onto the application icon stage II
			if (bDragDropHasFonts)
			{
				// Disable controls
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), FALSE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), FALSE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), FALSE);
				EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), FALSE);
				EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

				// Update StatusBarFontInfo
				HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
				SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(L"Loading fonts..."));

				// Set ProgressBarFont
				HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
				RECT rcStatusBarFontInfo{};
				GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
				int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
				SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
				RECT rcProgressBarFont{};
				SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
				SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);

				SendMessage(hWndProgressBarFont, PBM_SETRANGE, 0, MAKELPARAM(0, FontList.size()));

				// Start worker thread
				HANDLE hThreadDragDropWorkerThreadProc{ reinterpret_cast<HANDLE>(_beginthread(DragDropWorkerThreadProc, 0, nullptr)) };
				assert(hThreadDragDropWorkerThreadProc);

				bDragDropHasFonts = false;
			}

			ret = DefWindowProc(hWnd, Message, wParam, lParam);
		}
		break;
	case WM_COMMAND:
		{
			switch (static_cast<ID>(LOWORD(wParam)))
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

							HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
							HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };

							const DWORD cchBuffer{ 1024 };
							OPENFILENAME ofn{ sizeof(ofn), hWnd, NULL, L"Font Files(*.ttf;*.ttc;*.otf)\0*.ttf;*.ttc;*.otf\0", NULL, 0, 0, new WCHAR[cchBuffer]{}, cchBuffer * sizeof(WCHAR), NULL, 0, NULL, L"Select fonts", OFN_EXPLORER | OFN_ENABLESIZING | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_ENABLEHOOK, 0, 0, NULL, reinterpret_cast<LPARAM>(&ofn),
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
											lpofn = reinterpret_cast<LPOPENFILENAME>(lParam);

											// Get the HWND to open dialog
											hWndOpenDialog = GetAncestor(hWndOpenDialogChild, GA_PARENT);

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

											ret = static_cast<UINT_PTR>(FALSE);
										}
										break;
									case WM_NOTIFY:
										{
											switch (reinterpret_cast<LPOFNOTIFY>(lParam)->hdr.code)
											{
												// Adjust buffer size
											case CDN_SELCHANGE:
												{
													DWORD cchLength{ static_cast<DWORD>(CommDlg_OpenSave_GetSpec(hWndOpenDialog, NULL, 0) + MAX_PATH) };
													if (lpofn->nMaxFile < cchLength)
													{
														delete[] lpofn->lpstrFile;
														lpofn->lpstrFile = new WCHAR[cchLength]{};
														lpofn->nMaxFile = cchLength * sizeof(WCHAR);
													}

													ret = static_cast<UINT_PTR>(TRUE);
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
										bool bIsFontDuplicate{ false };
										for (const auto& i : FontList)
										{
											if (i.GetFontName() == lpszPath)
											{
												bIsFontDuplicate = true;

												break;
											}
										}
										if (!bIsFontDuplicate)
										{
											FontList.push_back(lpszPath);
											flcs.lpszFontName = lpszPath;
											SendMessage(hWnd, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::OPENED), reinterpret_cast<LPARAM>(&flcs));

											flcs.iItem++;
										}
									} while (*lpszFileName);
								}
								else
								{
									FontList.push_back(ofn.lpstrFile);

									flcs.lpszFontName = ofn.lpstrFile;
									SendMessage(hWnd, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::OPENED), reinterpret_cast<LPARAM>(&flcs));
								}
								int cchMessageLength{ Edit_GetTextLength(hWndEditMessage) };
								Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
								Edit_ReplaceSel(hWndEditMessage, L"\r\n");

								if (bIsFontListEmptyBefore)
								{
									EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), TRUE);
									EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), TRUE);
									EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), TRUE);
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
							SendMessage(GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)), SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(strFontInfo.c_str()));

							// Update syatem tray icon tip
							if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
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
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), FALSE);
							EnableWindow(reinterpret_cast<HWND>(lParam), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), FALSE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							// Update StatusBarFontInfo
							HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
							SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(L"Unloading and closing fonts..."));

							// Set ProgressBarFont
							HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
							RECT rcStatusBarFontInfo{};
							GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
							int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
							SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
							RECT rcProgressBarFont{};
							SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
							SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);

							HWND hWndListViewFontList{ (GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList))) };
							int iSelectedItemCount{}, iItemCount{ ListView_GetItemCount(hWndListViewFontList) };
							for (int i = 0; i < iItemCount; i++)
							{
								if (ListView_GetItemState(hWndListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED)
								{
									iSelectedItemCount++;
								}
							}
							SendMessage(hWndProgressBarFont, PBM_SETRANGE, 0, MAKELPARAM(0, iSelectedItemCount));

							// Update syatem tray icon tip
							if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
							{
								NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP, 0, NULL, L"Unloading and closing fonts..." };
								Shell_NotifyIcon(NIM_MODIFY, &nid);
							}

							// Create synchronization object and start worker thread
							hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);
							assert(hEventWorkerThreadReadyToTerminate);
							hThreadButtonCloseWorkerThreadProc = reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, ButtonCloseWorkerThreadProc, reinterpret_cast<void*>(ID::ListViewFontList), 0, nullptr));
							assert(hThreadButtonCloseWorkerThreadProc);
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
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), FALSE);
							EnableWindow(reinterpret_cast<HWND>(lParam), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), FALSE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							// Update StatusBarFontInfo
							HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
							SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(L"Loading fonts..."));

							// Set ProgressBarFont
							HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
							RECT rcStatusBarFontInfo{};
							GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
							int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
							SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
							RECT rcProgressBarFont{};
							SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
							SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);

							HWND hWndListViewFontList{ (GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList))) };
							int iSelectedItemCount{}, iItemCount{ ListView_GetItemCount(hWndListViewFontList) };
							for (int i = 0; i < iItemCount; i++)
							{
								if (ListView_GetItemState(hWndListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED)
								{
									iSelectedItemCount++;
								}
							}
							SendMessage(hWndProgressBarFont, PBM_SETRANGE, 0, MAKELPARAM(0, iSelectedItemCount));

							// Update syatem tray icon tip
							if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
							{
								NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP, 0, NULL, L"Loading fonts..." };
								Shell_NotifyIcon(NIM_MODIFY, &nid);
							}

							// Create synchronization object and start worker thread
							hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);
							assert(hEventWorkerThreadReadyToTerminate);
							hThreadButtonLoadWorkerThreadProc = reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, ButtonLoadWorkerThreadProc, reinterpret_cast<void*>(ID::ListViewFontList), 0, nullptr));
							assert(hThreadButtonLoadWorkerThreadProc);
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
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), FALSE);
							EnableWindow(reinterpret_cast<HWND>(lParam), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), FALSE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							// Update StatusBarFontInfo
							HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
							SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(L"Unloading fonts..."));

							// Set ProgressBarFont
							HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
							RECT rcStatusBarFontInfo{};
							GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
							int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
							SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
							RECT rcProgressBarFont{};
							SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
							SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);

							HWND hWndListViewFontList{ (GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList))) };
							int iSelectedItemCount{}, iItemCount{ ListView_GetItemCount(hWndListViewFontList) };
							for (int i = 0; i < iItemCount; i++)
							{
								if (ListView_GetItemState(hWndListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED)
								{
									iSelectedItemCount++;
								}
							}
							SendMessage(hWndProgressBarFont, PBM_SETRANGE, 0, MAKELPARAM(0, iSelectedItemCount));

							// Update syatem tray icon tip
							if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
							{
								NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP, 0, NULL, L"Unloading fonts..." };
								Shell_NotifyIcon(NIM_MODIFY, &nid);
							}

							// Create synchronization object and start worker thread
							hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);
							assert(hEventWorkerThreadReadyToTerminate);
							hThreadButtonUnloadWorkerThreadProc = reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, ButtonUnloadWorkerThreadProc, reinterpret_cast<void*>(ID::ListViewFontList), 0, nullptr));
							assert(hThreadButtonUnloadWorkerThreadProc);
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
							std::wstringstream ssTipEdit{};
							std::wstring strTipEdit{};

							BOOL bIsConverted{};
							DWORD dwTimeoutTemp{ static_cast<DWORD>(GetDlgItemInt(hWnd, static_cast<int>(ID::EditTimeout), &bIsConverted, FALSE)) };
							if (!bIsConverted)
							{
								if (Edit_GetTextLength(reinterpret_cast<HWND>(lParam)) == 0)
								{
									ssTipEdit << L"Empty text will be treated as infinite.";
									strTipEdit = ssTipEdit.str();
									EDITBALLOONTIP ebt{ sizeof(EDITBALLOONTIP), L"Infinite", strTipEdit.c_str(), TTI_INFO };
									Edit_ShowBalloonTip(reinterpret_cast<HWND>(lParam), &ebt);

									dwTimeout = INFINITE;
								}
								else
								{
									DWORD dwCaretIndex{ Edit_GetCaretIndex(reinterpret_cast<HWND>(lParam)) };
									SetDlgItemInt(hWnd, static_cast<int>(ID::EditTimeout), static_cast<UINT>(dwTimeout), FALSE);
									Edit_SetCaretIndex(reinterpret_cast<HWND>(lParam), dwCaretIndex);

									EDITBALLOONTIP ebt{ sizeof(EDITBALLOONTIP), L"Out of range", L"Valid timeout value is 0 ~ 4294967295.", TTI_ERROR };
									Edit_ShowBalloonTip(reinterpret_cast<HWND>(lParam), &ebt);
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
									Edit_ShowBalloonTip(reinterpret_cast<HWND>(lParam), &ebt);
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
							HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };

							static bool bIsSeDebugPrivilegeEnabled{ false };

							std::wstringstream ssMessage{};
							std::wstring strMessage{};
							int cchMessageLength{};

							// Enable SeDebugPrivilege
							if (!bIsSeDebugPrivilegeEnabled)
							{
								if (!EnableDebugPrivilege())
								{
									ssMessage << L"Failed to enable " << SE_DEBUG_NAME << L" for " << szAppName << L"\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

									break;
								}
								bIsSeDebugPrivilegeEnabled = true;
							}

							// Select process
							ProcessInfo* pi{ reinterpret_cast<ProcessInfo*>(DialogBox(NULL, MAKEINTRESOURCE(IDD_DIALOG1), hWnd, DialogProc)) };

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
									COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::PULLDLL), 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
									WaitForSingleObject(hEventProxyDllPullingFinished, INFINITE);
									CloseHandle(hEventProxyDllPullingFinished);
									switch (ProxyDllPullingResult)
									{
									case PROXYDLLPULL::SUCCESSFUL:
										{
											ssMessage << szInjectionDllNameByProxy << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
											ssMessage.str(L"");
										}
										goto continue_B9A25A68;
									case PROXYDLLPULL::FAILED:
										{
											ssMessage << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
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
									DWORD dwExitCodeMessageThread{};
									GetExitCodeThread(hThreadMessage, &dwExitCodeMessageThread);
									if (dwExitCodeMessageThread)
									{
										ssMessage << L"Message thread exited abnormally with code " << dwExitCodeMessageThread << L"\r\n\r\n";
										strMessage = ssMessage.str();
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
										ssMessage.str(L"");
									}
									CloseHandle(hThreadMessage);

									// Terminate proxy process
									COPYDATASTRUCT cds2{ static_cast<ULONG_PTR>(COPYDATA::TERMINATE), 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
									WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
									ssMessage << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") successfully terminated\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
									CloseHandle(ProxyProcessInfo.hProcess);
									ProxyProcessInfo.hProcess = NULL;

									// Close the handle to target process, duplicated handles and synchronization objects
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
										ssMessage << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
										ssMessage.str(L"");
									}
									else
									{
										ssMessage << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
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

									// Close the handle to target process
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
								}

								// Get HANDLE to target process
								SelectedProcessInfo.hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, SelectedProcessInfo.dwProcessID);
								if (!SelectedProcessInfo.hProcess)
								{
									ssMessage << L"Failed to open process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

									break;
								}

								// Determine current and target process machine architecture
								USHORT usCurrentProcessMachineArchitecture{}, usTargetProcessMachineArchitecture{}, usNativeMachineArchitecture{};
								// Get the function pointer to IsWow64Process2() for compatibility
								typedef BOOL(__stdcall *pfnIsWow64Process2)(HANDLE hProcess, USHORT *pProcessMachine, USHORT *pNativeMachine);
								pfnIsWow64Process2 IsWow64Process2_{ reinterpret_cast<pfnIsWow64Process2>(GetProcAddress(GetModuleHandle(L"Kernel32.dll"), "IsWow64Process2")) };
								// For Win10 1709+, IsWow64Process2 exists
								if (IsWow64Process2_)
								{
									BOOL bRetIsWow64Process21{ IsWow64Process2_(GetCurrentProcess(), &usCurrentProcessMachineArchitecture, &usNativeMachineArchitecture) };
									assert(bRetIsWow64Process21);
									BOOL bRetIsWow64Process22{ IsWow64Process2_(SelectedProcessInfo.hProcess, &usTargetProcessMachineArchitecture, &usNativeMachineArchitecture) };
									assert(bRetIsWow64Process22);
									if (usCurrentProcessMachineArchitecture == IMAGE_FILE_MACHINE_UNKNOWN)
									{
										usCurrentProcessMachineArchitecture = usNativeMachineArchitecture;
									}
									if (usTargetProcessMachineArchitecture == IMAGE_FILE_MACHINE_UNKNOWN)
									{
										usTargetProcessMachineArchitecture = usNativeMachineArchitecture;
									}
								}
								// Else call bIsWow64Process() instead
								else
								{
									BOOL bIsCurrentProcessWOW64{}, bIsTargetProcessWOW64{};
									BOOL bRetIsWow64Process{ IsWow64Process(SelectedProcessInfo.hProcess, &bIsTargetProcessWOW64) };
									assert(bRetIsWow64Process);
#ifdef _WIN64
									usCurrentProcessMachineArchitecture = IMAGE_FILE_MACHINE_AMD64;
									usNativeMachineArchitecture = IMAGE_FILE_MACHINE_AMD64;
									if (bIsTargetProcessWOW64)
									{
										usTargetProcessMachineArchitecture = IMAGE_FILE_MACHINE_I386;
									}
									else
									{
										usTargetProcessMachineArchitecture = IMAGE_FILE_MACHINE_AMD64;
									}
#else
									usCurrentProcessMachineArchitecture = IMAGE_FILE_MACHINE_I386;
									if (bIsCurrentProcessWOW64)
									{
										usNativeMachineArchitecture = IMAGE_FILE_MACHINE_AMD64;
									}
									else
									{
										usNativeMachineArchitecture = IMAGE_FILE_MACHINE_I386;
									}
									if (bIsTargetProcessWOW64)
									{
										usTargetProcessMachineArchitecture = IMAGE_FILE_MACHINE_I386;
									}
									else
									{
										usTargetProcessMachineArchitecture = IMAGE_FILE_MACHINE_AMD64;
									}
#endif // _WIN64
								}

								// If process architectures are different(one is WOW64 and another isn't), launch FontLoaderExProxy.exe to inject dll
								if (((usCurrentProcessMachineArchitecture == IMAGE_FILE_MACHINE_I386) && (usTargetProcessMachineArchitecture == IMAGE_FILE_MACHINE_AMD64)) || ((usCurrentProcessMachineArchitecture == IMAGE_FILE_MACHINE_AMD64) && (usTargetProcessMachineArchitecture == IMAGE_FILE_MACHINE_I386)))
								{
									// Create synchronization objects
									SECURITY_ATTRIBUTES sa{ sizeof(SECURITY_ATTRIBUTES), NULL, TRUE };
									hEventParentProcessRunning = CreateEvent(NULL, TRUE, FALSE, L"FontLoaderEx_EventParentProcessRunning_B980D8A4-C487-4306-9D17-3BA6A2CCA4A4");
									assert(hEventParentProcessRunning);
									hEventMessageThreadNotReady = CreateEvent(NULL, TRUE, FALSE, NULL);
									assert(hEventMessageThreadNotReady);
									hEventMessageThreadReady = CreateEvent(&sa, TRUE, FALSE, NULL);
									assert(hEventMessageThreadReady);
									hEventProxyProcessReady = CreateEvent(&sa, TRUE, FALSE, NULL);
									assert(hEventProxyProcessReady);
									hEventProxyProcessDebugPrivilegeEnablingFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
									assert(hEventProxyProcessDebugPrivilegeEnablingFinished);
									hEventProxyProcessHWNDRevieved = CreateEvent(NULL, TRUE, FALSE, NULL);
									assert(hEventProxyProcessHWNDRevieved);
									hEventProxyDllInjectionFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
									assert(hEventProxyDllInjectionFinished);
									hEventProxyDllPullingFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
									assert(hEventProxyDllPullingFinished);

									// Start message thread
									hThreadMessage = reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, MessageThreadProc, nullptr, 0, nullptr));
									HANDLE handles[]{ hEventMessageThreadNotReady, hEventMessageThreadReady };
									switch (WaitForMultipleObjects(2, handles, FALSE, INFINITE))
									{
									case WAIT_OBJECT_0:
										{
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, L"Failed to create message-only window\r\n\r\n");

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
										{
											assert(0 && "WaitForMultipleObjects() failed");
										}
										break;
									}
								break_721EFBC1:
									break;
								continue_721EFBC1:

									// Launch proxy process, send HANDLE to current process and target process, HWND to message window, HANDLE to synchronization objects and timeout as arguments to proxy process
									constexpr WCHAR szProxyAppName[]{ L"FontLoaderExProxy.exe" };

									BOOL bRetDuplicateHandle1{ DuplicateHandle(GetCurrentProcess(), GetCurrentProcess(), GetCurrentProcess(), &hProcessCurrentDuplicated, 0, TRUE, DUPLICATE_SAME_ACCESS) };
									assert(bRetDuplicateHandle1);
									BOOL bRetDuplicateHandle2{ DuplicateHandle(GetCurrentProcess(), SelectedProcessInfo.hProcess, GetCurrentProcess(), &hProcessTargetDuplicated, 0, TRUE, DUPLICATE_SAME_ACCESS) };
									assert(bRetDuplicateHandle2);
									std::wstringstream ssParams{};
#ifdef _WIN64
									ssParams << HandleToHandle32(hProcessCurrentDuplicated) << L" " << HandleToHandle32(hProcessTargetDuplicated) << L" " << HandleToHandle32(hWndMessage) << L" " << HandleToHandle32(hEventMessageThreadReady) << L" " << HandleToHandle32(hEventProxyProcessReady) << L" " << dwTimeout;
#else
									ssParams << HandleToHandle64(hProcessCurrentDuplicated) << L" " << HandleToHandle64(hProcessTargetDuplicated) << L" " << HandleToHandle64(hWndMessage) << L" " << HandleToHandle64(hEventMessageThreadReady) << L" " << HandleToHandle64(hEventProxyProcessReady) << L" " << dwTimeout;
#endif // _WIN64
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

										ssMessage << L"Failed to launch " << szProxyAppName << L"\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
										ssMessage.str(L"");

										// Terminate message thread
										PostMessage(hWndMessage, WM_CLOSE, 0, 0);
										WaitForSingleObject(hThreadMessage, INFINITE);
										DWORD dwExitCodeMessageThread{};
										GetExitCodeThread(hThreadMessage, &dwExitCodeMessageThread);
										if (dwExitCodeMessageThread)
										{
											ssMessage << L"Message thread exited abnormally with code " << dwExitCodeMessageThread << L"\r\n\r\n";
											strMessage = ssMessage.str();
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
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

									// Wait for message-only window to receive HWND to proxy process
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
											ssMessage << L"Failed to enable " << SE_DEBUG_NAME << L" for " << szProxyAppName;
											strMessage = ssMessage.str();
											MessageBoxCentered(NULL, strMessage.c_str(), szAppName, MB_ICONERROR);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
											ssMessage.str(L"");

											// Terminate proxy process
											COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::TERMINATE), 0, NULL };
											FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
											WaitForSingleObject(piProxyProcess.hProcess, INFINITE);

											ssMessage << szProxyAppName << L"(" << piProxyProcess.dwProcessId << L") successfully terminated\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
											ssMessage.str(L"");

											CloseHandle(piProxyProcess.hThread);
											CloseHandle(piProxyProcess.hProcess);
											piProxyProcess.hProcess = NULL;

											// Terminate message thread
											PostMessage(hWndMessage, WM_CLOSE, 0, 0);
											WaitForSingleObject(hThreadMessage, INFINITE);
											DWORD dwExitCodeMessageThread{};
											GetExitCodeThread(hThreadMessage, &dwExitCodeMessageThread);
											if (dwExitCodeMessageThread)
											{
												ssMessage << L"Message thread exited abnormally with code " << dwExitCodeMessageThread << L"\r\n\r\n";
												strMessage = ssMessage.str();
												Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
												Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
											}
											CloseHandle(hThreadMessage);

											// Close HANDLE to selected process and duplicated handles
											CloseHandle(SelectedProcessInfo.hProcess);
											CloseHandle(hProcessCurrentDuplicated);
											CloseHandle(hProcessTargetDuplicated);

											CloseHandle(hEventProxyProcessReady);
											CloseHandle(hEventMessageThreadReady);
											CloseHandle(hEventParentProcessRunning);
											CloseHandle(hEventProxyDllInjectionFinished);
											CloseHandle(hEventProxyDllPullingFinished);
										}
										goto break_90567013;
									default:
										break;
									}
								break_90567013:
									break;
								continue_90567013:

									// Wait for proxy process to ready
									WaitForSingleObject(hEventProxyProcessReady, INFINITE);
									CloseHandle(hEventProxyProcessReady);
									CloseHandle(hEventMessageThreadReady);
									CloseHandle(hEventParentProcessRunning);

									ssMessage << szProxyAppName << L"(" << piProxyProcess.dwProcessId << L") succesfully launched\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
									ssMessage.str(L"");

									// Begin dll injection
									COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::INJECTDLL), 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);

									// Wait for proxy process to inject dll into target process
									WaitForSingleObject(hEventProxyDllInjectionFinished, INFINITE);
									CloseHandle(hEventProxyDllInjectionFinished);
									switch (ProxyDllInjectionResult)
									{
									case PROXYDLLINJECTION::SUCCESSFUL:
										{
											ssMessage << szInjectionDllNameByProxy << L" successfully injected into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

											// Register proxy AddFont() and RemoveFont() procedure and create synchronization objects
											FontResource::RegisterAddRemoveFontProc(ProxyAddFontProc, ProxyRemoveFontProc);
											hEventProxyAddFontFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
											assert(hEventProxyAddFontFinished);
											hEventProxyRemoveFontFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
											assert(hEventProxyRemoveFontFinished);

											// Disable EditTimeout and ButtonBroadcastWM_FONTCHANGE
											EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::EditTimeout)), FALSE);
											EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), FALSE);

											// Change the caption of ButtonSelectProcess
											std::wstringstream Caption{};
											Caption << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")";
											Button_SetText(reinterpret_cast<HWND>(lParam), Caption.str().c_str());

											// Set TargetProcessInfo and ProxyProcessInfo
											TargetProcessInfo = SelectedProcessInfo;
											ProxyProcessInfo = { piProxyProcess.hProcess, szProxyAppName, piProxyProcess.dwProcessId };

											// Create synchronization object and start watch thread
											hEventTerminateWatchThread = CreateEvent(NULL, TRUE, FALSE, NULL);
											hThreadWatch = reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, ProxyAndTargetProcessWatchThreadProc, nullptr, 0, nullptr));
											assert(hThreadWatch);
										}
										break;
									case PROXYDLLINJECTION::FAILED:
										{
											ssMessage << L"Failed to inject " << szInjectionDllNameByProxy << L" into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::FAILEDTOENUMERATEMODULES:
										{
											ssMessage << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::GDI32NOTLOADED:
										{
											ssMessage << L"gdi32.dll not loaded by target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::MODULENOTFOUND:
										{
											ssMessage << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
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
											COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::TERMINATE), 0, NULL };
											FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
											WaitForSingleObject(piProxyProcess.hProcess, INFINITE);

											ssMessage << szProxyAppName << L"(" << piProxyProcess.dwProcessId << L") successfully terminated\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
											ssMessage.str(L"");

											CloseHandle(piProxyProcess.hProcess);
											piProxyProcess.hProcess = NULL;

											// Terminate message thread
											PostMessage(hWndMessage, WM_CLOSE, 0, 0);
											WaitForSingleObject(hThreadMessage, INFINITE);
											DWORD dwExitCodeMessageThread{};
											GetExitCodeThread(hThreadMessage, &dwExitCodeMessageThread);
											if (dwExitCodeMessageThread)
											{
												ssMessage << L"Message thread exited abnormally with code " << dwExitCodeMessageThread << L"\r\n\r\n";
												strMessage = ssMessage.str();
												Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
												Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
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
										{
											assert(0 && "invalid ProxyDllInjectionResult");
										}
										break;
									}
								}
								// Else DIY
								else if (((usCurrentProcessMachineArchitecture == IMAGE_FILE_MACHINE_I386) && (usTargetProcessMachineArchitecture == IMAGE_FILE_MACHINE_I386)) || ((usCurrentProcessMachineArchitecture == IMAGE_FILE_MACHINE_AMD64) && (usTargetProcessMachineArchitecture == IMAGE_FILE_MACHINE_AMD64)))
								{
									// Check whether target process loads gdi32.dll as AddFontResourceEx() and RemoveFontResourceEx() are in it
									HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, SelectedProcessInfo.dwProcessID) };
									assert(hModuleSnapshot);
									MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
									bool bIsGDI32Loaded{ false };
									if (!Module32First(hModuleSnapshot, &me32))
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot);

										ssMessage << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
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

										ssMessage << L"gdi32.dll not loaded by target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
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

										ssMessage << L"Failed to inject " << szInjectionDllName << L" into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

										break;
									}
									ssMessage << szInjectionDllName << L" successfully injected into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
									ssMessage.str(L"");

									// Get base address of FontLoaderExInjectionDll(64).dll in target process
									HANDLE hModuleSnapshot2{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, SelectedProcessInfo.dwProcessID) };
									assert(hModuleSnapshot2);
									MODULEENTRY32 me322{ sizeof(MODULEENTRY32) };
									BYTE* lpModBaseAddr{};
									if (!Module32First(hModuleSnapshot2, &me322))
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot2);

										ssMessage << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
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

										ssMessage << szInjectionDllName << " not found in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

										break;
									}
									CloseHandle(hModuleSnapshot2);

									// Calculate the addresses of AddFont() and RemoveFont() in target process
									HMODULE hModInjectionDll{ LoadLibrary(szInjectionDllName) };
									assert(hModInjectionDll);
									void* lpLocalAddFontProcAddr{ GetProcAddress(hModInjectionDll, "AddFont") };
									assert(lpLocalAddFontProcAddr);
									void* lpLocalRemoveFontProcAddr{ GetProcAddress(hModInjectionDll, "RemoveFont") };
									assert(lpLocalRemoveFontProcAddr);
									FreeLibrary(hModInjectionDll);
									lpRemoteAddFontProcAddr = reinterpret_cast<void*>(reinterpret_cast<UINT_PTR>(lpModBaseAddr) + (reinterpret_cast<UINT_PTR>(lpLocalAddFontProcAddr) - reinterpret_cast<UINT_PTR>(hModInjectionDll)));
									lpRemoteRemoveFontProcAddr = reinterpret_cast<void*>(reinterpret_cast<UINT_PTR>(lpModBaseAddr) + (reinterpret_cast<UINT_PTR>(lpLocalRemoveFontProcAddr) - reinterpret_cast<UINT_PTR>(hModInjectionDll)));

									// Register remote AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(RemoteAddFontProc, RemoteRemoveFontProc);

									// Disable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::EditTimeout)), FALSE);
									EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), FALSE);

									// Change the caption of ButtonSelectProcess
									std::wstringstream Caption{};
									Caption << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")";
									Button_SetText(reinterpret_cast<HWND>(lParam), Caption.str().c_str());

									// Set TargetProcessInfo
									TargetProcessInfo = SelectedProcessInfo;

									// Create synchronization object and start watch thread
									hEventTerminateWatchThread = CreateEvent(NULL, TRUE, FALSE, NULL);
									assert(hEventTerminateWatchThread);
									hThreadWatch = reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, TargetProcessWatchThreadProc, nullptr, 0, nullptr));
									assert(hThreadWatch);
								}
								// Else prompt user target process architecture is not supported
								else
								{
									CloseHandle(SelectedProcessInfo.hProcess);

									// ARM and AArch64 is not supported
									if (usTargetProcessMachineArchitecture == IMAGE_FILE_MACHINE_ARM)
									{
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, L"Process of ARM architecture not supported\r\n\r\n");

										break;
									}
									if (usTargetProcessMachineArchitecture == IMAGE_FILE_MACHINE_ARM64)
									{
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, L"Process of AArch64 architecture not supported\r\n\r\n");

										break;
									}
									else
									{
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, L"Target process architecture unknown\r\n\r\n");
									}
								}
							}
							// If p == nullptr, clear selected process
							else
							{
								// If loaded via proxy
								if (ProxyProcessInfo.hProcess)
								{
									// Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
									COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::PULLDLL), 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
									WaitForSingleObject(hEventProxyDllPullingFinished, INFINITE);
									CloseHandle(hEventProxyDllPullingFinished);
									switch (ProxyDllPullingResult)
									{
									case PROXYDLLPULL::SUCCESSFUL:
										{
											ssMessage << szInjectionDllNameByProxy << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
											ssMessage.str(L"");
										}
										goto continue_0F70B465;
									case PROXYDLLPULL::FAILED:
										{
											ssMessage << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
											strMessage = ssMessage.str();
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
										}
										goto break_0F70B465;
									default:
										{
											assert(0 && "ProxyDllPullingResult");
										}
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
									DWORD dwExitCodeMessageThread{};
									GetExitCodeThread(hThreadMessage, &dwExitCodeMessageThread);
									if (dwExitCodeMessageThread)
									{
										ssMessage << L"Message thread exited abnormally with code " << dwExitCodeMessageThread << L"\r\n\r\n";
										strMessage = ssMessage.str();
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
										ssMessage.str(L"");
									}
									CloseHandle(hThreadMessage);

									// Terminate proxy process
									COPYDATASTRUCT cds2{ static_cast<ULONG_PTR>(COPYDATA::TERMINATE), 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
									WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);

									ssMessage << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") successfully terminated\r\n\r\n";
									strMessage = ssMessage.str();
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

									CloseHandle(ProxyProcessInfo.hProcess);
									ProxyProcessInfo.hProcess = NULL;

									// Close the handle to target process, duplicated handles and synchronization objects
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
									CloseHandle(hProcessCurrentDuplicated);
									CloseHandle(hProcessTargetDuplicated);
									CloseHandle(hEventProxyAddFontFinished);
									CloseHandle(hEventProxyRemoveFontFinished);

									// Register global AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

									// Enable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::EditTimeout)), TRUE);
									EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), TRUE);

									// Revert the caption of ButtonSelectProcess to default
									Button_SetText(reinterpret_cast<HWND>(lParam), L"Select process");
								}
								// Else DIY
								else if (TargetProcessInfo.hProcess)
								{
									// Unload FontLoaderExInjectionDll(64).dll from target process
									if (PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										ssMessage << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
										strMessage = ssMessage.str();
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
									}
									else
									{
										ssMessage << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
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

									// Close the handle to target process
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;

									// Register global AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

									// Enable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::EditTimeout)), TRUE);
									EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), TRUE);

									// Revert the caption of ButtonSelectProcess to default
									Button_SetText(reinterpret_cast<HWND>(lParam), L"Select process");
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
							if (Button_GetCheck(reinterpret_cast<HWND>(lParam)) == BST_UNCHECKED)
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
					SendMessage(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), BM_CLICK, 0, 0);
				}
				break;
				// "Unload" menu item in hMenuContextListViewFontList
			case ID_MENU_UNLOAD:
				{
					// Simulate clicking "Unload" button
					SendMessage(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), BM_CLICK, 0, 0);
				}
				break;
				// "Close" menu item in hMenuContextListViewFontList
			case ID_MENU_CLOSE:
				{
					// Simulate clicking "Close" button
					SendMessage(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), BM_CLICK, 0, 0);
				}
				break;
				// "Select All" menu item in hMenuContextListViewFontList
			case ID_MENU_SELECTALL:
				{
					// Select all items in ListViewFontList
					ListView_SetItemState(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), -1, LVIS_SELECTED, LVIS_SELECTED);
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
					Button_SetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray)), BST_UNCHECKED);
					PostMessage(hWnd, WM_SYSCOMMAND, static_cast<WPARAM>(SC_CLOSE), 0);
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_SYSCOMMAND:
		{
			switch (static_cast<MENUITEMID>(wParam & 0xFFF0))
			{
			case MENUITEMID::SC_RESETSIZE:
				{
					// Reset window size and control position
					WINDOWPLACEMENT wp{ sizeof(WINDOWPLACEMENT) };
					GetWindowPlacement(hWnd, &wp);
					if ((wp.showCmd == SW_MAXIMIZE) || (wp.showCmd == SW_SHOWMAXIMIZED))
					{
						wp.showCmd = SW_RESTORE;
					}
					SetWindowPlacement(hWnd, &wp);
					RECT rcMainWindow{};
					GetWindowRect(hWnd, &rcMainWindow);
					MONITORINFO mi{ sizeof(MONITORINFO) };
					GetMonitorInfo(MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST), &mi);
					POINT ptMainWindow{ rcMainWindow.left, rcMainWindow.top };
					SIZE sizeMainWindow{ 700, 700 };
					CalculatePopupWindowPosition(&ptMainWindow, &sizeMainWindow, TPM_LEFTALIGN | TPM_TOPALIGN | TPM_WORKAREA, NULL, &rcMainWindow);
					SetWindowPos(hWnd, NULL, rcMainWindow.left, rcMainWindow.top, rcMainWindow.right - rcMainWindow.left, rcMainWindow.bottom - rcMainWindow.top, SWP_NOZORDER | SWP_NOACTIVATE);

					HWND hWndSplitter{ GetDlgItem(hWnd, static_cast<int>(ID::Splitter)) };
					RECT rcButtonOpen{}, rcStatusBarFontInfo{}, rcSplitter{};
					GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), &rcButtonOpen);
					GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)), &rcStatusBarFontInfo);
					GetWindowRect(hWndSplitter, &rcSplitter);
					MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
					MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);
					FORWARD_WM_MOVE(hWndSplitter, 0, ((rcStatusBarFontInfo.top + rcButtonOpen.bottom) - (rcSplitter.bottom - rcSplitter.top)) / 2, SendMessage);
				}
				break;
			}

			switch (wParam & 0xFFF0)
			{
			case SC_MINIMIZE:
				{
					// If ButtonMinimizeToTray is checked, minimize to system tray 
					if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
					{
						NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWnd, 0, NIF_MESSAGE | NIF_ICON | NIF_TIP | NIF_SHOWTIP, static_cast<UINT>(USERMESSAGE::TRAYNOTIFYICON), hIconApplication };
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
					if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
					{
						NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWnd, 0, NIF_MESSAGE | NIF_ICON | NIF_TIP | NIF_SHOWTIP, static_cast<UINT>(USERMESSAGE::TRAYNOTIFYICON), hIconApplication };
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
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonSelectProcess)), FALSE);
							EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), FALSE);
							EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							// Update StatusBarFontInfo
							HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
							SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(L"Unloading and closing fonts..."));

							// Set ProgressBarFont
							HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
							RECT rcStatusBarFontInfo{};
							GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
							int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
							SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
							RECT rcProgressBarFont{};
							SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
							SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_SHOWWINDOW);

							SendMessage(hWndProgressBarFont, PBM_SETRANGE, 0, MAKELPARAM(0, FontList.size()));

							// Create synchronization object and start worker thread
							hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);
							assert(hEventWorkerThreadReadyToTerminate);
							hThreadCloseWorkerThreadProc = reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, CloseWorkerThreadProc, nullptr, 0, nullptr));
							assert(hThreadCloseWorkerThreadProc);
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
			switch (static_cast<ID>(wParam))
			{
			case ID::ListViewFontList:
				{
					switch (reinterpret_cast<LPNMHDR>(lParam)->code)
					{
						// Set empty text
					case LVN_GETEMPTYMARKUP:
						{
							reinterpret_cast<NMLVEMPTYMARKUP*>(lParam)->dwFlags = 0;
							wcscpy_s(reinterpret_cast<NMLVEMPTYMARKUP*>(lParam)->szMarkup, LR"(Click "Open" or drag-drop font files here to add fonts.)");

							ret = static_cast<LRESULT>(TRUE);
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
			HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };

			if (reinterpret_cast<HWND>(wParam) == hWndListViewFontList)
			{
				POINT ptAnchorContextMenuListViewFontList{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
				if (ptAnchorContextMenuListViewFontList.x == -1 && ptAnchorContextMenuListViewFontList.y == -1)
				{
					int iSelectionMark{ ListView_GetSelectionMark(hWndListViewFontList) };
					if (iSelectionMark == -1)
					{
						RECT rcListViewFontListClient{};
						GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
						MapWindowRect(hWndListViewFontList, HWND_DESKTOP, &rcListViewFontListClient);
						ptAnchorContextMenuListViewFontList = { rcListViewFontListClient.left, rcListViewFontListClient.top };
					}
					else
					{
						ListView_EnsureVisible(hWndListViewFontList, iSelectionMark, FALSE);
						ListView_GetItemPosition(hWndListViewFontList, iSelectionMark, &ptAnchorContextMenuListViewFontList);
						ClientToScreen(hWndListViewFontList, &ptAnchorContextMenuListViewFontList);
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
				BOOL bRetTrackPopupMenu{ TrackPopupMenu(hMenuContextListViewFontList, uFlags | TPM_TOPALIGN | TPM_RIGHTBUTTON, ptAnchorContextMenuListViewFontList.x, ptAnchorContextMenuListViewFontList.y, 0, hWnd, NULL) };
				assert(bRetTrackPopupMenu);
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
			// and change the line color of Splitter to grayed text color
			// From https://social.msdn.microsoft.com/Forums/vstudio/en-US/7b6d1815-87e3-4f47-b5d5-fd4caa0e0a89/why-is-wmctlcolorstatic-sent-for-a-button-instead-of-wmctlcolorbtn?forum=vclanguage
			// "WM_CTLCOLORSTATIC is sent by any control that displays text which would be displayed using the default dialog/window background color. 
			// This includes check boxes, radio buttons, group boxes, static text, read-only or disabled edit controls, and disabled combo boxes (all styles)."
			if ((reinterpret_cast<HWND>(lParam) == GetDlgItem(hWnd, static_cast<int>(ID::ButtonBroadcastWM_FONTCHANGE))) || (reinterpret_cast<HWND>(lParam) == GetDlgItem(hWnd, static_cast<int>(ID::StaticTimeout))) || (reinterpret_cast<HWND>(lParam) == GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) || reinterpret_cast<HWND>(lParam) == GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)))
			{
				ret = reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_WINDOW));
			}
			else if (reinterpret_cast<HWND>(lParam) == GetDlgItem(hWnd, static_cast<int>(ID::Splitter)))
			{
				SelectPen(reinterpret_cast<HDC>(wParam), hPenSplitter);
			}
			else
			{
				ret = DefWindowProc(hWnd, Message, wParam, lParam);
			}
		}
		break;
		/*
		   Sizing related messasge sequences:
		   1. Drag window edge: WS_SIZING( SizingEdge = WMSZ_* ) -> WM_WINDOWPOSCHANGING -> WS_SIZED
		   2. Maximimze/Restore: WM_WINDOWPOSCHANGING -> WS_SIZED
		   3. Cornor/Border sticky: WM_WINDOWPOSCHANGING -> WS_SIZED, SizingEdge = 0
		*/
	case WM_WINDOWPOSCHANGING:
		{
			// Get StatusBarFontList window rectangle before main window changes size
			GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)), &rcStatusBarFontInfoOld);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfoOld);

			ret = DefWindowProc(hWnd, Message, wParam, lParam);
		}
		break;
	case WM_WINDOWPOSCHANGED:
		{
			// Confine splitter to a specific range
			/*

			┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
			┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃
			┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨
			┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃
			┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃
			┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃
			┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────╂────────
			┃ Font Name                                                         │ State        ┃        ↑
			┠───────────────────────────────────────────────────────────────────┼──────────────┨        │ cyListViewFontListMin
			┃                                                                   ┆              ┃        ↓
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄
			┃                                                                   ┆              ┃                    ↑
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨          ptSpliitterRange.top
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
			┃ 1.Click "Select process", select a process.                                  │   ┃          ptSpliitterRange.bottom
			┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │   ┃                    ↓
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼───╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄
			┃ view, then click "Load" button.                                              │ ↓ ┃        } cyEditMessageMin
			┠──────────────────────────────────────────────────────────────────────────────┴───┨────────
			┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃        } cyStatusBarFontInfo
			┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛────────
			*/

			// Calculate the minimal height of ListViewFontList
			HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
			RECT rcListViewFontList{}, rcListViewFontListClient{}, rcHeaderListViewFontList{};
			GetWindowRect(hWndListViewFontList, &rcListViewFontList);
			GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
			GetWindowRect(ListView_GetHeader(hWndListViewFontList), &rcHeaderListViewFontList);
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
			LONG cyListViewFontListMin{ (rcListViewFontListItem.bottom - rcListViewFontListItem.top) + (rcHeaderListViewFontList.bottom - rcHeaderListViewFontList.top) + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

			// Calculate the minimal height of EditMessage
			HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
			HDC hDCEditMessage{ GetDC(hWndEditMessage) };
			SelectFont(hDCEditMessage, GetWindowFont(hWndEditMessage));
			TEXTMETRIC tm{};
			BOOL bRetGetTextMetrics{ GetTextMetrics(hDCEditMessage, &tm) };
			assert(bRetGetTextMetrics);
			ReleaseDC(hWndEditMessage, hDCEditMessage);
			RECT rcEditMessage{}, rcEditMessageClient{};
			GetWindowRect(hWndEditMessage, &rcEditMessage);
			GetClientRect(hWndEditMessage, &rcEditMessageClient);
			LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + cyEditMessageTextMargin };

			// Calculate the Splitter range
			RECT rcMainClient{}, rcButtonOpen{}, rcStatusBarFontInfo{};
			GetClientRect(hWnd, &rcMainClient);
			GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), &rcButtonOpen);
			GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)), &rcStatusBarFontInfo);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);
			SendMessage(GetDlgItem(hWnd, static_cast<int>(ID::Splitter)), SPM_SETRANGE, MAKEWPARAM(rcMainClient.top + (rcButtonOpen.bottom - rcButtonOpen.top) + cyListViewFontListMin, rcMainClient.bottom - (rcStatusBarFontInfo.bottom - rcStatusBarFontInfo.top) - cyEditMessageMin), 0);

			ret = DefWindowProc(hWnd, Message, wParam, lParam);

#ifdef DBGPRINTWNDPOSINFO
			LPWINDOWPOS lpwp{ reinterpret_cast<LPWINDOWPOS>(lParam) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			ssMessage << L"* Main revieved WM_WINDOWPOSCHANGED\r\n"
				<< L"hwndInsertAfter: " << lpwp->hwndInsertAfter << L" "
				<< L"hwnd: " << lpwp->hwnd << L" "
				<< L"x: " << lpwp->x << L" "
				<< L"y: " << lpwp->y << L" "
				<< L"cx: " << lpwp->cx << L" "
				<< L"cy: " << lpwp->cy << L" "
				<< L"flags: ";
			if (lpwp->flags & SWP_NOSIZE)
			{
				ssMessage << L"SWP_NOSIZE | ";
			}
			if (lpwp->flags & SWP_NOMOVE)
			{
				ssMessage << L"SWP_NOMOVE | ";
			}
			if (lpwp->flags & SWP_NOZORDER)
			{
				ssMessage << L"SWP_NOZORDER | ";
			}
			if (lpwp->flags & SWP_NOREDRAW)
			{
				ssMessage << L"SWP_NOREDRAW | ";
			}
			if (lpwp->flags & SWP_NOACTIVATE)
			{
				ssMessage << L"SWP_NOACTIVATE | ";
			}
			if (lpwp->flags & SWP_FRAMECHANGED)
			{
				ssMessage << L"SWP_FRAMECHANGED | ";
			}
			if (lpwp->flags & SWP_SHOWWINDOW)
			{
				ssMessage << L"SWP_SHOWWINDOW | ";
			}
			if (lpwp->flags & SWP_HIDEWINDOW)
			{
				ssMessage << L"SWP_HIDEWINDOW | ";
			}
			if (lpwp->flags & SWP_NOCOPYBITS)
			{
				ssMessage << L"SWP_NOCOPYBITS | ";
			}
			if (lpwp->flags & SWP_NOOWNERZORDER)
			{
				ssMessage << L"SWP_NOOWNERZORDER | ";
			}
			if (lpwp->flags & SWP_NOSENDCHANGING)
			{
				ssMessage << L"SWP_NOSENDCHANGING | ";
			}
			ssMessage << L"\r\n";
			strMessage = ssMessage.str();
			std::size_t pos{ strMessage.find_last_of(L'|') };
			if (pos != std::wstring::npos)
			{
				strMessage.erase(strMessage.find_last_of(L'|') - 1, 3);
			}
			OutputDebugString(strMessage.c_str());
#endif // SHOWPOSINFO
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
								// Proportionally scale the position of splitter, caused by cornor/border sticky
							case 0:
								{
									/*
									┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
									┃                                                                                                                                                                      ┃
									┃                                                                                                                                                                      ┃
									┃	┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓                                                                               ┃← Desktop
									┃	┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃                                                                               ┃
									┃	┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨                                                                               ┃
									┃	┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃                                                                               ┃
									┃	┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃                                                                               ┃
									┃	┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃                                                                               ┃
									┃	┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨                                                                               ┃
									┃	┃ Font Name                                                         │ State        ┃                                                                               ┃
									┃	┠───────────────────────────────────────────────────────────────────┼──────────────┨                                                                               ┃
									┃	┃                                                                   ┆              ┃                                                                               ┃
									┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                                                                               ┃
									┃	┃                                                                   ┆              ┃                                                                               ┃
									┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                                                                               ┃
									┃	┃                                                                   ┆              ┃                                                                               ┃
									┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                                                                               ┃
									┃	┃                                                                   ┆              ┃                                                                               ┃
									┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                                                                               ┃
									┃	┠───────────────────────────────────────────────────────────────────┴──────────────┨                                                                               ┃
									┃	┠──────────────────────────────────────────────────────────────────────────────┬───┨                                                                               ┃  =>
									┃	┃ Temporarily load fonts to Windows or specific process                        │ ↑ ┃                                                                               ┃
									┃	┃                                                                              ├───┨                                                                               ┃
									┃	┃ How to load fonts to Windows:                                                │▓▓▓┃                                                                               ┃
									┃	┃ 1.Drag-drop font files onto the icon of this application.                    │▓▓▓┃                                                                               ┃
									┃	┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  │▓▓▓┃                                                                               ┃
									┃	┃  view, then click "Load" button.                                             │▓▓▓┃                                                                               ┃
									┃	┃                                                                              ├───┨                                                                               ┃
									┃	┃ How to unload fonts from Windows:                                            │   ┃                                                                               ┃
									┃	┃ Select all fonts then click "Unload" or "Close" button or the X at the       │   ┃                                                                               ┃
									┃	┃ upper-right cornor.                                                          │   ┃                                                                               ┃
									┃	┃                                                                              │   ┃                                                                               ┃
									┃	┃ How to load fonts to process:                                                │   ┃                                                                               ┃
									┃	┃ 1.Click "Select process", select a process.                                  │   ┃                                                                               ┃
									┃	┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨                                                                               ┃
									┃	┃ view, then click "Load" button.                                              │ ↓ ┃                                                                               ┃
									┃	┠──────────────────────────────────────────────────────────────────────────────┴───┨                                                                               ┃
									┃	┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃                                                                               ┃
									┃	┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛                                                                               ┃
									┃	                                                                                                                                                                   ┃
									┃                                                                                                                                                                      ┃
									┃                                                                                                                                                                      ┃
									┃                                                                                                                                                                      ┃
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛

									┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┳━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
									┃FontLoaderEx                                                          ┃_  ┃ □ ┃ x ┃                                                                                   ┃
									┠────────┬────────┬────────┬────────┬──────────────────────────────────┸┬──┸───┸──┬┨                                                                                   ┃
									┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000     │┃                                                                                   ┃← Desktop
									┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └─────────┘┃                                                                                   ┃
									┃        │        │        │        │     Select Process     │  □ Minimize to tray ┃                                                                                   ┃
									┠────────┴────────┴────────┴────────┴────────────────────────┴──────┬──────────────┨                                                                                   ┃
									┃ Font Name                                                         │ State        ┃                                                                                   ┃
									┠───────────────────────────────────────────────────────────────────┼──────────────┨                                                                                   ┃
									┃                                                                   ┆              ┃                                                                                   ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                                                                                   ┃
									┃                                                                   ┆              ┃                                                                                   ┃
									┠───────────────────────────────────────────────────────────────────┴──────────────┨                                                                                   ┃
									┠──────────────────────────────────────────────────────────────────────────────┬───┨                                                                                   ┃
									┃ Temporarily load fonts to Windows or specific process	                       │ ↑ ┃                                                                                   ┃
									┃	                                                                           ├───┨                                                                                   ┃
									┃ How to load fonts to Windows:                                                │▓▓▓┃                                                                                   ┃
									┃ 1.Drag-drop font files onto the icon of this application.	                   ├───┨                                                                                   ┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨                                                                                   ┃
									┃  view, then click "Load" button.                                             │ ↓ ┃                                                                                   ┃
									┠──────────────────────────────────────────────────────────────────────────────┴───┨                                                                                   ┃
									┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃                                                                                   ┃
									┣━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃                                                                                                                                                                      ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃	                                                                                                                                                                   ┃
									┃                                                                                                                                                                      ┃
									┃                                                                                                                                                                      ┃
									┃                                                                                                                                                                      ┃
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
									*/

									// Resize StatusBarFontInfo
									HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
									FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);
									RECT rcStatusBarFontInfo{};
									GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

									// Resize ProgressBarFont
									HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
									if (IsWindowVisible(hWndProgressBarFont))
									{
										RECT rcStatusBarFontInfo{};
										GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
										int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
										SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
										RECT rcProgressBarFont{};
										SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
										SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, static_cast<int>(ID::Splitter)) };
									RECT rcButtonOpen{}, rcListViewFontList{}, rcSplitter{}, rcEditMessage{};
									GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), &rcButtonOpen);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
									GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), &rcListViewFontList);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									GetWindowRect(hWndSplitter, &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)), &rcEditMessage);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);

									LONG yPosSplitterTopNew{ (((rcSplitter.top - rcButtonOpen.bottom) + (rcSplitter.bottom - rcButtonOpen.bottom)) * (rcStatusBarFontInfo.top - rcButtonOpen.bottom)) / ((rcStatusBarFontInfoOld.top - rcButtonOpen.bottom) * 2) - ((rcSplitter.bottom - rcSplitter.top) / 2) + rcButtonOpen.bottom };
									FORWARD_WM_MOVE(hWndSplitter, 0, yPosSplitterTopNew, SendMessage);

									// Resize ListViewFontList
									HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
									SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), yPosSplitterTopNew - rcButtonOpen.bottom, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);

									// Resize EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
									SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, yPosSplitterTopNew + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcStatusBarFontInfo.top - (yPosSplitterTopNew + (rcSplitter.bottom - rcSplitter.top)), SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
								}
								break;
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
									┃ 1.Click "Select process", select a process.                                  │   ┃      ┃ 1.Click "Select process", select a process.                                  │   ┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨
									┃ view, then click "Load" button.                                              │ ↓ ┃      ┃ view, then click "Load" button.                                              │ ↓ ┃
									┠──────────────────────────────────────────────────────────────────────────────┴───┨      ┠──────────────────────────────────────────────────────────────────────────────┴───┨
									┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃      ┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛      ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
									*/

									// Resize StatusBarFontInfo
									HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
									FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);
									RECT rcStatusBarFontInfo{};
									GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

									// Resize ProgressBarFont
									HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
									if (IsWindowVisible(hWndProgressBarFont))
									{
										RECT rcStatusBarFontInfo{};
										GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
										int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
										SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
										RECT rcProgressBarFont{};
										SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
										SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}

									// Calculate the minimal height of ListViewFontList
									HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
									RECT rcListViewFontList{}, rcListViewFontListClient{}, rcHeaderListViewFontList{};
									GetWindowRect(hWndListViewFontList, &rcListViewFontList);
									GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
									GetWindowRect(ListView_GetHeader(hWndListViewFontList), &rcHeaderListViewFontList);
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
									LONG cyListViewFontListMin{ (rcListViewFontListItem.bottom - rcListViewFontListItem.top) + (rcHeaderListViewFontList.bottom - rcHeaderListViewFontList.top) + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

									// Resize ListViewFontList
									RECT rcButtonOpen{}, rcSplitter{}, rcEditMessage{};
									GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), &rcButtonOpen);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
									GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::Splitter)), &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)), &rcEditMessage);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);

									bool bIsListViewFontListMinimized{ false };
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									if (rcStatusBarFontInfo.top - rcButtonOpen.bottom - (rcSplitter.bottom - rcSplitter.top) - (rcEditMessage.bottom - rcEditMessage.top) < cyListViewFontListMin)
									{
										bIsListViewFontListMinimized = true;

										SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), cyListViewFontListMin, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}
									else
									{
										SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcStatusBarFontInfo.top - rcButtonOpen.bottom - (rcSplitter.bottom - rcSplitter.top) - (rcEditMessage.bottom - rcEditMessage.top), SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, static_cast<int>(ID::Splitter)) };
									if (bIsListViewFontListMinimized)
									{
										FORWARD_WM_MOVE(hWndSplitter, 0, rcButtonOpen.bottom + cyListViewFontListMin, SendMessage);
									}
									else
									{
										FORWARD_WM_MOVE(hWndSplitter, 0, rcStatusBarFontInfo.top - (rcSplitter.bottom - rcSplitter.top) - (rcEditMessage.bottom - rcEditMessage.top), SendMessage);
									}

									// Resize EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
									if (bIsListViewFontListMinimized)
									{
										SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, rcButtonOpen.bottom + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcStatusBarFontInfo.top - rcButtonOpen.bottom - cyListViewFontListMin - (rcSplitter.bottom - rcSplitter.top), SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}
									else
									{
										SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, rcStatusBarFontInfo.top - (rcEditMessage.bottom - rcEditMessage.top), LOWORD(lParam), rcEditMessage.bottom - rcEditMessage.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
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
									┃ 1.Click "Select process", select a process.                                  │   ┃      ┃ 1.Click "Select process", select a process.                                  │   ┃
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
									HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
									FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);
									RECT rcStatusBarFontInfo{};
									GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

									// Resize ProgressBarFont
									HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
									if (IsWindowVisible(hWndProgressBarFont))
									{
										RECT rcStatusBarFontInfo{};
										GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
										int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
										SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
										RECT rcProgressBarFont{};
										SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
										SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}

									// Calculate the minimal height of EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
									HDC hDCEditMessage{ GetDC(hWndEditMessage) };
									SelectFont(hDCEditMessage, GetWindowFont(hWndEditMessage));
									TEXTMETRIC tm{};
									BOOL bRetGetTextMetrics{ GetTextMetrics(hDCEditMessage, &tm) };
									assert(bRetGetTextMetrics);
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

										SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, rcStatusBarFontInfo.top - cyEditMessageMin, LOWORD(lParam), cyEditMessageMin, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}
									else
									{
										SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, rcEditMessage.top, LOWORD(lParam), rcStatusBarFontInfo.top - rcEditMessage.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, static_cast<int>(ID::Splitter)) };
									RECT rcSplitter{};
									GetWindowRect(hWndSplitter, &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									if (bIsEditMessageMinimized)
									{
										FORWARD_WM_MOVE(hWndSplitter, 0, rcStatusBarFontInfo.top - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top), SendMessage);
									}
									else
									{
										FORWARD_WM_MOVE(hWndSplitter, 0, rcSplitter.top, SendMessage);
									}

									// Resize ListViewFontList
									HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
									RECT rcButtonOpen{}, rcListViewFontList{};
									GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), &rcButtonOpen);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
									GetWindowRect(hWndListViewFontList, &rcListViewFontList);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									if (bIsEditMessageMinimized)
									{
										SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcStatusBarFontInfo.top - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top) - rcButtonOpen.bottom, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}
									else
									{
										SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcListViewFontList.bottom - rcListViewFontList.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
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
									┃ How to load fonts to process:                                                │   ┃      ┃ 1.Click "Select process", select a process.                                                │   ┃
									┃ 1.Click "Select process", select a process.                                  │   ┃      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list view, then     │   ┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨      ┃  click "Load" button.                                                                      ├───┨
									┃ view, then click "Load" button.                                              │ ↓ ┃      ┃                                                                                            │ ↓ ┃
									┠──────────────────────────────────────────────────────────────────────────────┴───┨      ┠────────────────────────────────────────────────────────────────────────────────────────────┴───┨
									┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃      ┃ 0 font(s) opened, 0 font(s) loaded.                                                            ┃
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛      ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
									*/

									// Resize StatusBarFontInfo
									HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
									FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);

									// Resize ProgressBarFont
									HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
									if (IsWindowVisible(hWndProgressBarFont))
									{
										RECT rcStatusBarFontInfo{};
										GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
										int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
										SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
										RECT rcProgressBarFont{};
										SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
										SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
									}

									// Resize ListViewFontList
									HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
									RECT rcListViewFontList{};
									GetWindowRect(hWndListViewFontList, &rcListViewFontList);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcListViewFontList.bottom - rcListViewFontList.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, static_cast<int>(ID::Splitter)) };
									RECT rcSplitter{};
									GetWindowRect(hWndSplitter, &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									FORWARD_WM_MOVE(hWndSplitter, 0, rcSplitter.top, SendMessage);

									// Resize EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
									RECT rcEditMessage{};
									GetWindowRect(hWndEditMessage, &rcEditMessage);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);
									SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, rcEditMessage.top, LOWORD(lParam), rcEditMessage.bottom - rcEditMessage.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
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
							┃ 1.Click "Select process", select a process.                                  │   ┃
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
							┃ 1.Click "Select process", select a process.                                                                               │   ┃
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
							HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
							FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);
							RECT rcStatusBarFontInfo{};
							GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

							// Resize ProgressBarFont
							HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
							if (IsWindowVisible(hWndProgressBarFont))
							{
								RECT rcStatusBarFontInfo{};
								GetClientRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
								int aiStatusBarFontInfoParts[]{ rcStatusBarFontInfo.right - 150, -1 };
								SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 2, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));
								RECT rcProgressBarFont{};
								SendMessage(hWndStatusBarFontInfo, SB_GETRECT, 1, reinterpret_cast<LPARAM>(&rcProgressBarFont));
								SetWindowPos(hWndProgressBarFont, NULL, rcProgressBarFont.left, rcProgressBarFont.top, rcProgressBarFont.right - rcProgressBarFont.left, rcProgressBarFont.bottom - rcProgressBarFont.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
							}

							// Resize Splitter
							HWND hWndSplitter{ GetDlgItem(hWnd, static_cast<int>(ID::Splitter)) };
							RECT rcButtonOpen{}, rcListViewFontList{}, rcSplitter{}, rcEditMessage{};
							GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), &rcButtonOpen);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
							GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), &rcListViewFontList);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
							GetWindowRect(hWndSplitter, &rcSplitter);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
							GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)), &rcEditMessage);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);

							LONG yPosSplitterTopNew{ (((rcSplitter.top - rcButtonOpen.bottom) + (rcSplitter.bottom - rcButtonOpen.bottom)) * (rcStatusBarFontInfo.top - rcButtonOpen.bottom)) / ((rcStatusBarFontInfoOld.top - rcButtonOpen.bottom) * 2) - ((rcSplitter.bottom - rcSplitter.top) / 2) + rcButtonOpen.bottom };
							FORWARD_WM_MOVE(hWndSplitter, 0, yPosSplitterTopNew, SendMessage);

							// Resize ListViewFontList
							HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
							SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), yPosSplitterTopNew - rcButtonOpen.bottom, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);

							// Resize EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
							SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, yPosSplitterTopNew + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcStatusBarFontInfo.top - (yPosSplitterTopNew + (rcSplitter.bottom - rcSplitter.top)), SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);

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
							┃ 1.Click "Select process", select a process.                                                                               │   ┃
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
							┃ 1.Click "Select process", select a process.                                  │   ┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list  ├───┨
							┃ view, then click "Load" button.                                              │ ↓ ┃
							┠──────────────────────────────────────────────────────────────────────────────┴───┨
							┃ 0 font(s) opened, 0 font(s) loaded.                                              ┃
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
							*/

							// Resize StatusBarFontInfo
							HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
							FORWARD_WM_SIZE(hWndStatusBarFontInfo, 0, 0, 0, SendMessage);
							RECT rcStatusBarFontInfo{};
							GetWindowRect(hWndStatusBarFontInfo, &rcStatusBarFontInfo);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcStatusBarFontInfo);

							// Calculate the minimal height of ListViewFontList
							HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
							RECT rcListViewFontList{}, rcListViewFontListClient{}, rcHeaderListViewFontList{};
							GetWindowRect(hWndListViewFontList, &rcListViewFontList);
							GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
							GetWindowRect(ListView_GetHeader(hWndListViewFontList), &rcHeaderListViewFontList);
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
							LONG cyListViewFontListMin{ (rcListViewFontListItem.bottom - rcListViewFontListItem.top) + (rcHeaderListViewFontList.bottom - rcHeaderListViewFontList.top) + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

							// Calculate the minimal height of EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
							HDC hDCEditMessage{ GetDC(hWndEditMessage) };
							SelectFont(hDCEditMessage, GetWindowFont(hWndEditMessage));
							TEXTMETRIC tm{};
							BOOL bRetGetTextMetrics{ GetTextMetrics(hDCEditMessage, &tm) };
							assert(bRetGetTextMetrics);
							ReleaseDC(hWndEditMessage, hDCEditMessage);
							RECT rcEditMessage{}, rcEditMessageClient{};
							GetWindowRect(hWndEditMessage, &rcEditMessage);
							GetClientRect(hWndEditMessage, &rcEditMessageClient);
							LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + cyEditMessageTextMargin };

							// Calculate new position of splitter
							HWND hWndSplitter{ GetDlgItem(hWnd, static_cast<int>(ID::Splitter)) };
							RECT rcButtonOpen{}, rcSplitter{};
							GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), &rcButtonOpen);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
							GetWindowRect(hWndSplitter, &rcSplitter);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
							LONG yPosSplitterTopNew{ (((rcSplitter.top - rcButtonOpen.bottom) + (rcSplitter.bottom - rcButtonOpen.bottom)) * (rcStatusBarFontInfo.top - rcButtonOpen.bottom)) / ((rcStatusBarFontInfoOld.top - rcButtonOpen.bottom) * 2) - ((rcSplitter.bottom - rcSplitter.top) / 2) + rcButtonOpen.bottom };

							MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);
							LONG cyListViewFontListNew{ yPosSplitterTopNew - rcButtonOpen.bottom };
							LONG cyEditMessageNew{ rcStatusBarFontInfo.top - yPosSplitterTopNew - (rcSplitter.bottom - rcSplitter.top) };
							// If cyListViewFontListNew < cyListViewFontListMin, keep the minimal height of ListViewFontList
							if (cyListViewFontListNew < cyListViewFontListMin)
							{
								// Resize ListViewFontList
								HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
								SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), cyListViewFontListMin, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);

								// Resize Splitter
								FORWARD_WM_MOVE(hWndSplitter, 0, rcButtonOpen.bottom + cyListViewFontListMin, SendMessage);

								// Resize EditMessage
								HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
								SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, rcButtonOpen.bottom + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcStatusBarFontInfo.top - rcButtonOpen.bottom - cyListViewFontListMin - (rcSplitter.bottom - rcSplitter.top), SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
							}
							// If cyEditMessageNew < cyEditMessageMin, keep the minimal height of EditMessage
							else if (cyEditMessageNew < cyEditMessageMin)
							{
								// Resize EditMessage
								HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
								SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, rcStatusBarFontInfo.top - cyEditMessageMin, LOWORD(lParam), cyEditMessageMin, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);

								// Resize Splitter
								FORWARD_WM_MOVE(hWndSplitter, 0, rcStatusBarFontInfo.top - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top), SendMessage);

								// Resize ListViewFontList
								HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
								SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcStatusBarFontInfo.top - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top) - rcButtonOpen.bottom, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
							}
							// Else resize as usual
							else
							{
								// Resize Splitter
								FORWARD_WM_MOVE(hWndSplitter, 0, yPosSplitterTopNew, SendMessage);

								// Resize ListViewFontList
								HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
								SetWindowPos(hWndListViewFontList, NULL, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), cyListViewFontListNew, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);

								// Resize EditMessage
								HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
								SetWindowPos(hWndEditMessage, NULL, rcEditMessage.left, yPosSplitterTopNew + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), cyEditMessageNew, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
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

#ifdef _DEBUG
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			ssMessage << L"* Main revieved WM_SIZE\r\n"
				<< L"cxClient: " << LOWORD(lParam) << L" "
				<< L"cyClient: " << HIWORD(lParam) << L" "
				<< L"PreviousShowCmd: ";
			switch (PreviousShowCmd)
			{
			case SW_HIDE:
				{
					ssMessage << L"SW_HIDE";
				}
				break;
			case SW_SHOWNORMAL:
				{
					ssMessage << L"SW_SHOWNORMAL";
				}
				break;
			case SW_SHOWMINIMIZED:
				{
					ssMessage << L"SW_SHOWMINIMIZED";
				}
				break;
			case SW_SHOWMAXIMIZED:
				{
					ssMessage << L"SW_SHOWMAXIMIZED";
				}
				break;
			case SW_SHOWNOACTIVATE:
				{
					ssMessage << L" SW_SHOWNOACTIVATE";
				}
				break;
			case SW_SHOW:
				{
					ssMessage << L"SW_SHOW";
				}
				break;
			case SW_MINIMIZE:
				{
					ssMessage << L"SW_MINIMIZE";
				}
				break;
			case SW_SHOWMINNOACTIVE:
				{
					ssMessage << L"SW_SHOWMINNOACTIVE";
				}
				break;
			case SW_SHOWNA:
				{
					ssMessage << L"SW_SHOWNA";
				}
				break;
			case SW_RESTORE:
				{
					ssMessage << L"SW_RESTORE";
				}
				break;
			case SW_SHOWDEFAULT:
				{
					ssMessage << L"SW_SHOWDEFAULT";
				}
				break;
			case SW_FORCEMINIMIZE:
				{
					ssMessage << L"SW_FORCEMINIMIZE";
				}
				break;
			default:
				break;
			}
			ssMessage << L" SizingEdge: ";
			switch (SizingEdge)
			{
			case 0:
				{
					ssMessage << L"None";
				}
				break;
			case WMSZ_LEFT:
				{
					ssMessage << L"WMSZ_LEFT";
				}
				break;
			case WMSZ_RIGHT:
				{
					ssMessage << L"WMSZ_RIGHT";
				}
				break;
			case WMSZ_TOP:
				{
					ssMessage << L"WMSZ_TOP";
				}
				break;
			case WMSZ_TOPLEFT:
				{
					ssMessage << L"WMSZ_TOPLEFT";
				}
				break;
			case WMSZ_TOPRIGHT:
				{
					ssMessage << L"WMSZ_TOPRIGHT";
				}
				break;
			case WMSZ_BOTTOM:
				{
					ssMessage << L"WMSZ_BOTTOM";
				}
				break;
			case WMSZ_BOTTOMLEFT:
				{
					ssMessage << L"WMSZ_BOTTOMLEFT";
				}
				break;
			case WMSZ_BOTTOMRIGHT:
				{
					ssMessage << L"WMSZ_BOTTOMRIGHT";
				}
				break;
			default:
				break;
			}
			ssMessage << L" flag: ";
			switch (wParam)
			{
			case SIZE_RESTORED:
				{
					ssMessage << L"SIZE_RESTORED";
				}
				break;
			case SIZE_MINIMIZED:
				{
					ssMessage << L"SIZE_MINIMIZED";
				}
				break;
			case SIZE_MAXSHOW:
				{
					ssMessage << L"SIZE_MAXSHOW";
				}
				break;
			case SIZE_MAXIMIZED:
				{
					ssMessage << L"SIZE_MAXIMIZED";
				}
				break;
			case SIZE_MAXHIDE:
				{
					ssMessage << L"SIZE_MAXHIDE";
				}
				break;
			default:
				break;
			}
			ssMessage << L"\r\n";
			strMessage = ssMessage.str();
			OutputDebugString(strMessage.c_str());
#endif

			// Clear sizing edge
			SizingEdge = 0;

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
																							  │
																							  │ ← rcEditTimeout.right
																							  │
			*/

			// Get ButtonOpen, Splitter, StatusBarFontInfo and EditTimeout window rectangle
			RECT rcButtonOpen{}, rcSplitter{}, rcStatusBarFontInfo{}, rcEditTimeout{};
			GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), &rcButtonOpen);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
			GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::Splitter)), &rcSplitter);
			GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)), &rcStatusBarFontInfo);
			GetWindowRect(GetDlgItem(hWnd, static_cast<int>(ID::EditTimeout)), &rcEditTimeout);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcEditTimeout);

			// Calculate the minimal height of ListViewFontList
			HWND hWndListViewFontList{ GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)) };
			RECT rcListViewFontList{}, rcListViewFontListClient{}, rcHeaderListViewFontList{};
			GetWindowRect(hWndListViewFontList, &rcListViewFontList);
			GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
			GetWindowRect(ListView_GetHeader(hWndListViewFontList), &rcHeaderListViewFontList);
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
			LONG cyListViewFontListMin{ (rcListViewFontListItem.bottom - rcListViewFontListItem.top) + (rcHeaderListViewFontList.bottom - rcHeaderListViewFontList.top) + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

			// Calculate the minimal height of Editmessage
			HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };
			HDC hDCEditMessage{ GetDC(hWndEditMessage) };
			SelectFont(hDCEditMessage, GetWindowFont(hWndEditMessage));
			TEXTMETRIC tm{};
			BOOL bRetGetTextMetrics{ GetTextMetrics(hDCEditMessage, &tm) };
			assert(bRetGetTextMetrics);
			ReleaseDC(hWndEditMessage, hDCEditMessage);
			RECT rcEditMessage{}, rcEditMessageClient{};
			GetWindowRect(hWndEditMessage, &rcEditMessage);
			GetClientRect(hWndEditMessage, &rcEditMessageClient);
			LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + cyEditMessageTextMargin };

			// Calculate the minimal window size
			RECT rcMainMin{ 0, 0, rcEditTimeout.right, (rcButtonOpen.bottom - rcButtonOpen.top) + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top) + cyEditMessageMin + (rcStatusBarFontInfo.bottom - rcStatusBarFontInfo.top) };
			AdjustWindowRect(&rcMainMin, GetWindowStyle(hWnd), FALSE);
			reinterpret_cast<LPMINMAXINFO>(lParam)->ptMinTrackSize = { rcMainMin.right - rcMainMin.left, rcMainMin.bottom - rcMainMin.top };
		}
		break;
	case WM_CLOSE:
		{
			// Do cleanup
			HWND hWndEditMessage{ GetDlgItem(hWnd, static_cast<int>(ID::EditMessage)) };

			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			int cchMessageLength{};

			// If loaded via proxy
			if (ProxyProcessInfo.hProcess)
			{
				// Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
				COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::PULLDLL), 0, NULL };
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
						ssMessage << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
						strMessage = ssMessage.str();
						cchMessageLength = Edit_GetTextLength(hWndEditMessage);
						Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
						Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
						ssMessage.str(L"");

						// Re-enable controls
						EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), TRUE);
						EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), TRUE);
						EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
						EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), TRUE);
						EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), TRUE);
						EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), TRUE);

						// Update StatusBarFontInfo
						HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
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
						SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(strFontInfo.c_str()));

						// Set ProgressBarFont
						HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
						SendMessage(hWndProgressBarFont, PBM_SETPOS, 0, 0);
						ShowWindow(hWndProgressBarFont, SW_HIDE);

						int aiStatusBarFontInfoParts[]{ -1 };
						SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));

						// Update syatem tray icon tip
						if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
						{
							NOTIFYICONDATA nid{ sizeof(NOTIFYICONDATA), hWndMain, 0, NIF_TIP | NIF_SHOWTIP };
							wcscpy_s(nid.szTip, strFontInfo.c_str());
							Shell_NotifyIcon(NIM_MODIFY, &nid);
						}
					}
					goto break_69504405;
				default:
					{
						assert(0 && "invalid ProxyDllPullingResult");
					}
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
				DWORD dwExitCodeMessageThread{};
				GetExitCodeThread(hThreadMessage, &dwExitCodeMessageThread);
				if (dwExitCodeMessageThread)
				{
					ssMessage << L"Message thread exited abnormally with code " << dwExitCodeMessageThread << L"\r\n\r\n";
					strMessage = ssMessage.str();
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());
				}
				CloseHandle(hThreadMessage);

				// Terminate proxy process
				COPYDATASTRUCT cds2{ static_cast<ULONG_PTR>(COPYDATA::TERMINATE), 0, NULL };
				FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
				WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
				CloseHandle(ProxyProcessInfo.hProcess);
				ProxyProcessInfo.hProcess = NULL;

				// Close the handle to target process and duplicated handles
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
					ssMessage << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L")\r\n\r\n";
					strMessage = ssMessage.str();
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, strMessage.c_str());

					// Re-enable controls
					EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonOpen)), TRUE);
					EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ListViewFontList)), TRUE);
					EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
					EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonClose)), TRUE);
					EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonLoad)), TRUE);
					EnableWindow(GetDlgItem(hWnd, static_cast<int>(ID::ButtonUnload)), TRUE);

					// Update StatusBarFontInfo
					HWND hWndStatusBarFontInfo{ GetDlgItem(hWnd, static_cast<int>(ID::StatusBarFontInfo)) };
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
					SendMessage(hWndStatusBarFontInfo, SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(strFontInfo.c_str()));

					// Set ProgressBarFont
					HWND hWndProgressBarFont{ GetDlgItem(hWndStatusBarFontInfo, static_cast<int>(ID::ProgressBarFont)) };
					SendMessage(hWndProgressBarFont, PBM_SETPOS, 0, 0);
					ShowWindow(hWndProgressBarFont, SW_HIDE);

					int aiStatusBarFontInfoParts[]{ -1 };
					SendMessage(hWndStatusBarFontInfo, SB_SETPARTS, 1, reinterpret_cast<LPARAM>(aiStatusBarFontInfoParts));

					// Update syatem tray icon tip
					if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
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

				// Close the handle to target process
				CloseHandle(TargetProcessInfo.hProcess);
				TargetProcessInfo.hProcess = NULL;
			}

			// Remove the icon from system tray
			if (Button_GetCheck(GetDlgItem(hWnd, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
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
			DeletePen(hPenSplitter);

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
					assert(hClipboardData);
					LPCWSTR lpszText{ static_cast<LPCWSTR>(GlobalLock(hClipboardData)) };
					assert(lpszText);

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
			if (wParam == 0x41u)	// Virtual key code of 'A' key
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
			HWND hWndParent{ GetAncestor(hWndListViewFontList, GA_PARENT) };
			HWND hWndEditMessage{ GetDlgItem(hWndParent, static_cast<int>(ID::EditMessage)) };

			bool bIsFontListEmptyBefore{ FontList.empty() };

			UINT nFileCount{ DragQueryFile(reinterpret_cast<HDROP>(wParam), 0xFFFFFFFF, NULL, 0) };
			FONTLISTCHANGEDSTRUCT flcs{ ListView_GetItemCount(hWndListViewFontList) };
			for (UINT i = 0; i < nFileCount; i++)
			{
				WCHAR szFileName[MAX_PATH]{};
				DragQueryFile(reinterpret_cast<HDROP>(wParam), i, szFileName, MAX_PATH);
				if (PathIsDirectory(szFileName))
				{
					WCHAR szFontFileName[MAX_PATH]{};
					if (PathCombine(szFontFileName, szFileName, L"*"))
					{
						WIN32_FIND_DATA w32fd{};
						HANDLE hFindFile{ FindFirstFile(szFontFileName, &w32fd) };
						if (hFindFile)
						{
							do
							{
								if (PathMatchSpec(w32fd.cFileName, L"*.ttf") || PathMatchSpec(w32fd.cFileName, L"*.ttc") || PathMatchSpec(w32fd.cFileName, L"*.otf"))
								{
									PathCombine(szFontFileName, szFileName, w32fd.cFileName);

									bool bIsFontDuplicate{ false };
									for (const auto& j : FontList)
									{
										if (j.GetFontName() == szFontFileName)
										{
											bIsFontDuplicate = true;

											break;
										}
									}
									if (!bIsFontDuplicate)
									{
										FontList.push_back(szFontFileName);

										flcs.lpszFontName = szFontFileName;
										SendMessage(GetAncestor(hWndListViewFontList, GA_PARENT), static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::OPENED), reinterpret_cast<LPARAM>(&flcs));

										flcs.iItem++;
									}
								}
							} while (FindNextFile(hFindFile, &w32fd));

							FindClose(hFindFile);
						}
					}
				}
				else
				{
					if (PathMatchSpec(szFileName, L"*.ttf") || PathMatchSpec(szFileName, L"*.ttc") || PathMatchSpec(szFileName, L"*.otf"))
					{
						bool bIsFontDuplicate{ false };
						for (const auto& j : FontList)
						{
							if (j.GetFontName() == szFileName)
							{
								bIsFontDuplicate = true;

								break;
							}
						}
						if (!bIsFontDuplicate)
						{
							FontList.push_back(szFileName);

							flcs.lpszFontName = szFileName;
							SendMessage(GetAncestor(hWndListViewFontList, GA_PARENT), static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::OPENED), reinterpret_cast<LPARAM>(&flcs));

							flcs.iItem++;
						}
					}
				}
			}
			int cchMessageLength{ Edit_GetTextLength(hWndEditMessage) };
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, L"\r\n");

			DragFinish(reinterpret_cast<HDROP>(wParam));

			if (bIsFontListEmptyBefore && (!FontList.empty()))
			{
				EnableWindow(GetDlgItem(hWndMain, static_cast<int>(ID::ButtonClose)), TRUE);
				EnableWindow(GetDlgItem(hWndMain, static_cast<int>(ID::ButtonLoad)), TRUE);
				EnableWindow(GetDlgItem(hWndMain, static_cast<int>(ID::ButtonUnload)), TRUE);
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
			SendMessage(GetDlgItem(hWndParent, static_cast<int>(ID::StatusBarFontInfo)), SB_SETTEXT, MAKEWPARAM(MAKEWORD(0, 0), 0), reinterpret_cast<LPARAM>(strFontInfo.c_str()));

			// Update syatem tray icon tip
			if (Button_GetCheck(GetDlgItem(hWndParent, static_cast<int>(ID::ButtonMinimizeToTray))) == BST_CHECKED)
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

			PostMessage(GetAncestor(hWndListViewFontList, GA_PARENT), static_cast<UINT>(USERMESSAGE::CHILDWINDOWPOSCHANGED), reinterpret_cast<WPARAM>(hWndListViewFontList), static_cast<LPARAM>(reinterpret_cast<LPWINDOWPOS>(lParam)->flags));

#ifdef DBGPRINTWNDPOSINFO
			LPWINDOWPOS lpwp{ reinterpret_cast<LPWINDOWPOS>(lParam) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			ssMessage << L"* ListViewFontList revieved WM_WINDOWPOSCHANGED\r\n"
				<< L"hwndInsertAfter: " << lpwp->hwndInsertAfter << L" "
				<< L"hwnd: " << lpwp->hwnd << L" "
				<< L"x: " << lpwp->x << L" "
				<< L"y: " << lpwp->y << L" "
				<< L"cx: " << lpwp->cx << L" "
				<< L"cy: " << lpwp->cy << L" "
				<< L"flags: ";
			if (lpwp->flags & SWP_NOSIZE)
			{
				ssMessage << L"SWP_NOSIZE | ";
			}
			if (lpwp->flags & SWP_NOMOVE)
			{
				ssMessage << L"SWP_NOMOVE | ";
			}
			if (lpwp->flags & SWP_NOZORDER)
			{
				ssMessage << L"SWP_NOZORDER | ";
			}
			if (lpwp->flags & SWP_NOREDRAW)
			{
				ssMessage << L"SWP_NOREDRAW | ";
			}
			if (lpwp->flags & SWP_NOACTIVATE)
			{
				ssMessage << L"SWP_NOACTIVATE | ";
			}
			if (lpwp->flags & SWP_FRAMECHANGED)
			{
				ssMessage << L"SWP_FRAMECHANGED | ";
			}
			if (lpwp->flags & SWP_SHOWWINDOW)
			{
				ssMessage << L"SWP_SHOWWINDOW | ";
			}
			if (lpwp->flags & SWP_HIDEWINDOW)
			{
				ssMessage << L"SWP_HIDEWINDOW | ";
			}
			if (lpwp->flags & SWP_NOCOPYBITS)
			{
				ssMessage << L"SWP_NOCOPYBITS | ";
			}
			if (lpwp->flags & SWP_NOOWNERZORDER)
			{
				ssMessage << L"SWP_NOOWNERZORDER | ";
			}
			if (lpwp->flags & SWP_NOSENDCHANGING)
			{
				ssMessage << L"SWP_NOSENDCHANGING | ";
			}
			ssMessage << L"\r\n";
			strMessage = ssMessage.str();
			std::size_t pos{ strMessage.find_last_of(L'|') };
			if (pos != std::wstring::npos)
			{
				strMessage.erase(strMessage.find_last_of(L'|') - 1, 3);
			}
			OutputDebugString(strMessage.c_str());
#endif // SHOWPOSINFO
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
			if (wParam == 0x41u)	// Virtual key code of 'A' key
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
			static HWND hWndContextMenuOwner{};
			static POINT ptAnchorContextMenu{};
			hWndContextMenuOwner = reinterpret_cast<HWND>(wParam);
			ptAnchorContextMenu = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };

			HWINEVENTHOOK hWinEventHook{ SetWinEventHook(EVENT_SYSTEM_MENUPOPUPSTART, EVENT_SYSTEM_MENUPOPUPSTART, NULL,
				[](HWINEVENTHOOK hWinEventHook, DWORD Event, HWND hWnd, LONG idObject, LONG idChild, DWORD idEventThread, DWORD dwmsEventTime)
				{
					if (idObject == OBJID_CLIENT && idChild == CHILDID_SELF)
					{
						// Test whether context menu is triggered by mouse or by keyboard
						if (ptAnchorContextMenu.x == -1 && ptAnchorContextMenu.y == -1)
						{
							Edit_ScrollCaret(hWndContextMenuOwner);
							LRESULT lRet{ SendMessage(hWndContextMenuOwner, EM_POSFROMCHAR, Edit_GetCaretIndex(hWndContextMenuOwner), NULL) };
							ptAnchorContextMenu = { GET_X_LPARAM(lRet), GET_Y_LPARAM(lRet) };
							ClientToScreen(hWndContextMenuOwner, &ptAnchorContextMenu);
						}

						// Test whether context menu pops up in the client area of EditMessage
						RECT rcEditMessageClient{};
						GetClientRect(hWndContextMenuOwner, &rcEditMessageClient);
						MapWindowRect(hWndContextMenuOwner, HWND_DESKTOP, &rcEditMessageClient);
						if (PtInRect(&rcEditMessageClient, ptAnchorContextMenu))
						{
							HMENU hMenuContextEditMessage{ reinterpret_cast<HMENU>(SendMessage(hWnd, MN_GETHMENU, 0, 0)) };
							assert(hMenuContextEditMessage);

							// Menu item identifiers in Edit control context menu is the same as corresponding Windows messages
							DeleteMenu(hMenuContextEditMessage, WM_UNDO, MF_BYCOMMAND);	// Undo
							DeleteMenu(hMenuContextEditMessage, WM_CUT, MF_BYCOMMAND);		// Cut
							DeleteMenu(hMenuContextEditMessage, WM_PASTE, MF_BYCOMMAND);	// Paste
							DeleteMenu(hMenuContextEditMessage, WM_CLEAR, MF_BYCOMMAND);	// Clear
							DeleteMenu(hMenuContextEditMessage, 0, MF_BYPOSITION);			// Seperator
							InsertMenu(hMenuContextEditMessage, 0, MF_BYPOSITION | MF_SEPARATOR, 0, NULL);
							InsertMenu(hMenuContextEditMessage, 0, MF_BYPOSITION | MF_STRING, WM_USER + 100, L"C&lear log");

							// Adjust context menu position
							RECT rcContextMenuEditMessage{};
							GetWindowRect(hWnd, &rcContextMenuEditMessage);
							MONITORINFO mi{ sizeof(MONITORINFO) };
							GetMonitorInfo(MonitorFromPoint(ptAnchorContextMenu, MONITOR_DEFAULTTONEAREST), &mi);
							SIZE sizeContextMenuEditMessage{ rcContextMenuEditMessage.right - rcContextMenuEditMessage.left, rcContextMenuEditMessage.bottom - rcContextMenuEditMessage.top };
							UINT uFlags{ TPM_WORKAREA };
							if (ptAnchorContextMenu.y > mi.rcWork.bottom - sizeContextMenuEditMessage.cy)
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
							CalculatePopupWindowPosition(&ptAnchorContextMenu, &sizeContextMenuEditMessage, uFlags, NULL, &rcContextMenuEditMessage);
							SetWindowPos(hWnd, NULL, rcContextMenuEditMessage.left, rcContextMenuEditMessage.top, rcContextMenuEditMessage.right - rcContextMenuEditMessage.left, rcContextMenuEditMessage.bottom - rcContextMenuEditMessage.top, SWP_NOZORDER | SWP_NOREDRAW | SWP_NOACTIVATE);
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

			PostMessage(GetAncestor(hWndEditMessage, GA_PARENT), static_cast<UINT>(USERMESSAGE::CHILDWINDOWPOSCHANGED), reinterpret_cast<WPARAM>(hWndEditMessage), static_cast<LPARAM>(reinterpret_cast<LPWINDOWPOS>(lParam)->flags));

#ifdef DBGPRINTWNDPOSINFO
			LPWINDOWPOS lpwp{ reinterpret_cast<LPWINDOWPOS>(lParam) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			ssMessage << L"* EditMessage revieved WM_WINDOWPOSCHANGED\r\n"
				<< L"hwndInsertAfter: " << lpwp->hwndInsertAfter << L" "
				<< L"hwnd: " << lpwp->hwnd << L" "
				<< L"x: " << lpwp->x << L" "
				<< L"y: " << lpwp->y << L" "
				<< L"cx: " << lpwp->cx << L" "
				<< L"cy: " << lpwp->cy << L" "
				<< L"flags: ";
			if (lpwp->flags & SWP_NOSIZE)
			{
				ssMessage << L"SWP_NOSIZE | ";
			}
			if (lpwp->flags & SWP_NOMOVE)
			{
				ssMessage << L"SWP_NOMOVE | ";
			}
			if (lpwp->flags & SWP_NOZORDER)
			{
				ssMessage << L"SWP_NOZORDER | ";
			}
			if (lpwp->flags & SWP_NOREDRAW)
			{
				ssMessage << L"SWP_NOREDRAW | ";
			}
			if (lpwp->flags & SWP_NOACTIVATE)
			{
				ssMessage << L"SWP_NOACTIVATE | ";
			}
			if (lpwp->flags & SWP_FRAMECHANGED)
			{
				ssMessage << L"SWP_FRAMECHANGED | ";
			}
			if (lpwp->flags & SWP_SHOWWINDOW)
			{
				ssMessage << L"SWP_SHOWWINDOW | ";
			}
			if (lpwp->flags & SWP_HIDEWINDOW)
			{
				ssMessage << L"SWP_HIDEWINDOW | ";
			}
			if (lpwp->flags & SWP_NOCOPYBITS)
			{
				ssMessage << L"SWP_NOCOPYBITS | ";
			}
			if (lpwp->flags & SWP_NOOWNERZORDER)
			{
				ssMessage << L"SWP_NOOWNERZORDER | ";
			}
			if (lpwp->flags & SWP_NOSENDCHANGING)
			{
				ssMessage << L"SWP_NOSENDCHANGING | ";
			}
			ssMessage << L"\r\n";
			strMessage = ssMessage.str();
			std::size_t pos{ strMessage.find_last_of(L'|') };
			if (pos != std::wstring::npos)
			{
				strMessage.erase(strMessage.find_last_of(L'|') - 1, 3);
			}
			OutputDebugString(strMessage.c_str());
#endif // SHOWPOSINFO
		}
		break;
		// "Clear log" custom message
	case WM_USER + 100:
		{
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
				R"(1.Click "Select process", select a process.)""\r\n"
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
				R"("Minimize to tray": If checked, click the minimize or close button in the upper-right cornor will minimize the window to system tray.)""\r\n"
				R"("Font Name": Names of the fonts added to the list view.)""\r\n"
				R"("State": State of the font. There are five states, "Not loaded", "Loaded", "Load failed", "Unloaded" and "Unload failed".)""\r\n"
				"\r\n"
			);
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

#ifdef DBGPRINTWNDPOSINFO
LRESULT CALLBACK SplitterSubclassProc(HWND hWndSplitter, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData)
{
	LRESULT ret{};

	switch (Message)
	{
	case WM_WINDOWPOSCHANGED:
		{
			ret = DefSubclassProc(hWndSplitter, Message, wParam, lParam);

			LPWINDOWPOS lpwp{ reinterpret_cast<LPWINDOWPOS>(lParam) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			ssMessage << L"* Splitter revieved WM_WINDOWPOSCHANGED\r\n"
				<< L"hwndInsertAfter: " << lpwp->hwndInsertAfter << L" "
				<< L"hwnd: " << lpwp->hwnd << L" "
				<< L"x: " << lpwp->x << L" "
				<< L"y: " << lpwp->y << L" "
				<< L"cx: " << lpwp->cx << L" "
				<< L"cy: " << lpwp->cy << L" "
				<< L"flags: ";
			if (lpwp->flags & SWP_NOSIZE)
			{
				ssMessage << L"SWP_NOSIZE | ";
			}
			if (lpwp->flags & SWP_NOMOVE)
			{
				ssMessage << L"SWP_NOMOVE | ";
			}
			if (lpwp->flags & SWP_NOZORDER)
			{
				ssMessage << L"SWP_NOZORDER | ";
			}
			if (lpwp->flags & SWP_NOREDRAW)
			{
				ssMessage << L"SWP_NOREDRAW | ";
			}
			if (lpwp->flags & SWP_NOACTIVATE)
			{
				ssMessage << L"SWP_NOACTIVATE | ";
			}
			if (lpwp->flags & SWP_FRAMECHANGED)
			{
				ssMessage << L"SWP_FRAMECHANGED | ";
			}
			if (lpwp->flags & SWP_SHOWWINDOW)
			{
				ssMessage << L"SWP_SHOWWINDOW | ";
			}
			if (lpwp->flags & SWP_HIDEWINDOW)
			{
				ssMessage << L"SWP_HIDEWINDOW | ";
			}
			if (lpwp->flags & SWP_NOCOPYBITS)
			{
				ssMessage << L"SWP_NOCOPYBITS | ";
			}
			if (lpwp->flags & SWP_NOOWNERZORDER)
			{
				ssMessage << L"SWP_NOOWNERZORDER | ";
			}
			if (lpwp->flags & SWP_NOSENDCHANGING)
			{
				ssMessage << L"SWP_NOSENDCHANGING | ";
			}
			ssMessage << L"\r\n";
			strMessage = ssMessage.str();
			std::size_t pos{ strMessage.find_last_of(L'|') };
			if (pos != std::wstring::npos)
			{
				strMessage.erase(strMessage.find_last_of(L'|') - 1, 3);
			}
			OutputDebugString(strMessage.c_str());
		}
		break;
	default:
		{
			ret = DefSubclassProc(hWndSplitter, Message, wParam, lParam);
		}
		break;
	}

	return ret;
}

LRESULT CALLBACK ProgressBarFontSubclassProc(HWND hWndProgressBarFont, UINT Message, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData)
{
	LRESULT ret{};

	switch (Message)
	{
	case WM_WINDOWPOSCHANGED:
		{
			ret = DefSubclassProc(hWndProgressBarFont, Message, wParam, lParam);

			LPWINDOWPOS lpwp{ reinterpret_cast<LPWINDOWPOS>(lParam) };
			std::wstringstream ssMessage{};
			std::wstring strMessage{};
			ssMessage << L"* ProgressBarFont revieved WM_WINDOWPOSCHANGED\r\n"
				<< L"hwndInsertAfter: " << lpwp->hwndInsertAfter << L" "
				<< L"hwnd: " << lpwp->hwnd << L" "
				<< L"x: " << lpwp->x << L" "
				<< L"y: " << lpwp->y << L" "
				<< L"cx: " << lpwp->cx << L" "
				<< L"cy: " << lpwp->cy << L" "
				<< L"flags: ";
			if (lpwp->flags & SWP_NOSIZE)
			{
				ssMessage << L"SWP_NOSIZE | ";
			}
			if (lpwp->flags & SWP_NOMOVE)
			{
				ssMessage << L"SWP_NOMOVE | ";
			}
			if (lpwp->flags & SWP_NOZORDER)
			{
				ssMessage << L"SWP_NOZORDER | ";
			}
			if (lpwp->flags & SWP_NOREDRAW)
			{
				ssMessage << L"SWP_NOREDRAW | ";
			}
			if (lpwp->flags & SWP_NOACTIVATE)
			{
				ssMessage << L"SWP_NOACTIVATE | ";
			}
			if (lpwp->flags & SWP_FRAMECHANGED)
			{
				ssMessage << L"SWP_FRAMECHANGED | ";
			}
			if (lpwp->flags & SWP_SHOWWINDOW)
			{
				ssMessage << L"SWP_SHOWWINDOW | ";
			}
			if (lpwp->flags & SWP_HIDEWINDOW)
			{
				ssMessage << L"SWP_HIDEWINDOW | ";
			}
			if (lpwp->flags & SWP_NOCOPYBITS)
			{
				ssMessage << L"SWP_NOCOPYBITS | ";
			}
			if (lpwp->flags & SWP_NOOWNERZORDER)
			{
				ssMessage << L"SWP_NOOWNERZORDER | ";
			}
			if (lpwp->flags & SWP_NOSENDCHANGING)
			{
				ssMessage << L"SWP_NOSENDCHANGING | ";
			}
			ssMessage << L"\r\n";
			strMessage = ssMessage.str();
			std::size_t pos{ strMessage.find_last_of(L'|') };
			if (pos != std::wstring::npos)
			{
				strMessage.erase(strMessage.find_last_of(L'|') - 1, 3);
			}
			OutputDebugString(strMessage.c_str());
		}
		break;
	default:
		{
			ret = DefSubclassProc(hWndProgressBarFont, Message, wParam, lParam);
		}
		break;
	}

	return ret;
}
#endif // SHOWPOSINFO

INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT Message, WPARAM wParam, LPARAM lParam)
{
	INT_PTR ret{};

	static HFONT hFontDialog{};

	static std::vector<ProcessInfo> ProcessList{};

	static bool bOrderByProcessAscending{ true };
	static bool bOrderByPIDAscending{ true };

	static HMENU hMenuContextListViewProcessList{};

	switch (Message)
	{
	case WM_INITDIALOG:
		{
			bOrderByProcessAscending = true;
			bOrderByPIDAscending = true;

			NONCLIENTMETRICS ncm{ sizeof(NONCLIENTMETRICS) };
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
			hFontDialog = CreateFontIndirect(&ncm.lfMessageFont);

			HWND hWndOwner{ GetWindow(hWndDialog, GW_OWNER) };
			assert(hWndOwner);
			RECT rcOwner{}, rcDialog{};
			GetWindowRect(hWndOwner, &rcOwner);
			GetWindowRect(hWndDialog, &rcDialog);
			SetWindowPos(hWndDialog, NULL, rcOwner.left + ((rcOwner.right - rcOwner.left) - (rcDialog.right - rcDialog.left)) / 2, rcOwner.top + ((rcOwner.bottom - rcOwner.top) - (rcDialog.bottom - rcDialog.top)) / 2, 0, 0, SWP_NOZORDER | SWP_NOSIZE | SWP_NOACTIVATE);

			// Initialize ListViewProcessList
			HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };
			ListView_SetExtendedListViewStyle(hWndListViewProcessList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
			SetWindowFont(hWndListViewProcessList, hFontDialog, TRUE);

			LVCOLUMN lvc1{ LVCF_FMT | LVCF_TEXT, LVCFMT_LEFT, 0, (LPWSTR)L"Process" };
			LVCOLUMN lvc2{ LVCF_FMT | LVCF_TEXT, LVCFMT_LEFT, 0, (LPWSTR)L"PID" };
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
					lvi.pszText = const_cast<LPWSTR>(str.c_str());
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

			//Get HMENU to context menu
			hMenuContextListViewProcessList = GetSubMenu(LoadMenu(GetModuleHandle(NULL), MAKEINTRESOURCE(IDR_CONTEXTMENU1)), 2);

			ret = static_cast<INT_PTR>(TRUE);
		}
		break;
	case WM_COMMAND:
		{
			switch (LOWORD(wParam))
			{
				// Refresh process list
			case ID_DIALOGBOXMENU_REFRESH:
				{
					HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };
					HWND hWndHeaderListViewProcessList{ ListView_GetHeader(hWndListViewProcessList) };

					// Clear process list
					ListView_DeleteAllItems(hWndListViewProcessList);
					ProcessList.clear();

					// Refill process list
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
							lvi.pszText = const_cast<LPWSTR>(str.c_str());
							ListView_SetItem(hWndListViewProcessList, &lvi);
							lvi.iItem++;

							ProcessList.push_back({ NULL, pe32.szExeFile, pe32.th32ProcessID });
						} while (Process32Next(hProcessSnapshot, &pe32));
					}

					// Remove the arrow in the header
					HDITEM hdi{ HDI_FORMAT };

					Header_GetItem(hWndHeaderListViewProcessList, 0, &hdi);
					hdi.fmt = hdi.fmt & (~(HDF_SORTDOWN | HDF_SORTUP));
					Header_SetItem(hWndHeaderListViewProcessList, 0, &hdi);
					Header_GetItem(hWndHeaderListViewProcessList, 1, &hdi);
					hdi.fmt = hdi.fmt & (~(HDF_SORTDOWN | HDF_SORTUP));
					Header_SetItem(hWndHeaderListViewProcessList, 1, &hdi);

					bOrderByProcessAscending = true;
					bOrderByPIDAscending = true;

					ret = static_cast<INT_PTR>(TRUE);
				}
				break;
				// Select process
			case ID_DIALOGBOXMENU_SELECT:
				{
					// Simulate clicking "OK" button
					SendMessage(GetDlgItem(hWndDialog, IDOK), BM_CLICK, 0, 0);

					ret = static_cast<INT_PTR>(TRUE);
				}
				break;
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
						ProcessInfo* lpTargetProcessInfo = new ProcessInfo{ ProcessList[iSelected] };
						assert(lpTargetProcessInfo);

						EndDialog(hWndDialog, reinterpret_cast<INT_PTR>(lpTargetProcessInfo));
					}

					ret = static_cast<INT_PTR>(TRUE);
				}
				break;
				// Return null
			case IDCANCEL:
				{
					EndDialog(hWndDialog, NULL);

					ret = static_cast<INT_PTR>(TRUE);
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
					switch (reinterpret_cast<LPNMHDR>(lParam)->code)
					{
						// Sort items
					case LVN_COLUMNCLICK:
						{
							HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };
							HWND hWndHeaderListViewProcessList{ ListView_GetHeader(hWndListViewProcessList) };

							HDITEM hdi{ HDI_FORMAT };

							switch (reinterpret_cast<LPNMLISTVIEW>(lParam)->iSubItem)
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
								lvi.pszText = const_cast<LPWSTR>(i.strProcessName.c_str());
								ListView_SetItem(hWndListViewProcessList, &lvi);
								lvi.iSubItem = 1;
								std::wstring strPID{ std::to_wstring(i.dwProcessID) };
								lvi.pszText = const_cast<LPWSTR>(strPID.c_str());
								ListView_SetItem(hWndListViewProcessList, &lvi);
								lvi.iItem++;
							}

							ret = static_cast<INT_PTR>(TRUE);
						}
						break;
						// Select double-clicked item
					case NM_DBLCLK:
						{
							// Simulate clicking "OK" button
							SendMessage(GetDlgItem(hWndDialog, IDOK), BM_CLICK, 0, 0);

							ret = static_cast<INT_PTR>(TRUE);
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
			// Show context menu in ListViewProcessList
			HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };

			if (reinterpret_cast<HWND>(wParam) == hWndListViewProcessList)
			{
				POINT ptAnchorContextMenuListViewProcessList{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
				if (ptAnchorContextMenuListViewProcessList.x == -1 && ptAnchorContextMenuListViewProcessList.y == -1)
				{
					int iSelectionMark{ ListView_GetSelectionMark(hWndListViewProcessList) };
					if (iSelectionMark == -1)
					{
						RECT rcListViewFontListClient{};
						GetClientRect(hWndListViewProcessList, &rcListViewFontListClient);
						MapWindowRect(hWndListViewProcessList, HWND_DESKTOP, &rcListViewFontListClient);
						ptAnchorContextMenuListViewProcessList = { rcListViewFontListClient.left, rcListViewFontListClient.top };
					}
					else
					{
						ListView_EnsureVisible(hWndListViewProcessList, iSelectionMark, FALSE);
						ListView_GetItemPosition(hWndListViewProcessList, iSelectionMark, &ptAnchorContextMenuListViewProcessList);
						ClientToScreen(hWndListViewProcessList, &ptAnchorContextMenuListViewProcessList);
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
				BOOL bRetTrackPopupMenu{ TrackPopupMenu(hMenuContextListViewProcessList, uFlags | TPM_TOPALIGN | TPM_RIGHTBUTTON, ptAnchorContextMenuListViewProcessList.x, ptAnchorContextMenuListViewProcessList.y, 0, hWndDialog, NULL) };
				assert(bRetTrackPopupMenu);

				ret = static_cast<INT_PTR>(TRUE);
			}
			else
			{
				ret = static_cast<INT_PTR>(FALSE);
			}
		}
		break;
	case WM_CTLCOLORDLG:
		{
			if (reinterpret_cast<HWND>(lParam) == hWndDialog)
			{
				ret = reinterpret_cast<INT_PTR>(GetSysColorBrush(COLOR_WINDOW));
			}
		}
		break;
	case WM_DESTROY:
		{
			DeleteFont(hFontDialog);

			ret = static_cast<INT_PTR>(FALSE);
		}
		break;
	default:
		break;
	}

	return ret;
}

int MessageBoxCentered(HWND hWnd, LPCTSTR lpText, LPCTSTR lpCaption, UINT uType)
{
	// Center message box at its owner window
	static HHOOK hHookCBT{};
	hHookCBT = SetWindowsHookEx(WH_CBT,
		[](int nCode, WPARAM wParam, LPARAM lParam) -> LRESULT
		{
			switch (nCode)
			{
			case HCBT_CREATEWND:
				{
					if (reinterpret_cast<LPCBT_CREATEWND>(lParam)->lpcs->lpszClass == reinterpret_cast<LPWSTR>(static_cast<ATOM>(32770)))	// #32770 = dialog box class
					{
						if (reinterpret_cast<LPCBT_CREATEWND>(lParam)->lpcs->hwndParent)
						{
							RECT rcParent{};
							GetWindowRect(reinterpret_cast<LPCBT_CREATEWND>(lParam)->lpcs->hwndParent, &rcParent);
							reinterpret_cast<LPCBT_CREATEWND>(lParam)->lpcs->x = rcParent.left + ((rcParent.right - rcParent.left) - reinterpret_cast<LPCBT_CREATEWND>(lParam)->lpcs->cx) / 2;
							reinterpret_cast<LPCBT_CREATEWND>(lParam)->lpcs->y = rcParent.top + ((rcParent.bottom - rcParent.top) - reinterpret_cast<LPCBT_CREATEWND>(lParam)->lpcs->cy) / 2;
						}
					}
				}
				break;
			default:
				break;
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

	// Make full path to module
	WCHAR szDllPath[MAX_PATH]{};
	GetModuleFileName(NULL, szDllPath, MAX_PATH);
	PathRemoveFileSpec(szDllPath);
	PathAppend(szDllPath, szModuleName);

	// Call LoadLibraryW with module full path to inject dll into hProcess
	DWORD dwRemoteThreadExitCode{};
	if (CallRemoteProc(hProcess, GetProcAddress(GetModuleHandle(L"Kernel32"), "LoadLibraryW"), szDllPath, (std::wcslen(szDllPath) + 1) * sizeof(WCHAR), dwTimeout, &dwRemoteThreadExitCode))
	{
		if (dwRemoteThreadExitCode)
		{
			bRet = true;
		}
		else
		{
			bRet = false;
		}
	}
	else
	{
		bRet = false;
	}

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
		DWORD dwRemoteThreadExitCode{};
		if (CallRemoteProc(hProcess, GetProcAddress(GetModuleHandle(L"Kernel32"), "FreeLibrary"), hModInjectionDll, 0, dwTimeout, &dwRemoteThreadExitCode))
		{
			if (dwRemoteThreadExitCode)
			{
				bRet = true;
			}
			else
			{
				bRet = false;
			}
		}
		else
		{
			bRet = false;
		}
	} while (false);

	return bRet;
}

bool CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, std::size_t cbParamSize, DWORD dwTimeout, LPDWORD lpdwRemoteThreadExitCode)
{
	bool bRet{};

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
				bRet = false;

				break;
			}

			// Write parameter to remote buffer
			if (!WriteProcessMemory(hProcess, lpRemoteBuffer, lpParameter, cbParamSize, NULL))
			{
				VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

				bRet = false;

				break;
			}
		}

		// Create remote thread to call function
		HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, static_cast<LPTHREAD_START_ROUTINE>(lpRemoteProcAddr), lpRemoteBuffer, 0, NULL) };
		if (!hRemoteThread)
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;

			break;
		}

		// Wait for remote thread to terminate with timeout
		DWORD dwWaitResult{ WaitForSingleObject(hRemoteThread, dwTimeout) };
		if (dwWaitResult == WAIT_OBJECT_0)
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);
		}
		else
		{
			CloseHandle(hRemoteThread);
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;

			break;
		}

		// Get exit code of remote thread
		DWORD dwRemoteThreadExitCode{};
		if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
		{
			CloseHandle(hRemoteThread);

			bRet = false;

			break;
		}
		CloseHandle(hRemoteThread);

		bRet = true;

		if (lpdwRemoteThreadExitCode)
		{
			*lpdwRemoteThreadExitCode = dwRemoteThreadExitCode;
		}
	} while (false);

	return bRet;
}

std::wstring GetUniqueName(LPCWSTR lpszString, Scope scope)
{
	assert((scope == Scope::Machine) || (scope == Scope::User) || (scope == Scope::Session) || (scope == Scope::WindowStation) || (scope == Scope::Desktop));

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
			ssRet << reinterpret_cast<LPWSTR>(lpBuffer2.get());

			if (scope == Scope::WindowStation)
			{
				CloseHandle(hTokenProcess);

				break;
			}

			// On the same desktop
			ssRet << L"--";

			DWORD dwLength3{};
			HDESK hDeskProcess{ GetThreadDesktop(GetCurrentThreadId()) };
			GetUserObjectInformation(hDeskProcess, UOI_NAME, NULL, 0, &dwLength3);
			std::unique_ptr<BYTE[]> lpBuffer3{ new BYTE[dwLength3] };
			GetUserObjectInformation(hDeskProcess, UOI_NAME, lpBuffer3.get(), dwLength3, &dwLength3);
			ssRet << reinterpret_cast<LPWSTR>(lpBuffer3.get());

			if (scope == Scope::Desktop)
			{
				CloseHandle(hTokenProcess);

				break;
			}

			CloseHandle(hTokenProcess);
		} while (false);
	}

	return ssRet.str();
}