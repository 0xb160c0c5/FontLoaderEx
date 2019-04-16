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

LRESULT CALLBACK WndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);

const WCHAR szWindowCaption[]{ L"FontLoaderEx" };

std::list<FontResource> FontList{};

HWND hWndMain{};
HMENU hMenuContextListViewFontList{};

bool bDragDropHasFonts{ false };

// Create an unique string by scope
enum class Scope { Machine, User, Session, WindowStation, Desktop };
std::wstring GetUniqueName(LPCWSTR lpszString, Scope scope);

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
	// Check Windows version
	if (!IsWindows7OrGreater())
	{
		MessageBox(NULL, L"Windows 7 or higher required.", szWindowCaption, MB_ICONWARNING);
		return 0;
	}

	// Prevent multiple instances of FontLoaderEx in the same session
	Scope scope{ Scope::Session };
	std::wstring strMutexName{ GetUniqueName(L"FontLoaderEx-656A8394-5AB8-4061-8882-2FE2E7940C2E", scope) };
	HANDLE hMutexOneInstance{ CreateMutex(NULL, FALSE, strMutexName.c_str()) };
	if (!hMutexOneInstance)
	{
		MessageBox(NULL, L"Failed to create singleton mutex.", szWindowCaption, MB_ICONERROR);

		return -1;
	}
	else
	{
		if (GetLastError() == ERROR_ALREADY_EXISTS || GetLastError() == ERROR_ACCESS_DENIED)
		{
			std::wstringstream strMessage{};
			strMessage << L"An instance of " << szWindowCaption << L" is already running ";
			switch (scope)
			{
			case Scope::Machine:
				{
					strMessage << L"on the same machine.";
				}
				break;
			case Scope::User:
				{
					strMessage << L"by the same user.";
				}
				break;
			case Scope::Session:
				{
					strMessage << L"in the same session.";
				}
				break;
			case Scope::WindowStation:
				{
					strMessage << L"in the same window station.";
				}
				break;
			case Scope::Desktop:
				{
					strMessage << L"on the same desktop.";
				}
				break;
			default:
				break;
			}
			MessageBox(NULL, strMessage.str().c_str(), szWindowCaption, MB_ICONWARNING);

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

	// Create window
	WNDCLASS wc{ CS_HREDRAW | CS_VREDRAW, WndProc, 0, 0, hInstance, LoadIcon(NULL, IDI_APPLICATION), LoadCursor(NULL, IDC_ARROW), GetSysColorBrush(COLOR_WINDOW), NULL, szWindowCaption };
	if (!RegisterClass(&wc))
	{
		return -1;
	}
	if (!(hWndMain = CreateWindow(szWindowCaption, szWindowCaption, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 700, 700, NULL, NULL, hInstance, NULL)))
	{
		return -1;
	}
	ShowWindow(hWndMain, nShowCmd);
	UpdateWindow(hWndMain);

	// Get HMENU to context menu of ListViewFontList
	hMenuContextListViewFontList = GetSubMenu(LoadMenu(hInstance, MAKEINTRESOURCE(IDR_CONTEXTMENU1)), 0);

	MSG Msg{};
	BOOL bRet{};
	while ((bRet = GetMessage(&Msg, NULL, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			return (int)GetLastError();
		}
		else
		{
			if (!IsDialogMessage(hWndMain, &Msg))
			{
				TranslateMessage(&Msg);
				DispatchMessage(&Msg);
			}
		}
	}

	CloseHandle(hMutexOneInstance);

	return (int)Msg.wParam;
}

LRESULT CALLBACK ListViewFontListSubclassProc(HWND hWndListViewFontList, UINT Msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
LRESULT CALLBACK EditMessageSubclassProc(HWND hWndEditMessage, UINT Msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData);
INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT Msg, WPARAM wParam, LPARAM IParam);

void* pfnRemoteAddFontProc{};
void* pfnRemoteRemoveFontProc{};

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

DWORD dwTimeout{};

int MessageBoxCentered(HWND hWnd, LPCTSTR lpText, LPCTSTR lpCaption, UINT uType);

bool EnableDebugPrivilege();
bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout);
bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout);

#ifdef _WIN64
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll64.dll" };
const WCHAR szInjectionDllNameByProxy[]{ L"FontLoaderExInjectionDll.dll" };
#else
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll.dll" };
const WCHAR szInjectionDllNameByProxy[]{ L"FontLoaderExInjectionDll64.dll" };
#endif // _WIN64

LRESULT CALLBACK WndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	switch ((USERMESSAGE)Msg)
	{
	case USERMESSAGE::CLOSEWORKERTHREADTERMINATED:
		{
			HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

			// If unloading is interrupted, just re-enable and re-disable controls
			if (lParam)
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::EditTimeout), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
				EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
			}
			// If unloading is not interrupted
			else
			{
				// If unloading successful, do cleanup
				if (wParam)
				{
					std::wstringstream Message{};
					int cchMessageLength{};

					// If loaded via proxy
					if (ProxyProcessInfo.hProcess)
					{
						// Terminate watch thread
						SetEvent(hEventTerminateWatchThread);
						WaitForSingleObject(hThreadWatch, INFINITE);
						CloseHandle(hEventTerminateWatchThread);
						CloseHandle(hThreadWatch);

						// Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::PULLDLL, 0, NULL };
						FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
						WaitForSingleObject(hEventProxyDllPullingFinished, INFINITE);
						CloseHandle(hEventProxyDllPullingFinished);
						switch (ProxyDllPullingResult)
						{
						case PROXYDLLPULL::SUCCESSFUL:
							goto continue_DAA249E0;
						case PROXYDLLPULL::FAILED:
							{
								Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
								cchMessageLength = Edit_GetTextLength(hWndEditMessage);
								Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
								Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

								EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
								EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
								EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
								EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
								EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
								EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
								if (!TargetProcessInfo.hProcess)
								{
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
								}
							}
							goto break_DAA249E0;
						default:
							break;
						}
					break_DAA249E0:
						break;
					continue_DAA249E0:

						// Terminate proxy process
						COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
						FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
						WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
						CloseHandle(ProxyProcessInfo.hProcess);
						ProxyProcessInfo.hProcess = NULL;

						// Terminate message thread
						PostMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
						WaitForSingleObject(hThreadMessage, INFINITE);
						DWORD dwMessageThreadExitCode{};
						GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
						if (dwMessageThreadExitCode)
						{
							std::wstringstream Message{};
							Message << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
							MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
						}
						CloseHandle(hThreadMessage);

						// Close handle to target process and duplicated handles
						CloseHandle(TargetProcessInfo.hProcess);
						TargetProcessInfo.hProcess = NULL;
						CloseHandle(hProcessCurrentDuplicated);
						CloseHandle(hProcessTargetDuplicated);
					}

					// Else DIY
					if (TargetProcessInfo.hProcess)
					{
						// Terminate watch thread
						SetEvent(hEventTerminateWatchThread);
						WaitForSingleObject(hThreadWatch, INFINITE);
						CloseHandle(hEventTerminateWatchThread);
						CloseHandle(hThreadWatch);

						// Unload FontLoaderExInjectionDll(64).dll from target process
						if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
						{
							Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
							cchMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
							EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
							if (!TargetProcessInfo.hProcess)
							{
								EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
							}
							break;
						}

						// Close handle to target process
						CloseHandle(TargetProcessInfo.hProcess);
						TargetProcessInfo.hProcess = NULL;
					}

					DestroyWindow(hWnd);
				}
				else
				{
					// Else, prompt user whether inisit to exit
					switch (MessageBoxCentered(hWnd, L"Some fonts are not successfully unloaded.\r\n\r\nDo you still want to exit?", szWindowCaption, MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON1 | MB_APPLMODAL))
					{
					case IDYES:
						{
							std::wstringstream Message{};
							int cchMessageLength{};

							// If loaded via proxy
							if (ProxyProcessInfo.hProcess)
							{
								// Terminate watch thread
								SetEvent(hEventTerminateWatchThread);
								WaitForSingleObject(hThreadWatch, INFINITE);
								CloseHandle(hEventTerminateWatchThread);
								CloseHandle(hThreadWatch);

								// Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
								COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::PULLDLL, 0, NULL };
								FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
								WaitForSingleObject(hEventProxyDllPullingFinished, INFINITE);
								CloseHandle(hEventProxyDllPullingFinished);
								switch (ProxyDllPullingResult)
								{
								case PROXYDLLPULL::SUCCESSFUL:
									goto continue_C82EA5C2;
								case PROXYDLLPULL::FAILED:
									{
										Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

										EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
										EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
										EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
										EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
										EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
										EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
										if (!TargetProcessInfo.hProcess)
										{
											EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
										}
									}
									goto break_C82EA5C2;
								default:
									break;
								}
							break_C82EA5C2:
								break;
							continue_C82EA5C2:

								// Terminate proxy process
								COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
								FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
								WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
								CloseHandle(ProxyProcessInfo.hProcess);
								ProxyProcessInfo.hProcess = NULL;

								// Terminate message thread
								PostMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
								WaitForSingleObject(hThreadMessage, INFINITE);
								DWORD dwMessageThreadExitCode{};
								GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
								if (dwMessageThreadExitCode)
								{
									std::wstringstream Message{};
									Message << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
									MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
								}
								CloseHandle(hThreadMessage);

								// Close handle to target process and duplicated handles
								CloseHandle(TargetProcessInfo.hProcess);
								TargetProcessInfo.hProcess = NULL;
								CloseHandle(hProcessCurrentDuplicated);
								CloseHandle(hProcessTargetDuplicated);
							}

							// Else DIY
							if (TargetProcessInfo.hProcess)
							{
								// Terminate watch thread
								SetEvent(hEventTerminateWatchThread);
								WaitForSingleObject(hThreadWatch, INFINITE);
								CloseHandle(hEventTerminateWatchThread);
								CloseHandle(hThreadWatch);

								// Unload FontLoaderExInjectionDll(64).dll from target process
								if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
								{
									Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
									EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
									if (!TargetProcessInfo.hProcess)
									{
										EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
									}
									break;
								}

								// Close handle to target process
								CloseHandle(TargetProcessInfo.hProcess);
								TargetProcessInfo.hProcess = NULL;
							}

							DestroyWindow(hWnd);
						}
						break;
					case IDNO:
						{
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
							EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
							if (!TargetProcessInfo.hProcess)
							{
								EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
							}
						}
						break;
					default:
						break;
					}
				}
			}
		}
		break;
	case USERMESSAGE::BUTTONCLOSEWORKERTHREADTERMINATED:
		{
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
			EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
			if (FontList.empty())
			{
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
			}
			else
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
			}
			bool bIsSomeFontsLoaded{};
			for (auto& i : FontList)
			{
				if (i.IsLoaded())
				{
					bIsSomeFontsLoaded = true;
					break;
				}
			}
			if (!bIsSomeFontsLoaded)
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), TRUE);
			}
		}
		break;
	case USERMESSAGE::DRAGDROPWORKERTHREADTERMINATED:
	case USERMESSAGE::BUTTONLOADWORKERTHREADTERMINATED:
	case USERMESSAGE::BUTTONUNLOADWORKERTHREADTERMINATED:
		{
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
			EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
			if (!TargetProcessInfo.hProcess)
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
			}
			bool bIsSomeFontsLoaded{};
			for (auto& i : FontList)
			{
				if (i.IsLoaded())
				{
					bIsSomeFontsLoaded = true;
					break;
				}
			}
			if (!bIsSomeFontsLoaded)
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), TRUE);
			}
		}
		break;
	case USERMESSAGE::WATCHTHREADTERMINATED:
		{
			CloseHandle(hThreadWatch);

			EnableWindow(GetDlgItem(hWndMain, (int)ID::EditTimeout), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
		}
		break;
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
	default:
		break;
	}

	static HFONT hFontMain{};

	static LONG EditMessageTextMarginY{};

	static UINT_PTR SizingEdge{};
	static RECT rcMainClientOld{};
	static int PreviousShowCmd{};

	switch (Msg)
	{
	case WM_CREATE:
		{
			/*
			┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
			┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃
			┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸─┬─┨
			┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │ ┃
			┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘ ┃
			┃        │        │        │        │     Select Process     │                    ┃
			┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────┨
			┃ Font Name                                                        │ State        ┃
			┠──────────────────────────────────────────────────────────────────┼──────────────┨
			┃                                                                  ┆              ┃
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
			┃                                                                  ┆              ┃
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
			┃                                                                  ┆              ┃
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
			┃                                                                  ┆              ┃
			┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
			┠──────────────────────────────────────────────────────────────────┴──────────────┨
			┠─────────────────────────────────────────────────────────────────────────────┬───┨
			┃ Temporarily load fonts to Windows or specific process                       │ ↑ ┃
			┃                                                                             ├───┨
			┃ How to load fonts to Windows:                                               │▓▓▓┃
			┃ 1.Drag-drop font files onto the icon of this application.                   │▓▓▓┃
			┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │▓▓▓┃
			┃  view, then click "Load" button.                                            │▓▓▓┃
			┃                                                                             ├───┨
			┃ How to unload fonts from Windows:                                           │   ┃
			┃ Select all fonts then click "Unload" or "Close" button or the X at the      │   ┃
			┃ upper-right cornor.                                                         │   ┃
			┃                                                                             │   ┃
			┃ How to load fonts to process:                                               │   ┃
			┃ 1.Click "Click to select process", select a process.                        │   ┃
			┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list ├───┨
			┃ view, then click "Load" button.                                             │ ↓ ┃
			┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛
			*/

			// Get initial window state
			WINDOWPLACEMENT wp{ sizeof(WINDOWPLACEMENT) };
			GetWindowPlacement(hWnd, &wp);
			PreviousShowCmd = wp.showCmd;

			RECT rcClientMain{};
			GetClientRect(hWnd, &rcClientMain);

			NONCLIENTMETRICS ncm{ sizeof(NONCLIENTMETRICS) };
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
			hFontMain = CreateFontIndirect(&ncm.lfMessageFont);

			// Initialize ButtonOpen
			HWND hWndButtonOpen{ CreateWindow(WC_BUTTON, L"&Open", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 50, 50, hWnd, (HMENU)ID::ButtonOpen, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonOpen, hFontMain, TRUE);

			// Initialize ButtonClose
			HWND hWndButtonClose{ CreateWindow(WC_BUTTON, L"&Close", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, 50, 0, 50, 50, hWnd, (HMENU)ID::ButtonClose, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonClose, hFontMain, TRUE);

			// Initialize ButtonLoad
			HWND hWndButtonLoad{ CreateWindow(WC_BUTTON, L"&Load", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, 100, 0, 50, 50, hWnd, (HMENU)ID::ButtonLoad, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonLoad, hFontMain, TRUE);

			// Initialize ButtonUnload
			HWND hWndButtonUnload{ CreateWindow(WC_BUTTON, L"&Unload", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, 150, 0, 50, 50, hWnd, (HMENU)ID::ButtonUnload, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonUnload, hFontMain, TRUE);

			// Initialize ButtonBroadcastWM_FONTCHANGE
			HWND hWndButtonBroadcastWM_FONTCHANGE{ CreateWindow(WC_BUTTON, L"&Broadcast WM_FONTCHANGE", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 200, 0, 250, 21, hWnd, (HMENU)ID::ButtonBroadcastWM_FONTCHANGE, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonBroadcastWM_FONTCHANGE, hFontMain, TRUE);

			// Initialize EditTimeout and its label
			HWND hWndStaticTimeout{ CreateWindow(WC_STATIC, L"&Timeout:", WS_CHILD | WS_VISIBLE | SS_LEFT , 470, 1, 50, 19, hWnd, (HMENU)ID::StaticTimeout, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			HWND hWndEditTimeout{ CreateWindow(WC_EDIT, L"5000", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_LEFT | ES_NUMBER | ES_AUTOHSCROLL | ES_NOHIDESEL, 520, 0, 80, 21, hWnd, (HMENU)ID::EditTimeout, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndStaticTimeout, hFontMain, TRUE);
			SetWindowFont(hWndEditTimeout, hFontMain, TRUE);

			Edit_LimitText(hWndEditTimeout, 10);

			dwTimeout = 5000;

			// Initialize ButtonSelectProcess
			HWND hWndButtonSelectProcess{ CreateWindow(WC_BUTTON, L"&Select process", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 200, 29, 250, 21, hWnd, (HMENU)ID::ButtonSelectProcess, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonSelectProcess, hFontMain, TRUE);

			// Initialize ListViewFontList
			HWND hWndListViewFontList{ CreateWindow(WC_LISTVIEW, L"FontList", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_NOSORTHEADER, 0, 50, rcClientMain.right - rcClientMain.left, 300, hWnd, (HMENU)ID::ListViewFontList, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
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
			ChangeWindowMessageFilterEx(hWndListViewFontList, 0x0049, MSGFLT_ALLOW, &cfs);	// WM_COPYGLOBALDATA

			// Initialize Splitter
			HWND hWndSplitter{ CreateWindow(UC_SPLITTER, NULL, WS_CHILD | WS_VISIBLE, 0, 350, rcClientMain.right - rcClientMain.left, 5, hWnd, (HMENU)ID::Splitter, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };

			// Initialize EditMessage
			HWND hWndEditMessage{ CreateWindow(WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | ES_READONLY | ES_LEFT | ES_MULTILINE | ES_NOHIDESEL, 0, 355, rcClientMain.right - rcClientMain.left, rcClientMain.bottom - rcClientMain.top - 355, hWnd, (HMENU)ID::EditMessage, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndEditMessage, hFontMain, TRUE);
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
				R"("Timeout": The time in milliseconds FontLoaderEx waits before reporting failure while injecting dll into target process via proxy process, the default value is 5000. Type 0 or clear to wait infinitely.)""\r\n"
				R"("Font Name": Names of the fonts added to the list view.)""\r\n"
				R"("State": State of the font. There are five states, "Not loaded", "Loaded", "Load failed", "Unloaded" and "Unload failed".)""\r\n"
				"\r\n"
			);

			SetWindowSubclass(hWndEditMessage, EditMessageSubclassProc, 0, NULL);

			// Get vertical margin of the formatting rcangle in EditMessage
			RECT rcEditMessageClient{}, rcEditMessageFormatting{};
			GetClientRect(hWndEditMessage, &rcEditMessageClient);
			Edit_GetRect(hWndEditMessage, &rcEditMessageFormatting);
			EditMessageTextMarginY = (rcEditMessageClient.bottom - rcEditMessageClient.top) - (rcEditMessageFormatting.bottom - rcEditMessageFormatting.top);
		}
		break;
	case WM_ACTIVATE:
		{
			SetFocus(GetDlgItem(hWnd, (int)ID::ButtonOpen));

			// Process drag-drop font files onto the application icon stage II
			if (bDragDropHasFonts)
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
				EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

				_beginthread(DragDropWorkerThreadProc, 0, nullptr);

				bDragDropHasFonts = false;
			}
		}
		break;
	case WM_CLOSE:
		{
			// Unload all fonts
			if (!FontList.empty())
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
				EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

				_beginthread(CloseWorkerThreadProc, 0, nullptr);
			}
			else
			{
				HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

				std::wstringstream Message{};
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
							Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
							cchMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
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
					PostMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
					WaitForSingleObject(hThreadMessage, INFINITE);
					DWORD dwMessageThreadExitCode{};
					GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
					if (dwMessageThreadExitCode)
					{
						std::wstringstream Message{};
						Message << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
						MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
					}
					CloseHandle(hThreadMessage);

					// Terminate proxy process
					COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
					FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
					WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
					CloseHandle(ProxyProcessInfo.hProcess);
					ProxyProcessInfo.hProcess = NULL;

					// Close handle to target process and duplicated handles
					CloseHandle(TargetProcessInfo.hProcess);
					TargetProcessInfo.hProcess = NULL;
					CloseHandle(hProcessCurrentDuplicated);
					CloseHandle(hProcessTargetDuplicated);
				}
				// Else DIY
				if (TargetProcessInfo.hProcess)
				{
					// Unload FontLoaderExInjectionDll(64).dll from target process
					if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
					{
						Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
						cchMessageLength = Edit_GetTextLength(hWndEditMessage);
						Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
						Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
						break;
					}

					// Terminate watch thread
					SetEvent(hEventTerminateWatchThread);
					WaitForSingleObject(hThreadWatch, INFINITE);
					CloseHandle(hEventTerminateWatchThread);
					CloseHandle(hThreadWatch);

					// Close handle to target process
					CloseHandle(TargetProcessInfo.hProcess);
					TargetProcessInfo.hProcess = NULL;
				}

				ret = DefWindowProc(hWnd, Msg, wParam, lParam);
			}
		}
		break;
	case WM_DESTROY:
		{
			DeleteFont(hFontMain);

			PostQuitMessage(0);
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

							HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
							HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

							const std::size_t cchBuffer{ 1024 };
							std::unique_ptr<WCHAR[]> lpszOpenFileNames{ new WCHAR[cchBuffer]{} };
							OPENFILENAME ofn{ sizeof(ofn), hWnd, NULL, L"Font Files(*.ttf;*.ttc;*.otf)\0*.ttf;*.ttc;*.otf\0", NULL, 0, 0, lpszOpenFileNames.get(), cchBuffer * sizeof(WCHAR), NULL, 0, NULL, L"Select fonts", OFN_EXPLORER | OFN_ENABLESIZING | OFN_HIDEREADONLY | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT | OFN_ENABLEHOOK, 0, 0, NULL, (LPARAM)&ofn,
								[](HWND hWndOpenDialogChild, UINT Msg, WPARAM wParam, LPARAM lParam) -> UINT_PTR
								{
									UINT_PTR ret{};

									static HFONT hFontOpenDialog{};

									static LPOPENFILENAME lpofn{};

									switch (Msg)
									{
									case WM_INITDIALOG:
										{
											// Get the pointer to original ofn
											lpofn = (LPOPENFILENAME)lParam;

											// Change default font
											NONCLIENTMETRICS ncm{ sizeof(NONCLIENTMETRICS) };
											SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
											hFontOpenDialog = CreateFontIndirect(&ncm.lfMessageFont);

											HWND hWndDialogOpen{ GetParent(hWndOpenDialogChild) };
											SetWindowFont(GetDlgItem(hWndDialogOpen, chx1), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, cmb1), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, stc2), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, cmb2), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, stc4), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, cmb13), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, edt1), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, stc3), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, lst1), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, stc1), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, IDOK), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, IDCANCEL), hFontOpenDialog, TRUE);
											SetWindowFont(GetDlgItem(hWndDialogOpen, pshHelp), hFontOpenDialog, TRUE);

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
													static HWND hWndDialogOpen{ GetParent(hWndOpenDialogChild) };

													DWORD cchLength{ (DWORD)CommDlg_OpenSave_GetSpec(hWndDialogOpen, NULL, 0) + MAX_PATH + 2 };
													if (lpofn->nMaxFile < cchLength)
													{
														delete[] lpofn->lpstrFile;
														lpofn->lpstrFile = new WCHAR[cchLength]{};
														lpofn->nMaxFile = cchLength;
													}

													ret = (UINT_PTR)TRUE;
												}
												break;
											default:
												break;
											}
										}
										break;
									case WM_DESTROY:
										{
											DeleteFont(hFontOpenDialog);

											ret = (UINT_PTR)FALSE;
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
								LVITEM lvi{ LVIF_TEXT, ListView_GetItemCount(hWndListViewFontList) };
								std::wstringstream Message{};
								int iMessageLenth{};
								if (PathIsDirectory(ofn.lpstrFile))
								{
									WCHAR* lpszFileName{ ofn.lpstrFile + ofn.nFileOffset };
									do
									{
										WCHAR lpszPath[MAX_PATH]{};
										PathCombine(lpszPath, ofn.lpstrFile, lpszFileName);
										lpszFileName += std::wcslen(lpszFileName) + 1;
										FontList.push_back(lpszPath);

										lvi.iSubItem = 0;
										lvi.pszText = (LPWSTR)lpszPath;
										ListView_InsertItem(hWndListViewFontList, &lvi);
										lvi.iSubItem = 1;
										lvi.pszText = (LPWSTR)L"Not loaded";
										ListView_SetItem(hWndListViewFontList, &lvi);
										ListView_SetItemState(hWndListViewFontList, lvi.iItem, LVIS_SELECTED, LVIS_SELECTED);
										lvi.iItem++;

										Message << lpszPath << L" opened\r\n";
										iMessageLenth = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLenth, iMessageLenth);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										Message.str(L"");
									} while (*lpszFileName);
								}
								else
								{
									FontList.push_back(ofn.lpstrFile);

									lvi.iSubItem = 0;
									lvi.pszText = (LPWSTR)ofn.lpstrFile;
									ListView_InsertItem(hWndListViewFontList, &lvi);
									lvi.iSubItem = 1;
									lvi.pszText = (LPWSTR)L"Not loaded";
									ListView_SetItem(hWndListViewFontList, &lvi);
									ListView_SetItemState(hWndListViewFontList, lvi.iItem, LVIS_SELECTED, LVIS_SELECTED);

									Message << ofn.lpstrFile << L" opened\r\n";
									iMessageLenth = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, iMessageLenth, iMessageLenth);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
								}
								iMessageLenth = Edit_GetTextLength(hWndEditMessage);
								Edit_SetSel(hWndEditMessage, iMessageLenth, iMessageLenth);
								Edit_ReplaceSel(hWndEditMessage, L"\r\n");

								if (bIsFontListEmptyBefore)
								{
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
									EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_ENABLED);
									EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_ENABLED);
									EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_ENABLED);
									EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_ENABLED);
								}
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
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
							EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							_beginthread(ButtonCloseWorkerThreadProc, 0, nullptr);
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
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
							EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							_beginthread(ButtonLoadWorkerThreadProc, 0, nullptr);
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
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
							EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
							EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

							_beginthread(ButtonUnloadWorkerThreadProc, 0, nullptr);
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

							if (Edit_GetTextLength(hWndEditTimeout))
							{
								BOOL bIsConverted{};
								DWORD dwTimeoutTemp{ (DWORD)GetDlgItemInt(hWnd, (int)ID::EditTimeout, &bIsConverted, FALSE) };
								if (!bIsConverted)
								{
									DWORD dwCaretIndex{ Edit_GetCaretIndex(hWndEditTimeout) };
									SetDlgItemInt(hWnd, (int)ID::EditTimeout, (UINT)dwTimeout, FALSE);
									Edit_SetCaretIndex(hWndEditTimeout, dwCaretIndex);

									EDITBALLOONTIP ebt{ sizeof(EDITBALLOONTIP), L"Out of range", L"Timeout out of range.", TTI_ERROR };
									SendMessage(hWndEditTimeout, EM_SHOWBALLOONTIP, NULL, (LPARAM)&ebt);
									MessageBeep(MB_ICONWARNING);
								}
								else
								{
									if (dwTimeoutTemp == 0)
									{
										dwTimeout = INFINITE;
									}
									else
									{
										dwTimeout = dwTimeoutTemp;
									}
								}
							}
							else
							{
								dwTimeout = INFINITE;
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
							HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
							HWND hWndButtonSelectProcess{ GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess) };

							static bool bIsSeDebugPrivilegeEnabled{ false };

							static HANDLE hCurrentProcessDuplicated{};
							static HANDLE hTargetProcessDuplicated{};

							std::wstringstream Message{};
							int cchMessageLength{};

							// Enable SeDebugPrivilege
							if (!bIsSeDebugPrivilegeEnabled)
							{
								if (!EnableDebugPrivilege())
								{
									Message << L"Failed to enable " << SE_DEBUG_NAME << L" for " << szWindowCaption << L".";
									MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
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
											Message << szInjectionDllNameByProxy << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
											Message.str(L"");
										}
										goto continue_B9A25A68;
									case PROXYDLLPULL::FAILED:
										{
											Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
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
									PostMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
									WaitForSingleObject(hThreadMessage, INFINITE);
									DWORD dwMessageThreadExitCode{};
									GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
									if (dwMessageThreadExitCode)
									{
										std::wstringstream Message{};
										Message << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
										MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
									}
									CloseHandle(hThreadMessage);

									// Terminate proxy process
									COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
									WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
									Message << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") successfully terminated.\r\n\r\n";
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									CloseHandle(ProxyProcessInfo.hProcess);
									ProxyProcessInfo.hProcess = NULL;

									// Close handles to target process, duplicated handles and synchronization objects
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
									CloseHandle(hCurrentProcessDuplicated);
									CloseHandle(hTargetProcessDuplicated);
									CloseHandle(hEventProxyAddFontFinished);
									CloseHandle(hEventProxyRemoveFontFinished);
								}

								// Else DIY
								if (TargetProcessInfo.hProcess)
								{
									// Unload FontLoaderExInjectionDll(64).dll from target process
									if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									else
									{
										Message << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										Message.str(L"");
									}

									// Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hThreadWatch, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hThreadWatch);

									// Close handle to target process
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
								}

								// Get handle to target process
								SelectedProcessInfo.hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, SelectedProcessInfo.dwProcessID);
								if (!SelectedProcessInfo.hProcess)
								{
									Message << L"Failed to open process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									break;
								}

								// Determine current and target process type
								BOOL bIsWOW64Current{}, bIsWOW64Target{};
								IsWow64Process(GetCurrentProcess(), &bIsWOW64Current);
								IsWow64Process(SelectedProcessInfo.hProcess, &bIsWOW64Target);

								// If process types are different, launch FontLoaderExProxy.exe to inject dll
								if (bIsWOW64Current != bIsWOW64Target)
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
											MessageBoxCentered(NULL, L"Failed to create message-only window.", szWindowCaption, MB_ICONERROR);

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

									// Launch proxy process, send handles to current process and target process, HWND to message window, handles to synchronization objects and timeout as arguments to proxy process
									const WCHAR szProxyProcessName[]{ L"FontLoaderExProxy.exe" };

									DuplicateHandle(GetCurrentProcess(), GetCurrentProcess(), GetCurrentProcess(), &hCurrentProcessDuplicated, 0, TRUE, DUPLICATE_SAME_ACCESS);
									DuplicateHandle(GetCurrentProcess(), SelectedProcessInfo.hProcess, GetCurrentProcess(), &hTargetProcessDuplicated, 0, TRUE, DUPLICATE_SAME_ACCESS);
									std::wstringstream ssParams{};
									ssParams << (UINT_PTR)hCurrentProcessDuplicated << L" " << (UINT_PTR)hTargetProcessDuplicated << L" " << (UINT_PTR)hWndMessage << L" " << (UINT_PTR)hEventMessageThreadReady << L" " << (UINT_PTR)hEventProxyProcessReady << L" " << dwTimeout;
									std::size_t cchParamLength{ ssParams.str().length() };
									std::unique_ptr<WCHAR[]> lpszParams{ new WCHAR[cchParamLength + 1]{} };
									wcsncpy_s(lpszParams.get(), cchParamLength + 1, ssParams.str().c_str(), cchParamLength);
									std::wstringstream ssProxyPath{};
									STARTUPINFO si{ sizeof(STARTUPINFO) };
									PROCESS_INFORMATION piProxyProcess{};
#ifdef _DEBUG
#ifdef _WIN64
									ssProxyPath << LR"(..\Debug\)" << szProxyProcessName;
#else
									ssProxyPath << LR"(..\x64\Debug\)" << szProxyProcessName;
#endif // _WIN64
#else
									ssProxyPath << szProxyProcessName;
#endif // _DEBUG
									if (!CreateProcess(ssProxyPath.str().c_str(), lpszParams.get(), NULL, NULL, TRUE, 0, NULL, NULL, &si, &piProxyProcess))
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hCurrentProcessDuplicated);
										CloseHandle(hTargetProcessDuplicated);

										Message << L"Failed to launch " << szProxyProcessName << L"." << L"\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

										// Terminate message thread
										PostMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
										WaitForSingleObject(hThreadMessage, INFINITE);
										DWORD dwMessageThreadExitCode{};
										GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
										if (dwMessageThreadExitCode)
										{
											std::wstringstream Message{};
											Message << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
											MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
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

									Message << szProxyProcessName << L"(" << piProxyProcess.dwProcessId << L") succesfully launched.\r\n\r\n";
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									Message.str(L"");

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
											Message << L"Failed to enable " << SE_DEBUG_NAME << L" for " << szProxyProcessName << L".";
											MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
											Message.str(L"");

											// Terminate proxy process
											COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
											FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
											WaitForSingleObject(piProxyProcess.hProcess, INFINITE);

											Message << szProxyProcessName << L"(" << piProxyProcess.dwProcessId << L") successfully terminated.\r\n\r\n";
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

											CloseHandle(piProxyProcess.hThread);
											CloseHandle(piProxyProcess.hProcess);
											piProxyProcess.hProcess = NULL;

											// Terminate message thread
											PostMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
											WaitForSingleObject(hThreadMessage, INFINITE);
											DWORD dwMessageThreadExitCode{};
											GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
											if (dwMessageThreadExitCode)
											{
												std::wstringstream Message{};
												Message << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
												MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
											}
											CloseHandle(hThreadMessage);

											// Close handle to selected process
											CloseHandle(SelectedProcessInfo.hProcess);
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
											Message << szInjectionDllNameByProxy << L" successfully injected into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

											// Register proxy AddFont() and RemoveFont() procedure and create synchronization objects
											FontResource::RegisterAddRemoveFontProc(ProxyAddFontProc, ProxyRemoveFontProc);
											hEventProxyAddFontFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
											hEventProxyRemoveFontFinished = CreateEvent(NULL, TRUE, FALSE, NULL);

											// Disable EditTimeout and ButtonBroadcastWM_FONTCHANGE
											EnableWindow(GetDlgItem(hWndMain, (int)ID::EditTimeout), FALSE);
											EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);

											// Change the caption of ButtonSelectProcess
											std::wstringstream Caption{};
											Caption << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L")";
											Button_SetText(hWndButtonSelectProcess, (LPCWSTR)Caption.str().c_str());

											// Set TargetProcessInfo and ProxyProcessInfo
											TargetProcessInfo = SelectedProcessInfo;
											ProxyProcessInfo = { piProxyProcess.hProcess, szProxyProcessName, piProxyProcess.dwProcessId };

											// Create synchronization object and start watch thread
											hEventTerminateWatchThread = CreateEvent(NULL, TRUE, FALSE, NULL);
											hThreadWatch = (HANDLE)_beginthreadex(nullptr, 0, ProxyAndTargetProcessWatchThreadProc, nullptr, 0, nullptr);
										}
										break;
									case PROXYDLLINJECTION::FAILED:
										{
											Message << L"Failed to inject " << szInjectionDllNameByProxy << L" into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::FAILEDTOENUMERATEMODULES:
										{
											Message << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::GDI32NOTLOADED:
										{
											Message << L"gdi32.dll not loaded by target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::MODULENOTFOUND:
										{
											Message << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									continue_DBEA36FE:
										{
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
											Message.str(L"");

											// Terminate proxy process
											COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
											FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
											WaitForSingleObject(piProxyProcess.hProcess, INFINITE);

											Message << szProxyProcessName << L"(" << piProxyProcess.dwProcessId << L") successfully terminated.\r\n\r\n";
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

											CloseHandle(piProxyProcess.hThread);
											CloseHandle(piProxyProcess.hProcess);
											piProxyProcess.hProcess = NULL;

											// Terminate message thread
											PostMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
											WaitForSingleObject(hThreadMessage, INFINITE);
											DWORD dwMessageThreadExitCode{};
											GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
											if (dwMessageThreadExitCode)
											{
												std::wstringstream Message{};
												Message << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
												MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
											}
											CloseHandle(hThreadMessage);

											// Close handle to selected process
											CloseHandle(SelectedProcessInfo.hProcess);
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

										Message << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
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

										Message << L"gdi32.dll not loaded by target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									CloseHandle(hModuleSnapshot);

									// Inject FontLoaderExInjectionDll(64).dll into target process
									if (!InjectModule(SelectedProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										CloseHandle(SelectedProcessInfo.hProcess);

										Message << L"Failed to inject " << szInjectionDllName << L" into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									Message << szInjectionDllName << L" successfully injected into target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									Message.str(L"");

									// Get base address of FontLoaderExInjectionDll(64).dll in target process
									HANDLE hModuleSnapshot2{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, SelectedProcessInfo.dwProcessID) };
									MODULEENTRY32 me322{ sizeof(MODULEENTRY32) };
									BYTE* pModBaseAddr{};
									if (!Module32First(hModuleSnapshot2, &me322))
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot2);

										Message << L"Failed to enumerate modules in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									do
									{
										if (!lstrcmpi(me322.szModule, szInjectionDllName))
										{
											pModBaseAddr = me322.modBaseAddr;
											break;
										}
									} while (Module32Next(hModuleSnapshot2, &me322));
									if (!pModBaseAddr)
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot2);

										Message << szInjectionDllName << " not found in target process " << SelectedProcessInfo.strProcessName << L"(" << SelectedProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									CloseHandle(hModuleSnapshot2);

									// Calculate addresses of AddFont() and RemoveFont() in target process
									HMODULE hModInjectionDll{ LoadLibrary(szInjectionDllName) };
									void* pLocalAddFontProcAddr{ GetProcAddress(hModInjectionDll, "AddFont") };
									void* pLocalRemoveFontProcAddr{ GetProcAddress(hModInjectionDll, "RemoveFont") };
									FreeLibrary(hModInjectionDll);
									UINT_PTR AddFontProcOffset{ (UINT_PTR)pLocalAddFontProcAddr - (UINT_PTR)hModInjectionDll };
									UINT_PTR RemoveFontProcOffset{ (UINT_PTR)pLocalRemoveFontProcAddr - (UINT_PTR)hModInjectionDll };
									pfnRemoteAddFontProc = (void*)(pModBaseAddr + AddFontProcOffset);
									pfnRemoteRemoveFontProc = (void*)(pModBaseAddr + RemoveFontProcOffset);

									// Register remote AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(RemoteAddFontProc, RemoteRemoveFontProc);

									// Disable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWndMain, (int)ID::EditTimeout), FALSE);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);

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
											Message << szInjectionDllNameByProxy << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
											Message.str(L"");
										}
										goto continue_0F70B465;
									case PROXYDLLPULL::FAILED:
										{
											Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
											cchMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
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
									PostMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
									WaitForSingleObject(hThreadMessage, INFINITE);
									DWORD dwMessageThreadExitCode{};
									GetExitCodeThread(hThreadMessage, &dwMessageThreadExitCode);
									if (dwMessageThreadExitCode)
									{
										std::wstringstream Message{};
										Message << L"Message thread exited abnormally with code " << dwMessageThreadExitCode << L".";
										MessageBoxCentered(NULL, Message.str().c_str(), szWindowCaption, MB_ICONERROR);
									}
									CloseHandle(hThreadMessage);

									// Terminate proxy process
									COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
									WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);

									Message << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") successfully terminated.\r\n\r\n";
									cchMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

									CloseHandle(ProxyProcessInfo.hProcess);
									ProxyProcessInfo.hProcess = NULL;

									// Close handles to target process, duplicated handles and synchronization objects
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
									CloseHandle(hCurrentProcessDuplicated);
									CloseHandle(hTargetProcessDuplicated);
									CloseHandle(hEventProxyAddFontFinished);
									CloseHandle(hEventProxyRemoveFontFinished);

									// Register global AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

									// Enable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWndMain, (int)ID::EditTimeout), TRUE);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);

									// Revert the caption of ButtonSelectProcess to default
									Button_SetText(hWndButtonSelectProcess, L"Select process");
								}

								// Else DIY
								if (TargetProcessInfo.hProcess)
								{
									// Unload FontLoaderExInjectionDll(64).dll from target process
									if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									else
									{
										Message << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L").\r\n\r\n";
										cchMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									}

									// Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hThreadWatch, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hThreadWatch);

									// Close handle to target process
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;

									// Register global AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

									// Enable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWndMain, (int)ID::EditTimeout), TRUE);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);

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
			default:
				break;
			}
			switch (LOWORD(wParam))
			{
				// "Load" menu item
			case ID_MENU_LOAD:
				{
					// Simulate clicking "Load" button
					SendMessage(GetDlgItem(hWnd, (int)ID::ButtonLoad), BM_CLICK, NULL, NULL);
				}
				break;
				// "Unload" menu item
			case ID_MENU_UNLOAD:
				{
					// Simulate clicking "Unload" button
					SendMessage(GetDlgItem(hWnd, (int)ID::ButtonUnload), BM_CLICK, NULL, NULL);
				}
				break;
				// "Close" menu item
			case ID_MENU_CLOSE:
				{
					// Simulate clicking "Close" button
					SendMessage(GetDlgItem(hWnd, (int)ID::ButtonClose), BM_CLICK, NULL, NULL);
				}
				break;
				// "Select All" menu item
			case ID_MENU_SELECTALL:
				{
					// Select all items in ListViewFontList
					ListView_SetItemState(GetDlgItem(hWndMain, (int)ID::ListViewFontList), -1, LVIS_SELECTED, LVIS_SELECTED);
				}
			default:
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
					static LONG CursorOffsetY{};

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
							CursorOffsetY = ((LPSPLITTERSTRUCT)lParam)->ptCursor.y - rcSplitterClient.top;
							MapWindowRect(hWndSplitter, HWND_DESKTOP, &rcSplitterClient);
							CursorOffsetY += rcSplitterClient.top - rcSplitter.top;

							// Confine cursor to a specific rcangle
							/*
							┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓
							┃                                                                                                                               ┃
							┃                                                                                                                               ┃
							┃	┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓                                         ┃← Desktop
							┃	┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃                                         ┃
							┃	┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸─┬─┨                                         ┃
							┃	┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │ ┃                                         ┃
							┃	┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘ ┃                                         ┃
							┃	┃        │        │        │        │     Select Process     │                    ┃                                         ┃
							┃	┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────╂────────                                 ┃
							┃	┃ Font Name                                                        │ State        ┃        ↑                                ┃
							┃	┠──────────────────────────────────────────────────────────────────┼──────────────┨        │ cyListViewFontListMin          ┃
							┃	┃                                                                  ┆              ┃        ↓                                ┃
							┠┄┄┄╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	┃                                                                  ┆              ┃                             ↑           ┃
							┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                     rcMouseClip.top     ┃
							┃	┃                                                                  ┆              ┃                                         ┃
							┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                                         ┃
							┃	┃                                                                  ┆              ┃                                         ┃
							┃	┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨                                         ┃
							┃	┠──────────────────────────────────────────────────────────────────┴──────────────┨                      rcMouseClip.right →┃
							┃	┠─────────────────────────────────────────────────────────────────────────────┬───┨                                         ┃
							┃	┃ Temporarily load fonts to Windows or specific process                       │ ↑ ┃                                         ┃
							┃	┃                                                                             ├───┨                                         ┃
							┃	┃ How to load fonts to Windows:                                               │▓▓▓┃                                         ┃
							┃	┃ 1.Drag-drop font files onto the icon of this application.                   │▓▓▓┃                                         ┃
							┃	┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │▓▓▓┃                                         ┃
							┃	┃  view, then click "Load" button.                                            │▓▓▓┃                                         ┃
							┃	┃                                                                             ├───┨                                         ┃
							┃	┃ How to unload fonts from Windows:                                           │   ┃                                         ┃
							┃	┃ Select all fonts then click "Unload" or "Close" button or the X at the      │   ┃                                         ┃
							┃	┃ upper-right cornor.                                                         │   ┃                                         ┃
							┃	┃                                                                             │   ┃                                         ┃
							┃	┃ How to load fonts to process:                                               │   ┃                                         ┃
							┃	┃ 1.Click "Click to select process", select a process.                        │   ┃                      rcMouseClip.bottom ┃
							┃	┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │   ┃                             ↓           ┃
							┠┄┄┄╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼───╂┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃	┃ view, then click "Load" button.                                             │ ↓ ┃        } cyEditMessageMin               ┃
							┃	┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┹────────                                 ┃
							┃                                                                                                                               ┃
							┃                                                                                                                               ┃
							┃← rcMouseClip.left                                                                                                             ┃
							┃                                                                                                                               ┃
							┃                                                                                                                               ┃
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛
							*/

							// Calculate the minimal heights of ListViewFontList
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
							LONG cyListViewFontListMin{ rcListViewFontListItem.bottom + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

							// Calculate the minimal heights of EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
							HDC hDCEditMessage{ GetDC(hWndEditMessage) };
							HDC hDCEditMessageMemory{ CreateCompatibleDC(hDCEditMessage) };
							SelectFont(hDCEditMessageMemory, GetWindowFont(hWndEditMessage));
							TEXTMETRIC tm{};
							GetTextMetrics(hDCEditMessageMemory, &tm);
							DeleteDC(hDCEditMessageMemory);
							ReleaseDC(hWndEditMessage, hDCEditMessage);
							RECT rcEditMessage{}, rcEditMessageClient{};
							GetWindowRect(hWndEditMessage, &rcEditMessage);
							GetClientRect(hWndEditMessage, &rcEditMessageClient);
							LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + EditMessageTextMarginY };

							// Calculate confine rcangle
							RECT rcMainClient{}, rcDesktop{}, rcButtonOpen{};
							GetClientRect(hWnd, &rcMainClient);
							MapWindowRect(hWnd, HWND_DESKTOP, &rcMainClient);
							GetWindowRect(GetDesktopWindow(), &rcDesktop);
							GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
							RECT rcMouseClip{ rcDesktop.left, rcMainClient.top + (rcButtonOpen.bottom - rcButtonOpen.top) + cyListViewFontListMin + CursorOffsetY, rcDesktop.right, rcMainClient.bottom - ((rcSplitter.bottom - rcSplitter.top) - CursorOffsetY) - cyEditMessageMin };

							ClipCursor(&rcMouseClip);
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
							MoveWindow(hWndSplitter, rcSplitter.left, ptCursor.y - CursorOffsetY, rcSplitter.right - rcSplitter.left, rcSplitter.bottom - rcSplitter.top, TRUE);

							// Resize ListViewFontList
							HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
							RECT rcListViewFontList{};
							GetWindowRect(hWndListViewFontList, &rcListViewFontList);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
							MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, rcListViewFontList.right - rcListViewFontList.left, ptCursor.y - CursorOffsetY - rcListViewFontList.top, TRUE);

							// Resize EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
							RECT rcEditMessage{}, rcMainClient{};
							GetWindowRect(hWndEditMessage, &rcEditMessage);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);
							GetClientRect(hWnd, &rcMainClient);
							MoveWindow(hWndEditMessage, rcEditMessage.left, ptCursor.y + (rcSplitter.bottom - rcSplitter.top) - CursorOffsetY, rcEditMessage.right - rcEditMessage.left, rcMainClient.bottom - rcSplitter.bottom, TRUE);
						}
						break;
						// End dragging Splitter
					case SPLITTERNOTIFICATION::DRAGEND:
						{
							// Free cursor
							ClipCursor(NULL);

							// Redraw controls
							InvalidateRect(hWnd, NULL, FALSE);
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
			if ((HWND)wParam == GetDlgItem(hWnd, (int)ID::ListViewFontList))
			{
				POINT ptMenu{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
				if (ptMenu.x == -1 && ptMenu.y == -1)
				{
					HWND hWndListViewFontList{ GetDlgItem(hWnd,(int)ID::ListViewFontList) };

					int iSelectionMark{ ListView_GetSelectionMark(hWndListViewFontList) };
					if (iSelectionMark == -1)
					{
						RECT rcListViewFontListClient{};
						GetClientRect(hWndListViewFontList, &rcListViewFontListClient);
						MapWindowRect(hWndListViewFontList, HWND_DESKTOP, &rcListViewFontListClient);
						ptMenu = { rcListViewFontListClient.left, rcListViewFontListClient.top };
					}
					else
					{
						POINT ptSelectionMark{};
						ListView_GetItemPosition(hWndListViewFontList, iSelectionMark, &ptSelectionMark);
						ClientToScreen(hWndListViewFontList, &ptSelectionMark);
						ptMenu = ptSelectionMark;
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

				TrackPopupMenu(hMenuContextListViewFontList, uFlags | TPM_RIGHTBUTTON, ptMenu.x, ptMenu.y, 0, hWnd, NULL);
			}
			else
			{
				ret = DefWindowProc(hWnd, Msg, wParam, lParam);
			}
		}
		break;
	case WM_WINDOWPOSCHANGING:
		{
			// Get client rcangle before main window changes size
			GetClientRect(((LPWINDOWPOS)lParam)->hwnd, &rcMainClientOld);

			ret = DefWindowProc(hWnd, Msg, wParam, lParam);
		}
		break;
	case WM_SIZING:
		{
			// Get sizing edge
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
																															 ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
																															 ┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃
																															 ┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸─┬─┨
									┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓      ┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │ ┃
									┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃      ┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘ ┃
									┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸─┬─┨      ┃        │        │        │        │     Select Process     │                    ┃
									┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │ ┃      ┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────┨
									┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘ ┃      ┃ Font Name                                                        │ State        ┃
									┃        │        │        │        │     Select Process     │                    ┃      ┠──────────────────────────────────────────────────────────────────┼──────────────┨
									┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────┨      ┃                                                                  ┆              ┃
									┃ Font Name                                                        │ State        ┃      ┠──────────────────────────────────────────────────────────────────┼──────────────┨
									┠──────────────────────────────────────────────────────────────────┼──────────────┨      ┃                                                                  ┆              ┃
									┃                                                                  ┆              ┃      ┠──────────────────────────────────────────────────────────────────┼──────────────┨
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┃                                                                  ┆              ┃
									┃                                                                  ┆              ┃      ┠──────────────────────────────────────────────────────────────────┼──────────────┨
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┃                                                                  ┆              ┃
									┃                                                                  ┆              ┃      ┠──────────────────────────────────────────────────────────────────┼──────────────┨
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┃                                                                  ┆              ┃
									┃                                                                  ┆              ┃      ┠──────────────────────────────────────────────────────────────────┼──────────────┨
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┃                                                                  ┆              ┃
									┠──────────────────────────────────────────────────────────────────┴──────────────┨  =>  ┠──────────────────────────────────────────────────────────────────┴──────────────┨
									┠─────────────────────────────────────────────────────────────────────────────┬───┨      ┠─────────────────────────────────────────────────────────────────────────────┬───┨
									┃ Temporarily load fonts to Windows or specific process                       │ ↑ ┃      ┃ Temporarily load fonts to Windows or specific process                       │ ↑ ┃
									┃                                                                             ├───┨      ┃                                                                             ├───┨
									┃ How to load fonts to Windows:                                               │▓▓▓┃      ┃ How to load fonts to Windows:                                               │▓▓▓┃
									┃ 1.Drag-drop font files onto the icon of this application.                   │▓▓▓┃      ┃ 1.Drag-drop font files onto the icon of this application.                   │▓▓▓┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │▓▓▓┃      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │▓▓▓┃
									┃  view, then click "Load" button.                                            │▓▓▓┃      ┃  view, then click "Load" button.                                            │▓▓▓┃
									┃                                                                             ├───┨      ┃                                                                             ├───┨
									┃ How to unload fonts from Windows:                                           │   ┃      ┃ How to unload fonts from Windows:                                           │   ┃
									┃ Select all fonts then click "Unload" or "Close" button or the X at the      │   ┃      ┃ Select all fonts then click "Unload" or "Close" button or the X at the      │   ┃
									┃ upper-right cornor.                                                         │   ┃      ┃ upper-right cornor.                                                         │   ┃
									┃                                                                             │   ┃      ┃                                                                             │   ┃
									┃ How to load fonts to process:                                               │   ┃      ┃ How to load fonts to process:                                               │   ┃
									┃ 1.Click "Click to select process", select a process.                        │   ┃      ┃ 1.Click "Click to select process", select a process.                        │   ┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list ├───┨      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list ├───┨
									┃ view, then click "Load" button.                                             │ ↓ ┃      ┃ view, then click "Load" button.                                             │ ↓ ┃
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛      ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛
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
									LONG cyListViewFontListMin{ rcListViewFontListItem.bottom + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

									// Resize ListViewFontList
									RECT rcButtonOpen{}, rcSplitter{}, rcEditMessage{};
									GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
									GetWindowRect(GetDlgItem(hWnd, (int)ID::Splitter), &rcSplitter);
									GetWindowRect(GetDlgItem(hWnd, (int)ID::EditMessage), &rcEditMessage);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);

									bool bIsListViewFontListMinimized{ false };
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									if (HIWORD(lParam) - rcButtonOpen.bottom - (rcSplitter.bottom - rcSplitter.top) - (rcEditMessage.bottom - rcEditMessage.top) < cyListViewFontListMin)
									{
										bIsListViewFontListMinimized = true;

										MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), cyListViewFontListMin, TRUE);
									}
									else
									{
										MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), HIWORD(lParam) - rcButtonOpen.bottom - (rcSplitter.bottom - rcSplitter.top) - (rcEditMessage.bottom - rcEditMessage.top), TRUE);
									}

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
									if (bIsListViewFontListMinimized)
									{
										MoveWindow(hWndSplitter, rcSplitter.left, rcButtonOpen.bottom + cyListViewFontListMin, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, TRUE);
									}
									else
									{
										MoveWindow(hWndSplitter, rcSplitter.left, HIWORD(lParam) - (rcSplitter.bottom - rcSplitter.top) - (rcEditMessage.bottom - rcEditMessage.top), LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, TRUE);
									}

									// Resize EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
									if (bIsListViewFontListMinimized)
									{
										MoveWindow(hWndEditMessage, rcEditMessage.left, rcButtonOpen.bottom + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), HIWORD(lParam) - (rcButtonOpen.bottom + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top)), TRUE);
									}
									else
									{
										MoveWindow(hWndEditMessage, rcEditMessage.left, HIWORD(lParam) - (rcEditMessage.bottom - rcEditMessage.top), LOWORD(lParam), rcEditMessage.bottom - rcEditMessage.top, TRUE);
									}
								}
								break;
								// Resize EditMessage and keep the height of ListViewFontList as far as possible
							case WMSZ_BOTTOM:
							case WMSZ_BOTTOMLEFT:
							case WMSZ_BOTTOMRIGHT:
								{
									/*
									┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓      ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
									┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃      ┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃
									┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸─┬─┨      ┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸─┬─┨
									┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │ ┃      ┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │ ┃
									┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘ ┃      ┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘ ┃
									┃        │        │        │        │     Select Process     │                    ┃      ┃        │        │        │        │     Select Process     │                    ┃
									┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────┨      ┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────┨
									┃ Font Name                                                        │ State        ┃      ┃ Font Name                                                        │ State        ┃
									┠──────────────────────────────────────────────────────────────────┼──────────────┨      ┠──────────────────────────────────────────────────────────────────┼──────────────┨
									┃                                                                  ┆              ┃      ┃                                                                  ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠──────────────────────────────────────────────────────────────────┼──────────────┨
									┃                                                                  ┆              ┃      ┃                                                                  ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                  ┆              ┃      ┃                                                                  ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                  ┆              ┃      ┃                                                                  ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┠──────────────────────────────────────────────────────────────────┴──────────────┨  =>  ┠──────────────────────────────────────────────────────────────────┴──────────────┨
									┠─────────────────────────────────────────────────────────────────────────────┬───┨      ┠─────────────────────────────────────────────────────────────────────────────┬───┨
									┃ Temporarily load fonts to Windows or specific process                       │ ↑ ┃      ┃ Temporarily load fonts to Windows or specific process                       │ ↑ ┃
									┃                                                                             ├───┨      ┃                                                                             ├───┨
									┃ How to load fonts to Windows:                                               │▓▓▓┃      ┃ How to load fonts to Windows:                                               │▓▓▓┃
									┃ 1.Drag-drop font files onto the icon of this application.                   │▓▓▓┃      ┃ 1.Drag-drop font files onto the icon of this application.                   │▓▓▓┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │▓▓▓┃      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │▓▓▓┃
									┃  view, then click "Load" button.                                            │▓▓▓┃      ┃  view, then click "Load" button.                                            │▓▓▓┃
									┃                                                                             ├───┨      ┃                                                                             │▓▓▓┃
									┃ How to unload fonts from Windows:                                           │   ┃      ┃ How to unload fonts from Windows:                                           │▓▓▓┃
									┃ Select all fonts then click "Unload" or "Close" button or the X at the      │   ┃      ┃ Select all fonts then click "Unload" or "Close" button or the X at the      ├───┨
									┃ upper-right cornor.                                                         │   ┃      ┃ upper-right cornor.                                                         │   ┃
									┃                                                                             │   ┃      ┃                                                                             │   ┃
									┃ How to load fonts to process:                                               │   ┃      ┃ How to load fonts to process:                                               │   ┃
									┃ 1.Click "Click to select process", select a process.                        │   ┃      ┃ 1.Click "Click to select process", select a process.                        │   ┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list ├───┨      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │   ┃
									┃ view, then click "Load" button.                                             │ ↓ ┃      ┃ view, then click "Load" button.                                             │   ┃
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛      ┃ How to unload fonts from process:                                           │   ┃
																															 ┃ Select all fonts then click "Unload" or "Close" button or the X at the      ├───┨
																															 ┃  upper-right cornor or terminate selected process.                          │ ↓ ┃
																															 ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛
									*/

									// Calculate the minimal height of EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
									HDC hDCEditMessage{ GetDC(hWndEditMessage) };
									HDC hDCEditMessageMemory{ CreateCompatibleDC(hDCEditMessage) };
									SelectFont(hDCEditMessageMemory, GetWindowFont(hWndEditMessage));
									TEXTMETRIC tm{};
									GetTextMetrics(hDCEditMessageMemory, &tm);
									DeleteDC(hDCEditMessageMemory);
									RECT rcEditMessage{}, rcEditMessageClient{};
									GetWindowRect(hWndEditMessage, &rcEditMessage);
									GetClientRect(hWndEditMessage, &rcEditMessageClient);
									LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + EditMessageTextMarginY };

									// Resize EditMessage
									bool bIsEditMessageMinimized{ false };
									MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);
									if (HIWORD(lParam) - rcEditMessage.top < cyEditMessageMin)
									{
										bIsEditMessageMinimized = true;

										MoveWindow(hWndEditMessage, rcEditMessage.left, HIWORD(lParam) - cyEditMessageMin, LOWORD(lParam), cyEditMessageMin, TRUE);
									}
									else
									{
										MoveWindow(hWndEditMessage, rcEditMessage.left, rcEditMessage.top, LOWORD(lParam), HIWORD(lParam) - rcEditMessage.top, TRUE);
									}

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
									RECT rcSplitter{};
									GetWindowRect(hWndSplitter, &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									if (bIsEditMessageMinimized)
									{
										MoveWindow(hWndSplitter, rcSplitter.left, HIWORD(lParam) - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, TRUE);
									}
									else
									{
										MoveWindow(hWndSplitter, rcSplitter.left, rcSplitter.top, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, TRUE);
									}

									// Resize ListViewFontList
									RECT rcButtonOpen{};
									GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
									HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
									RECT rcListViewFontList{};
									GetWindowRect(hWndListViewFontList, &rcListViewFontList);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									if (bIsEditMessageMinimized)
									{
										MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), HIWORD(lParam) - cyEditMessageMin - (rcSplitter.bottom - rcSplitter.top) - rcButtonOpen.bottom, TRUE);
									}
									else
									{
										MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcListViewFontList.bottom - rcListViewFontList.top, TRUE);
									}
								}
								break;
								// Just modify width
							case WMSZ_LEFT:
							case WMSZ_RIGHT:
								{
									/*
									┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓      ┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
									┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃      ┃FontLoaderEx                                                                       ┃_  ┃ □ ┃ x ┃
									┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸─┬─┨      ┠────────┬────────┬────────┬────────┬───────────────────────────────────┬───────┬───┸───┸───┸───┨
									┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │ ┃      ┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │               ┃
									┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘ ┃      ┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘               ┃
									┃        │        │        │        │     Select Process     │                    ┃      ┃        │        │        │        │     Select Process     │                                  ┃
									┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────┨      ┠────────┴────────┴────────┴────────┴────────────────────────┴───────────────────┬──────────────┨
									┃ Font Name                                                        │ State        ┃      ┃ Font Name                                                                      │ State        ┃
									┠──────────────────────────────────────────────────────────────────┼──────────────┨      ┠────────────────────────────────────────────────────────────────────────────────┼──────────────┨
									┃                                                                  ┆              ┃      ┃                                                                                ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                  ┆              ┃      ┃                                                                                ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                  ┆              ┃      ┃                                                                                ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┃                                                                  ┆              ┃      ┃                                                                                ┆              ┃
									┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨      ┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
									┠──────────────────────────────────────────────────────────────────┴──────────────┨  =>  ┠────────────────────────────────────────────────────────────────────────────────┴──────────────┨
									┠─────────────────────────────────────────────────────────────────────────────┬───┨      ┠───────────────────────────────────────────────────────────────────────────────────────────┬───┨
									┃ Temporarily load fonts to Windows or specific process                       │ ↑ ┃      ┃ Temporarily load fonts to Windows or specific process                                     │ ↑ ┃
									┃                                                                             ├───┨      ┃                                                                                           ├───┨
									┃ How to load fonts to Windows:                                               │▓▓▓┃      ┃ How to load fonts to Windows:                                                             │▓▓▓┃
									┃ 1.Drag-drop font files onto the icon of this application.                   │▓▓▓┃      ┃ 1.Drag-drop font files onto the icon of this application.                                 │▓▓▓┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │▓▓▓┃      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list view, then    │▓▓▓┃
									┃  view, then click "Load" button.                                            │▓▓▓┃      ┃  click "Load" button.                                                                     │▓▓▓┃
									┃                                                                             ├───┨      ┃                                                                                           ├───┨
									┃ How to unload fonts from Windows:                                           │   ┃      ┃ How to unload fonts from Windows:                                                         │   ┃
									┃ Select all fonts then click "Unload" or "Close" button or the X at the      │   ┃      ┃ Select all fonts then click "Unload" or "Close" button or the X at the upper-right cornor.│   ┃
									┃ upper-right cornor.                                                         │   ┃      ┃                                                                                           │   ┃
									┃                                                                             │   ┃      ┃ How to load fonts to process:                                                             │   ┃
									┃ How to load fonts to process:                                               │   ┃      ┃ 1.Click "Click to select process", select a process.                                      │   ┃
									┃ 1.Click "Click to select process", select a process.                        │   ┃      ┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list view, then    │   ┃
									┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list ├───┨      ┃  click "Load" button.                                                                     ├───┨
									┃ view, then click "Load" button.                                             │ ↓ ┃      ┃                                                                                           │ ↓ ┃
									┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛      ┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛
									*/

									// Resize ListViewFontList
									HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
									RECT rcListViewFontList{};
									GetWindowRect(hWndListViewFontList, &rcListViewFontList);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
									MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), rcListViewFontList.bottom - rcListViewFontList.top, TRUE);

									// Resize Splitter
									HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
									RECT rcSplitter{};
									GetWindowRect(hWndSplitter, &rcSplitter);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
									MoveWindow(hWndSplitter, rcSplitter.left, rcSplitter.top, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, TRUE);

									// Resize EditMessage
									HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
									RECT rcEditMessage{};
									GetWindowRect(hWndEditMessage, &rcEditMessage);
									MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);
									MoveWindow(hWndEditMessage, rcEditMessage.left, rcEditMessage.top, LOWORD(lParam), rcEditMessage.bottom - rcEditMessage.top, TRUE);
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
							┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
							┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃
							┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸─┬─┨
							┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │ ┃
							┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘ ┃
							┃        │        │        │        │     Select Process     │                    ┃
							┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────┨
							┃ Font Name                                                        │ State        ┃
							┠──────────────────────────────────────────────────────────────────┼──────────────┨
							┃                                                                  ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                  ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                  ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                  ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┠──────────────────────────────────────────────────────────────────┴──────────────┨  =>
							┠─────────────────────────────────────────────────────────────────────────────┬───┨
							┃ Temporarily load fonts to Windows or specific process                       │ ↑ ┃
							┃                                                                             ├───┨
							┃ How to load fonts to Windows:                                               │▓▓▓┃
							┃ 1.Drag-drop font files onto the icon of this application.                   │▓▓▓┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │▓▓▓┃
							┃  view, then click "Load" button.                                            │▓▓▓┃
							┃                                                                             ├───┨
							┃ How to unload fonts from Windows:                                           │   ┃
							┃ Select all fonts then click "Unload" or "Close" button or the X at the      │   ┃
							┃ upper-right cornor.                                                         │   ┃
							┃                                                                             │   ┃
							┃ How to load fonts to process:                                               │   ┃
							┃ 1.Click "Click to select process", select a process.                        │   ┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list ├───┨
							┃  view, then click "Load" button.                                            │ ↓ ┃
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛

							┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
							┃FontLoaderEx                                                                                                       ┃_  ┃ □ ┃ x ┃
							┠────────┬────────┬────────┬────────┬───────────────────────────────────┬───────┬───────────────────────────────────┸───┸───┸───┨
							┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │                                               ┃
							┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘                                               ┃
							┃        │        │        │        │     Select Process     │                                                                  ┃
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
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛
							*/

							// Resize Splitter
							HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
							RECT rcButtonOpen{}, rcListViewFontList{}, rcSplitter{}, rcEditMessage{};
							GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
							GetWindowRect(GetDlgItem(hWnd, (int)ID::ListViewFontList), &rcListViewFontList);
							GetWindowRect(hWndSplitter, &rcSplitter);
							GetWindowRect(GetDlgItem(hWnd, (int)ID::EditMessage), &rcEditMessage);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);

							LONG ySplitterTopNew{ (((rcSplitter.top - rcButtonOpen.bottom) + (rcSplitter.bottom - rcButtonOpen.bottom)) * (HIWORD(lParam) - rcButtonOpen.bottom)) / (((rcMainClientOld.bottom - rcMainClientOld.top) - rcButtonOpen.bottom) * 2) - ((rcSplitter.bottom - rcSplitter.top) / 2) + rcButtonOpen.bottom };
							MoveWindow(hWndSplitter, rcSplitter.left, ySplitterTopNew, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, TRUE);

							// Resize ListViewFontList
							HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
							MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), ySplitterTopNew - rcButtonOpen.bottom, TRUE);

							// Resize EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
							MoveWindow(hWndEditMessage, rcEditMessage.left, ySplitterTopNew + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), HIWORD(lParam) - (ySplitterTopNew + (rcSplitter.bottom - rcSplitter.top)), TRUE);

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
							┠────────┬────────┬────────┬────────┬───────────────────────────────────┬───────┬───────────────────────────────────┸───┸───┸───┨
							┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │                                               ┃
							┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘                                               ┃
							┃        │        │        │        │     Select Process     │                                                                  ┃
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
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛

							┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
							┃FontLoaderEx                                                         ┃_  ┃ □ ┃ x ┃
							┠────────┬────────┬────────┬────────┬─────────────────────────────────┸─┬─┸───┸─┬─┨
							┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   │ ┃
							┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┘ ┃
							┃        │        │        │        │     Select Process     │                    ┃
							┠────────┴────────┴────────┴────────┴────────────────────────┴─────┬──────────────┨
							┃ Font Name                                                        │ State        ┃
							┠──────────────────────────────────────────────────────────────────┼──────────────┨
							┃                                                                  ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                  ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                  ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┃                                                                  ┆              ┃
							┠┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┄┼┄┄┄┄┄┄┄┄┄┄┄┄┄┄┨
							┠──────────────────────────────────────────────────────────────────┴──────────────┨
							┠─────────────────────────────────────────────────────────────────────────────┬───┨
							┃ Temporarily load fonts to Windows or specific process                       │ ↑ ┃
							┃                                                                             ├───┨
							┃ How to load fonts to Windows:                                               │▓▓▓┃
							┃ 1.Drag-drop font files onto the icon of this application.                   │▓▓▓┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list │▓▓▓┃
							┃  view, then click "Load" button.                                            │▓▓▓┃
							┃                                                                             ├───┨
							┃ How to unload fonts from Windows:                                           │   ┃
							┃ Select all fonts then click "Unload" or "Close" button or the X at the      │   ┃
							┃ upper-right cornor.                                                         │   ┃
							┃                                                                             │   ┃
							┃ How to load fonts to process:                                               │   ┃
							┃ 1.Click "Click to select process", select a process.                        │   ┃
							┃ 2.Click "Open" button to select fonts or drag-drop font files onto the list ├───┨
							┃  view, then click "Load" button.                                            │ ↓ ┃
							┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━┛
							*/

							// Calculate minimal height of ListViewFontList
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
							LONG cyListViewFontListMin{ rcListViewFontListItem.bottom + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

							// Calculate minimal height of EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
							HDC hDCEditMessage{ GetDC(hWndEditMessage) };
							HDC hDCEditMessageMemory{ CreateCompatibleDC(hDCEditMessage) };
							SelectFont(hDCEditMessageMemory, GetWindowFont(hWndEditMessage));
							TEXTMETRIC tm{};
							GetTextMetrics(hDCEditMessageMemory, &tm);
							DeleteDC(hDCEditMessageMemory);
							RECT rcEditMessage{}, rcEditMessageClient{};
							GetWindowRect(hWndEditMessage, &rcEditMessage);
							GetClientRect(hWndEditMessage, &rcEditMessageClient);
							LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + EditMessageTextMarginY };

							// Calculate new position of splitter
							HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
							RECT rcButtonOpen{}, rcSplitter{};
							GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
							GetWindowRect(hWndSplitter, &rcSplitter);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcListViewFontList);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcSplitter);
							MapWindowRect(HWND_DESKTOP, hWnd, &rcEditMessage);

							LONG ySplitterTopNew{ (((rcSplitter.top - rcButtonOpen.bottom) + (rcSplitter.bottom - rcButtonOpen.bottom)) * (HIWORD(lParam) - rcButtonOpen.bottom)) / (((rcMainClientOld.bottom - rcMainClientOld.top) - rcButtonOpen.bottom) * 2) - ((rcSplitter.bottom - rcSplitter.top) / 2) + rcButtonOpen.bottom };
							LONG cyListViewFontListNew{ ySplitterTopNew - rcButtonOpen.bottom };
							LONG cyEditMessageNew{ HIWORD(lParam) - (ySplitterTopNew + (rcSplitter.bottom - rcSplitter.top)) };

							// If cyListViewFontListNew < cyListViewFontListMin, keep the minimal height of ListViewFontList
							if (cyListViewFontListNew < cyListViewFontListMin)
							{
								// Resize ListViewFontList
								HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
								MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), cyListViewFontListMin, TRUE);

								// Resize Splitter
								MoveWindow(hWndSplitter, rcSplitter.left, rcButtonOpen.bottom + cyListViewFontListMin, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, TRUE);

								// Resize EditMessage
								HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
								MoveWindow(hWndEditMessage, rcEditMessage.left, rcButtonOpen.bottom + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), HIWORD(lParam) - (rcButtonOpen.bottom + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top)), TRUE);
							}
							// If cyEditMessageNew < cyEditMessageMin, keep the minimal height of EditMessage
							else if (cyEditMessageNew < cyEditMessageMin)
							{
								// Resize EditMessage
								HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
								MoveWindow(hWndEditMessage, rcEditMessage.left, HIWORD(lParam) - cyEditMessageMin, LOWORD(lParam), cyEditMessageMin, TRUE);

								// Resize Splitter
								MoveWindow(hWndSplitter, rcSplitter.left, HIWORD(lParam) - (cyEditMessageMin + (rcSplitter.bottom - rcSplitter.top)), LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, TRUE);

								// Resize ListViewFontList
								HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
								MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), HIWORD(lParam) - (cyEditMessageMin + (rcSplitter.bottom - rcSplitter.top) + rcButtonOpen.bottom), TRUE);
							}
							// Else resize as usual
							else
							{
								// Resize Splitter
								MoveWindow(hWndSplitter, rcSplitter.left, ySplitterTopNew, LOWORD(lParam), rcSplitter.bottom - rcSplitter.top, TRUE);

								// Resize ListViewFontList
								HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
								MoveWindow(hWndListViewFontList, rcListViewFontList.left, rcListViewFontList.top, LOWORD(lParam), cyListViewFontListNew, TRUE);

								// Resize EditMessage
								HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
								MoveWindow(hWndEditMessage, rcEditMessage.left, ySplitterTopNew + (rcSplitter.bottom - rcSplitter.top), LOWORD(lParam), cyEditMessageNew, TRUE);
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
		}
		break;
	case WM_GETMINMAXINFO:
		{
			// Limit minimize window size
			/*
			┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┳━━━┳━━━┳━━━┓
			┃FontLoaderEx                                                       ┃_  ┃ □ ┃ x ┃
			┠────────┬────────┬────────┬────────┬───────────────────────────────┸───╀───┸───╂───────
			┃        │        │        │        │□ Broadcast WM_FONTCHANGE  Timeout:│5000   ┃       ↑
			┃  Open  │  Close │  Load  │ Unload ├────────────────────────┐          └───────┨       │ cyButtonOpen
			┃        │        │        │        │     Select Process     │                  ┃       ↓
			┠────────┴────────┴────────┴────────┴────────────────────────┴───┬──────────────╂───────
			┃ Font Name                                                      │ State        ┃       ↑
			┠────────────────────────────────────────────────────────────────┼──────────────┨       │ cyListViewFontListMin
			┃                                                                ┆              ┃       ↓
			┠────────────────────────────────────────────────────────────────┴──────────────╂───────
			┠───────────────────────────────────────────────────────────────────────────┬───╂───────
			┃ Temporarily load fonts to Windows or specific process                     ├───┨       } cyEditMessagemin
			┡━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┷━━━╉───────
			│                              rcEditTimeout.right                              │
			│←─────────────────────────────────────────────────────────────────────────────→│
			│                                                                               │
			*/

			// Get ButtonOpen window rcangle
			RECT rcButtonOpen{};
			GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rcButtonOpen);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcButtonOpen);

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
			LONG cyListViewFontListMin{ rcListViewFontListItem.bottom + ((rcListViewFontList.bottom - rcListViewFontList.top) - (rcListViewFontListClient.bottom - rcListViewFontListClient.top)) };

			// Get Splitter window rcangle
			RECT rcSplitter{};
			GetWindowRect(GetDlgItem(hWnd, (int)ID::Splitter), &rcSplitter);

			// Calculate the minimal height of Editmessage
			HWND hWndEditMessage{ GetDlgItem(hWnd,(int)ID::EditMessage) };
			HDC hDCEditMessage{ GetDC(hWndEditMessage) };
			HDC hDCEditMessageMemory{ CreateCompatibleDC(hDCEditMessage) };
			SelectFont(hDCEditMessageMemory, GetWindowFont(hWndEditMessage));
			TEXTMETRIC tm{};
			GetTextMetrics(hDCEditMessageMemory, &tm);
			DeleteDC(hDCEditMessageMemory);
			RECT rcEditMessage{}, rcEditMessageClient{};
			GetWindowRect(hWndEditMessage, &rcEditMessage);
			GetClientRect(hWndEditMessage, &rcEditMessageClient);
			LONG cyEditMessageMin{ tm.tmHeight + tm.tmExternalLeading * 2 + ((rcEditMessage.bottom - rcEditMessage.top) + (rcEditMessageClient.top - rcEditMessageClient.bottom)) + EditMessageTextMarginY };

			// Get EditTimeout window rcangle
			RECT rcEditTimeout{};
			GetWindowRect(GetDlgItem(hWnd, (int)ID::EditTimeout), &rcEditTimeout);
			MapWindowRect(HWND_DESKTOP, hWnd, &rcEditTimeout);

			// Calculate minimal window size
			RECT rcMainMin{ 0, 0, rcEditTimeout.right, (rcButtonOpen.bottom - rcButtonOpen.top) + cyListViewFontListMin + (rcSplitter.bottom - rcSplitter.top) + cyEditMessageMin };
			AdjustWindowRect(&rcMainMin, (DWORD)GetWindowLongPtr(hWnd, GWL_STYLE), FALSE);
			((LPMINMAXINFO)lParam)->ptMinTrackSize = { rcMainMin.right - rcMainMin.left, rcMainMin.bottom - rcMainMin.top };
		}
		break;
	case WM_CTLCOLORSTATIC:
		{
			// Change the background color of ButtonBroadcastWM_FONTCHANGE, StaticTimeout and EditMessage to default window background color
			// From https://social.msdn.microsoft.com/Forums/vstudio/en-US/7b6d1815-87e3-4f47-b5d5-fd4caa0e0a89/why-is-wmctlcolorstatic-sent-for-a-button-instead-of-wmctlcolorbtn?forum=vclanguage
			// "WM_CTLCOLORSTATIC is sent by any control that displays text which would be displayed using the default dialog/window background color. 
			// This includes check boxes, radio buttons, group boxes, static text, read-only or disabled edit controls, and disabled combo boxes (all styles)."
			if (((HWND)lParam == GetDlgItem(hWnd, (int)ID::ButtonBroadcastWM_FONTCHANGE)) || ((HWND)lParam == GetDlgItem(hWnd, (int)ID::StaticTimeout)) || ((HWND)lParam == GetDlgItem(hWnd, (int)ID::EditMessage)))
			{
				ret = (LRESULT)GetSysColorBrush(COLOR_WINDOW);
			}
		}
		break;
	default:
		{
			ret = DefWindowProc(hWnd, Msg, wParam, lParam);
		}
		break;
	}

	return ret;
}

LRESULT CALLBACK ListViewFontListSubclassProc(HWND hWndListViewFontList, UINT Msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData)
{
	LRESULT ret{};

	switch (Msg)
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

			ret = DefSubclassProc(hWndListViewFontList, Msg, wParam, lParam);
		}
		break;
	case WM_DROPFILES:
		{
			// Process drag-drop and open fonts
			bool bIsFontListEmptyBefore{ FontList.empty() };

			HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
			HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

			std::wstringstream Message{};
			int cchMessageLength{};

			UINT nFileCount{ DragQueryFile((HDROP)wParam, 0xFFFFFFFF, NULL, 0) };
			LVITEM lvi{ LVIF_TEXT, ListView_GetItemCount(hWndListViewFontList) };
			for (UINT i = 0; i < nFileCount; i++)
			{
				WCHAR szFileName[MAX_PATH]{};
				DragQueryFile((HDROP)wParam, i, szFileName, MAX_PATH);
				if (PathMatchSpec(szFileName, L"*.ttf") || PathMatchSpec(szFileName, L"*.ttc") || PathMatchSpec(szFileName, L"*.otf"))
				{
					FontList.push_back(szFileName);

					lvi.iSubItem = 0;
					lvi.pszText = szFileName;
					ListView_InsertItem(hWndListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = (LPWSTR)L"Not loaded";
					ListView_SetItem(hWndListViewFontList, &lvi);
					ListView_SetItemState(hWndListViewFontList, lvi.iItem, LVIS_SELECTED, LVIS_SELECTED);
					lvi.iItem++;

					Message << szFileName << L" opened\r\n";
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
					Message.str(L"");
				}
			}
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, L"\r\n");

			DragFinish((HDROP)wParam);

			if (bIsFontListEmptyBefore)
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_ENABLED);
			}
		}
		break;
	case WM_WINDOWPOSCHANGED:
		{
			// Post USERMESSAGE::CHILDWINDOWPOSCHANGED to parent window
			PostMessage(hWndMain, (UINT)USERMESSAGE::CHILDWINDOWPOSCHANGED, (WPARAM)GetDlgCtrlID(hWndListViewFontList), NULL);

			ret = DefSubclassProc(hWndListViewFontList, Msg, wParam, lParam);
		}
		break;
	default:
		{
			ret = DefSubclassProc(hWndListViewFontList, Msg, wParam, lParam);
		}
		break;
	}

	return ret;
}

LRESULT CALLBACK EditMessageSubclassProc(HWND hWndEditMessage, UINT Msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIDSubclass, DWORD_PTR dwRefData)
{
	LRESULT ret{};

	switch (Msg)
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

			ret = DefSubclassProc(hWndEditMessage, Msg, wParam, lParam);
		}
		break;
	case WM_CONTEXTMENU:
		{
			// Delete "Undo", "Cut", "Paste" and "Clear" menu items from context menu
			HWINEVENTHOOK hWinEventHook{ SetWinEventHook(EVENT_SYSTEM_MENUPOPUPSTART, EVENT_SYSTEM_MENUPOPUPSTART, NULL,
				[](HWINEVENTHOOK hWinEventHook, DWORD Event, HWND hWnd, LONG idObject, LONG idChild, DWORD idEventThread, DWORD dwmsEventTime)
				{
					if (idObject == OBJID_CLIENT && idChild == CHILDID_SELF)
					{
						HMENU hMenuContextEdit{ (HMENU)SendMessage(hWnd, MN_GETHMENU, NULL, NULL) };

						// Menu Item identifiers in Edit control context menu is the same as corresponding window messages
						DeleteMenu(hMenuContextEdit, WM_UNDO, MF_BYCOMMAND);	// Undo
						DeleteMenu(hMenuContextEdit, WM_CUT, MF_BYCOMMAND);	// Cut
						DeleteMenu(hMenuContextEdit, WM_PASTE, MF_BYCOMMAND);	// Paste
						DeleteMenu(hMenuContextEdit, WM_CLEAR, MF_BYCOMMAND);	// Clear
						DeleteMenu(hMenuContextEdit, 0, MF_BYPOSITION);	// Seperator

						RECT rcContextMenu{};
						GetWindowRect(hWnd, &rcContextMenu);
						POINT ptCursor{};
						GetCursorPos(&ptCursor);
						if (rcContextMenu.bottom <= ptCursor.y)
						{
							MoveWindow(hWnd, rcContextMenu.left, rcContextMenu.top + (ptCursor.y - rcContextMenu.bottom), rcContextMenu.right - rcContextMenu.left, rcContextMenu.bottom - rcContextMenu.top, FALSE);
						}
					}
				},
				GetCurrentProcessId(), GetCurrentThreadId(), WINEVENT_OUTOFCONTEXT) };

			ret = DefSubclassProc(hWndEditMessage, Msg, wParam, lParam);

			UnhookWinEvent(hWinEventHook);
		}
		break;
	case WM_WINDOWPOSCHANGED:
		{
			// Post USERMESSAGE::CHILDWINDOWPOSCHANGED to parent window
			PostMessage(hWndMain, (UINT)USERMESSAGE::CHILDWINDOWPOSCHANGED, (WPARAM)GetDlgCtrlID(hWndEditMessage), NULL);

			ret = DefSubclassProc(hWndEditMessage, Msg, wParam, lParam);
		}
		break;
	default:
		{
			ret = DefSubclassProc(hWndEditMessage, Msg, wParam, lParam);
		}
		break;
	}

	return ret;
}

INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	INT_PTR ret{};

	static HFONT hFontDialog{};

	static std::vector<ProcessInfo> ProcessList{};

	static bool bOrderByProcessAscending{ true };
	static bool bOrderByPIDAscending{ true };

	switch (Msg)
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

										// Add arrow to the header in ListViewProcessList
										Header_GetItem(hWndHeaderListViewProcessList, 1, &hdi);
										hdi.fmt = hdi.fmt & ~(HDF_SORTDOWN | HDF_SORTUP);
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

										// Add arrow to the header in ListViewProcessList
										Header_GetItem(hWndHeaderListViewProcessList, 0, &hdi);
										hdi.fmt = hdi.fmt & ~(HDF_SORTDOWN | HDF_SORTUP);
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

							// Reset contents of ListViewProcessList
							LVITEM lvi{ LVIF_TEXT, 0 };
							for (auto& i : ProcessList)
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
							SendMessage(GetDlgItem(hWndDialog, IDOK), BM_CLICK, NULL, NULL);

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

HHOOK hHookCBT{};

int MessageBoxCentered(HWND hWnd, LPCTSTR lpText, LPCTSTR lpCaption, UINT uType)
{
	// Center message box at its parent window
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
		if (!AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(tp), NULL, NULL))
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

			ssRet.str(L"");
		} while (false);
	}

	return ssRet.str();
}