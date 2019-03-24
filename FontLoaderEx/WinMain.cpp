#if !defined(UNICODE) || !defined(_UNICODE)
#error Unicode character set required
#endif

#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "ComCtl32.lib")
#pragma comment(lib, "Shlwapi.lib")

#include <windows.h>
#include <CommCtrl.h>
#include <shlwapi.h>
#include <tlhelp32.h>
#include <windowsx.h>
#include <process.h>
#include <string>
#include <cstring>
#include <sstream>
#include <list>
#include <vector>
#include <algorithm>
#include <cstddef>
#include "FontResource.h"
#include "Globals.h"
#include "Splitter.h"
#include "resource.h"

LRESULT CALLBACK WndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);

std::list<FontResource> FontList{};

HWND hWndMain{};
HMENU hMenuContextListViewFontList{};

bool bDragDropHasFonts{ false };

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
	//Register default AddFont() and RemoveFont() procedure
	FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);

	//Process drag-drop font files onto the application icon stage I
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

	//Initialize common controls
	InitCommonControls();

	//Create window
	WNDCLASS wc{ CS_HREDRAW | CS_VREDRAW, WndProc, 0, 0, hInstance, LoadIcon(NULL, IDI_APPLICATION), LoadCursor(NULL, IDC_ARROW), GetSysColorBrush(COLOR_WINDOW), NULL, L"FontLoaderEx" };

	if (!RegisterClass(&wc))
	{
		return -1;
	}

	if (!(hWndMain = CreateWindow(L"FontLoaderEx", L"FontLoaderEx", WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 700, 700, NULL, NULL, hInstance, NULL)))
	{
		return -1;
	}

	ShowWindow(hWndMain, nShowCmd);
	UpdateWindow(hWndMain);

	hMenuContextListViewFontList = GetSubMenu(LoadMenu(hInstance, MAKEINTRESOURCE(IDR_CONTEXTMENU1)), 0);

	HACCEL hAccel{ LoadAccelerators(hInstance, MAKEINTRESOURCE(IDR_ACCELERATOR1)) };

	MSG Msg{};
	BOOL bRet{}, b1{}, b2{};
	while ((bRet = GetMessage(&Msg, NULL, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			return (int)GetLastError();
		}
		else
		{
			//Avoid short-circuit evaluation
			b1 = IsDialogMessage(hWndMain, &Msg);
			b2 = TranslateAccelerator(hWndMain, hAccel, &Msg);
			if (!(b1 || b2))
			{
				TranslateMessage(&Msg);
				DispatchMessage(&Msg);
			}
		}
	}

	return (int)Msg.wParam;
}

LRESULT CALLBACK ListViewProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT Msg, WPARAM wParam, LPARAM IParam);

WNDPROC OldStaticProc;
WNDPROC OldListViewProc;

void* pfnRemoteAddFontProc{};
void* pfnRemoteRemoveFontProc{};

HANDLE hWatchThread{};
HANDLE hMessageThread{};
HANDLE hProxyProcess{};
HWND hWndProxy{};

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

ProcessInfo TargetProcessInfo{};
PROCESS_INFORMATION piProxyProcess{};

DWORD dwTimeout{};

bool EnableDebugPrivilege();
bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout);
bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout);

#ifdef _WIN64
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll64.dll" };
const WCHAR szInjectionDllNameByProxy[]{ L"FontLoaderExInjectionDll.dll" };
#else
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll.dll" };
const WCHAR szInjectionDllNameByProxy[]{ L"FontLoaderExInjectionDll64.dll" };
#endif

LRESULT CALLBACK WndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	switch ((USERMESSAGE)Msg)
	{
	case USERMESSAGE::CLOSEWORKINGTHREADTERMINATED:
		{
			HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

			if (wParam)
			{
				std::wstringstream Message{};
				int iMessageLength{};

				//If loaded via proxy
				if (piProxyProcess.hProcess)
				{
					//Terminate watch thread
					SetEvent(hEventTerminateWatchThread);
					WaitForSingleObject(hWatchThread, INFINITE);
					CloseHandle(hEventTerminateWatchThread);
					CloseHandle(hWatchThread);

					//Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
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
							Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
						}
						goto break_DAA249E0;
					default:
						break;
					}
				break_DAA249E0:
					break;
				continue_DAA249E0:

					//Terminate proxy process
					COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
					FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
					WaitForSingleObject(piProxyProcess.hProcess, INFINITE);
					CloseHandle(piProxyProcess.hThread);
					CloseHandle(piProxyProcess.hProcess);
					piProxyProcess.hProcess = NULL;

					//Terminate message thread
					SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
					WaitForSingleObject(hMessageThread, INFINITE);
					DWORD dwMessageThreadExitCode{};
					GetExitCodeThread(hMessageThread, &dwMessageThreadExitCode);
					if (dwMessageThreadExitCode)
					{
						MessageBox(NULL, L"Message thread exited abnormally.", L"FontLoaderEx", MB_ICONERROR);
					}
					CloseHandle(hMessageThread);

					//Close handle to target process
					CloseHandle(TargetProcessInfo.hProcess);
					TargetProcessInfo.hProcess = NULL;
				}

				//Else DIY
				if (TargetProcessInfo.hProcess)
				{
					//Terminate watch thread
					SetEvent(hEventTerminateWatchThread);
					WaitForSingleObject(hWatchThread, INFINITE);
					CloseHandle(hEventTerminateWatchThread);
					CloseHandle(hWatchThread);

					//Unload FontLoaderExInjectionDll(64).dll from target process
					if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
					{
						Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
						iMessageLength = Edit_GetTextLength(hWndEditMessage);
						Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
						Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
						break;
					}

					//Close handle to target process
					CloseHandle(TargetProcessInfo.hProcess);
					TargetProcessInfo.hProcess = NULL;
				}

				DestroyWindow(hWnd);
			}
			else
			{
				switch (MessageBox(hWnd, L"Some fonts are not successfully unloaded.\r\n\r\nDo you still want to exit?", L"FontLoaderEx", MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON1 | MB_APPLMODAL))
				{
				case IDYES:
					{
						std::wstringstream Message{};
						int iMessageLength{};

						//If loaded via proxy
						if (piProxyProcess.hProcess)
						{
							//Terminate watch thread
							SetEvent(hEventTerminateWatchThread);
							WaitForSingleObject(hWatchThread, INFINITE);
							CloseHandle(hEventTerminateWatchThread);
							CloseHandle(hWatchThread);

							//Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
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
									Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
									iMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
								}
								goto break_C82EA5C2;
							default:
								break;
							}
						break_C82EA5C2:
							break;
						continue_C82EA5C2:

							//Terminate proxy process
							COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
							FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
							WaitForSingleObject(piProxyProcess.hProcess, INFINITE);
							CloseHandle(piProxyProcess.hThread);
							CloseHandle(piProxyProcess.hProcess);
							piProxyProcess.hProcess = NULL;

							//Terminate message thread
							SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
							WaitForSingleObject(hMessageThread, INFINITE);
							DWORD dwMessageThreadExitCode{};
							GetExitCodeThread(hMessageThread, &dwMessageThreadExitCode);
							if (dwMessageThreadExitCode)
							{
								MessageBox(NULL, L"Message thread exited abnormally.", L"FontLoaderEx", MB_ICONERROR);
							}
							CloseHandle(hMessageThread);

							//Close handle to target process
							CloseHandle(TargetProcessInfo.hProcess);
							TargetProcessInfo.hProcess = NULL;
						}

						//Else DIY
						if (TargetProcessInfo.hProcess)
						{
							//Terminate watch thread
							SetEvent(hEventTerminateWatchThread);
							WaitForSingleObject(hWatchThread, INFINITE);
							CloseHandle(hEventTerminateWatchThread);
							CloseHandle(hWatchThread);

							//Unload FontLoaderExInjectionDll(64).dll from target process
							if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
							{
								Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
								iMessageLength = Edit_GetTextLength(hWndEditMessage);
								Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
								Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
								break;
							}

							//Close handle to target process
							CloseHandle(TargetProcessInfo.hProcess);
							TargetProcessInfo.hProcess = NULL;
						}

						DestroyWindow(hWnd);
					}
					break;
				case IDNO:
					{
						EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
						EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
						EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
						EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
						EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
						EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), TRUE);
						EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
						EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
					}
					break;
				default:
					break;
				}
			}
		}
		break;
	case USERMESSAGE::BUTTONCLOSEWORKINGTHREADTERMINATED:
		{
			if (FontList.empty())
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), TRUE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
				EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
				EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);
			}
		}
		break;
	case USERMESSAGE::DRAGDROPWORKINGTHREADTERMINATED:
	case USERMESSAGE::BUTTONLOADWORKINGTHREADTERMINATED:
	case USERMESSAGE::BUTTONUNLOADWORKINGTHREADTERMINATED:
		{
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
			EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
		}
		break;
	case USERMESSAGE::WATCHTHREADTERMINATED:
		{
			EnableWindow(GetDlgItem(hWndMain, (int)ID::EditTimeout), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
		}
		break;
	default:
		break;
	}
	switch (Msg)
	{
	case WM_CREATE:
		{
			RECT rectClientMain{};
			GetClientRect(hWnd, &rectClientMain);

			NONCLIENTMETRICS ncm{ sizeof(NONCLIENTMETRICS) };
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
			HFONT hFont{ CreateFontIndirect(&ncm.lfMessageFont) };

			//Initialize ButtonOpen
			HWND hWndButtonOpen{ CreateWindow(WC_BUTTON, L"Open", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 50, 50, hWnd, (HMENU)ID::ButtonOpen, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonOpen, hFont, TRUE);

			//Initialize ButtonClose
			HWND hWndButtonClose{ CreateWindow(WC_BUTTON, L"Close", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, 50, 0, 50, 50, hWnd, (HMENU)ID::ButtonClose, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonClose, hFont, TRUE);

			//Initialize ButtonLoad
			HWND hWndButtonLoad{ CreateWindow(WC_BUTTON, L"Load", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, 100, 0, 50, 50, hWnd, (HMENU)ID::ButtonLoad, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonLoad, hFont, TRUE);

			//Initialize ButtonUnload
			HWND hWndButtonUnload{ CreateWindow(WC_BUTTON, L"Unload", WS_CHILD | WS_VISIBLE | WS_DISABLED | WS_TABSTOP | BS_PUSHBUTTON, 150, 0, 50, 50, hWnd, (HMENU)ID::ButtonUnload, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonUnload, hFont, TRUE);

			//Initialize ButtonBroadcastWM_FONTCHANGE
			HWND hWndButtonBroadcastWM_FONTCHANGE{ CreateWindow(WC_BUTTON, L"Broadcast WM_FONTCHANGE", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_AUTOCHECKBOX, 200, 0, 250, 21, hWnd, (HMENU)ID::ButtonBroadcastWM_FONTCHANGE, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonBroadcastWM_FONTCHANGE, hFont, TRUE);

			//Initialize EditTimeout and label
			HWND hWndStaticTimeout{ CreateWindow(WC_STATIC, L"Timeout:", WS_CHILD | WS_VISIBLE | SS_LEFT , 470, 1, 50, 19, hWnd, (HMENU)ID::StaticTimeout, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			HWND hWndEditTimeout{ CreateWindow(WC_EDIT, L"5000", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | ES_LEFT | ES_NUMBER | ES_AUTOHSCROLL | ES_NOHIDESEL, 520, 0, 70, 21, hWnd, (HMENU)ID::EditTimeout, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndStaticTimeout, hFont, TRUE);
			SetWindowFont(hWndEditTimeout, hFont, TRUE);

			dwTimeout = 5000;

			//Initialize ButtonSelectProcess
			HWND hWndButtonSelectProcess{ CreateWindow(WC_BUTTON, L"Select process", WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON, 200, 29, 250, 21, hWnd, (HMENU)ID::ButtonSelectProcess, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndButtonSelectProcess, hFont, TRUE);

			//Initialize ListViewFontList
			HWND hWndListViewFontList{ CreateWindow(WC_LISTVIEW, L"FontList", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_TABSTOP | LVS_REPORT | LVS_SHOWSELALWAYS, 0, 50, rectClientMain.right - rectClientMain.left, 300, hWnd, (HMENU)ID::ListViewFontList, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			ListView_SetExtendedListViewStyle(hWndListViewFontList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
			SetWindowFont(hWndListViewFontList, hFont, TRUE);

			RECT rectListViewFontListClient{};
			GetClientRect(hWndListViewFontList, &rectListViewFontListClient);
			LVCOLUMN lvc1{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectListViewFontListClient.right - rectListViewFontListClient.left) * 4 / 5 , (LPWSTR)L"Font Name" };
			LVCOLUMN lvc2{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectListViewFontListClient.right - rectListViewFontListClient.left) * 1 / 5 , (LPWSTR)L"State" };
			ListView_InsertColumn(hWndListViewFontList, 0, &lvc1);
			ListView_InsertColumn(hWndListViewFontList, 1, &lvc2);

			DragAcceptFiles(hWndListViewFontList, TRUE);

			OldListViewProc = (WNDPROC)SetWindowLongPtr(hWndListViewFontList, GWLP_WNDPROC, (LONG_PTR)ListViewProc);

			CHANGEFILTERSTRUCT cfs{ sizeof(CHANGEFILTERSTRUCT) };
			ChangeWindowMessageFilterEx(hWndListViewFontList, WM_DROPFILES, MSGFLT_ALLOW, &cfs);
			ChangeWindowMessageFilterEx(hWndListViewFontList, WM_COPYDATA, MSGFLT_ALLOW, &cfs);
			ChangeWindowMessageFilterEx(hWndListViewFontList, 0x0049, MSGFLT_ALLOW, &cfs);	//WM_COPYGLOBALDATA

			//Initialize StaticSplitter
			InitSplitter();
			HWND hWndSplitter{ CreateWindow(UC_SPLITTER, NULL, WS_CHILD | WS_VISIBLE, 0, 350, rectClientMain.right - rectClientMain.left, 5, hWnd, (HMENU)ID::Splitter, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };

			//Initialize EditMessage
			HWND hWndEditMessage{ CreateWindow(WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_READONLY | ES_LEFT | ES_MULTILINE | ES_NOHIDESEL, 0, 355, rectClientMain.right - rectClientMain.left, rectClientMain.bottom - rectClientMain.top - 355, hWnd, (HMENU)ID::EditMessage, ((LPCREATESTRUCT)lParam)->hInstance, NULL) };
			SetWindowFont(hWndEditMessage, hFont, TRUE);
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
				R"("Timeout": The time in milliseconds FontLoaderEx waits before reporting failure while injecting dll into target process via proxy process, the default value is 5000.)""\r\n"
				R"("Font Name": Names of the fonts added to the list view.)""\r\n"
				R"("State": State of the font. There are five states, "Not loaded", "Loaded", "Load failed", "Unloaded" and "Unload failed".)""\r\n"
				"\r\n"
			);
		}
		break;
	case WM_ACTIVATE:
		{
			//Process drag-drop font files onto the application icon stage II
			if (bDragDropHasFonts)
			{
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
				EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
				EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

				_beginthread(DragDropWorkingThreadProc, 0, nullptr);
			}
		}
		break;
	case WM_CLOSE:
		{
			//Unload all fonts
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

				_beginthread(CloseWorkingThreadProc, 0, nullptr);
			}
			else
			{
				HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

				std::wstringstream Message{};
				int iMessageLength{};

				//If proxy process is running
				if (piProxyProcess.hProcess)
				{
					//Terminate watch thread
					SetEvent(hEventTerminateWatchThread);
					WaitForSingleObject(hWatchThread, INFINITE);
					CloseHandle(hEventTerminateWatchThread);
					CloseHandle(hWatchThread);

					//Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
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
							Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
						}
						goto break_69504405;
					default:
						break;
					}
				break_69504405:
					break;
				continue_69504405:

					//Terminate proxy process
					COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
					FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
					WaitForSingleObject(piProxyProcess.hProcess, INFINITE);
					CloseHandle(piProxyProcess.hThread);
					CloseHandle(piProxyProcess.hProcess);
					piProxyProcess.hProcess = NULL;

					//Terminate message thread
					SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
					WaitForSingleObject(hMessageThread, INFINITE);
					DWORD dwMessageThreadExitCode{};
					GetExitCodeThread(hMessageThread, &dwMessageThreadExitCode);
					if (dwMessageThreadExitCode)
					{
						MessageBox(NULL, L"Message thread exited abnormally.", L"FontLoaderEx", MB_ICONERROR);
					}
					CloseHandle(hMessageThread);

					//Close handle to target process
					CloseHandle(TargetProcessInfo.hProcess);
					TargetProcessInfo.hProcess = NULL;
				}
				//Else DIY
				if (TargetProcessInfo.hProcess)
				{
					//Terminate watch thread
					SetEvent(hEventTerminateWatchThread);
					WaitForSingleObject(hWatchThread, INFINITE);
					CloseHandle(hEventTerminateWatchThread);
					CloseHandle(hWatchThread);

					//Unload FontLoaderExInjectionDll(64).dll from target process
					if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
					{
						Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
						iMessageLength = Edit_GetTextLength(hWndEditMessage);
						Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
						Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
						break;
					}

					//Close handle to target process
					CloseHandle(TargetProcessInfo.hProcess);
					TargetProcessInfo.hProcess = NULL;
				}

				ret = DefWindowProc(hWnd, Msg, wParam, lParam);
			}
		}
		break;
	case WM_DESTROY:
		{
			PostQuitMessage(0);
		}
		break;
	case WM_COMMAND:
		{
			switch ((ID)LOWORD(wParam))
			{
				//"Open" Button
			case ID::ButtonOpen:
				{
					switch (HIWORD(wParam))
					{
						//Open dialog and add fonts to list view
					case BN_CLICKED:
						{
							HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
							HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

							WCHAR szOpenFileNames[32768]{};
							OPENFILENAME ofn{ sizeof(ofn), hWnd, NULL, L"Font Files(*.ttf;*.ttc;*.otf)\0*.ttf;*.ttc;*.otf\0", NULL, 0, 0, szOpenFileNames, sizeof(szOpenFileNames), NULL, 0, NULL, NULL, OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT };
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
					default:
						break;
					}
				}
				break;
				//"Close" button
			case ID::ButtonClose:
				{
					switch (HIWORD(wParam))
					{
						//Unload and close selected fonts
						//Won't close those failed to unload
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

							_beginthread(ButtonCloseWorkingThreadProc, 0, nullptr);
						}
						break;
					default:
						break;
					}
				}
				break;
				//"Load" button
			case ID::ButtonLoad:
				{
					switch (HIWORD(wParam))
					{
						//Load selected fonts
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

							_beginthread(ButtonLoadWorkingThreadProc, 0, nullptr);
						}
					default:
						break;
					}
				}
				break;
				//"Unload" button
			case ID::ButtonUnload:
				{
					switch (HIWORD(wParam))
					{
						//Unload selected fonts
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

							_beginthread(ButtonUnloadWorkingThreadProc, 0, nullptr);
						}
					default:
						break;
					}
				}
				break;
				//"Select process" button
			case ID::ButtonSelectProcess:
				{
					switch (HIWORD(wParam))
					{
						//Select a process to load fonts to
					case BN_CLICKED:
						{
							HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
							HWND hWndButtonSelectProcess{ GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess) };

							static bool bIsSeDebugPrivilegeEnabled{ false };

							std::wstringstream Message{};
							int iMessageLength{};

							//Convert text in EditTimeout to integer
							BOOL bIsConverted{};
							DWORD dwTimeoutTemp{ (DWORD)GetDlgItemInt(hWnd, (int)ID::EditTimeout, &bIsConverted, FALSE) };
							if (!bIsConverted)
							{
								MessageBox(hWnd, L"Timeout value invalid.", L"FontLoaderEx", MB_ICONEXCLAMATION);

								SetDlgItemInt(hWnd, (int)ID::EditTimeout, (UINT)dwTimeout, FALSE);
								break;
							}
							else
							{
								dwTimeout = dwTimeoutTemp;
							}

							//Enable SeDebugPrivilege
							if (!bIsSeDebugPrivilegeEnabled)
							{
								if (!EnableDebugPrivilege())
								{
									MessageBox(NULL, L"Failed to enable SeDebugPrivilige for FontLoaderEx.", L"FontLoaderEx", MB_ICONERROR);
									break;
								}
								bIsSeDebugPrivilegeEnabled = true;
							}

							//If some fonts are loaded to target process, prompt user to close all first
							if ((!FontList.empty()) && TargetProcessInfo.hProcess)
							{
								MessageBox(hWnd, L"Please close all fonts before selecting process.", L"FontLoaderEx", MB_ICONEXCLAMATION);
								break;
							}

							//Select process
							ProcessInfo* pi{ (ProcessInfo*)DialogBox(NULL, MAKEINTRESOURCE(IDD_DIALOG1), hWnd, DialogProc) };

							//If p != nullptr, select it
							if (pi)
							{
								ProcessInfo SelectedProcessInfo = *pi;
								delete pi;

								//Clear selected process
								//If loaded via proxy
								if (piProxyProcess.hProcess)
								{
									//Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hWatchThread, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hWatchThread);

									//Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
									COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::PULLDLL, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
									WaitForSingleObject(hEventProxyDllPullingFinished, INFINITE);
									CloseHandle(hEventProxyDllPullingFinished);
									switch (ProxyDllPullingResult)
									{
									case PROXYDLLPULL::SUCCESSFUL:
										{
											Message << szInjectionDllNameByProxy << L" successfully unloaded from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
											iMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
											Message.str(L"");
										}
										goto continue_B9A25A68;
									case PROXYDLLPULL::FAILED:
										{
											Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
											iMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										}
										goto break_B9A25A68;
									default:
										break;
									}
								break_B9A25A68:
									break;
								continue_B9A25A68:

									//Terminate proxy process
									COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
									WaitForSingleObject(piProxyProcess.hProcess, INFINITE);
									Message << L"FontLoaderExProxy(" << piProxyProcess.dwProcessId << L") successfully terminated.\r\n\r\n";
									iMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									CloseHandle(piProxyProcess.hThread);
									CloseHandle(piProxyProcess.hProcess);
									piProxyProcess.hProcess = NULL;

									//Terminate message thread
									SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
									WaitForSingleObject(hMessageThread, INFINITE);
									DWORD dwMessageThreadExitCode{};
									GetExitCodeThread(hMessageThread, &dwMessageThreadExitCode);
									if (dwMessageThreadExitCode)
									{
										MessageBox(NULL, L"Message thread exited abnormally.", L"FontLoaderEx", MB_ICONERROR);
									}
									CloseHandle(hMessageThread);

									//Close handle to target process and synchronization objects
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
									CloseHandle(hEventProxyAddFontFinished);
									CloseHandle(hEventProxyRemoveFontFinished);
								}

								//Else DIY
								if (TargetProcessInfo.hProcess)
								{
									//Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hWatchThread, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hWatchThread);

									//Unload FontLoaderExInjectionDll(64).dll from target process
									if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									else
									{
										Message << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										Message.str(L"");
									}

									//Close handle to target process
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
								}

								//Get handle to target process
								SelectedProcessInfo.hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, SelectedProcessInfo.ProcessID);
								if (!SelectedProcessInfo.hProcess)
								{
									Message << L"Failed to open process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
									iMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									break;
								}

								//Determine current and target process type
								BOOL bIsWOW64Target{}, bIsWOW64Current{};
								IsWow64Process(SelectedProcessInfo.hProcess, &bIsWOW64Target);
								IsWow64Process(GetCurrentProcess(), &bIsWOW64Current);

								//If process types are different, launch FontLoaderExProxy.exe to inject dll
								if (bIsWOW64Current != bIsWOW64Target)
								{
									//Create synchronization objects and message thread
									SECURITY_ATTRIBUTES sa{ sizeof(SECURITY_ATTRIBUTES), NULL, TRUE };
									hEventParentProcessRunning = CreateEvent(NULL, TRUE, FALSE, L"FontLoaderEx_EventParentProcessRunning");
									hEventMessageThreadNotReady = CreateEvent(&sa, TRUE, FALSE, NULL);
									hEventMessageThreadReady = CreateEvent(&sa, TRUE, FALSE, NULL);
									hEventProxyProcessReady = CreateEvent(&sa, TRUE, FALSE, NULL);
									hEventProxyProcessDebugPrivilegeEnablingFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
									hEventProxyProcessHWNDRevieved = CreateEvent(NULL, TRUE, FALSE, NULL);
									hEventProxyDllInjectionFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
									hEventProxyDllPullingFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
									hMessageThread = (HANDLE)_beginthreadex(nullptr, 0, MessageThreadProc, (void*)GetWindowLongPtr(hWnd, GWLP_HINSTANCE), 0, nullptr);
									HANDLE handles[]{ hEventMessageThreadNotReady, hEventMessageThreadReady };
									switch (WaitForMultipleObjects(2, handles, FALSE, INFINITE))
									{
									case WAIT_OBJECT_0:
										{
											MessageBox(NULL, L"Failed to create message-only window.", L"FontLoaderEx", MB_ICONERROR);

											WaitForSingleObject(hMessageThread, INFINITE);
											CloseHandle(hMessageThread);
										}
										goto break_721EFBC1;
									case WAIT_OBJECT_0 + 1:
										goto continue_721EFBC1;
									default:
										break;
									}
								break_721EFBC1:
									break;
								continue_721EFBC1:

									//Run proxy process, send handle to current process and target process, HWND to message window, handle to synchronization objects and timeout as arguments to proxy process
									HANDLE hCurrentProcessDuplicated{};
									HANDLE hTargetProcessDuplicated{};
									DuplicateHandle(GetCurrentProcess(), GetCurrentProcess(), GetCurrentProcess(), &hCurrentProcessDuplicated, 0, TRUE, DUPLICATE_SAME_ACCESS);
									DuplicateHandle(GetCurrentProcess(), SelectedProcessInfo.hProcess, GetCurrentProcess(), &hTargetProcessDuplicated, 0, TRUE, DUPLICATE_SAME_ACCESS);
									std::wstringstream strParams{};
									strParams << (UINT_PTR)hCurrentProcessDuplicated << L" " << (UINT_PTR)hTargetProcessDuplicated << L" " << (UINT_PTR)hWndMessage << L" " << (UINT_PTR)hEventMessageThreadReady << L" " << (UINT_PTR)hEventProxyProcessReady << L" " << dwTimeout;
									std::size_t nParamLength{ strParams.str().length() };
									std::unique_ptr<WCHAR[]> lpszParams{ new WCHAR[nParamLength + 1]{} };
									wcsncpy_s(lpszParams.get(), nParamLength + 1, strParams.str().c_str(), nParamLength);
									STARTUPINFO si{ sizeof(STARTUPINFO) };
#ifdef _DEBUG
#ifdef _WIN64
									if (!CreateProcess(LR"(..\Debug\FontLoaderExProxy.exe)", lpszParams.get(), NULL, NULL, TRUE, 0, NULL, NULL, &si, &piProxyProcess))
#else
									if (!CreateProcess(LR"(..\x64\Debug\FontLoaderExProxy.exe)", lpszParams.get(), NULL, NULL, TRUE, 0, NULL, NULL, &si, &piProxyProcess))
#endif
#else
									if (!CreateProcess(LR"(FontLoaderExProxy.exe)", lpszParams.get(), NULL, NULL, TRUE, 0, NULL, NULL, &si, &piProxyProcess))
#endif
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hCurrentProcessDuplicated);
										CloseHandle(hTargetProcessDuplicated);

										Message << L"Failed to launch FontLoaderExProxy." << L"\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}

									//Wait for proxy process to ready
									WaitForSingleObject(hEventProxyProcessReady, INFINITE);
									CloseHandle(hEventProxyProcessReady);
									CloseHandle(hEventMessageThreadReady);
									CloseHandle(hEventParentProcessRunning);

									Message << L"FontLoaderExProxy(" << piProxyProcess.dwProcessId << L") succesfully launched.\r\n\r\n";
									iMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									Message.str(L"");

									//Wait for message-only window to recieve HWND to proxy process
									WaitForSingleObject(hEventProxyProcessHWNDRevieved, INFINITE);
									CloseHandle(hEventProxyProcessHWNDRevieved);

									//Wait for proxy process to enable SeDebugPrivilege
									WaitForSingleObject(hEventProxyProcessDebugPrivilegeEnablingFinished, INFINITE);
									CloseHandle(hEventProxyProcessDebugPrivilegeEnablingFinished);
									switch (ProxyDebugPrivilegeEnablingResult)
									{
									case PROXYPROCESSDEBUGPRIVILEGEENABLING::SUCCESSFUL:
										goto continue_90567013;
									case PROXYPROCESSDEBUGPRIVILEGEENABLING::FAILED:
										{
											MessageBox(NULL, L"Failed to enable SeDebugPrivilige for FontLoaderExProxy.", L"FontLoaderEx", MB_ICONERROR);

											//Terminate proxy process
											COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
											FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
											WaitForSingleObject(piProxyProcess.hProcess, INFINITE);

											Message << L"FontLoaderExProxy(" << piProxyProcess.dwProcessId << L") successfully terminated.\r\n\r\n";
											iMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

											CloseHandle(piProxyProcess.hThread);
											CloseHandle(piProxyProcess.hProcess);
											piProxyProcess.hProcess = NULL;

											//Terminate message thread
											SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
											WaitForSingleObject(hMessageThread, INFINITE);
											DWORD dwMessageThreadExitCode{};
											GetExitCodeThread(hMessageThread, &dwMessageThreadExitCode);
											if (dwMessageThreadExitCode)
											{
												MessageBox(NULL, L"Message thread exited abnormally.", L"FontLoaderEx", MB_ICONERROR);
											}
											CloseHandle(hMessageThread);

											//Close handles to selected process
											CloseHandle(SelectedProcessInfo.hProcess);
										}
										goto break_90567013;
									default:
										break;
									}
								break_90567013:
									break;
								continue_90567013:

									//Begin dll injection
									COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::INJECTDLL, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);

									//Wait for proxy process to inject dll into target process
									WaitForSingleObject(hEventProxyDllInjectionFinished, INFINITE);
									CloseHandle(hEventProxyDllInjectionFinished);
									switch (ProxyDllInjectionResult)
									{
									case PROXYDLLINJECTION::SUCCESSFUL:
										{
											Message << szInjectionDllNameByProxy << L" successfully injected into target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
											iMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

											//Register proxy AddFont() and RemoveFont() procedure and create synchronization object
											FontResource::RegisterAddRemoveFontProc(ProxyAddFontProc, ProxyRemoveFontProc);
											hEventProxyAddFontFinished = CreateEvent(NULL, TRUE, FALSE, NULL);
											hEventProxyRemoveFontFinished = CreateEvent(NULL, TRUE, FALSE, NULL);

											//Disable EditTimeout and ButtonBroadcastWM_FONTCHANGE
											EnableWindow(GetDlgItem(hWndMain, (int)ID::EditTimeout), FALSE);
											EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);

											//Change caption
											std::wstringstream Caption{};
											Caption << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L")";
											Button_SetText(hWndButtonSelectProcess, (LPCWSTR)Caption.str().c_str());

											//Set TargetProcessInfo
											TargetProcessInfo = SelectedProcessInfo;

											//Start watch thread and create synchronization object
											hEventTerminateWatchThread = CreateEvent(NULL, TRUE, FALSE, NULL);
											hWatchThread = (HANDLE)_beginthreadex(nullptr, 0, ProxyAndTargetProcessWatchThreadProc, nullptr, 0, nullptr);
										}
										break;
									case PROXYDLLINJECTION::FAILED:
										{
											Message << L"Failed to inject " << szInjectionDllNameByProxy << L" into target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::FAILEDTOENUMERATEMODULES:
										{
											Message << L"Failed to enumerate modules in target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::GDI32NOTLOADED:
										{
											Message << L"gdi32.dll not loaded by target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									case PROXYDLLINJECTION::MODULENOTFOUND:
										{
											Message << L"Failed to enumerate modules in target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
										}
										goto continue_DBEA36FE;
									continue_DBEA36FE:
										{
											iMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

											//Terminate proxy process
											COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
											FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
											WaitForSingleObject(piProxyProcess.hProcess, INFINITE);

											Message << L"FontLoaderExProxy(" << piProxyProcess.dwProcessId << L") successfully terminated.\r\n\r\n";
											iMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

											CloseHandle(piProxyProcess.hThread);
											CloseHandle(piProxyProcess.hProcess);
											piProxyProcess.hProcess = NULL;

											//Terminate message thread
											SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
											WaitForSingleObject(hMessageThread, INFINITE);
											DWORD dwMessageThreadExitCode{};
											GetExitCodeThread(hMessageThread, &dwMessageThreadExitCode);
											if (dwMessageThreadExitCode)
											{
												MessageBox(NULL, L"Message thread exited abnormally.", L"FontLoaderEx", MB_ICONERROR);
											}
											CloseHandle(hMessageThread);

											//Close handles to selected process
											CloseHandle(SelectedProcessInfo.hProcess);
										}
										break;
									default:
										break;
									}
								}
								//Else DIY
								else
								{
									//Check whether target process loads gdi32.dll as AddFontResourceEx() and RemoveFontResourceEx() are in it
									HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, SelectedProcessInfo.ProcessID) };
									MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
									bool bIsGDI32Loaded{ false };
									if (!Module32First(hModuleSnapshot, &me32))
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot);

										Message << L"Failed to enumerate modules in target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
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

										Message << L"gdi32.dll not loaded by target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									CloseHandle(hModuleSnapshot);

									//Inject FontLoaderExInjectionDll(64).dll into target process
									if (!InjectModule(SelectedProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										CloseHandle(SelectedProcessInfo.hProcess);

										Message << L"Failed to inject " << szInjectionDllName << L" into target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									Message << szInjectionDllName << L" successfully injected into target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
									iMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									Message.str(L"");

									//Get base address of FontLoaderExInjectionDll(64).dll in target process
									HANDLE hModuleSnapshot2{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, SelectedProcessInfo.ProcessID) };
									MODULEENTRY32 me322{ sizeof(MODULEENTRY32) };
									BYTE* pModBaseAddr{};
									if (!Module32First(hModuleSnapshot2, &me322))
									{
										CloseHandle(SelectedProcessInfo.hProcess);
										CloseHandle(hModuleSnapshot2);

										Message << L"Failed to enumerate modules in target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
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

										Message << szInjectionDllName << " not found in target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									CloseHandle(hModuleSnapshot2);

									//Calculate addresses of AddFont() and RemoveFont() in target process
									HMODULE hModInjectionDll{ LoadLibrary(szInjectionDllName) };
									void* pLocalAddFontProcAddr{ GetProcAddress(hModInjectionDll, "AddFont") };
									void* pLocalRemoveFontProcAddr{ GetProcAddress(hModInjectionDll, "RemoveFont") };
									FreeLibrary(hModInjectionDll);
									UINT_PTR AddFontProcOffset{ (UINT_PTR)pLocalAddFontProcAddr - (UINT_PTR)hModInjectionDll };
									UINT_PTR RemoveFontProcOffset{ (UINT_PTR)pLocalRemoveFontProcAddr - (UINT_PTR)hModInjectionDll };
									pfnRemoteAddFontProc = (void*)(pModBaseAddr + AddFontProcOffset);
									pfnRemoteRemoveFontProc = (void*)(pModBaseAddr + RemoveFontProcOffset);

									//Register remote AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(RemoteAddFontProc, RemoteRemoveFontProc);

									//Change caption
									std::wstringstream Caption{};
									Caption << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L")";
									Button_SetText(hWndButtonSelectProcess, (LPCWSTR)Caption.str().c_str());

									//Set TargetProcessInfo
									TargetProcessInfo = SelectedProcessInfo;

									//Start watch thread and create synchronization object
									hEventTerminateWatchThread = CreateEvent(NULL, TRUE, FALSE, NULL);
									hWatchThread = (HANDLE)_beginthreadex(nullptr, 0, TargetProcessWatchThreadProc, nullptr, 0, nullptr);
								}
							}
							//If p == nullptr, clear selected process
							else
							{
								//If loaded via proxy
								if (piProxyProcess.hProcess)
								{
									//Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hWatchThread, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hWatchThread);

									//Unload FontLoaderExInjectionDll(64).dll from target process via proxy process
									COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::PULLDLL, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds, SendMessage);
									WaitForSingleObject(hEventProxyDllPullingFinished, INFINITE);
									CloseHandle(hEventProxyDllPullingFinished);
									switch (ProxyDllPullingResult)
									{
									case PROXYDLLPULL::SUCCESSFUL:
										{
											Message << szInjectionDllNameByProxy << L" successfully unloaded from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
											iMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
											Message.str(L"");
										}
										goto continue_0F70B465;
									case PROXYDLLPULL::FAILED:
										{
											Message << L"Failed to unload " << szInjectionDllNameByProxy << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
											iMessageLength = Edit_GetTextLength(hWndEditMessage);
											Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
											Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										}
										goto break_0F70B465;
									default:
										break;
									}
								break_0F70B465:
									break;
								continue_0F70B465:

									//Terminate proxy process
									COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
									FORWARD_WM_COPYDATA(hWndProxy, hWnd, &cds2, SendMessage);
									WaitForSingleObject(piProxyProcess.hProcess, INFINITE);

									Message << L"FontLoaderExProxy(" << piProxyProcess.dwProcessId << L") successfully terminated.\r\n\r\n";
									iMessageLength = Edit_GetTextLength(hWndEditMessage);
									Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
									Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

									CloseHandle(piProxyProcess.hThread);
									CloseHandle(piProxyProcess.hProcess);
									piProxyProcess.hProcess = NULL;

									//Terminate message thread
									SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
									WaitForSingleObject(hMessageThread, INFINITE);
									DWORD dwMessageThreadExitCode{};
									GetExitCodeThread(hMessageThread, &dwMessageThreadExitCode);
									if (dwMessageThreadExitCode)
									{
										MessageBox(NULL, L"Message thread exited abnormally.", L"FontLoaderEx", MB_ICONERROR);
									}
									CloseHandle(hMessageThread);

									//Close handle to target process and synchronization objects
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;
									CloseHandle(hEventProxyAddFontFinished);
									CloseHandle(hEventProxyRemoveFontFinished);

									//Register default AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);

									//Enable EditTimeout and ButtonBroadcastWM_FONTCHANGE
									EnableWindow(GetDlgItem(hWndMain, (int)ID::EditTimeout), TRUE);
									EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);

									//Revert to default caption
									Button_SetText(hWndButtonSelectProcess, L"Select process");
								}

								//Else DIY
								if (TargetProcessInfo.hProcess)
								{
									//Terminate watch thread
									SetEvent(hEventTerminateWatchThread);
									WaitForSingleObject(hWatchThread, INFINITE);
									CloseHandle(hEventTerminateWatchThread);
									CloseHandle(hWatchThread);

									//Unload FontLoaderExInjectionDll(64).dll from target process
									if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, dwTimeout))
									{
										Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
										break;
									}
									else
									{
										Message << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
										iMessageLength = Edit_GetTextLength(hWndEditMessage);
										Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
										Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
									}

									//Close handle to target process
									CloseHandle(TargetProcessInfo.hProcess);
									TargetProcessInfo.hProcess = NULL;

									//Register default AddFont() and RemoveFont() procedure
									FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);

									//Revert to default caption
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
				//"Load" menu item
			case ID_MENU_LOAD:
				{
					//Simulate clicking "Load" button
					SendMessage(GetDlgItem(hWnd, (int)ID::ButtonLoad), BM_CLICK, NULL, NULL);
				}
				break;
				//"Unload" menu item
			case ID_MENU_UNLOAD:
				{
					//Simulate clicking "Unload" button
					SendMessage(GetDlgItem(hWnd, (int)ID::ButtonUnload), BM_CLICK, NULL, NULL);
				}
				break;
				//"Close" menu item
			case ID_MENU_CLOSE:
				{
					//Simulate clicking "Close" button
					SendMessage(GetDlgItem(hWnd, (int)ID::ButtonClose), BM_CLICK, NULL, NULL);
				}
				break;
				//"Select All" menu item
			case ID_MENU_SELECTALL:
				{
					//Select all items in ListViewFontList
					ListView_SetItemState(GetDlgItem(hWndMain, (int)ID::ListViewFontList), -1, LVIS_SELECTED, LVIS_SELECTED);
				}
				//"Ctrl+A" accelerator
			case ID_ACCELERATOR_SELECTALL:
				{
					//Select all items in ListViewFontList
					if (GetFocus() == GetDlgItem(hWndMain, (int)ID::ListViewFontList))
					{
						ListView_SetItemState(GetDlgItem(hWndMain, (int)ID::ListViewFontList), -1, LVIS_SELECTED, LVIS_SELECTED);
					}
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_NOTIFY:
		{
			switch ((ID)wParam)
			{
			case ID::Splitter:
				{
					static LONG CursorOffsetY{};

					switch ((SPLITTERNOTIFICATION)((LPNMHDR)lParam)->code)
					{
						//Begin dragging Splitter
					case SPLITTERNOTIFICATION::DRAGBEGIN:
						{
							//Caculate cursor offset to window top in y orientation
							HWND hWndSplitter{ ((LPSPLITTERSTRUCT)lParam)->nmhdr.hwndFrom };
							RECT rectSplitter{}, rectSplitterClient{};
							GetWindowRect(hWndSplitter, &rectSplitter);
							GetClientRect(hWndSplitter, &rectSplitterClient);
							CursorOffsetY = ((LPSPLITTERSTRUCT)lParam)->ptCursor.y - rectSplitterClient.top;
							MapWindowRect(hWndSplitter, HWND_DESKTOP, &rectSplitterClient);
							CursorOffsetY += rectSplitterClient.top - rectSplitter.top;

							//Confine cursor to a specific rectangle
							RECT rectMainClient{}, rectDesktop{}, rectButtonOpen{};
							GetClientRect(hWnd, &rectMainClient);
							MapWindowRect(hWnd, HWND_DESKTOP, &rectMainClient);
							GetWindowRect(GetDesktopWindow(), &rectDesktop);
							GetWindowRect(GetDlgItem(hWnd, (int)ID::ButtonOpen), &rectButtonOpen);
							RECT rectMouseClip{ rectDesktop.left, rectMainClient.top + (rectButtonOpen.bottom - rectButtonOpen.top) + CursorOffsetY, rectDesktop.right, rectMainClient.bottom - ((rectSplitter.bottom - rectSplitter.top) - CursorOffsetY) };
							ClipCursor(&rectMouseClip);
						}
						break;
						//Dragging Splitter
					case SPLITTERNOTIFICATION::DRAGGING:
						{
							//Convert cursor coordinates from Splitter to main window
							POINT ptCursor{ ((LPSPLITTERSTRUCT)lParam)->ptCursor };
							MapWindowPoints(((LPSPLITTERSTRUCT)lParam)->nmhdr.hwndFrom, hWnd, &ptCursor, 1);

							//Move Splitter
							HWND hWndSplitter{ ((LPSPLITTERSTRUCT)lParam)->nmhdr.hwndFrom };
							RECT rectSplitter{};
							GetWindowRect(hWndSplitter, &rectSplitter);
							MapWindowRect(HWND_DESKTOP, hWnd, &rectSplitter);
							MoveWindow(hWndSplitter, 0, ptCursor.y - CursorOffsetY, rectSplitter.right - rectSplitter.left, rectSplitter.bottom - rectSplitter.top, TRUE);

							//Resize ListViewFontList
							HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
							RECT rectListViewFontList{};
							GetWindowRect(hWndListViewFontList, &rectListViewFontList);
							MapWindowRect(HWND_DESKTOP, hWnd, &rectListViewFontList);
							MoveWindow(hWndListViewFontList, 0, rectListViewFontList.top, rectListViewFontList.right - rectListViewFontList.left, ptCursor.y - CursorOffsetY - rectListViewFontList.top, TRUE);

							//Resize EditMessage
							HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
							RECT rectEditMessage{}, rectMainClient{};
							GetWindowRect(hWndEditMessage, &rectEditMessage);
							MapWindowRect(HWND_DESKTOP, hWnd, &rectEditMessage);
							GetClientRect(hWnd, &rectMainClient);
							MoveWindow(hWndEditMessage, 0, ptCursor.y + (rectSplitter.bottom - rectSplitter.top) - CursorOffsetY, rectEditMessage.right - rectEditMessage.left, rectMainClient.bottom - rectSplitter.bottom, TRUE);
						}
						break;
						//End dragging Splitter
					case SPLITTERNOTIFICATION::DRAGEND:
						{
							//Free cursor
							ClipCursor(NULL);

							//Redraw controls
							RECT rectClientMain{};
							GetClientRect(hWnd, &rectClientMain);
							InvalidateRect(hWnd, &rectClientMain, FALSE);
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
			//Show context menu on ListViewFontList
			if ((HWND)wParam == GetDlgItem(hWnd, (int)ID::ListViewFontList))
			{
				POINT ptMenu{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
				if (ptMenu.x == -1 && ptMenu.y == -1)
				{
					HWND hWndListViewFontList{ GetDlgItem(hWnd,(int)ID::ListViewFontList) };

					int iSelectionMark{ ListView_GetSelectionMark(hWndListViewFontList) };
					if (iSelectionMark == -1)
					{
						RECT rectListViewFontListClient{};
						GetClientRect(hWndListViewFontList, &rectListViewFontListClient);
						MapWindowRect(hWndListViewFontList, HWND_DESKTOP, &rectListViewFontListClient);
						ptMenu = { rectListViewFontListClient.left, rectListViewFontListClient.top };
					}
					else
					{
						POINT ptSelectionMark{};
						ListView_GetItemPosition(hWndListViewFontList, iSelectionMark, &ptSelectionMark);
						ClientToScreen(hWndListViewFontList, &ptSelectionMark);
						ptMenu = ptSelectionMark;
					}
				}

				TrackPopupMenu(hMenuContextListViewFontList, TPM_LEFTALIGN | TPM_RIGHTBUTTON, ptMenu.x, ptMenu.y, 0, hWnd, NULL);
			}
		}
		break;
	case WM_SIZE:
		{
			switch (wParam)
			{
				//Resize controls
			case SIZE_MAXIMIZED:
			case SIZE_RESTORED:
				{
					//Resize Splitter
					HWND hWndSplitter{ GetDlgItem(hWnd, (int)ID::Splitter) };
					RECT rectSplitter{};
					GetWindowRect(hWndSplitter, &rectSplitter);
					MapWindowRect(HWND_DESKTOP, hWnd, &rectSplitter);
					MoveWindow(hWndSplitter, rectSplitter.left, rectSplitter.top, LOWORD(lParam), rectSplitter.bottom - rectSplitter.top, TRUE);

					//Resize ListViewFontList
					HWND hWndListViewFontList{ GetDlgItem(hWnd, (int)ID::ListViewFontList) };
					RECT rectListViewFontList{};
					GetWindowRect(hWndListViewFontList, &rectListViewFontList);
					MapWindowRect(HWND_DESKTOP, hWnd, &rectListViewFontList);
					MoveWindow(hWndListViewFontList, rectListViewFontList.left, rectListViewFontList.top, LOWORD(lParam), rectListViewFontList.bottom - rectListViewFontList.top, TRUE);

					RECT rectListViewFontListClient{};
					GetClientRect(hWndListViewFontList, &rectListViewFontListClient);
					ListView_SetColumnWidth(hWndListViewFontList, 0, rectListViewFontListClient.right - rectListViewFontListClient.left - ListView_GetColumnWidth(hWndListViewFontList, 1));

					//Resize EditMessage
					HWND hWndEditMessage{ GetDlgItem(hWnd, (int)ID::EditMessage) };
					RECT rectEditMessage{};
					GetWindowRect(hWndEditMessage, &rectEditMessage);
					MapWindowRect(HWND_DESKTOP, hWnd, &rectEditMessage);
					MoveWindow(hWndEditMessage, rectEditMessage.left, rectEditMessage.top, LOWORD(lParam), HIWORD(lParam) - rectSplitter.bottom, TRUE);
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_GETMINMAXINFO:
		{
			//Limit minimize window size
			((LPMINMAXINFO)lParam)->ptMinTrackSize = { 700, 700 };
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

LRESULT CALLBACK ListViewProc(HWND hWndListView, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	switch (Msg)
	{
	case WM_DROPFILES:
		{
			//Process drag-drop and open fonts
			HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
			HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

			std::wstringstream Message{};
			int iMessageLength{};

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
					iMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
					Message.str(L"");
				}
			}
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, L"\r\n");

			DragFinish((HDROP)wParam);

			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
			EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
			EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_ENABLED);
			EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_ENABLED);
			EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_ENABLED);
			EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_ENABLED);
		}
		break;
	default:
		{
			ret = CallWindowProc(OldListViewProc, hWndListView, Msg, wParam, lParam);
		}
		break;
	}

	return ret;
}

INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	INT_PTR ret{};

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
			HFONT hFont{ CreateFontIndirect(&ncm.lfMessageFont) };

			//Initialize ListViewProcessList
			HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };
			SetWindowFont(hWndListViewProcessList, hFont, TRUE);
			SetWindowLongPtr(hWndListViewProcessList, GWL_STYLE, GetWindowLongPtr(hWndListViewProcessList, GWL_STYLE) | LVS_REPORT | LVS_SINGLESEL);

			RECT rectListViewProcessListClient{};
			GetClientRect(hWndListViewProcessList, &rectListViewProcessListClient);
			LVCOLUMN lvc1{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectListViewProcessListClient.right - rectListViewProcessListClient.left) * 4 / 5 , (LPWSTR)L"Process" };
			LVCOLUMN lvc2{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectListViewProcessListClient.right - rectListViewProcessListClient.left) * 1 / 5 , (LPWSTR)L"PID" };
			ListView_InsertColumn(hWndListViewProcessList, 0, &lvc1);
			ListView_InsertColumn(hWndListViewProcessList, 1, &lvc2);
			ListView_SetExtendedListViewStyle(hWndListViewProcessList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

			//Initialize ButtonOK
			HWND hWndButtonOK{ GetDlgItem(hWndDialog, IDOK) };
			SetWindowFont(hWndButtonOK, hFont, TRUE);

			//Initialize ButtonCancel
			HWND hWndButtonCancel{ GetDlgItem(hWndDialog, IDCANCEL) };
			SetWindowFont(hWndButtonCancel, hFont, TRUE);

			//Fill ProcessList
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

			ret = (INT_PTR)TRUE;
		}
		break;
	case WM_COMMAND:
		{
			switch (LOWORD(wParam))
			{
				//Return selected process
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
				//Return null
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
						//Sort items
					case LVN_COLUMNCLICK:
						{
							HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };
							HWND hWndHeader{ ListView_GetHeader(hWndListViewProcessList) };

							HDITEM hdi{ HDI_FORMAT };

							switch (((LPNMLISTVIEW)lParam)->iSubItem)
							{
							case 0:
								{
									//Sort items by Process
									bOrderByProcessAscending ?
										std::sort(ProcessList.begin(), ProcessList.end(),
											[](const ProcessInfo& value1, const ProcessInfo& value2) -> bool
											{
												int i{ lstrcmpi(value1.ProcessName.c_str(), value2.ProcessName.c_str()) };
												if (i < 0)
												{
													return true;
												}
												if (i > 0)
												{
													return false;
												}
												else
												{
													return false;
												}
											})
										:
										std::sort(ProcessList.begin(), ProcessList.end(),
											[](const ProcessInfo& value1, const ProcessInfo& value2) -> bool
											{
												int i{ lstrcmpi(value2.ProcessName.c_str(), value1.ProcessName.c_str()) };
												if (i < 0)
												{
													return true;
												}
												if (i >= 0)
												{
													return false;
												}
												else
												{
													return false;
												}
											})
										;

									//Add arrow to the header in list view
									Header_GetItem(hWndHeader, 1, &hdi);
									hdi.fmt = hdi.fmt & ~(HDF_SORTDOWN | HDF_SORTUP);
									Header_SetItem(hWndHeader, 1, &hdi);
									Header_GetItem(hWndHeader, 0, &hdi);
									bOrderByProcessAscending ? hdi.fmt = (hdi.fmt & ~HDF_SORTDOWN) | HDF_SORTUP : hdi.fmt = (hdi.fmt & ~HDF_SORTUP) | HDF_SORTDOWN;
									Header_SetItem(hWndHeader, 0, &hdi);

									bOrderByProcessAscending = !bOrderByProcessAscending;
								}
								break;
							case 1:
								{
									//Sort items by PID
									bOrderByPIDAscending ?
										std::sort(ProcessList.begin(), ProcessList.end(),
											[](const ProcessInfo& value1, const ProcessInfo& value2) -> bool
									{
										return value1.ProcessID < value2.ProcessID;
									})
										:
										std::sort(ProcessList.begin(), ProcessList.end(),
											[](const ProcessInfo& value1, const ProcessInfo& value2) -> bool
									{
										return value1.ProcessID > value2.ProcessID;
									})
										;

									//Add arrow to the header in list view
									Header_GetItem(hWndHeader, 0, &hdi);
									hdi.fmt = hdi.fmt & ~(HDF_SORTDOWN | HDF_SORTUP);
									Header_SetItem(hWndHeader, 0, &hdi);
									Header_GetItem(hWndHeader, 1, &hdi);
									bOrderByPIDAscending ? hdi.fmt = (hdi.fmt & ~HDF_SORTDOWN) | HDF_SORTUP : hdi.fmt = (hdi.fmt & ~HDF_SORTUP) | HDF_SORTDOWN;
									Header_SetItem(hWndHeader, 1, &hdi);

									bOrderByPIDAscending = !bOrderByPIDAscending;
								}
								break;
							default:
								break;
							}

							//Reset contents of list view
							LVITEM lvi{ LVIF_TEXT, 0 };
							for (auto& i : ProcessList)
							{
								lvi.iSubItem = 0;
								lvi.pszText = (LPWSTR)i.ProcessName.c_str();
								ListView_SetItem(hWndListViewProcessList, &lvi);
								lvi.iSubItem = 1;
								std::wstring strPID{ std::to_wstring(i.ProcessID) };
								lvi.pszText = (LPWSTR)strPID.c_str();
								ListView_SetItem(hWndListViewProcessList, &lvi);
								lvi.iItem++;
							}

							ret = (INT_PTR)TRUE;
						}
						break;
						//Select double-clicked item
					case NM_DBLCLK:
						{
							//Simulate clicking "OK" button
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
	default:
		break;
	}

	return ret;
}

bool EnableDebugPrivilege()
{
	bool bRet{};

	//Enable SeDebugPrivilege
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

		TOKEN_PRIVILEGES tkp{ 1 , {luid, SE_PRIVILEGE_ENABLED} };
		if (!AdjustTokenPrivileges(hToken, FALSE, &tkp, sizeof(tkp), NULL, NULL))
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

bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout)
{
	bool bRet{};

	//Inject dll into target process
	do
	{
		WCHAR szDllPath[MAX_PATH]{};
		GetModuleFileName(NULL, szDllPath, MAX_PATH);
		PathRemoveFileSpec(szDllPath);
		PathAppend(szDllPath, szModuleName);

		LPVOID lpRemoteBuffer{ VirtualAllocEx(hProcess, NULL, (std::wcslen(szDllPath) + 1) * sizeof(WCHAR), MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE) };
		if (!lpRemoteBuffer)
		{
			bRet = false;
			break;
		}

		if (!WriteProcessMemory(hProcess, lpRemoteBuffer, (LPVOID)szDllPath, (std::wcslen(szDllPath) + 1) * sizeof(WCHAR), NULL))
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;
			break;
		}

		HMODULE hModule{ GetModuleHandle(L"Kernel32") };
		if (!hModule)
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;
			break;
		}

		LPTHREAD_START_ROUTINE addr{ (LPTHREAD_START_ROUTINE)GetProcAddress(hModule, "LoadLibraryW") };
		HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, addr, lpRemoteBuffer, 0, NULL) };
		if (!hRemoteThread)
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;
			break;
		}

		if (WaitForSingleObject(hRemoteThread, Timeout) == WAIT_TIMEOUT)
		{
			CloseHandle(hRemoteThread);
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;
			break;
		}
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

		//Get exit code of remote thread
		DWORD dwRemoteThreadExitCode{};
		if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}

		if (!dwRemoteThreadExitCode)
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}
		CloseHandle(hRemoteThread);

		bRet = true;
	} while (false);

	return bRet;
}

bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout)
{
	bool bRet{};

	//Find HMODULE of szModuleName in target process
	do
	{
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

		//Unload FontLoaderExInjectionDll(64).dll from target process
		HMODULE hModule{ GetModuleHandle(L"Kernel32") };
		if (!hModule)
		{
			bRet = false;
			break;
		}

		LPTHREAD_START_ROUTINE addr{ (LPTHREAD_START_ROUTINE)GetProcAddress(hModule, "FreeLibrary") };
		HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, addr, (LPVOID)hModInjectionDll, 0, NULL) };
		if (!hRemoteThread)
		{
			bRet = false;
			break;
		}
		if (WaitForSingleObject(hRemoteThread, Timeout) == WAIT_TIMEOUT)
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}

		//Get exit code of remote thread
		DWORD dwRemoteThreadExitCode{};
		if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}
		if (!dwRemoteThreadExitCode)
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}
		CloseHandle(hRemoteThread);

		bRet = true;
	} while (false);

	return bRet;
}