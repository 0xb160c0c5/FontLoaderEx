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
#include <memory>
#include <algorithm>
#include <cctype>
#include "FontResource.h"
#include "Globals.h"
#include "resource.h"

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);
bool EnableDebugPrivilege();

std::list<FontResource> FontList{};
HWND hWndMainWindow{};
bool DragDropHasFonts{ false };

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
	//Register default AddFont() and RemoveFont() procedure
	FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);

	//Process drag-drop font files onto the application icon stage I
	int argc{};
	LPWSTR* argv = CommandLineToArgvW(GetCommandLine(), &argc);
	if (argc > 1)
	{
		DragDropHasFonts = true;
		for (int i = 1; i < argc; i++)
		{
			if (PathMatchSpec(argv[i], L"*.ttf") || PathMatchSpec(argv[i], L"*.ttc") || PathMatchSpec(argv[i], L"*.otf"))
			{
				FontList.push_back(argv[i]);
			}
		}
	}
	LocalFree(argv);

	InitCommonControls();

	//Create window
	MSG Msg{};
	WNDCLASS wndclass{ CS_HREDRAW | CS_VREDRAW, WndProc, 0, 0, hInstance, LoadIcon(NULL, IDI_APPLICATION), LoadCursor(NULL, IDC_ARROW), GetSysColorBrush(COLOR_WINDOW), NULL, L"FontLoaderEx" };

	if (!RegisterClass(&wndclass))
	{
		MessageBox(NULL, L"Failed to register window class.", L"FontLoaderEx", MB_ICONERROR);
		return -1;
	}

	if (!(hWndMainWindow = CreateWindow(L"FontLoaderEx", L"FontLoaderEx", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_BORDER | WS_MINIMIZEBOX, CW_USEDEFAULT, CW_USEDEFAULT, 700, 700, NULL, NULL, hInstance, NULL)))
	{
		MessageBox(NULL, L"Failed to create window.", L"FontLoaderEx", MB_ICONERROR);
		return -1;
	}

	ShowWindow(hWndMainWindow, nShowCmd);
	UpdateWindow(hWndMainWindow);

	BOOL bRet{};
	int ret{};
	while ((bRet = GetMessage(&Msg, NULL, 0, 0)))
	{
		if (bRet == 0)
		{
			ret = (int)Msg.wParam;
			break;
		}
		else if (ret == -1)
		{
			ret = (int)GetLastError();
			break;
		}
		else
		{
			if (!IsDialogMessage(hWndMainWindow, &Msg))
			{
				TranslateMessage(&Msg);
				DispatchMessage(&Msg);
			}
		}

	}

	return ret;
}

HWND hWndButtonOpen{};
HWND hWndButtonClose{};
HWND hWndButtonCloseAll{};
HWND hWndButtonLoad{};
HWND hWndButtonLoadAll{};
HWND hWndButtonUnload{};
HWND hWndButtonUnloadAll{};
HWND hWndButtonBroadcastWM_FONTCHANGE{};
HWND hWndButtonSelectProcess{};
HWND hWndListViewFontList{};
HWND hWndEditMessage{};
const unsigned int idButtonOpen{ 0 };
const unsigned int idButtonClose{ 1 };
const unsigned int idButtonCloseAll{ 2 };
const unsigned int idButtonLoad{ 3 };
const unsigned int idButtonLoadAll{ 4 };
const unsigned int idButtonUnload{ 5 };
const unsigned int idButtonUnloadAll{ 6 };
const unsigned int idButtonBroadcastWM_FONTCHANGE{ 7 };
const unsigned int idButtonSelectProcess{ 8 };
const unsigned int idListViewFontList{ 9 };
const unsigned int idEditMessage{ 10 };

LRESULT ButtonOpenProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam);
LRESULT ButtonCloseProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam);
LRESULT ButtonCloseAllProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam);
LRESULT ButtonLoadProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam);
LRESULT ButtonLoadAllProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam);
LRESULT ButtonUnloadProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam);
LRESULT ButtonUnloadAllProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam);
LRESULT ButtonSelectProcessProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK ListViewProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);
INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT UMessage, WPARAM wParam, LPARAM IParam);

LONG_PTR OldListViewProc;

void EnableAllButtons();
void DisableAllButtons();

void* lpRemoteAddFontProc{};
void* lpRemoteRemoveFontProc{};
HANDLE hTargetProcessWatchThread{};

ProcessInfo TargetProcessInfo{};

bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout);
bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout);

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};
	switch (Message)
	{
	case WM_CREATE:
		{
			RECT rectClient{};
			GetClientRect(hWnd, &rectClient);
			NONCLIENTMETRICS ncm{ sizeof(NONCLIENTMETRICS) };
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
			HFONT hFont{ CreateFontIndirect(&ncm.lfMessageFont) };

			//Initialize ButtonOpen
			hWndButtonOpen = CreateWindow(WC_BUTTON, L"Open", WS_CHILD | WS_VISIBLE | WS_GROUP | WS_TABSTOP | BS_PUSHBUTTON, 0, 0, 50, 50, hWnd, (HMENU)(idButtonOpen | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonOpen, hFont, TRUE);

			//Initialize ButtonClose
			hWndButtonClose = CreateWindow(WC_BUTTON, L"Close", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 50, 0, 50, 50, hWnd, (HMENU)(idButtonClose | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonClose, hFont, TRUE);

			//Initialize ButtonCloseAll
			hWndButtonCloseAll = CreateWindow(WC_BUTTON, L"Close\r\nAll", WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_PUSHBUTTON, 100, 0, 50, 50, hWnd, (HMENU)(idButtonCloseAll | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonCloseAll, hFont, TRUE);

			//Initialize ButtonLoad
			hWndButtonLoad = CreateWindow(WC_BUTTON, L"Load", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 150, 0, 50, 50, hWnd, (HMENU)(idButtonLoad | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonLoad, hFont, TRUE);

			//Initialize ButtonLoadAll
			hWndButtonLoadAll = CreateWindow(WC_BUTTON, L"Load\r\nAll", WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_PUSHBUTTON, 200, 0, 50, 50, hWnd, (HMENU)(idButtonLoadAll | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonLoadAll, hFont, TRUE);

			//Initialize ButtonUnload
			hWndButtonUnload = CreateWindow(WC_BUTTON, L"Unload", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 250, 0, 50, 50, hWnd, (HMENU)(idButtonUnload | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonUnload, hFont, TRUE);

			//Initialize ButtonUnloadAll
			hWndButtonUnloadAll = CreateWindow(WC_BUTTON, L"Unload\r\nAll", WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_PUSHBUTTON, 300, 0, 50, 50, hWnd, (HMENU)(idButtonUnloadAll | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonUnloadAll, hFont, TRUE);

			//Initialize ButtonBroadcastWM_FONTCHANGE
			hWndButtonBroadcastWM_FONTCHANGE = CreateWindow(WC_BUTTON, L"Broadcast WM_FONTCHANGE", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 350, 0, 250, 21, hWnd, (HMENU)(idButtonBroadcastWM_FONTCHANGE | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonBroadcastWM_FONTCHANGE, hFont, TRUE);

			//Initialize ButtonSelectProcess
			hWndButtonSelectProcess = CreateWindow(WC_BUTTON, L"Click to select process", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 350, 29, rectClient.right - rectClient.left - 350, 21, hWnd, (HMENU)(idButtonSelectProcess | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonSelectProcess, hFont, TRUE);

			//Initialize ListViewFontList
			hWndListViewFontList = CreateWindow(WC_LISTVIEW, L"FontList", WS_CHILD | WS_VISIBLE | WS_BORDER | WS_GROUP | LVS_REPORT, 0, 50, rectClient.right - rectClient.left, 300, hWnd, (HMENU)(idListViewFontList | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndListViewFontList, hFont, TRUE);
			DragAcceptFiles(hWndListViewFontList, TRUE);
			OldListViewProc = GetWindowLongPtr(hWndListViewFontList, GWLP_WNDPROC);
			SetWindowLongPtr(hWndListViewFontList, GWLP_WNDPROC, (LONG_PTR)ListViewProc);
			LVCOLUMN lvcolumn1{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectClient.right - rectClient.left) * 4 / 5 , (LPWSTR)L"Font Name" };
			ListView_InsertColumn(hWndListViewFontList, 0, &lvcolumn1);
			LVCOLUMN lvcolumn2{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectClient.right - rectClient.left) * 1 / 5 , (LPWSTR)L"State" };
			ListView_InsertColumn(hWndListViewFontList, 1, &lvcolumn2);
			ListView_SetExtendedListViewStyle(hWndListViewFontList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);
			CHANGEFILTERSTRUCT cfs{ sizeof(CHANGEFILTERSTRUCT) };
			ChangeWindowMessageFilterEx(hWndListViewFontList, WM_DROPFILES, MSGFLT_ALLOW, &cfs);
			ChangeWindowMessageFilterEx(hWndListViewFontList, WM_COPYDATA, MSGFLT_ALLOW, &cfs);
			ChangeWindowMessageFilterEx(hWndListViewFontList, 0x0049, MSGFLT_ALLOW, &cfs);	//WM_COPYGLOBALDATA

			//Initialize EditMessage
			hWndEditMessage = CreateWindow(WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_READONLY | ES_LEFT | ES_MULTILINE, 0, 350, rectClient.right - rectClient.left, rectClient.bottom - rectClient.top - 350, hWnd, (HMENU)(idEditMessage | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndEditMessage, hFont, TRUE);
			Edit_SetText(hWndEditMessage,
				L"Temporary load fonts to system font table.\r\n"
				"\r\n"
				"How to load fonts:\r\n"
				"1.Drag-drop font files onto the icon of this application.\r\n"
				"2.Click \"Open\" button to select fonts or drag-drop font files onto the list, then click \"Load All\" button.\r\n"
				"\r\n"
				"How to unload fonts:\r\n"
				"Click \"Unload All\" or \"Close All\" button or the X at the upper-right cornor.\r\n"
				"\r\n"
				"How to load fonts to process:\r\n"
				"1.Click \"Click to select process\", select a process.\r\n"
				"2.Click \"Open\" button to select fonts or drag-drop font files onto the list, then click \"Load All\" button.\r\n"
				"\r\n"
				"How to unload fonts from process:\r\n"
				"Click or \"Close All\" button or the X at the upper-right cornor or terminate selected process.\r\n"
				"\r\n"
				"UI description:\r\n"
				"\"Open\": Add fonts to the list.\r\n"
				"\"Close\": Remove selected fonts from system font table or target process and the list.\r\n"
				"\"Close All\": Remove all fonts from system font table or target process and the list.\r\n"
				"\"Load\": Add selected fonts to system font table or target process.\r\n"
				"\"Load All\": Add all fonts to system font table or target process.\r\n"
				"\"Unload\": Remove selected fonts from system font table or target process.\r\n"
				"\"Unload All\": Remove all fonts from system font table or target process.\r\n"
				"\"Broadcast WM_FONTCHANGE\": If checked, broadcast WM_FONTCHANGE message to all top windows when loading or unloading fonts.\r\n"
				"\"Click to select process\": click and select a process to only load fonts into selected process.\r\n"
				"\"Font Name\": Names of the fonts added to the list.\r\n"
				"\"State\": State of the font. There are five states, \"Not loaded\", \"Loaded\", \"Load failed\", \"Unloaded\" and \"Unload failed\".\r\n"
				"\r\n"
			);
			ret = 0;
		}
		break;
	case WM_ACTIVATE:
		{
			if (DragDropHasFonts)
			{
				//Process drag-drop font files onto the application icon stage II
				EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
				DisableAllButtons();
				_beginthread(DragDropWorkingThreadProc, 0, nullptr);

				ret = 0;
			}
		}
		break;
	case WM_CLOSE:
		{
			//Unload all fonts
			EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
			DisableAllButtons();
			_beginthread(CloseWorkingThreadProc, 0, nullptr);

			ret = 0;
		}
		break;
	case WM_DESTROY:
		{
			PostQuitMessage(0);
			ret = 0;
		}
		break;
	case UM_WORKINGTHREADTERMINATED:
		{
			EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
			EnableAllButtons();

			ret = 0;
		}
		break;
	case UM_CLOSEWORKINGTHREADTERMINATED:
		{
			if (wParam)
			{
#ifdef _WIN64
				const WCHAR szDllName[]{ L"FontLoaderExInjectionDll64.dll" };
#else
				const WCHAR szDllName[]{ L"FontLoaderExInjectionDll.dll" };
#endif

				//If loaded, unload FontLoaderExInjectionDll(64).dll from target process
				if (TargetProcessInfo.hProcess)
				{
					std::wstringstream Message{};
					int iMessageLength{};

					//Unload FontLoaderExInjectionDll(64).dll from target process
					if (!PullModule(TargetProcessInfo.hProcess, szDllName, 5000))
					{
						Message << L"Failed to unload " << szDllName << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
						iMessageLength = Edit_GetTextLength(hWndEditMessage);
						Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
						Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
						break;
					}
					Message << szDllName << L" successfully unloaded from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
					iMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

					//Close TargetProcessInfo.hProcess
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
					DestroyWindow(hWnd);
					break;
				case IDNO:
					EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
					EnableAllButtons();
					break;
				default:
					break;
				}
			}

			ret = 0;
		}
		break;
	case UM_WATCHPROCESSTERMINATED:
		{
			hTargetProcessWatchThread = NULL;

			EnableMenuItem(GetSystemMenu(hWnd, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
			EnableAllButtons();
		}
		break;
	default:
		{
			ret = ButtonOpenProc(hWnd, Message, wParam, lParam);
			ret = ButtonCloseProc(hWnd, Message, wParam, lParam);
			ret = ButtonCloseAllProc(hWnd, Message, wParam, lParam);
			ret = ButtonLoadProc(hWnd, Message, wParam, lParam);
			ret = ButtonLoadAllProc(hWnd, Message, wParam, lParam);
			ret = ButtonUnloadProc(hWnd, Message, wParam, lParam);
			ret = ButtonUnloadAllProc(hWnd, Message, wParam, lParam);
			ret = ButtonSelectProcessProc(hWnd, Message, wParam, lParam);

			ret = DefWindowProc(hWnd, Message, wParam, lParam);
		}
		break;
	}
	return ret;
}

LRESULT ButtonOpenProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam)
{
	switch (Message)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonOpen)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Open Dialog
					std::unique_ptr<WCHAR[]> lpszOpenFileNames{ new WCHAR[MAX_PATH * MAX_PATH]{} };
					std::vector<std::wstring> NewFontList;
					OPENFILENAME ofn{ sizeof(ofn), hWndParent, NULL, L"Font Files(*.ttf;*.ttc;*.otf)\0*.ttf;*.ttc;*.otf\0", NULL, 0, 0, lpszOpenFileNames.get(), MAX_PATH * MAX_PATH, NULL, 0, NULL, NULL, OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT, 0, 0, NULL, NULL, NULL, NULL, nullptr, 0, 0 };
					if (GetOpenFileName(&ofn))
					{
						if (PathIsDirectory(ofn.lpstrFile))
						{
							WCHAR* lpszFileName{ ofn.lpstrFile + ofn.nFileOffset };
							do
							{
								std::unique_ptr<WCHAR[]> lpszPath{ new WCHAR[std::wcslen(ofn.lpstrFile) + std::wcslen(lpszFileName) + 2]{} };
								PathCombine(lpszPath.get(), ofn.lpstrFile, lpszFileName);
								NewFontList.push_back(lpszPath.get());
								lpszFileName += std::wcslen(lpszFileName) + 1;
							} while (*lpszFileName);
						}
						else
						{
							NewFontList.push_back(ofn.lpstrFile);
						}

						//Insert items to ListViewFontList
						LVITEM lvi{ LVIF_TEXT, ListView_GetItemCount(hWndListViewFontList) };
						std::wstringstream Message{};
						int iMessageLenth{};
						for (int i = 0; i < (int)NewFontList.size(); i++)
						{
							lvi.iSubItem = 0;
							lvi.pszText = (LPWSTR)(NewFontList[i].c_str());
							ListView_InsertItem(hWndListViewFontList, &lvi);
							lvi.iSubItem = 1;
							lvi.pszText = (LPWSTR)L"Not loaded";
							ListView_SetItem(hWndListViewFontList, &lvi);
							lvi.iItem++;
							FontList.push_back(NewFontList[i]);
							Message.str(L"");
							Message << NewFontList[i] << L" successfully opened\r\n";
							iMessageLenth = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLenth, iMessageLenth);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
						}
						iMessageLenth = Edit_GetTextLength(hWndEditMessage);
						Edit_SetSel(hWndEditMessage, iMessageLenth, iMessageLenth);
						Edit_ReplaceSel(hWndEditMessage, L"\r\n");
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
	return 0;
}

LRESULT ButtonCloseProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam)
{
	switch (Message)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonClose)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Unload and close selected fonts
					//Won't close those failed to unload
					EnableMenuItem(GetSystemMenu(hWndParent, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
					DisableAllButtons();
					_beginthread(ButtonCloseWorkingThreadProc, 0, nullptr);
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
	return 0;
}

LRESULT ButtonCloseAllProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam)
{
	switch (Message)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonCloseAll)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Unload and close all fonts
					//Won't close those failed to unload
					EnableMenuItem(GetSystemMenu(hWndParent, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
					DisableAllButtons();

					int iCount = ListView_GetItemCount(hWndListViewFontList);
					for (int i = 0; i < iCount; i++)
					{
						ListView_SetItemState(hWndListViewFontList, i, LVIS_SELECTED, LVIS_SELECTED);
					}
					_beginthread(ButtonCloseWorkingThreadProc, 0, nullptr);
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
	return 0;
}

LRESULT ButtonLoadProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam)
{
	switch (Message)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonLoad)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Load selected fonts
					EnableMenuItem(GetSystemMenu(hWndParent, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
					DisableAllButtons();
					_beginthread(ButtonLoadWorkingThreadProc, 0, nullptr);
				}
			default:
				break;
			}
		}
		break;
	default:
		break;
	}
	return 0;
}

LRESULT ButtonLoadAllProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam)
{
	switch (Message)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonLoadAll)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Load all fonts
					EnableMenuItem(GetSystemMenu(hWndParent, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
					DisableAllButtons();

					int iCount = ListView_GetItemCount(hWndListViewFontList);
					for (int i = 0; i < iCount; i++)
					{
						ListView_SetItemState(hWndListViewFontList, i, LVIS_SELECTED, LVIS_SELECTED);
					}
					_beginthread(ButtonLoadWorkingThreadProc, 0, nullptr);
				}
			default:
				break;
			}
		}
		break;
	default:
		break;
	}
	return 0;
}

LRESULT ButtonUnloadProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam)
{
	switch (Message)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonUnload)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Unload selected fonts;
					EnableMenuItem(GetSystemMenu(hWndParent, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
					DisableAllButtons();
					_beginthread(ButtonUnloadWorkingThreadProc, 0, nullptr);
				}
			default:
				break;
			}
		}
		break;
	default:
		break;
	}
	return 0;
}

LRESULT ButtonUnloadAllProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam)
{
	switch (Message)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonUnloadAll)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Unload all fonts
					EnableMenuItem(GetSystemMenu(hWndParent, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
					DisableAllButtons();

					int iCount = ListView_GetItemCount(hWndListViewFontList);
					for (int i = 0; i < iCount; i++)
					{
						ListView_SetItemState(hWndListViewFontList, i, LVIS_SELECTED, LVIS_SELECTED);
					}
					_beginthread(ButtonUnloadWorkingThreadProc, 0, nullptr);
				}
			default:
				break;
			}
		}
		break;
	default:
		break;
	}
	return 0;
}

LRESULT ButtonSelectProcessProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam)
{
	switch (Message)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonSelectProcess)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
#ifdef _WIN64
					const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll64.dll" };
#else
					const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll.dll" };
#endif

					std::wstringstream Message{};
					int iMessageLength{};
					
					static bool bIsSeDebugPrivilegeGranted{ false };

					//Elevate privilege
					if (!bIsSeDebugPrivilegeGranted)
					{
						if (!EnableDebugPrivilege())
						{
							MessageBox(NULL, L"Failed to grant FontLoaderEx SeDebugPrivilige.", L"FontLoaderEx", MB_ICONERROR);
							return -1;
						}
						bIsSeDebugPrivilegeGranted = true;
					}

					//If some fonts are loaded into the target process, prompt user to close all first
					if ((!FontList.empty()) && TargetProcessInfo.hProcess)
					{
						MessageBox(hWndParent, L"Please close all fonts before selecting process.", L"FontLoaderEx", MB_ICONEXCLAMATION);
						break;
					}

					//Select process
					ProcessInfo* p{ (ProcessInfo*)DialogBox(NULL, MAKEINTRESOURCE(IDD_DIALOG1), hWndParent, DialogProc) };
					
					//If p != nullptr, select it
					if (p)
					{
						ProcessInfo SelectedProcessInfo = *p;
						delete p;

						//Terminate watch thread if started
						if (hTargetProcessWatchThread)
						{
							PostThreadMessage(GetThreadId(hTargetProcessWatchThread), UM_TERMINATEWATCHPROCESS, NULL, NULL);
							WaitForSingleObject(hTargetProcessWatchThread, INFINITE);
							hTargetProcessWatchThread = NULL;
						}

						//If loaded, unload FontLoaderExInjectionDll(64).dll from target process
						if (TargetProcessInfo.hProcess)
						{
							if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, 5000))
							{
								Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
								iMessageLength = Edit_GetTextLength(hWndEditMessage);
								Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
								Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
								break;
							}
							Message << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
							Message.str(L"");

							//Close PreviousTargetProcessInfo.ProcessName
							CloseHandle(TargetProcessInfo.hProcess);
							TargetProcessInfo.hProcess = NULL;
						}

						//Determine target process type
						SelectedProcessInfo.hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, SelectedProcessInfo.ProcessID);
						if (!SelectedProcessInfo.hProcess)
						{
							Message << L"Failed to open process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
							break;
						}
						BOOL bIsWOW64{};
						IsWow64Process(SelectedProcessInfo.hProcess, &bIsWOW64);
#ifdef _WIN64
						if (bIsWOW64)
						{
							CloseHandle(SelectedProcessInfo.hProcess);
							Message << L"Cannot inject 64-bit dll " << szInjectionDllName << L" into 32-bit process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
							break;
						}
#else
						if (!bIsWOW64)
						{
							CloseHandle(SelectedProcessInfo.hProcess);
							Message << L"Cannot inject 32-bit dll " << szInjectionDllName << L" into 64-bit process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
							break;
						}
#endif

						//Check whether target process loads gdi32.dll as AddFontResourceEx() and RemoveFontResourceEx() are in it
						HANDLE hModuleSnapshot1{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, SelectedProcessInfo.ProcessID) };
						MODULEENTRY32 me321{ sizeof(MODULEENTRY32) };
						bool bIsGDI32Loaded{ false };
						if (!Module32First(hModuleSnapshot1, &me321))
						{
							CloseHandle(SelectedProcessInfo.hProcess);
							CloseHandle(hModuleSnapshot1);
							Message << L"Failed to enumerate modules in target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
							break;
						}
						do
						{
							if (!lstrcmpi(me321.szModule, L"gdi32.dll"))
							{
								bIsGDI32Loaded = true;
								break;
							}
						} while (Module32Next(hModuleSnapshot1, &me321));
						if (!bIsGDI32Loaded)
						{
							CloseHandle(SelectedProcessInfo.hProcess);
							CloseHandle(hModuleSnapshot1);
							Message << L"gdi32.dll not loaded by target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
							break;
						}
						CloseHandle(hModuleSnapshot1);

						//Inject FontLoaderExInjectionDll(64).dll into target process
						if (!InjectModule(SelectedProcessInfo.hProcess, szInjectionDllName, 5000))
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

						//Get base address of FontLoaderExInjectionDll(64).dll in target process
						HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, SelectedProcessInfo.ProcessID) };
						MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
						BYTE* lpModBaseAddress{};
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
							if (!lstrcmpi(me32.szModule, szInjectionDllName))
							{
								lpModBaseAddress = me32.modBaseAddr;
								break;
							}
						} while (Module32Next(hModuleSnapshot, &me32));
						if (!lpModBaseAddress)
						{
							CloseHandle(SelectedProcessInfo.hProcess);
							CloseHandle(hModuleSnapshot);
							Message << szInjectionDllName << " not found in target process " << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
							break;
						}
						CloseHandle(hModuleSnapshot);

						//Calculate addresses of AddFont() and RemoveFont() in target process
						HMODULE hModInjectionDll{ LoadLibrary(szInjectionDllName) };
						void* lpLocalAddFontProcAddr{ GetProcAddress(hModInjectionDll, "AddFont") };
						void* lpLocalRemoveFontProcAddr{ GetProcAddress(hModInjectionDll, "RemoveFont") };
						INT_PTR AddFontProcOffset{ (INT_PTR)lpLocalAddFontProcAddr - (INT_PTR)hModInjectionDll };
						INT_PTR RemoveFontProcOffset{ (INT_PTR)lpLocalRemoveFontProcAddr - (INT_PTR)hModInjectionDll };
						FreeLibrary(hModInjectionDll);
						lpRemoteAddFontProc = lpModBaseAddress + AddFontProcOffset;
						lpRemoteRemoveFontProc = lpModBaseAddress + RemoveFontProcOffset;

						//Register remote AddFont() and RemoveFont() procedure
						FontResource::RegisterAddRemoveFontProc(RemoteAddFontProc, RemoteRemoveFontProc);

						//Change caption and set TargetProcessInfo
						std::wstringstream Caption{};
						Caption << SelectedProcessInfo.ProcessName << L"(" << SelectedProcessInfo.ProcessID << L")";
						Button_SetText(hWndButtonSelectProcess, (LPCWSTR)Caption.str().c_str());
						TargetProcessInfo = SelectedProcessInfo;

						//Start watch thread
						hTargetProcessWatchThread = (HANDLE)_beginthread(TargetProcessWatchThreadProc, 0, nullptr);
					}
					//If p == nullptr, clear selected process
					else
					{
						//If loaded, unload FontLoaderExInjectionDll(64).dll from target process
						if (TargetProcessInfo.hProcess)
						{
							//Terminate watch thread
							PostThreadMessage(GetThreadId(hTargetProcessWatchThread), UM_TERMINATEWATCHPROCESS, NULL, NULL);
							WaitForSingleObject(hTargetProcessWatchThread, INFINITE);
							hTargetProcessWatchThread = NULL;

							//Unload FontLoaderExInjectionDll(64).dll from target process
							if (!PullModule(TargetProcessInfo.hProcess, szInjectionDllName, 5000))
							{
								Message << L"Failed to unload " << szInjectionDllName << L" from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
								iMessageLength = Edit_GetTextLength(hWndEditMessage);
								Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
								Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
								break;
							}
							Message << szInjectionDllName << L" successfully unloaded from target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L").\r\n\r\n";
							iMessageLength = Edit_GetTextLength(hWndEditMessage);
							Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
							Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

							//Close TargetProcessInfo.hProcess
							CloseHandle(TargetProcessInfo.hProcess);
							TargetProcessInfo.hProcess = NULL;

							//Register default AddFont() and RemoveFont() procedure
							FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);

							//Revert to default caption
							Button_SetText(hWndButtonSelectProcess, L"Click to select process");
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
	return 0;
}

LRESULT CALLBACK ListViewProc(HWND hWndParent, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{ 0 };
	switch (Message)
	{
	case WM_DROPFILES:
		{
			//Process drag-drop and open fonts
			LVITEM lvi{ LVIF_TEXT, ListView_GetItemCount(hWndListViewFontList) };
			UINT nFileCount{ DragQueryFile((HDROP)wParam, 0xFFFFFFFF, NULL, 0) };
			std::wstringstream Message{};
			int iMessageLength{};
			for (UINT i = 0; i < nFileCount; i++)
			{
				UINT nSize{ DragQueryFile((HDROP)wParam, i, NULL, 0) + 1 };
				std::unique_ptr<WCHAR[]> lpszFileName(new WCHAR[nSize]{});
				DragQueryFile((HDROP)wParam, i, lpszFileName.get(), nSize);
				if (PathMatchSpec(lpszFileName.get(), L"*.ttf") || PathMatchSpec(lpszFileName.get(), L"*.ttc") || PathMatchSpec(lpszFileName.get(), L"*.otf"))
				{
					lvi.iSubItem = 0;
					lvi.pszText = lpszFileName.get();
					ListView_InsertItem(hWndListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = (LPWSTR)L"Not loaded";
					ListView_SetItem(hWndListViewFontList, &lvi);
					lvi.iItem++;
					FontList.push_back(lpszFileName.get());
					Message.str(L"");
					Message << lpszFileName.get() << L" successfully opened\r\n";
					iMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
				}
			}
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, L"\r\n");
			DragFinish((HDROP)wParam);

			ret = 0;
		}
		break;
	default:
		{
			ret = CallWindowProc((WNDPROC)OldListViewProc, hWndParent, Message, wParam, lParam);
		}
		break;
	}
	return ret;
}

bool CompareByProcessNameAscending(const ProcessInfo& value1, const ProcessInfo& value2)
{
	std::wstring s1(value1.ProcessName.size(), L'\0'), s2(value2.ProcessName.size(), L'\0');
	std::transform(value1.ProcessName.begin(), value1.ProcessName.end(), s1.begin(), std::tolower);
	std::transform(value2.ProcessName.begin(), value2.ProcessName.end(), s2.begin(), std::tolower);
	return s1 < s2;
}

bool CompareByProcessNameDescending(const ProcessInfo& value1, const ProcessInfo& value2)
{
	std::wstring s1(value1.ProcessName.size(), L'\0'), s2(value2.ProcessName.size(), L'\0');
	std::transform(value1.ProcessName.begin(), value1.ProcessName.end(), s1.begin(), std::tolower);
	std::transform(value2.ProcessName.begin(), value2.ProcessName.end(), s2.begin(), std::tolower);
	return s1 > s2;
}

bool CompareByProcessIDAscending(const ProcessInfo& value1, const ProcessInfo& value2)
{
	return value1.ProcessID < value2.ProcessID;
}

bool CompareByProcessIDDescending(const ProcessInfo& value1, const ProcessInfo& value2)
{
	return value1.ProcessID > value2.ProcessID;
}

INT_PTR CALLBACK DialogProc(HWND hWndDialog, UINT Message, WPARAM wParam, LPARAM lParam)
{
	INT_PTR ret{};
	static HANDLE hProcessSnapshot{};
	static std::vector<ProcessInfo> ProcessList{};
	static bool bOrder{ true };

	switch (Message)
	{
	case WM_INITDIALOG:
		{
			bOrder = true;
			NONCLIENTMETRICS ncm{ sizeof(NONCLIENTMETRICS) };
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
			HFONT hFont{ CreateFontIndirect(&ncm.lfMessageFont) };
			HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };
			HWND hWndOK{ GetDlgItem(hWndDialog, IDOK) };
			HWND hWndCancel{ GetDlgItem(hWndDialog, IDCANCEL) };
			SetWindowFont(hWndListViewProcessList, hFont, TRUE);
			SetWindowFont(hWndOK, hFont, TRUE);
			SetWindowFont(hWndCancel, hFont, TRUE);
			RECT rectListViewClient{};
			GetClientRect(hWndListViewProcessList, &rectListViewClient);
			SetWindowLongPtr(hWndListViewProcessList, GWL_STYLE, GetWindowLongPtr(hWndListViewProcessList, GWL_STYLE) | LVS_REPORT | LVS_SINGLESEL);
			LVCOLUMN lvcolumn1{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectListViewClient.right - rectListViewClient.left) * 4 / 5 , (LPWSTR)L"Process" };
			ListView_InsertColumn(hWndListViewProcessList, 0, &lvcolumn1);
			LVCOLUMN lvcolumn2{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectListViewClient.right - rectListViewClient.left) * 1 / 5 , (LPWSTR)L"PID" };
			ListView_InsertColumn(hWndListViewProcessList, 1, &lvcolumn2);
			ListView_SetExtendedListViewStyle(hWndListViewProcessList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

			ProcessList.clear();
			PROCESSENTRY32 pe32{ sizeof(PROCESSENTRY32) };
			hProcessSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
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
		switch (LOWORD(wParam))
		{
		case IDOK:
			{
				int iSelected{ ListView_GetSelectionMark(GetDlgItem(hWndDialog, IDC_LIST1)) };
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
		case IDCANCEL:
			{
				EndDialog(hWndDialog, NULL);
				ret = (INT_PTR)TRUE;
			}
			break;
		default:
			break;
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
					case LVN_COLUMNCLICK:
						{
							// Sort item by Process or PID
							switch (((LPNMLISTVIEW)lParam)->iSubItem)
							{
							case 0:
								{
									bOrder ? std::sort(ProcessList.begin(), ProcessList.end(), CompareByProcessNameAscending) : std::sort(ProcessList.begin(), ProcessList.end(), CompareByProcessNameDescending);
									bOrder = !bOrder;
								}
								break;
							case 1:
								{
									bOrder ? std::sort(ProcessList.begin(), ProcessList.end(), CompareByProcessIDAscending) : std::sort(ProcessList.begin(), ProcessList.end(), CompareByProcessIDDescending);
									bOrder = !bOrder;
								}
								break;
							default:
								break;
							}
							LVITEM lvi{ LVIF_TEXT, 0 };
							HWND hWndListViewProcessList{ GetDlgItem(hWndDialog, IDC_LIST1) };
							for (auto& i : ProcessList)
							{
								lvi.iSubItem = 0;
								lvi.pszText = (LPWSTR)i.ProcessName.c_str();
								ListView_SetItem(hWndListViewProcessList, &lvi);
								lvi.iSubItem = 1;
								std::wstring str{ std::to_wstring(i.ProcessID) };
								lvi.pszText = (LPWSTR)str.c_str();
								ListView_SetItem(hWndListViewProcessList, &lvi);
								lvi.iItem++;
							}
						}
						break;
					case NM_DBLCLK:
						{
							//Select double-clicked item
							SendMessage(GetDlgItem(hWndDialog, IDOK), BM_CLICK, NULL, NULL);
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
	ret = (INT_PTR)FALSE;

	return ret;
}

void EnableAllButtons()
{
	EnableWindow(hWndButtonOpen, TRUE);
	EnableWindow(hWndButtonClose, TRUE);
	EnableWindow(hWndButtonCloseAll, TRUE);
	EnableWindow(hWndButtonLoad, TRUE);
	EnableWindow(hWndButtonLoadAll, TRUE);
	EnableWindow(hWndButtonUnload, TRUE);
	EnableWindow(hWndButtonUnloadAll, TRUE);
	EnableWindow(hWndButtonBroadcastWM_FONTCHANGE, TRUE);
	EnableWindow(hWndButtonSelectProcess, TRUE);
	EnableWindow(hWndListViewFontList, TRUE);
}

void DisableAllButtons()
{
	EnableWindow(hWndButtonOpen, FALSE);
	EnableWindow(hWndButtonClose, FALSE);
	EnableWindow(hWndButtonCloseAll, FALSE);
	EnableWindow(hWndButtonLoad, FALSE);
	EnableWindow(hWndButtonLoadAll, FALSE);
	EnableWindow(hWndButtonUnload, FALSE);
	EnableWindow(hWndButtonUnloadAll, FALSE);
	EnableWindow(hWndButtonBroadcastWM_FONTCHANGE, FALSE);
	EnableWindow(hWndButtonSelectProcess, FALSE);
	EnableWindow(hWndListViewFontList, FALSE);
}

bool EnableDebugPrivilege()
{
	//Elevate privillege
	HANDLE hToken{};
	LUID luid{};
	if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken))
	{
		return false;
	}
	if (!LookupPrivilegeValue(NULL, SE_DEBUG_NAME, &luid))
	{
		CloseHandle(hToken);
		return false;
	}
	TOKEN_PRIVILEGES tkp{ 1 };
	tkp.Privileges[0].Luid = luid;
	tkp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;
	if (!AdjustTokenPrivileges(hToken, FALSE, &tkp, sizeof(tkp), NULL, NULL))
	{
		CloseHandle(hToken);
		return false;
	}

	CloseHandle(hToken);
	return true;
}

bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout)
{
	//Inject dll into target process
	std::unique_ptr<WCHAR[]> szDllPath{ new WCHAR[MAX_PATH * MAX_PATH]{} };
	GetModuleFileName(NULL, szDllPath.get(), MAX_PATH);
	PathRemoveFileSpec(szDllPath.get());
	PathAppend(szDllPath.get(), szModuleName);
	LPVOID lpRemoteBuffer{ VirtualAllocEx(hProcess, NULL, (std::wcslen(szDllPath.get()) + 1) * sizeof(WCHAR), MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE) };
	if (!lpRemoteBuffer)
	{
		return false;
	}
	if (!WriteProcessMemory(hProcess, lpRemoteBuffer, (LPVOID)szDllPath.get(), (std::wcslen(szDllPath.get()) + 1) * sizeof(WCHAR), NULL))
	{
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);
		return false;
	}
	HMODULE hModule{ GetModuleHandle(L"Kernel32") };
	LPTHREAD_START_ROUTINE addr{ (LPTHREAD_START_ROUTINE)GetProcAddress(hModule, "LoadLibraryW") };
	HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, addr, lpRemoteBuffer, 0, NULL) };
	if (!hRemoteThread)
	{
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);
		return false;
	}
	if (WaitForSingleObject(hRemoteThread, Timeout) == WAIT_TIMEOUT)
	{
		CloseHandle(hRemoteThread);
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);
		return false;
	}
	VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

	//Get exit code of remote thread
	DWORD dwRemoteThreadExitCode{};
	if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
	{
		CloseHandle(hRemoteThread);
		return false;
	}
	if (!dwRemoteThreadExitCode)
	{
		CloseHandle(hRemoteThread);
		return false;
	}
	CloseHandle(hRemoteThread);

	return true;
}

bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout)
{
	//Find HMODULE of szModuleName in target process
	HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetProcessId(hProcess)) };
	MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
	HMODULE hModInjectionDll{};
	if (!Module32First(hModuleSnapshot, &me32))
	{
		CloseHandle(hModuleSnapshot);
		return false;
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
		return false;
	}
	CloseHandle(hModuleSnapshot);

	//Unload FontLoaderExInjectionDll(64).dll from target process
	HMODULE hModule{ GetModuleHandle(L"Kernel32") };
	LPTHREAD_START_ROUTINE addr{ (LPTHREAD_START_ROUTINE)GetProcAddress(hModule, "FreeLibrary") };
	HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, addr, (LPVOID)hModInjectionDll, 0, NULL) };
	if (!hRemoteThread)
	{
		return false;
	}
	if (WaitForSingleObject(hRemoteThread, Timeout) == WAIT_TIMEOUT)
	{
		CloseHandle(hRemoteThread);
		return false;
	}

	//Get exit code of remote thread
	DWORD dwRemoteThreadExitCode{};
	if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
	{
		CloseHandle(hRemoteThread);
		return false;
	}
	if (!dwRemoteThreadExitCode)
	{
		CloseHandle(hRemoteThread);
		return false;
	}
	CloseHandle(hRemoteThread);

	return true;
}
