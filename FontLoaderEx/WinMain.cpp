#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "ComCtl32.lib")
#pragma comment(lib, "Shlwapi.lib")

#include <windows.h>
#include <CommCtrl.h>
#include <shlwapi.h>
#include <windowsx.h>
#include <process.h>
#include <string>
#include <cstring>
#include <sstream>
#include <climits>
#include <list>
#include <vector>
#include <memory>
#include "FontResource.h"
#include "Globals.h"

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

std::list<FontResource> FontList{};

HWND hWndMainWindow{};
CRITICAL_SECTION CriticalSection{};
bool DragDropHasFonts{ false };

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
	InitializeCriticalSection(&CriticalSection);

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
	MSG msg{};
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
	while ((bRet = GetMessage(&msg, NULL, 0, 0)))
	{
		if (bRet == 0)
		{
			ret = (int)msg.wParam;
			break;
		}
		else if (ret == -1)
		{
			ret = (int)GetLastError();
			break;
		}
		else
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	DeleteCriticalSection(&CriticalSection);

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
const unsigned int idListViewFontList{ 8 };
const unsigned int idEditMessage{ 9 };

extern const UINT UM_WORKINGTHREADTERMINATED = WM_USER + 0x100;
extern const UINT UM_CLOSEWORKINGTHREADTERMINATED = WM_USER + 0x101;

LRESULT ButtonOpenProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonCloseProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonCloseAllProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonLoadProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonLoadAllProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonUnloadProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonUnloadAllProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK ListViewProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

LONG_PTR OldListViewProc;

void EnableAllButtons();
void DisableAllButtons();

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};
	switch (msg)
	{
	case WM_CREATE:
		{
			RECT rectClient{};
			GetClientRect(hWnd, &rectClient);
			NONCLIENTMETRICS ncm{ sizeof(NONCLIENTMETRICS) };
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(NONCLIENTMETRICS), &ncm, 0);
			HFONT hFont = CreateFontIndirect(&ncm.lfMessageFont);

			//Initialize ButtonOpen
			hWndButtonOpen = CreateWindow(WC_BUTTON, L"Open", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 50, 50, hWnd, (HMENU)(idButtonOpen | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
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

			//Initialize ButtonBroadcastMsg
			hWndButtonBroadcastWM_FONTCHANGE = CreateWindow(WC_BUTTON, L"Broadcast WM_FONTCHANGE", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 350, 0, 250, 20, hWnd, (HMENU)(idButtonBroadcastWM_FONTCHANGE | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndButtonBroadcastWM_FONTCHANGE, hFont, TRUE);

			//Initialize ListViewFontList
			hWndListViewFontList = CreateWindow(WC_LISTVIEW, L"FontList", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT, 0, 50, rectClient.right - rectClient.left, 300, hWnd, (HMENU)(idListViewFontList | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hWndListViewFontList, hFont, TRUE);
			DragAcceptFiles(hWndListViewFontList, TRUE);
			OldListViewProc = GetWindowLongPtr(hWndListViewFontList, GWLP_WNDPROC);
			SetWindowLongPtr(hWndListViewFontList, GWLP_WNDPROC, (LONG_PTR)ListViewProc);
			LVCOLUMN lvcolumn1{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectClient.right - rectClient.left) * 7 / 10 , (LPWSTR)L"Font Name" };
			ListView_InsertColumn(hWndListViewFontList, 0, &lvcolumn1);
			LVCOLUMN lvcolumn2{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectClient.right - rectClient.left) * 3 / 10 , (LPWSTR)L"State" };
			ListView_InsertColumn(hWndListViewFontList, 1, &lvcolumn2);
			ListView_SetExtendedListViewStyle(hWndListViewFontList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

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
				"UI description:\r\n"
				"\"Open\": Add fonts to the list.\r\n"
				"\"Close\": Remove selected fonts from system font table and the list.\r\n"
				"\"Close All\": Remove all fonts from system font table and the list.\r\n"
				"\"Load\": Add selected fonts to system font table.\r\n"
				"\"Load All\": Add all fonts to system font table.\r\n"
				"\"Unload\": Remove selected fonts from system font table.\r\n"
				"\"Unload All\": Remove all fonts from system font table.\r\n"
				"\"Broadcast WM_FONTCHANGE\": If checked, broadcast WM_FONTCHANGE message to all top windows when loading or unloading fonts.\r\n"
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
				DisableAllButtons();
				_beginthread((_beginthread_proc_type)DragDropWorkingThreadProc, 0, nullptr);

				ret = 0;
			}
		}
		break;
	case WM_CLOSE:
		{
			//Unload all fonts
			DisableAllButtons();
			_beginthread((_beginthread_proc_type)CloseWorkingThreadProc, 0, nullptr);

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
			EnableAllButtons();

			ret = 0;
		}
		break;
	case UM_CLOSEWORKINGTHREADTERMINATED:
		{
			if (wParam)
			{
				DestroyWindow(hWndMainWindow);
			}
			else
			{
				switch (MessageBox(hWndMainWindow, L"Some fonts are not successfully unloaded.\r\n\r\nDo you still want to exit?", L"FontLoaderEx", MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON1 | MB_APPLMODAL))
				{
				case IDYES:
					DestroyWindow(hWndMainWindow);
					break;
				case IDNO:
					EnableAllButtons();
					break;
				default:
					break;
				}
			}

			ret = 0;
		}
		break;
	default:
		{
			ret = ButtonOpenProc(hWnd, msg, wParam, lParam);
			ret = ButtonCloseProc(hWnd, msg, wParam, lParam);
			ret = ButtonCloseAllProc(hWnd, msg, wParam, lParam);
			ret = ButtonLoadProc(hWnd, msg, wParam, lParam);
			ret = ButtonLoadAllProc(hWnd, msg, wParam, lParam);
			ret = ButtonUnloadProc(hWnd, msg, wParam, lParam);
			ret = ButtonUnloadAllProc(hWnd, msg, wParam, lParam);

			ret = DefWindowProc(hWnd, msg, wParam, lParam);
		}
		break;
	}
	return ret;
}

LRESULT ButtonOpenProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
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
					OPENFILENAME ofn{ sizeof(ofn), hWnd, NULL, L"Font Files(*.ttf;*.ttc;*.otf)\0*.ttf;*.ttc;*.otf\0", NULL, 0, 0, lpszOpenFileNames.get(), MAX_PATH * MAX_PATH, NULL, 0, NULL, NULL, OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT, 0, 0, NULL, NULL, NULL, NULL, nullptr, 0, 0 };
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

						//Insert items to lstFontList
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
	default:
		break;
	}
	return 0;
}

LRESULT ButtonCloseProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
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
					DisableAllButtons();
					_beginthread((_beginthread_proc_type)ButtonCloseWorkingThreadProc, 0, nullptr);
				}
				break;
			default:
				break;
			}
		}
	default:
		break;
	}
	return 0;
}

LRESULT ButtonCloseAllProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
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
					DisableAllButtons();
					_beginthread((_beginthread_proc_type)ButtonCloseAllWorkingThreadProc, 0, nullptr);
				}
				break;
			default:
				break;
			}
		}
	default:
		break;
	}
	return 0;
}

LRESULT ButtonLoadProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonLoad)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Load selected fonts
					DisableAllButtons();
					_beginthread((_beginthread_proc_type)ButtonLoadWorkingThreadProc, 0, nullptr);
				}
			default:
				break;
			}
		}
	default:
		break;
	}
	return 0;
}

LRESULT ButtonLoadAllProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonLoadAll)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Load all fonts
					DisableAllButtons();
					_beginthread((_beginthread_proc_type)ButtonLoadAllWorkingThreadProc, 0, nullptr);
				}
			default:
				break;
			}
		}
	default:
		break;
	}
	return 0;
}

LRESULT ButtonUnloadProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonUnload)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Unload selected fonts;
					DisableAllButtons();
					_beginthread((_beginthread_proc_type)ButtonUnloadWorkingThreadProc, 0, nullptr);
				}
			default:
				break;
			}
		}
	default:
		break;
	}
	return 0;
}

LRESULT ButtonUnloadAllProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	switch (msg)
	{
	case WM_COMMAND:
		if (LOWORD(wParam) == idButtonUnloadAll)
		{
			switch (HIWORD(wParam))
			{
			case BN_CLICKED:
				{
					//Unload all fonts
					DisableAllButtons();
					_beginthread((_beginthread_proc_type)ButtonUnloadAllWorkingThreadProc, 0, nullptr);
				}
			default:
				break;
			}
		}
	default:
		break;
	}
	return 0;
}

LRESULT CALLBACK ListViewProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{ 0 };
	switch (msg)
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
				UINT nSize = DragQueryFile((HDROP)wParam, i, NULL, 0) + 1;
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
			ret = CallWindowProc((WNDPROC)OldListViewProc, hWnd, msg, wParam, lParam);
		}
		break;
	}
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
	EnableWindow(hWndListViewFontList, TRUE);
	EnableMenuItem(GetSystemMenu(hWndMainWindow, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
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
	EnableWindow(hWndListViewFontList, FALSE);
	EnableMenuItem(GetSystemMenu(hWndMainWindow, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
}
