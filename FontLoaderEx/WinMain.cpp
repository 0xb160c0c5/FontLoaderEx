#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "ComCtl32.lib")
#pragma comment(lib, "Shlwapi.lib")

#include <windows.h>
#include <CommCtrl.h>
#include <shlwapi.h>
#include <windowsx.h>
#include <string>
#include <cstring>
#include <list>
#include <vector>
#include <memory>
#include "FontResource.h"

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

std::list<FontResource> FontList{};

bool DragDropHasFonts = false;

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
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
		for (auto& i : FontList)
		{
			i.Load();
		}
	}
	LocalFree(argv);

	//Create window
	HWND hWnd{};
	MSG msg{};
	WNDCLASS wndclass{ CS_HREDRAW | CS_VREDRAW, WndProc, 0, 0, hInstance, LoadIcon(NULL, IDI_APPLICATION), LoadCursor(NULL, IDC_ARROW), GetSysColorBrush(COLOR_WINDOW), NULL, L"FontLoaderEx" };

	InitCommonControls();

	if (!RegisterClass(&wndclass))
	{
		MessageBox(NULL, L"Failed to register window class.", L"FontLoaderEx", MB_ICONERROR);
		return -1;
	}

	if (!(hWnd = CreateWindow(L"FontLoaderEx", L"FontLoaderEx", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_BORDER | WS_MINIMIZEBOX, CW_USEDEFAULT, CW_USEDEFAULT, 700, 700, NULL, NULL, hInstance, NULL)))
	{
		MessageBox(NULL, L"Failed to create window.", L"FontLoaderEx", MB_ICONERROR);
		return -1;
	}

	ShowWindow(hWnd, nShowCmd);
	UpdateWindow(hWnd);

	BOOL bRet{};
	while ((bRet = GetMessage(&msg, NULL, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			return (int)GetLastError();
		}
		else
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return (int)msg.wParam;
}

HWND hButtonOpen;
HWND hButtonClose;
HWND hButtonCloseAll;
HWND hButtonLoad;
HWND hButtonLoadAll;
HWND hButtonUnload;
HWND hButtonUnloadAll;
HWND hButtonBroadcastMsg;
HWND hListViewFontList;
HWND hEditMessage;
const unsigned int idButtonOpen{ 0 };
const unsigned int idButtonClose{ 1 };
const unsigned int idButtonCloseAll{ 2 };
const unsigned int idButtonLoad{ 3 };
const unsigned int idButtonLoadAll{ 4 };
const unsigned int idButtonUnload{ 5 };
const unsigned int idButtonUnloadAll{ 6 };
const unsigned int idButtonBroadcastMsg{ 7 };
const unsigned int idListViewFontList{ 8 };
const unsigned int idEditMessage{ 9 };

LRESULT ButtonOpenProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonCloseProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonCloseAllProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonLoadProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonLoadAllProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonUnloadProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT ButtonUnloadAllProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK ListViewProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);

LONG_PTR OldListViewProc;

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
			hButtonOpen = CreateWindow(WC_BUTTON, L"Open", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 0, 0, 50, 50, hWnd, (HMENU)(idButtonOpen | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hButtonOpen, hFont, TRUE);

			//Initialize ButtonClose
			hButtonClose = CreateWindow(WC_BUTTON, L"Close", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 50, 0, 50, 50, hWnd, (HMENU)(idButtonClose | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hButtonClose, hFont, TRUE);

			//Initialize ButtonCloseAll
			hButtonCloseAll = CreateWindow(WC_BUTTON, L"Close\r\nAll", WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_PUSHBUTTON, 100, 0, 50, 50, hWnd, (HMENU)(idButtonCloseAll | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hButtonCloseAll, hFont, TRUE);

			//Initialize ButtonLoad
			hButtonLoad = CreateWindow(WC_BUTTON, L"Load", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 150, 0, 50, 50, hWnd, (HMENU)(idButtonLoad | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hButtonLoad, hFont, TRUE);

			//Initialize ButtonLoadAll
			hButtonLoadAll = CreateWindow(WC_BUTTON, L"Load\r\nAll", WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_PUSHBUTTON, 200, 0, 50, 50, hWnd, (HMENU)(idButtonLoadAll | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hButtonLoadAll, hFont, TRUE);

			//Initialize ButtonUnload
			hButtonUnload = CreateWindow(WC_BUTTON, L"Unload", WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON, 250, 0, 50, 50, hWnd, (HMENU)(idButtonUnload | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hButtonUnload, hFont, TRUE);

			//Initialize ButtonUnloadAll
			hButtonUnloadAll = CreateWindow(WC_BUTTON, L"Unload\r\nAll", WS_CHILD | WS_VISIBLE | BS_MULTILINE | BS_PUSHBUTTON, 300, 0, 50, 50, hWnd, (HMENU)(idButtonUnloadAll | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hButtonUnloadAll, hFont, TRUE);

			//Initialize ButtonBroadcastMsg
			hButtonBroadcastMsg = CreateWindow(WC_BUTTON, L"Broadcast WM_FONTCHANGE", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 350, 0, 250, 20, hWnd, (HMENU)(idButtonBroadcastMsg | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hButtonBroadcastMsg, hFont, TRUE);

			//Initialize ListViewFontList
			hListViewFontList = CreateWindow(WC_LISTVIEW, L"FontList", WS_CHILD | WS_VISIBLE | WS_BORDER | LVS_REPORT, 0, 50, rectClient.right - rectClient.left, 300, hWnd, (HMENU)(idListViewFontList | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hListViewFontList, hFont, TRUE);
			DragAcceptFiles(hListViewFontList, TRUE);
			OldListViewProc = GetWindowLongPtr(hListViewFontList, GWLP_WNDPROC);
			SetWindowLongPtr(hListViewFontList, GWLP_WNDPROC, (LONG_PTR)ListViewProc);
			LVCOLUMN lvcolumn1{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectClient.right - rectClient.left) * 7 / 10 , (LPWSTR)L"Font Name" };
			ListView_InsertColumn(hListViewFontList, 0, &lvcolumn1);
			LVCOLUMN lvcolumn2{ LVCF_FMT | LVCF_WIDTH | LVCF_TEXT, LVCFMT_LEFT, (rectClient.right - rectClient.left) * 3 / 10 , (LPWSTR)L"State" };
			ListView_InsertColumn(hListViewFontList, 1, &lvcolumn2);
			ListView_SetExtendedListViewStyle(hListViewFontList, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);

			//Initialize EditMessage
			hEditMessage = CreateWindow(WC_EDIT, NULL, WS_CHILD | WS_VISIBLE | WS_VSCROLL | ES_READONLY | ES_LEFT | ES_MULTILINE, 0, 350, rectClient.right - rectClient.left, rectClient.bottom - rectClient.top - 350, hWnd, (HMENU)(idEditMessage | (UINT_PTR)0), ((LPCREATESTRUCT)lParam)->hInstance, NULL);
			SetWindowFont(hEditMessage, hFont, TRUE);
			Edit_SetText(hEditMessage,
				L"Temporary load and unload fonts to system font table.\r\n"
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
				"\"Broadcast WM_FONTCHANGE\": If checked, broadcast WM_FONTCHANGE message to all top windows.\r\n"
				"\"Font Name\": Names of the fonts added to the list.\r\n"
				"\"State\": State of the font. There are four states, \"Loaded\", \"Load failed\", \"Unloaded\" and \"Unload failed\".\r\n"
				"\r\n"
			);
			ret = 0;	
		}
		break;
	case WM_ACTIVATE:
		{
			if (DragDropHasFonts)	//Process drag-drop font files onto the application icon stage II
			{
				LVITEM lvi{ LVIF_TEXT };
				std::list<FontResource>::iterator iter = FontList.begin();
				for (int i = 0; i < (int)FontList.size(); i++)
				{
					lvi.iItem = i;
					lvi.iSubItem = 0;
					lvi.pszText = (LPWSTR)iter->GetFontPath().c_str();
					ListView_InsertItem(hListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = NULL;
					ListView_SetItem(hListViewFontList, &lvi);
					if (iter->IsLoaded())
					{
						lvi.pszText = (LPWSTR)L"Loaded";
						ListView_SetItem(hListViewFontList, &lvi);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L" successfully opened and loaded\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					}
					else
					{
						lvi.pszText = (LPWSTR)L"Load failed";
						ListView_SetItem(hListViewFontList, &lvi);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"Failed to load ");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					}
					iter++;
				}
				Edit_ReplaceSel(hEditMessage, L"\r\n");
				Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
				DragDropHasFonts = false;
			}
			ret = 0;
		}
		break;
	case WM_COMMAND: 
		{
			ret = ButtonOpenProc(hWnd, msg, wParam, lParam);
			ret = ButtonCloseProc(hWnd, msg, wParam, lParam);
			ret = ButtonCloseAllProc(hWnd, msg, wParam, lParam);
			ret = ButtonLoadProc(hWnd, msg, wParam, lParam);
			ret = ButtonLoadAllProc(hWnd, msg, wParam, lParam);
			ret = ButtonUnloadProc(hWnd, msg, wParam, lParam);
			ret = ButtonUnloadAllProc(hWnd, msg, wParam, lParam);
		}
		break;
	case WM_CLOSE:
		{
			//Unload all fonts
			bool bIsFontListChanged{ false };
			bool bIsUnloadSuccessful{ true };
			LVITEM lvi{ LVIF_TEXT, 0, 1 };
			FontList.reverse();
			std::list<FontResource>::iterator iter = FontList.begin();
			for (int i = ListView_GetItemCount(hListViewFontList) - 1; i >= 0; i--)
			{
				if (iter->IsLoaded())
				{
					if (iter->Unload())
					{
						bIsFontListChanged = true;
						ListView_DeleteItem(hListViewFontList, i);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L" successfully unloaded and closed\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						iter = FontList.erase(iter);
						continue;
					}
					else
					{
						bIsUnloadSuccessful = false;
						lvi.iItem = i;
						lvi.pszText = (LPWSTR)L"Unload failed";
						ListView_SetItem(hListViewFontList, &lvi);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"Failed to unload ");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					}
				}
				else
				{
					ListView_DeleteItem(hListViewFontList, i);
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, L" successfully closed\r\n");
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					iter = FontList.erase(iter);
					continue;
				}
				iter++;
			}
			FontList.reverse();
			if ((Button_GetCheck(hButtonBroadcastMsg) == BST_CHECKED && bIsFontListChanged))
			{
				FORWARD_WM_FONTCHANGE(HWND_BROADCAST, SendMessage);
				Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
				Edit_ReplaceSel(hEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
				Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
			}
			Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
			Edit_ReplaceSel(hEditMessage, L"\r\n");
			Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));

			//If some fonts are not unloaded, prompt user whether inisit to exit.
			if (bIsUnloadSuccessful)
			{
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
					break;
				default:
					break;
				}
			}
		}
		break;
	case WM_DESTROY:
		{
			PostQuitMessage(0);
			ret = 0;
		}
		break;
	default:
		{
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
					OPENFILENAME ofn{ sizeof(ofn), hWnd, NULL, L"Font Files(*.ttf;*.ttc;*.otf)\0*.ttf;*.ttc;*.otf\0", NULL, 0, 0, lpszOpenFileNames.get(), MAX_PATH * MAX_PATH, NULL, 0, NULL,NULL,OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST | OFN_ALLOWMULTISELECT, 0, 0, NULL, NULL, NULL, NULL, nullptr, 0, 0 };
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
						LVITEM lvi{ LVIF_TEXT };
						for (int i = 0; i < (int)NewFontList.size(); i++)
						{
							lvi.iItem = ListView_GetItemCount(hListViewFontList);
							lvi.iSubItem = 0;
							lvi.pszText = (LPWSTR)(NewFontList[i].c_str());
							ListView_InsertItem(hListViewFontList, &lvi);
							lvi.iSubItem = 1;
							lvi.pszText = NULL;
							ListView_SetItem(hListViewFontList, &lvi);
							FontList.push_back(NewFontList[i]);
							Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							Edit_ReplaceSel(hEditMessage, NewFontList[i].c_str());
							Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							Edit_ReplaceSel(hEditMessage, L" successfully opened\r\n");
							Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						}
						Edit_ReplaceSel(hEditMessage, L"\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
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
					bool bIsFontListChanged{ false };
					LVITEM lvi{ LVIF_TEXT, 0, 1 };
					FontList.reverse();
					std::list<FontResource>::iterator iter = FontList.begin();
					for (int i = ListView_GetItemCount(hListViewFontList) - 1; i >= 0; i--)
					{
						if ((ListView_GetItemState(hListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED))
						{
							if (iter->IsLoaded())
							{
								if (iter->Unload())
								{
									bIsFontListChanged = true;
									ListView_DeleteItem(hListViewFontList, i);
									Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
									Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
									Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
									Edit_ReplaceSel(hEditMessage, L" successfully unloaded and closed\r\n");
									Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
									iter = FontList.erase(iter);
									continue;
								}
								else
								{
									lvi.iItem = i;
									lvi.pszText = (LPWSTR)L"Unload failed";
									ListView_SetItem(hListViewFontList, &lvi);
									Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
									Edit_ReplaceSel(hEditMessage, L"Failed to unload ");
									Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
									Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
									Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
									Edit_ReplaceSel(hEditMessage, L"\r\n");
									Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								}
							}
							else
							{
								ListView_DeleteItem(hListViewFontList, i);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L" successfully closed\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								iter = FontList.erase(iter);
								continue;
							}
						}
						iter++;
					}
					FontList.reverse();
					if ((Button_GetCheck(hButtonBroadcastMsg) == BST_CHECKED && bIsFontListChanged))
					{
						FORWARD_WM_FONTCHANGE(HWND_BROADCAST, SendMessage);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					}
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, L"\r\n");
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
				}
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
					bool bIsFontListChanged{ false };
					LVITEM lvi{ LVIF_TEXT, 0, 1 };
					FontList.reverse();
					std::list<FontResource>::iterator iter = FontList.begin();
					for (int i = ListView_GetItemCount(hListViewFontList) - 1; i >= 0; i--)
					{
						if (iter->IsLoaded())
						{
							if (iter->Unload())
							{
								bIsFontListChanged = true;
								ListView_DeleteItem(hListViewFontList, i);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L" successfully unloaded and closed\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								iter = FontList.erase(iter);
								continue;
							}
							else
							{
								lvi.iItem = i;
								lvi.pszText = (LPWSTR)L"Unload failed";
								ListView_SetItem(hListViewFontList, &lvi);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"Failed to unload ");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							}
						}
						else
						{
							ListView_DeleteItem(hListViewFontList, i);
							Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
							Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							Edit_ReplaceSel(hEditMessage, L" successfully closed\r\n");
							Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							iter = FontList.erase(iter);
							continue;
						}
						iter++;
					}
					FontList.reverse();
					if ((Button_GetCheck(hButtonBroadcastMsg) == BST_CHECKED && bIsFontListChanged))
					{
						FORWARD_WM_FONTCHANGE(HWND_BROADCAST, SendMessage);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					}
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, L"\r\n");
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
				}
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
					bool bIsFontListChanged{ false };
					LVITEM lvi{ LVIF_TEXT, 0, 1 };
					std::list<FontResource>::iterator iter = FontList.begin();
					for (int i = 0; i < ListView_GetItemCount(hListViewFontList); i++)
					{
						if ((ListView_GetItemState(hListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED) && (!(iter->IsLoaded())))
						{
							lvi.iItem = i;
							if (iter->Load())
							{
								bIsFontListChanged = true;
								lvi.pszText = (LPWSTR)L"Loaded";
								ListView_SetItem(hListViewFontList, &lvi);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L" successfully loaded\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							}
							else
							{
								lvi.pszText = (LPWSTR)L"Load failed";
								ListView_SetItem(hListViewFontList, &lvi);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"Failed to load ");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							}

						}
						iter++;
					}
					if ((Button_GetCheck(hButtonBroadcastMsg) == BST_CHECKED && bIsFontListChanged))
					{
						FORWARD_WM_FONTCHANGE(HWND_BROADCAST, SendMessage);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					}
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, L"\r\n");
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
				}
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
					bool bIsFontListChanged{ false };
					LVITEM lvi{ LVIF_TEXT, 0, 1 };
					std::list<FontResource>::iterator iter = FontList.begin();
					for (int i = 0; i < (int)FontList.size(); i++)
					{
						lvi.iItem = i;
						if (!(iter->IsLoaded()))
						{
							if (iter->Load())
							{
								bIsFontListChanged = true;
								lvi.pszText = (LPWSTR)L"Loaded";
								ListView_SetItem(hListViewFontList, &lvi);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L" successfully loaded\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							}
							else
							{
								lvi.pszText = (LPWSTR)L"Load failed";
								ListView_SetItem(hListViewFontList, &lvi);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"Failed to load ");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							}
						}
						iter++;
					}
					if ((Button_GetCheck(hButtonBroadcastMsg) == BST_CHECKED && bIsFontListChanged))
					{
						FORWARD_WM_FONTCHANGE(HWND_BROADCAST, SendMessage);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					}
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, L"\r\n");
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
				}
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
					bool bIsFontListChanged{ false };
					LVITEM lvi{ LVIF_TEXT, 0, 1 };
					std::list<FontResource>::iterator iter = FontList.begin();
					for (int i = 0; i < ListView_GetItemCount(hListViewFontList); i++)
					{
						if ((ListView_GetItemState(hListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED) && (iter->IsLoaded()))
						{
							lvi.iItem = i;
							if (iter->Unload())
							{
								bIsFontListChanged = true;
								lvi.pszText = (LPWSTR)L"Unloaded";
								ListView_SetItem(hListViewFontList, &lvi);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L" successfully unloaded\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							}
							else
							{
								lvi.pszText = (LPWSTR)L"Unload failed";
								ListView_SetItem(hListViewFontList, &lvi);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"Failed to unload ");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							}
						}
						iter++;
					}
					if ((Button_GetCheck(hButtonBroadcastMsg) == BST_CHECKED && bIsFontListChanged))
					{
						FORWARD_WM_FONTCHANGE(HWND_BROADCAST, SendMessage);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					}
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, L"\r\n");
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
				}
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
					bool bIsFontListChanged{ false };
					LVITEM lvi{ LVIF_TEXT, 0, 1 };
					std::list<FontResource>::iterator iter = FontList.begin();
					for (int i = 0; i < (int)FontList.size(); i++)
					{
						lvi.iItem = i;
						if (iter->IsLoaded())
						{
							if (iter->Unload())
							{
								bIsFontListChanged = true;
								lvi.pszText = (LPWSTR)L"Unloaded";
								ListView_SetItem(hListViewFontList, &lvi);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L" successfully unloaded\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							}
							else
							{
								lvi.pszText = (LPWSTR)L"Unload failed";
								ListView_SetItem(hListViewFontList, &lvi);
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"Failed to unload ");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, iter->GetFontPath().c_str());
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
								Edit_ReplaceSel(hEditMessage, L"\r\n");
								Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
							}
						}
						iter++;
					}
					if ((Button_GetCheck(hButtonBroadcastMsg) == BST_CHECKED && bIsFontListChanged))
					{
						FORWARD_WM_FONTCHANGE(HWND_BROADCAST, SendMessage);
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
						Edit_ReplaceSel(hEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
						Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					}
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, L"\r\n");
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
				}
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
			LVITEM lvi{ LVIF_TEXT };
			UINT nFileCount{ DragQueryFile((HDROP)wParam, 0xFFFFFFFF, NULL, 0) };
			for (UINT i = 0; i < nFileCount; i++)
			{
				UINT nSize = DragQueryFile((HDROP)wParam, i, NULL, 0) + 1;
				std::unique_ptr<WCHAR[]> lpszFileName(new WCHAR[nSize]{});
				DragQueryFile((HDROP)wParam, i, lpszFileName.get(), nSize);
				if (PathMatchSpec(lpszFileName.get(), L"*.ttf") || PathMatchSpec(lpszFileName.get(), L"*.ttc") || PathMatchSpec(lpszFileName.get(), L"*.otf"))
				{
					lvi.iItem = ListView_GetItemCount(hListViewFontList);
					lvi.iSubItem = 0;
					lvi.pszText = lpszFileName.get();
					ListView_InsertItem(hListViewFontList, &lvi);
					lvi.iSubItem = 1;
					lvi.pszText = NULL;
					ListView_SetItem(hListViewFontList, &lvi);
					FontList.push_back(lpszFileName.get());
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, lpszFileName.get());
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
					Edit_ReplaceSel(hEditMessage, L" successfully opened\r\n");
					Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
				}
			}
			Edit_ReplaceSel(hEditMessage, L"\r\n");
			Edit_SetSel(hEditMessage, Edit_GetTextLength(hEditMessage), Edit_GetTextLength(hEditMessage));
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
