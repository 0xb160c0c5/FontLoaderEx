#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <list>
#include "FontResource.h"
#include "Globals.h"

DWORD DragDropWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	LVITEM lvi{ LVIF_TEXT };
	std::list<FontResource>::iterator iter = FontList.begin();
	for (int i = 0; i < (int)FontList.size(); i++)
	{
		lvi.iItem = i;
		lvi.iSubItem = 0;
		lvi.pszText = (LPWSTR)(iter->GetFontPath().c_str());
		ListView_InsertItem(hWndListViewFontList, &lvi);
		lvi.iSubItem = 1;
		if (iter->Load())
		{
			lvi.pszText = (LPWSTR)L"Loaded";
			ListView_SetItem(hWndListViewFontList, &lvi);
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			Edit_ReplaceSel(hWndEditMessage, L" successfully opened and loaded\r\n");
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
		}
		else
		{
			lvi.pszText = (LPWSTR)L"Load failed";
			ListView_SetItem(hWndListViewFontList, &lvi);
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			Edit_ReplaceSel(hWndEditMessage, L"Failed to load ");
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			Edit_ReplaceSel(hWndEditMessage, L"\r\n");
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
		}
		iter++;
	}
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	DragDropHasFonts = false;

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD CloseWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	bool bIsUnloadSuccessful{ true };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	FontList.reverse();
	std::list<FontResource>::iterator iter = FontList.begin();
	for (int i = ListView_GetItemCount(hWndListViewFontList) - 1; i >= 0; i--)
	{
		if (iter->IsLoaded())
		{
			if (iter->Unload())
			{
				bIsFontListChanged = true;
				ListView_DeleteItem(hWndListViewFontList, i);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L" successfully unloaded and closed\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				iter = FontList.erase(iter);
				continue;
			}
			else
			{
				bIsUnloadSuccessful = false;
				lvi.iItem = i;
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"Failed to unload ");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}
		}
		else
		{
			ListView_DeleteItem(hWndListViewFontList, i);
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			Edit_ReplaceSel(hWndEditMessage, L" successfully closed\r\n");
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			iter = FontList.erase(iter);
			continue;
		}
		iter++;
	}
	FontList.reverse();
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	}
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));

	//If some fonts are not unloaded, prompt user whether inisit to exit.
	if (bIsUnloadSuccessful)
	{
		PostMessage(hWndMainWindow, UM_CLOSEWORKINGTHREADTERMINATED, (WPARAM)TRUE, NULL);
	}
	else
	{
		PostMessage(hWndMainWindow, UM_CLOSEWORKINGTHREADTERMINATED, (WPARAM)FALSE, NULL);
	}

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonCloseWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	FontList.reverse();
	std::list<FontResource>::iterator iter = FontList.begin();
	for (int i = ListView_GetItemCount(hWndListViewFontList) - 1; i >= 0; i--)
	{
		if ((ListView_GetItemState(hWndListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED))
		{
			if (iter->IsLoaded())
			{
				if (iter->Unload())
				{
					bIsFontListChanged = true;
					ListView_DeleteItem(hWndListViewFontList, i);
					Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
					Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
					Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
					Edit_ReplaceSel(hWndEditMessage, L" successfully unloaded and closed\r\n");
					Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
					iter = FontList.erase(iter);
					continue;
				}
				else
				{
					lvi.iItem = i;
					lvi.pszText = (LPWSTR)L"Unload failed";
					ListView_SetItem(hWndListViewFontList, &lvi);
					Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
					Edit_ReplaceSel(hWndEditMessage, L"Failed to unload ");
					Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
					Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
					Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
					Edit_ReplaceSel(hWndEditMessage, L"\r\n");
					Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				}
			}
			else
			{
				ListView_DeleteItem(hWndListViewFontList, i);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L" successfully closed\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				iter = FontList.erase(iter);
				continue;
			}
		}
		iter++;
	}
	FontList.reverse();
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	}
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonCloseAllWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	FontList.reverse();
	std::list<FontResource>::iterator iter = FontList.begin();
	for (int i = ListView_GetItemCount(hWndListViewFontList) - 1; i >= 0; i--)
	{
		if (iter->IsLoaded())
		{
			if (iter->Unload())
			{
				bIsFontListChanged = true;
				ListView_DeleteItem(hWndListViewFontList, i);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L" successfully unloaded and closed\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				iter = FontList.erase(iter);
				continue;
			}
			else
			{
				lvi.iItem = i;
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"Failed to unload ");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}
		}
		else
		{
			ListView_DeleteItem(hWndListViewFontList, i);
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			Edit_ReplaceSel(hWndEditMessage, L" successfully closed\r\n");
			Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			iter = FontList.erase(iter);
			continue;
		}
		iter++;
	}
	FontList.reverse();
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	}
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonLoadWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::list<FontResource>::iterator iter = FontList.begin();
	for (int i = 0; i < ListView_GetItemCount(hWndListViewFontList); i++)
	{
		if ((ListView_GetItemState(hWndListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED) && (!(iter->IsLoaded())))
		{
			lvi.iItem = i;
			if (iter->Load())
			{
				bIsFontListChanged = true;
				lvi.pszText = (LPWSTR)L"Loaded";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L" successfully loaded\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Load failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"Failed to load ");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}

		}
		iter++;
	}
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	}
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonLoadAllWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

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
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L" successfully loaded\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Load failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"Failed to load ");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}
		}
		iter++;
	}
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	}
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonUnloadWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::list<FontResource>::iterator iter = FontList.begin();
	for (int i = 0; i < ListView_GetItemCount(hWndListViewFontList); i++)
	{
		if ((ListView_GetItemState(hWndListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED) && (iter->IsLoaded()))
		{
			lvi.iItem = i;
			if (iter->Unload())
			{
				bIsFontListChanged = true;
				lvi.pszText = (LPWSTR)L"Unloaded";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L" successfully unloaded\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"Failed to unload ");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}
		}
		iter++;
	}
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	}
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonUnloadAllWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

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
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L" successfully unloaded\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"Failed to unload ");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, iter->GetFontPath().c_str());
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");
				Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
			}
		}
		iter++;
	}
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
		Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	}
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
	Edit_SetSel(hWndEditMessage, Edit_GetTextLength(hWndEditMessage), Edit_GetTextLength(hWndEditMessage));

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}
