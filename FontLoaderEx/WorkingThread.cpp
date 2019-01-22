#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <list>
#include <sstream>
#include "FontResource.h"
#include "Globals.h"

DWORD DragDropWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	LVITEM lvi{ LVIF_TEXT };
	std::list<FontResource>::iterator iter = FontList.begin();
	std::wstringstream Message{};
	int iMessageLength{};
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
			Message.str(L"");
			Message << iter->GetFontPath() << L" successfully opened and loaded\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		}
		else
		{
			lvi.pszText = (LPWSTR)L"Load failed";
			ListView_SetItem(hWndListViewFontList, &lvi);
			Message.str(L"");
			Message << L"Failed to load " << iter->GetFontPath() << L"\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		}
		iter++;
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
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
	std::wstringstream Message{};
	int iMessageLength{};
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
				Message.str(L"");
				Message << iter->GetFontPath() << L" successfully unloaded and closed\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
				iter = FontList.erase(iter);
				continue;
			}
			else
			{
				bIsUnloadSuccessful = false;
				lvi.iItem = i;
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Message.str(L"");
				Message << L"Failed to unload " << iter->GetFontPath() << L"\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
		}
		else
		{
			ListView_DeleteItem(hWndListViewFontList, i);
			Message.str(L"");
			Message << iter->GetFontPath() << L" successfully closed\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			iter = FontList.erase(iter);
			continue;
		}
		iter++;
	}
	FontList.reverse();
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

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
	std::wstringstream Message{};
	int iMessageLength{};
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
					Message.str(L"");
					Message << iter->GetFontPath() << L" successfully unloaded and closed\r\n";
					iMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
					iter = FontList.erase(iter);
					continue;
				}
				else
				{
					lvi.iItem = i;
					lvi.pszText = (LPWSTR)L"Unload failed";
					ListView_SetItem(hWndListViewFontList, &lvi);
					Message.str(L"");
					Message << L"Failed to unload " << iter->GetFontPath() << L"\r\n";
					iMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
				}
			}
			else
			{
				ListView_DeleteItem(hWndListViewFontList, i);
				Message.str(L"");
				Message << iter->GetFontPath() << L" successfully closed\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
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
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonCloseAllWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::wstringstream Message{};
	int iMessageLength{};
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
				Message.str(L"");
				Message << iter->GetFontPath() << L" successfully unloaded and closed\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
				iter = FontList.erase(iter);
				continue;
			}
			else
			{
				lvi.iItem = i;
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Message.str(L"");
				Message << L"Failed to unload " << iter->GetFontPath() << L"\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
		}
		else
		{
			ListView_DeleteItem(hWndListViewFontList, i);
			Message.str(L"");
			Message << iter->GetFontPath() << L" successfully closed\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			iter = FontList.erase(iter);
			continue;
		}
		iter++;
	}
	FontList.reverse();
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonLoadWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::wstringstream Message{};
	int iMessageLength{};
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
				Message.str(L"");
				Message << iter->GetFontPath() << L" successfully loaded\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Load failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Message.str(L"");
				Message << L"Failed to load " << iter->GetFontPath() << L"\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}

		}
		iter++;
	}
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonLoadAllWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::wstringstream Message{};
	int iMessageLength{};
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
				Message.str(L"");
				Message << iter->GetFontPath() << L" successfully loaded\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Load failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Message.str(L"");
				Message << L"Failed to load " << iter->GetFontPath() << L"\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
		}
		iter++;
	}
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonUnloadWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::wstringstream Message{};
	int iMessageLength{};
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
				Message.str(L"");
				Message << iter->GetFontPath() << L" successfully unloaded\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Message.str(L"");
				Message << L"Failed to unload " << iter->GetFontPath() << L"\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
		}
		iter++;
	}
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}

DWORD ButtonUnloadAllWorkingThreadProc(void* lpParameter)
{
	EnterCriticalSection(&CriticalSection);

	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::wstringstream Message{};
	int iMessageLength{};
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
				Message.str(L"");
				Message << iter->GetFontPath() << L" successfully unloaded\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Message.str(L"");
				Message << L"Failed to unload " << iter->GetFontPath() << L"\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
		}
		iter++;
	}
	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMainWindow, UM_WORKINGTHREADTERMINATED, NULL, NULL);

	LeaveCriticalSection(&CriticalSection);
	return 0;
}
