#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <process.h>
#include <list>
#include <sstream>
#include "FontResource.h"
#include "Globals.h"

//Process drag-drop font files onto the application icon stage II working thread
void DragDropWorkingThreadProc(void* lpParameter)
{
	LVITEM lvi{ LVIF_TEXT };
	std::list<FontResource>::iterator iter{ FontList.begin() };
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
			ListView_SetItemState(hWndListViewFontList, i, LVIS_SELECTED, LVIS_SELECTED);
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
}

//Unload all fonts working thread
void CloseWorkingThreadProc(void* lpParameter)
{
	bool bIsFontListChanged{ false };
	bool bIsUnloadSuccessful{ true };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::wstringstream Message{};
	int iMessageLength{};
	FontList.reverse();
	std::list<FontResource>::iterator iter{ FontList.begin() };
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
		if (hTargetProcessWatchThread)
		{
#pragma warning(suppress: 6387)
			PostThreadMessage(GetThreadId(hTargetProcessWatchThread), UM_TERMINATEWATCHPROCESS, NULL, NULL);
			WaitForSingleObject(hTargetProcessWatchThread, INFINITE);
		}
		PostMessage(hWndMainWindow, UM_CLOSEWORKINGTHREADTERMINATED, (WPARAM)TRUE, NULL);
	}
	else
	{
		PostMessage(hWndMainWindow, UM_CLOSEWORKINGTHREADTERMINATED, (WPARAM)FALSE, NULL);
	}
}

//Unload and close selected fonts working thread
void ButtonCloseWorkingThreadProc(void* lpParameter)
{
	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::wstringstream Message{};
	int iMessageLength{};
	FontList.reverse();
	std::list<FontResource>::iterator iter{ FontList.begin() };
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
}

//Load selected fonts working thread
void ButtonLoadWorkingThreadProc(void* lpParameter)
{
	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::wstringstream Message{};
	int iMessageLength{};
	std::list<FontResource>::iterator iter{ FontList.begin() };
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
}

//Unload selected fonts working thread
void ButtonUnloadWorkingThreadProc(void* lpParameter)
{
	bool bIsFontListChanged{ false };
	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::wstringstream Message{};
	int iMessageLength{};
	std::list<FontResource>::iterator iter{ FontList.begin() };
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
}

//Target process watch thread
void TargetProcessWatchThreadProc(void* lpParameter)
{
	//Force create message queue for current thread
	MSG message;
	PeekMessage(&message, NULL, WM_USER, WM_USER, PM_NOREMOVE);

	//Wait for target process or termination message
	try
	{
		while (true)
		{
			switch (MsgWaitForMultipleObjects(1, &TargetProcessInfo.hProcess, FALSE, INFINITE, QS_ALLPOSTMESSAGE))
			{
			case WAIT_OBJECT_0:
				throw 1;
				break;
			case WAIT_OBJECT_0 + 1:
				{
					while (PeekMessage(&message, NULL, WM_USER + 0x100, WM_USER + 0x105, PM_REMOVE))
					{
						switch (message.message)
						{
						case UM_TERMINATEWATCHPROCESS:
							{
								_endthread();
							}
							break;
						default:
							break;
						}
					}
				}
				break;
			default:
				break;
			}
		}
	}
	catch (const int)
	{
	}
	
	//If target process terminates, clear FontList and ListViewFontList
	EnableMenuItem(GetSystemMenu(hWndMainWindow, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
	DisableAllButtons();

	FontResource::RegisterAddRemoveFontProc(NullAddFontProc, NullRemoveFontProc);
	FontList.clear();
	ListView_DeleteAllItems(hWndListViewFontList);

	std::wstringstream Message{};
	int iMessageLength{};

	Message << L"Target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L") terminated.\r\n\r\n";
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

	FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);
	Button_SetText(hWndButtonSelectProcess, L"Click to select process");
	
	//Close TargetProcessInfo.hProcess
	CloseHandle(TargetProcessInfo.hProcess);
	TargetProcessInfo.hProcess = NULL;

	SendMessage(hWndMainWindow, UM_WATCHPROCESSTERMINATED, NULL, NULL);
}
