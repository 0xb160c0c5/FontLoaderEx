#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <list>
#include <sstream>
#include "FontResource.h"
#include "Globals.h"
#include "resource.h"

// Process drag-drop font files onto the application icon stage II working thread
void DragDropWorkerThreadProc(void* lpParameter)
{
	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

	std::wstringstream Message{};
	int iMessageLength{};

	LVITEM lvi{ LVIF_TEXT };
	std::list<FontResource>::iterator iter{ FontList.begin() };
	for (int i = 0; i < (int)FontList.size(); i++)
	{
		lvi.iItem = i;
		lvi.iSubItem = 0;
		lvi.pszText = (LPWSTR)(iter->GetFontName().c_str());
		ListView_InsertItem(hWndListViewFontList, &lvi);
		lvi.iSubItem = 1;
		if (iter->Load())
		{
			lvi.pszText = (LPWSTR)L"Loaded";
			ListView_SetItem(hWndListViewFontList, &lvi);
			ListView_SetItemState(hWndListViewFontList, i, LVIS_SELECTED, LVIS_SELECTED);

			Message << iter->GetFontName() << L" opened and successfully loaded\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			Message.str(L"");
		}
		else
		{
			lvi.pszText = (LPWSTR)L"Load failed";
			ListView_SetItem(hWndListViewFontList, &lvi);

			Message << L"Opened but failed to load " << iter->GetFontName() << L"\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			Message.str(L"");
		}
		iter++;
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	bDragDropHasFonts = false;

	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), TRUE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), TRUE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), TRUE);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_ENABLED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_ENABLED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_ENABLED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_ENABLED);

	PostMessage(hWndMain, (UINT)USERMESSAGE::DRAGDROPWORKINGTHREADTERMINATED, NULL, NULL);
}

// Close working thread
void CloseWorkerThreadProc(void* lpParameter)
{
	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	HWND hWndButtonBroadcastWM_FONTCHANGE{ GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE) };

	std::wstringstream Message{};
	int iMessageLength{};

	bool bIsUnloadSuccessful{ true };
	bool bIsFontListChanged{ false };

	LVITEM lvi{ LVIF_TEXT, 0, 1 };
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

				Message << iter->GetFontName() << L" successfully unloaded and closed\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
				Message.str(L"");

				iter = FontList.erase(iter);
			}
			else
			{
				bIsUnloadSuccessful = false;

				lvi.iItem = i;
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);

				Message << L"Failed to unload " << iter->GetFontName() << L"\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
				Message.str(L"");

				iter++;
			}
		}
		else
		{
			ListView_DeleteItem(hWndListViewFontList, i);

			Message << iter->GetFontName() << L" closed\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			Message.str(L"");

			iter = FontList.erase(iter);
		}
	}
	FontList.reverse();

	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged && !TargetProcessInfo.hProcess)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	// If some fonts are not unloaded, prompt user whether inisit to exit.
	PostMessage(hWndMain, (UINT)USERMESSAGE::CLOSEWORKINGTHREADTERMINATED, (WPARAM)bIsUnloadSuccessful, NULL);
}

// Unload and close selected fonts working thread
void ButtonCloseWorkerThreadProc(void* lpParameter)
{
	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	HWND hWndButtonBroadcastWM_FONTCHANGE{ GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE) };

	std::wstringstream Message{};
	int iMessageLength{};

	bool bIsFontListChanged{ false };

	LVITEM lvi{ LVIF_TEXT, 0, 1 };
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

					Message << iter->GetFontName() << L" successfully unloaded and closed\r\n";
					iMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
					Message.str(L"");

					iter = FontList.erase(iter);
				}
				else
				{
					lvi.iItem = i;
					lvi.pszText = (LPWSTR)L"Unload failed";
					ListView_SetItem(hWndListViewFontList, &lvi);

					Message << L"Failed to unload " << iter->GetFontName() << L"\r\n";
					iMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
					Message.str(L"");

					iter++;
				}
			}
			else
			{
				ListView_DeleteItem(hWndListViewFontList, i);

				Message << iter->GetFontName() << L" closed\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
				Message.str(L"");

				iter = FontList.erase(iter);
			}
		}
		else
		{
			iter++;
		}
	}
	FontList.reverse();

	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged && !TargetProcessInfo.hProcess)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMain, (UINT)USERMESSAGE::BUTTONCLOSEWORKINGTHREADTERMINATED, NULL, NULL);
}

// Load selected fonts working thread
void ButtonLoadWorkerThreadProc(void* lpParameter)
{
	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	HWND hWndButtonBroadcastWM_FONTCHANGE{ GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE) };

	bool bIsFontListChanged{ false };

	std::wstringstream Message{};
	int iMessageLength{};

	LVITEM lvi{ LVIF_TEXT, 0, 1 };
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

				Message << iter->GetFontName() << L" successfully loaded\r\n";
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Load failed";
				ListView_SetItem(hWndListViewFontList, &lvi);

				Message << L"Failed to load " << iter->GetFontName() << L"\r\n";
			}
		}
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		Message.str(L"");

		iter++;
	}

	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged && !TargetProcessInfo.hProcess)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMain, (UINT)USERMESSAGE::BUTTONLOADWORKINGTHREADTERMINATED, NULL, NULL);
}

// Unload selected fonts working thread
void ButtonUnloadWorkerThreadProc(void* lpParameter)
{
	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	HWND hWndButtonBroadcastWM_FONTCHANGE{ GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE) };

	bool bIsFontListChanged{ false };

	std::wstringstream Message{};
	int iMessageLength{};

	LVITEM lvi{ LVIF_TEXT, 0, 1 };
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

				Message << iter->GetFontName() << L" successfully unloaded\r\n";
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);

				Message << L"Failed to unload " << iter->GetFontName() << L"\r\n";
			}
		}
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		Message.str(L"");

		iter++;
	}

	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged && !TargetProcessInfo.hProcess)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMain, (UINT)USERMESSAGE::BUTTONUNLOADWORKINGTHREADTERMINATED, NULL, NULL);
}

// Target process watch thread
unsigned int __stdcall TargetProcessWatchThreadProc(void* lpParameter)
{
	// Wait for target process or termination event
	HANDLE handles[]{ TargetProcessInfo.hProcess, hEventTerminateWatchThread };
	switch (WaitForMultipleObjects(2, handles, FALSE, INFINITE))
	{
	case WAIT_OBJECT_0:
		break;
	case WAIT_OBJECT_0 + 1:
		{
			return 0;
		}
		break;
	default:
		break;
	}

	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

	// Disable controls
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
	EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

	// Clear FontList and ListViewFontList
	FontResource::RegisterAddRemoveFontProc(NullAddFontProc, NullRemoveFontProc);
	FontList.clear();
	ListView_DeleteAllItems(GetDlgItem(hWndMain, (int)ID::ListViewFontList));

	std::wstringstream Message{};
	int iMessageLength{};

	Message << L"Target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L") terminated.\r\n\r\n";
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

	// Register default AddFont() and RemoveFont() procedures
	FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);

	// Revert to default caption
	Button_SetText(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), L"Select process");

	// Close handle to target process and duplicated handles
	CloseHandle(TargetProcessInfo.hProcess);
	TargetProcessInfo.hProcess = NULL;
	CloseHandle(hCurrentProcessDuplicated);
	CloseHandle(hTargetProcessDuplicated);

	// Enable controls
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), TRUE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
	EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);

	PostMessage(hWndMain, (UINT)USERMESSAGE::WATCHTHREADTERMINATED, NULL, NULL);

	return 0;
}

// Proxy process and target process watch thread
unsigned int __stdcall ProxyAndTargetProcessWatchThreadProc(void* lpParameter)
{
	// Wait for proxy process or target process or termination event
	enum class Termination { Proxy, Target };
	Termination t{};

	HANDLE handles[]{ piProxyProcess.hProcess, TargetProcessInfo.hProcess, hEventTerminateWatchThread };
	switch (WaitForMultipleObjects(3, handles, FALSE, INFINITE))
	{
	case WAIT_OBJECT_0:
		{
			t = Termination::Proxy;
		}
		break;
	case WAIT_OBJECT_0 + 1:
		{
			t = Termination::Target;
		}
		break;
	case WAIT_OBJECT_0 + 2:
		{
			return 0;
		}
		break;
	default:
		break;
	}

	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

	// Disable controls
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
	EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);

	// Clear FontList and ListViewFontList
	FontResource::RegisterAddRemoveFontProc(NullAddFontProc, NullRemoveFontProc);
	FontList.clear();
	ListView_DeleteAllItems(GetDlgItem(hWndMain, (int)ID::ListViewFontList));

	std::wstringstream Message{};
	int iMessageLength{};

	switch (t)
	{
		// If proxy process terminates, just print message
	case Termination::Proxy:
		{
			Message << L"FontLoaderExProxy(" << piProxyProcess.dwProcessId << L") terminated.\r\n\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		}
		break;
		// If target process termiantes, print message and terminate proxy process
	case Termination::Target:
		{
			Message << L"Target process " << TargetProcessInfo.ProcessName << L"(" << TargetProcessInfo.ProcessID << L") terminated.\r\n\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			Message.str(L"");

			COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
			FORWARD_WM_COPYDATA(hWndProxy, hWndMain, &cds, SendMessage);
			WaitForSingleObject(piProxyProcess.hProcess, INFINITE);
			Message << L"FontLoaderExProxy(" << piProxyProcess.dwProcessId << L") successfully terminated.\r\n\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		}
		break;
	default:
		break;
	}

	// Terminate message thread
	SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
	WaitForSingleObject(hMessageThread, INFINITE);

	// Register default AddFont() and RemoveFont() procedures
	FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);

	// Revert to default caption
	Button_SetText(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), L"Click to select process");

	// Close handles to proxy process and target process, duplicated handles and synchronization objects
	CloseHandle(piProxyProcess.hThread);
	CloseHandle(piProxyProcess.hProcess);
	CloseHandle(TargetProcessInfo.hProcess);
	piProxyProcess.hProcess = NULL;
	TargetProcessInfo.hProcess = NULL;
	CloseHandle(hCurrentProcessDuplicated);
	CloseHandle(hTargetProcessDuplicated);
	CloseHandle(hEventProxyAddFontFinished);
	CloseHandle(hEventProxyRemoveFontFinished);

	// Enable controls
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), TRUE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), TRUE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), TRUE);
	EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), TRUE);
	EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_ENABLED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_LOAD, MF_BYCOMMAND | MF_GRAYED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_UNLOAD, MF_BYCOMMAND | MF_GRAYED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_CLOSE, MF_BYCOMMAND | MF_GRAYED);
	EnableMenuItem(hMenuContextListViewFontList, ID_MENU_SELECTALL, MF_BYCOMMAND | MF_GRAYED);

	PostMessage(hWndMain, (UINT)USERMESSAGE::WATCHTHREADTERMINATED, NULL, NULL);

	return 0;
}

LRESULT CALLBACK MessageWndProc(HWND hWnd, UINT MsgMessage, WPARAM wParam, LPARAM lParam);

HWND hWndMessage{};

// Message thread
unsigned int __stdcall MessageThreadProc(void* lpParameter)
{
	// Force windows to create message queue for current thread
	MSG Msg{};
	PeekMessage(&Msg, NULL, 0, 0, PM_NOREMOVE);

	// Create message-only window
	WNDCLASS wc{ 0, MessageWndProc, 0, 0, (HINSTANCE)lpParameter, NULL, NULL, NULL, NULL, L"FontLoaderExMessage" };
	if (!RegisterClass(&wc))
	{
		SetEvent(hEventMessageThreadNotReady);

		return 0xffffffff;
	}
	if (!(hWndMessage = CreateWindow(L"FontLoaderExMessage", L"FontLoaderExMessage", NULL, 0, 0, 0, 0, HWND_MESSAGE, NULL, (HINSTANCE)lpParameter, NULL)))
	{
		SetEvent(hEventMessageThreadNotReady);

		return 0xffffffff;
	}

	SetEvent(hEventMessageThreadReady);

	BOOL bRet{};
	while ((bRet = GetMessage(&Msg, NULL, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			return (unsigned int)GetLastError();
		}
		else
		{
			DispatchMessage(&Msg);
		}
	}
	UnregisterClass(L"FontLoaderExMessage", (HINSTANCE)lpParameter);
	return (unsigned int)Msg.wParam;
}

LRESULT CALLBACK MessageWndProc(HWND hWndMessage, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	switch ((USERMESSAGE)Msg)
	{
	case USERMESSAGE::TERMINATEMESSAGETHREAD:
		{
			DestroyWindow(hWndMessage);
		}
		break;
	default:
		break;
	}
	switch (Msg)
	{
	case WM_COPYDATA:
		{
			switch ((COPYDATA)((PCOPYDATASTRUCT)lParam)->dwData)
			{
				// Get proxy SeDebugPrivilege enabling result
			case COPYDATA::PROXYPROCESSDEBUGPRIVILEGEENABLINGFINISHED:
				{
					ProxyDebugPrivilegeEnablingResult = *(PROXYPROCESSDEBUGPRIVILEGEENABLING*)((PCOPYDATASTRUCT)lParam)->lpData;
					SetEvent(hEventProxyProcessDebugPrivilegeEnablingFinished);
				}
				break;
				// Recieve HWND to proxy process
			case COPYDATA::PROXYPROCESSHWNDSENT:
				{
					hWndProxy = *(HWND*)((PCOPYDATASTRUCT)lParam)->lpData;
					SetEvent(hEventProxyProcessHWNDRevieved);
				}
				break;
				// Get proxy dll injection result
			case COPYDATA::DLLINJECTIONFINISHED:
				{
					ProxyDllInjectionResult = *(PROXYDLLINJECTION*)((PCOPYDATASTRUCT)lParam)->lpData;
					SetEvent(hEventProxyDllInjectionFinished);
				}
				break;
				// Get proxy dll pull result
			case COPYDATA::DLLPULLINGFINISHED:
				{
					ProxyDllPullingResult = *(PROXYDLLPULL*)((PCOPYDATASTRUCT)lParam)->lpData;
					SetEvent(hEventProxyDllPullingFinished);
				}
				break;
				// Get add font result
			case COPYDATA::ADDFONTFINISHED:
				{
					ProxyAddFontResult = *(ADDFONT*)((PCOPYDATASTRUCT)lParam)->lpData;
					SetEvent(hEventProxyAddFontFinished);
				}
				break;
				// Get remove font result
			case COPYDATA::REMOVEFONTFINISHED:
				{
					ProxyRemoveFontResult = *(REMOVEFONT*)((PCOPYDATASTRUCT)lParam)->lpData;
					SetEvent(hEventProxyRemoveFontFinished);
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_DESTROY:
		{
			PostQuitMessage(0);
		}
		break;
	default:
		{
			ret = DefWindowProc(hWndMessage, Msg, wParam, lParam);
		}
		break;
	}

	return ret;
}