#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
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
		lvi.pszText = (LPWSTR)(iter->GetFontName().c_str());
		ListView_InsertItem(hWndListViewFontList, &lvi);
		lvi.iSubItem = 1;
		if (iter->Load())
		{
			lvi.pszText = (LPWSTR)L"Loaded";
			ListView_SetItem(hWndListViewFontList, &lvi);
			ListView_SetItemState(hWndListViewFontList, i, LVIS_SELECTED, LVIS_SELECTED);
			Message.str(L"");
			Message << iter->GetFontName() << L" opened and successfully loaded\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		}
		else
		{
			lvi.pszText = (LPWSTR)L"Load failed";
			ListView_SetItem(hWndListViewFontList, &lvi);
			Message.str(L"");
			Message << L"Failed to load " << iter->GetFontName() << L"\r\n";
			iMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		}
		iter++;
	}
	iMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");
	bDragDropHasFonts = false;

	PostMessage(hWndMain, (UINT)USERMESSAGE::WORKINGTHREADTERMINATED, NULL, NULL);
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
				Message << iter->GetFontName() << L" successfully unloaded and closed\r\n";
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
				Message << L"Failed to unload " << iter->GetFontName() << L"\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
		}
		else
		{
			ListView_DeleteItem(hWndListViewFontList, i);
			Message.str(L"");
			Message << iter->GetFontName() << L" closed\r\n";
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
		PostMessage(hWndMain, (UINT)USERMESSAGE::CLOSEWORKINGTHREADTERMINATED, (WPARAM)TRUE, NULL);
	}
	else
	{
		PostMessage(hWndMain, (UINT)USERMESSAGE::CLOSEWORKINGTHREADTERMINATED, (WPARAM)FALSE, NULL);
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
					Message << iter->GetFontName() << L" successfully unloaded and closed\r\n";
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
					Message << L"Failed to unload " << iter->GetFontName() << L"\r\n";
					iMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
				}
			}
			else
			{
				ListView_DeleteItem(hWndListViewFontList, i);
				Message.str(L"");
				Message << iter->GetFontName() << L" closed\r\n";
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

	PostMessage(hWndMain, (UINT)USERMESSAGE::WORKINGTHREADTERMINATED, NULL, NULL);
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
				Message << iter->GetFontName() << L" successfully loaded\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Load failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Message.str(L"");
				Message << L"Failed to load " << iter->GetFontName() << L"\r\n";
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

	PostMessage(hWndMain, (UINT)USERMESSAGE::WORKINGTHREADTERMINATED, NULL, NULL);
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
				Message << iter->GetFontName() << L" successfully unloaded\r\n";
				iMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
				Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			}
			else
			{
				lvi.pszText = (LPWSTR)L"Unload failed";
				ListView_SetItem(hWndListViewFontList, &lvi);
				Message.str(L"");
				Message << L"Failed to unload " << iter->GetFontName() << L"\r\n";
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

	PostMessage(hWndMain, (UINT)USERMESSAGE::WORKINGTHREADTERMINATED, NULL, NULL);
}

//Target process watch thread
unsigned int __stdcall TargetProcessWatchThreadProc(void* lpParameter)
{
	//Wait for target process or termination event
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

	//If target process terminates, clear FontList and ListViewFontList
	EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
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

	//Register default AddFont() and RemoveFont() procedure
	FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);

	//Revert to default caption
	Button_SetText(hWndButtonSelectProcess, L"Select process");

	//Close handle to target process
	CloseHandle(TargetProcessInfo.hProcess);
	TargetProcessInfo.hProcess = NULL;

	PostMessage(hWndMain, (UINT)USERMESSAGE::WATCHTHREADTERMINATED, NULL, NULL);

	return 0;
}

//Proxy process and target process watch thread
unsigned int __stdcall ProxyAndTargetProcessWatchThreadProc(void* lpParameter)
{
	//Wait for proxy process or target process or termination event
	bool bProxyOrTarget{};
	HANDLE handles[]{ piProxyProcess.hProcess, TargetProcessInfo.hProcess, hEventTerminateWatchThread };
	switch (WaitForMultipleObjects(3, handles, FALSE, INFINITE))
	{
	case WAIT_OBJECT_0:
		{
			bProxyOrTarget = false;
		}
		break;
	case WAIT_OBJECT_0 + 1:
		{
			bProxyOrTarget = true;
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

	//Clear FontList and ListViewFontList
	EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
	DisableAllButtons();

	FontResource::RegisterAddRemoveFontProc(NullAddFontProc, NullRemoveFontProc);
	FontList.clear();
	ListView_DeleteAllItems(hWndListViewFontList);

	std::wstringstream Message{};
	int iMessageLength{};

	//If proxy process terminates, just print message
	if (!bProxyOrTarget)
	{
		Message << L"FontLoaderExProxy(" << piProxyProcess.dwProcessId << L") terminated.\r\n\r\n";
		iMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, iMessageLength, iMessageLength);
		Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
	}
	//If target process termiantes, print message and terminate proxy process
	else
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

	//Terminate message thread
	SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
	WaitForSingleObject(hMessageThread, INFINITE);

	//Register default AddFont() and RemoveFont() procedure
	FontResource::RegisterAddRemoveFontProc(DefaultAddFontProc, DefaultRemoveFontProc);

	//Revert to default caption
	Button_SetText(hWndButtonSelectProcess, L"Click to select process");

	//Close handles to proxy and target process and synchronization objects
	CloseHandle(piProxyProcess.hThread);
	CloseHandle(piProxyProcess.hProcess);
	CloseHandle(TargetProcessInfo.hProcess);
	piProxyProcess.hProcess = NULL;
	TargetProcessInfo.hProcess = NULL;
	CloseHandle(hEventProxyAddFontFinished);
	CloseHandle(hEventProxyRemoveFontFinished);

	PostMessage(hWndMain, (UINT)USERMESSAGE::WATCHTHREADTERMINATED, NULL, NULL);

	return 0;
}

LRESULT CALLBACK MsgWndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);

HWND hWndMessage{};

//Message thread
unsigned int __stdcall MessageThreadProc(void* lpParameter)
{
	//Create message-only window
	WNDCLASS wc{ 0, MsgWndProc, 0, 0, (HINSTANCE)lpParameter, NULL, NULL, NULL, NULL, L"FontLoaderExMessage" };

	if (!RegisterClass(&wc))
	{
		return 0;
	}

	if (!(hWndMessage = CreateWindow(L"FontLoaderExMessage", L"FontLoaderExMessage", NULL, 0, 0, 0, 0, HWND_MESSAGE, NULL, (HINSTANCE)lpParameter, NULL)))
	{
		return 0;
	}

	SetEvent(hEventMessageThreadReady);

	MSG Msg{};
	BOOL bRet{};
	while ((bRet = GetMessage(&Msg, hWndMessage, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			return 0;
		}
		else
		{
			DispatchMessage(&Msg);
		}
	}
	UnregisterClass(L"FontLoaderExMessage", (HINSTANCE)lpParameter);

	return 0;
}

LRESULT CALLBACK MsgWndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};
	switch ((USERMESSAGE)Msg)
	{
	case USERMESSAGE::TERMINATEMESSAGETHREAD:
		{
			DestroyWindow(hWnd);
		}
		break;
	}
	switch (Msg)
	{
	case WM_COPYDATA:
		{
			COPYDATASTRUCT* pcds{ (PCOPYDATASTRUCT)lParam };
			switch ((COPYDATA)pcds->dwData)
			{
				//Get proxy SeDebugPrivilege enabling result
			case COPYDATA::PROXYPROCESSDEBUGPRIVILEGEENABLINGFINISHED:
				{
					ProxyDebugPrivilegeEnablingResult = *(PROXYPROCESSDEBUGPRIVILEGEENABLING*)pcds->lpData;
					SetEvent(hEventProxyProcessDebugPrivilegeEnablingFinished);
				}
				break;
				//Recieve HWND to proxy process
			case COPYDATA::PROXYPROCESSHWNDSENT:
				{
					hWndProxy = *(HWND*)pcds->lpData;
					SetEvent(hEventProxyProcessHWNDRevieved);
				}
				break;
				//Get proxy dll injection result
			case COPYDATA::DLLINJECTIONFINISHED:
				{
					ProxyDllInjectionResult = *(PROXYDLLINJECTION*)pcds->lpData;
					SetEvent(hEventProxyDllInjectionFinished);
				}
				break;
				//Get proxy dll pull result
			case COPYDATA::DLLPULLFINISHED:
				{
					ProxyDllPullResult = *(PROXYDLLPULL*)pcds->lpData;
					SetEvent(hEventProxyDllPullFinished);
				}
				break;
				//Get add font result
			case COPYDATA::ADDFONTFINISHED:
				{
					ProxyAddFontResult = *(ADDFONT*)pcds->lpData;
					SetEvent(hEventProxyAddFontFinished);
				}
				break;
				//Get remove font result
			case COPYDATA::REMOVEFONTFINISHED:
				{
					ProxyRemoveFontResult = *(REMOVEFONT*)pcds->lpData;
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
			ret = DefWindowProc(hWnd, Msg, wParam, lParam);
		}
		break;
	}

	return ret;
}
