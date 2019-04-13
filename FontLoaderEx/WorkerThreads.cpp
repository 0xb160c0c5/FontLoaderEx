#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <list>
#include <sstream>
#include <atomic>
#include "FontResource.h"
#include "Globals.h"
#include "resource.h"

std::atomic<bool> bIsWorkerThreadRunning{ false };
std::atomic<bool> bIsTargetProcessTerminated{ false };
HANDLE hEventWorkerThreadReadyToTerminate{};

// Process drag-drop font files onto the application icon stage II worker thread
void DragDropWorkerThreadProc(void* lpParameter)
{
	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };

	std::wstringstream Message{};
	int cchMessageLength{};

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
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			Message.str(L"");
		}
		else
		{
			lvi.pszText = (LPWSTR)L"Load failed";
			ListView_SetItem(hWndListViewFontList, &lvi);

			Message << L"Opened but failed to load " << iter->GetFontName() << L"\r\n";
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			Message.str(L"");
		}
		iter++;
	}
	cchMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMain, (UINT)USERMESSAGE::DRAGDROPWORKERTHREADTERMINATED, NULL, NULL);
}

// Close worker thread
void CloseWorkerThreadProc(void* lpParameter)
{
	bIsWorkerThreadRunning = true;
	hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);

	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	HWND hWndButtonBroadcastWM_FONTCHANGE{ GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE) };

	std::wstringstream Message{};
	int cchMessageLength{};

	bool bIsUnloadingSuccessful{ true };
	bool bIsUnloadingInterrupted{ false };
	bool bIsFontListChanged{ false };

	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	FontList.reverse();
	std::list<FontResource>::iterator iter{ FontList.begin() };
	for (int i = ListView_GetItemCount(hWndListViewFontList) - 1; i >= 0; i--)
	{
		if (iter->IsLoaded())
		{
			if (bIsTargetProcessTerminated)
			{
				// If target process terminated, wait for watch thread to terminate first
				cchMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");

				SetEvent(hEventWorkerThreadReadyToTerminate);
				WaitForSingleObject(hThreadWatch, INFINITE);

				bIsUnloadingInterrupted = true;
				break;
			}
			else
			{
				if (iter->Unload())
				{
					bIsFontListChanged = true;

					ListView_DeleteItem(hWndListViewFontList, i);

					Message << iter->GetFontName() << L" successfully unloaded and closed\r\n";
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
					Message.str(L"");

					iter = FontList.erase(iter);
				}
				else
				{
					bIsUnloadingSuccessful = false;

					lvi.iItem = i;
					lvi.pszText = (LPWSTR)L"Unload failed";
					ListView_SetItem(hWndListViewFontList, &lvi);

					Message << L"Failed to unload " << iter->GetFontName() << L"\r\n";
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
					Message.str(L"");

					iter++;
				}
			}
		}
		else
		{
			ListView_DeleteItem(hWndListViewFontList, i);

			Message << iter->GetFontName() << L" closed\r\n";
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			Message.str(L"");

			iter = FontList.erase(iter);
		}
	}
	FontList.reverse();

	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged && !TargetProcessInfo.hProcess)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		cchMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	cchMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMain, (UINT)USERMESSAGE::CLOSEWORKERTHREADTERMINATED, (WPARAM)bIsUnloadingSuccessful, (LPARAM)bIsUnloadingInterrupted);

	bIsWorkerThreadRunning = false;
	SetEvent(hEventWorkerThreadReadyToTerminate);
	CloseHandle(hEventWorkerThreadReadyToTerminate);
}

// Unload and close selected fonts worker thread
void ButtonCloseWorkerThreadProc(void* lpParameter)
{
	bIsWorkerThreadRunning = true;
	hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);

	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	HWND hWndButtonBroadcastWM_FONTCHANGE{ GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE) };

	std::wstringstream Message{};
	int cchMessageLength{};

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
				if (bIsTargetProcessTerminated)
				{
					// If target process terminated, wait for watch thread to terminate first
					cchMessageLength = Edit_GetTextLength(hWndEditMessage);
					Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
					Edit_ReplaceSel(hWndEditMessage, L"\r\n");

					SetEvent(hEventWorkerThreadReadyToTerminate);
					WaitForSingleObject(hThreadWatch, INFINITE);
					break;
				}
				else
				{
					if (iter->Unload())
					{
						bIsFontListChanged = true;

						ListView_DeleteItem(hWndListViewFontList, i);

						Message << iter->GetFontName() << L" successfully unloaded and closed\r\n";
						cchMessageLength = Edit_GetTextLength(hWndEditMessage);
						Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
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
						cchMessageLength = Edit_GetTextLength(hWndEditMessage);
						Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
						Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
						Message.str(L"");

						iter++;
					}
				}
			}
			else
			{
				ListView_DeleteItem(hWndListViewFontList, i);

				Message << iter->GetFontName() << L" closed\r\n";
				cchMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
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
		cchMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	cchMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMain, (UINT)USERMESSAGE::BUTTONCLOSEWORKERTHREADTERMINATED, NULL, NULL);

	bIsWorkerThreadRunning = false;
	SetEvent(hEventWorkerThreadReadyToTerminate);
	CloseHandle(hEventWorkerThreadReadyToTerminate);
}

// Load selected fonts worker thread
void ButtonLoadWorkerThreadProc(void* lpParameter)
{
	bIsWorkerThreadRunning = true;
	hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);

	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	HWND hWndButtonBroadcastWM_FONTCHANGE{ GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE) };

	bool bIsFontListChanged{ false };

	std::wstringstream Message{};
	int cchMessageLength{};

	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::list<FontResource>::iterator iter{ FontList.begin() };
	for (int i = 0; i < ListView_GetItemCount(hWndListViewFontList); i++)
	{
		if ((ListView_GetItemState(hWndListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED) && (!(iter->IsLoaded())))
		{
			if (bIsTargetProcessTerminated)
			{
				// If target process terminated, wait for watch thread to terminate first
				cchMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");

				SetEvent(hEventWorkerThreadReadyToTerminate);
				WaitForSingleObject(hThreadWatch, INFINITE);
				break;
			}
			else
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

		}
		cchMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
		Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		Message.str(L"");

		iter++;
	}

	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged && !TargetProcessInfo.hProcess)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		cchMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	cchMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMain, (UINT)USERMESSAGE::BUTTONLOADWORKERTHREADTERMINATED, NULL, NULL);

	bIsWorkerThreadRunning = false;
	SetEvent(hEventWorkerThreadReadyToTerminate);
	CloseHandle(hEventWorkerThreadReadyToTerminate);
}

// Unload selected fonts worker thread
void ButtonUnloadWorkerThreadProc(void* lpParameter)
{
	bIsWorkerThreadRunning = true;
	hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);

	HWND hWndListViewFontList{ GetDlgItem(hWndMain, (int)ID::ListViewFontList) };
	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	HWND hWndButtonBroadcastWM_FONTCHANGE{ GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE) };

	bool bIsFontListChanged{ false };

	std::wstringstream Message{};
	int cchMessageLength{};

	LVITEM lvi{ LVIF_TEXT, 0, 1 };
	std::list<FontResource>::iterator iter{ FontList.begin() };
	for (int i = 0; i < ListView_GetItemCount(hWndListViewFontList); i++)
	{
		if ((ListView_GetItemState(hWndListViewFontList, i, LVIS_SELECTED) & LVIS_SELECTED) && (iter->IsLoaded()))
		{
			if (bIsTargetProcessTerminated)
			{
				// If target process terminated, wait for watch thread to terminate first
				cchMessageLength = Edit_GetTextLength(hWndEditMessage);
				Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
				Edit_ReplaceSel(hWndEditMessage, L"\r\n");

				SetEvent(hEventWorkerThreadReadyToTerminate);
				WaitForSingleObject(hThreadWatch, INFINITE);
				break;
			}
			else
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
		}
		cchMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
		Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		Message.str(L"");

		iter++;
	}

	if ((Button_GetCheck(hWndButtonBroadcastWM_FONTCHANGE) == BST_CHECKED) && bIsFontListChanged && !TargetProcessInfo.hProcess)
	{
		FORWARD_WM_FONTCHANGE(HWND_BROADCAST, PostMessage);
		cchMessageLength = Edit_GetTextLength(hWndEditMessage);
		Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
		Edit_ReplaceSel(hWndEditMessage, L"WM_FONTCHANGE broadcasted to all top windows.\r\n");
	}
	cchMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
	Edit_ReplaceSel(hWndEditMessage, L"\r\n");

	PostMessage(hWndMain, (UINT)USERMESSAGE::BUTTONUNLOADWORKERTHREADTERMINATED, NULL, NULL);

	bIsWorkerThreadRunning = false;
	SetEvent(hEventWorkerThreadReadyToTerminate);
	CloseHandle(hEventWorkerThreadReadyToTerminate);
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

	// Singal worker thread and wait for worker thread to ready to exit
	if (bIsWorkerThreadRunning)
	{
		bIsTargetProcessTerminated = true;
		WaitForSingleObject(hEventWorkerThreadReadyToTerminate, INFINITE);
	}

	// Disable controls
	if (!bIsWorkerThreadRunning)
	{
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
		EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
	}

	// Clear FontList and ListViewFontList
	FontResource::RegisterAddRemoveFontProc(NullAddFontProc, NullRemoveFontProc);
	FontList.clear();
	ListView_DeleteAllItems(GetDlgItem(hWndMain, (int)ID::ListViewFontList));

	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	std::wstringstream Message{};
	int cchMessageLength{};
	Message << L"Target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L") terminated.\r\n\r\n";
	cchMessageLength = Edit_GetTextLength(hWndEditMessage);
	Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
	Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());

	// Register global AddFont() and RemoveFont() procedures
	FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

	// Revert the caption of ButtonSelectProcess to default
	Button_SetText(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), L"Select process");

	// Close handle to target process and duplicated handles
	CloseHandle(TargetProcessInfo.hProcess);
	TargetProcessInfo.hProcess = NULL;
	CloseHandle(hProcessCurrentDuplicated);
	CloseHandle(hProcessTargetDuplicated);

	// Enable controls
	if (!bIsWorkerThreadRunning)
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

	PostMessage(hWndMain, (UINT)USERMESSAGE::WATCHTHREADTERMINATED, NULL, NULL);

	// Singal worker thread to ready to continue
	if (bIsWorkerThreadRunning)
	{
		bIsTargetProcessTerminated = false;
	}

	return 0;
}

// Proxy process and target process watch thread
unsigned int __stdcall ProxyAndTargetProcessWatchThreadProc(void* lpParameter)
{
	// Wait for proxy process or target process or termination event
	enum class Termination { Proxy, Target };
	Termination t{};

	HANDLE handles[]{ ProxyProcessInfo.hProcess, TargetProcessInfo.hProcess, hEventTerminateWatchThread };
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

	// Singal worker thread and wait for worker thread to ready to exit
	if (bIsWorkerThreadRunning)
	{
		bIsTargetProcessTerminated = true;
		WaitForSingleObject(hEventWorkerThreadReadyToTerminate, INFINITE);
	}

	// Disable controls
	if (!bIsWorkerThreadRunning)
	{
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonOpen), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonClose), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonLoad), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonUnload), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonBroadcastWM_FONTCHANGE), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), FALSE);
		EnableWindow(GetDlgItem(hWndMain, (int)ID::ListViewFontList), FALSE);
		EnableMenuItem(GetSystemMenu(hWndMain, FALSE), SC_CLOSE, MF_BYCOMMAND | MF_GRAYED);
	}

	// Clear FontList and ListViewFontList
	FontResource::RegisterAddRemoveFontProc(NullAddFontProc, NullRemoveFontProc);
	FontList.clear();
	ListView_DeleteAllItems(GetDlgItem(hWndMain, (int)ID::ListViewFontList));

	HWND hWndEditMessage{ GetDlgItem(hWndMain, (int)ID::EditMessage) };
	std::wstringstream Message{};
	int cchMessageLength{};
	switch (t)
	{
		// If proxy process terminates, just print message
	case Termination::Proxy:
		{
			Message << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") terminated.\r\n\r\n";
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		}
		break;
		// If target process termiantes, print message and terminate proxy process
	case Termination::Target:
		{
			Message << L"Target process " << TargetProcessInfo.strProcessName << L"(" << TargetProcessInfo.dwProcessID << L") terminated.\r\n\r\n";
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
			Message.str(L"");

			COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::TERMINATE, 0, NULL };
			FORWARD_WM_COPYDATA(hWndProxy, hWndMain, &cds, SendMessage);
			WaitForSingleObject(ProxyProcessInfo.hProcess, INFINITE);
			Message << ProxyProcessInfo.strProcessName << L"(" << ProxyProcessInfo.dwProcessID << L") successfully terminated.\r\n\r\n";
			cchMessageLength = Edit_GetTextLength(hWndEditMessage);
			Edit_SetSel(hWndEditMessage, cchMessageLength, cchMessageLength);
			Edit_ReplaceSel(hWndEditMessage, Message.str().c_str());
		}
		break;
	default:
		break;
	}

	// Terminate message thread
	SendMessage(hWndMessage, (UINT)USERMESSAGE::TERMINATEMESSAGETHREAD, NULL, NULL);
	WaitForSingleObject(hThreadMessage, INFINITE);

	// Register global AddFont() and RemoveFont() procedures
	FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

	// Revert the caption of ButtonSelectProcess to default
	Button_SetText(GetDlgItem(hWndMain, (int)ID::ButtonSelectProcess), L"Click to select process");

	// Close handles to proxy process and target process, duplicated handles and synchronization objects
	CloseHandle(ProxyProcessInfo.hProcess);
	CloseHandle(TargetProcessInfo.hProcess);
	ProxyProcessInfo.hProcess = NULL;
	TargetProcessInfo.hProcess = NULL;
	CloseHandle(hProcessCurrentDuplicated);
	CloseHandle(hProcessTargetDuplicated);
	CloseHandle(hEventProxyAddFontFinished);
	CloseHandle(hEventProxyRemoveFontFinished);

	// Enable controls
	if (!bIsWorkerThreadRunning)
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

	PostMessage(hWndMain, (UINT)USERMESSAGE::WATCHTHREADTERMINATED, NULL, NULL);

	// Singal worker thread to ready to continue
	if (bIsWorkerThreadRunning)
	{
		bIsTargetProcessTerminated = false;
	}

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
	WNDCLASS wc{ 0, MessageWndProc, 0, 0, (HINSTANCE)GetModuleHandle(NULL), NULL, NULL, NULL, NULL, L"FontLoaderExMessage" };
	if (!RegisterClass(&wc))
	{
		SetEvent(hEventMessageThreadNotReady);

		return 0xffffffff;
	}
	if (!(hWndMessage = CreateWindow(L"FontLoaderExMessage", L"FontLoaderExMessage", NULL, 0, 0, 0, 0, HWND_MESSAGE, NULL, (HINSTANCE)GetModuleHandle(NULL), NULL)))
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
	UnregisterClass(L"FontLoaderExMessage", (HINSTANCE)GetModuleHandle(NULL));

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