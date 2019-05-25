#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <list>
#include <sstream>
#include <atomic>
#include <climits>
#include <cassert>
#include "FontResource.h"
#include "Globals.h"
#include "resource.h"

std::atomic_flag IsTargetProcessTerminated{};

// Process drag-drop font files onto the application icon stage II worker thread
void DragDropWorkerThreadProc(void* lpParameter)
{
	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{};
	for (flcs.iItem = 0; flcs.iItem < static_cast<int>(FontList.size()); flcs.iItem++)
	{
		flcs.lpszFontName = iter->GetFontName().c_str();
		if (iter->Load())
		{
			SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::OPENED_LOADED), reinterpret_cast<LPARAM>(&flcs));
		}
		else
		{
			SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::OPENED_NOTLOADED), reinterpret_cast<LPARAM>(&flcs));
		}

		iter++;
	}

	PostMessage(hWndMain, static_cast<UINT>(USERMESSAGE::DRAGDROPWORKERTHREADTERMINATED), static_cast<WPARAM>(false), 0);
}

// Close worker thread
unsigned int __stdcall CloseWorkerThreadProc(void* lpParameter)
{
	bool bIsUnloadingSuccessful{ true };
	bool bIsUnloadingInterrupted{ false };
	bool bIsFontListChanged{ false };

	FontList.reverse();
	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{};
	for (flcs.iItem = static_cast<int>(FontList.size()) - 1; flcs.iItem >= 0; flcs.iItem--)
	{
		// If target process terminated, wait for watch thread to terminate first
		if (IsTargetProcessTerminated.test_and_set())
		{
			SetEvent(hEventWorkerThreadReadyToTerminate);
			WaitForSingleObject(hThreadWatch, INFINITE);

			bIsUnloadingInterrupted = true;

			break;
		}
		// Else do as usual
		else
		{
			IsTargetProcessTerminated.clear();
			std::wstring strTemp{ iter->GetFontName() };
			flcs.lpszFontName = strTemp.c_str();
			if (iter->IsLoaded())
			{
				if (iter->Unload())
				{
					SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::UNLOADED_CLOSED), reinterpret_cast<LPARAM>(&flcs));

					bIsFontListChanged = true;

					iter = FontList.erase(iter);
				}
				else
				{
					SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::NOTUNLOADED), reinterpret_cast<LPARAM>(&flcs));

					bIsUnloadingSuccessful = false;

					iter++;
				}
			}
			else
			{
				SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::CLOSED), reinterpret_cast<LPARAM>(&flcs));

				iter = FontList.erase(iter);
			}
		}
	}
	FontList.reverse();

	PostMessage(hWndMain, static_cast<UINT>(USERMESSAGE::CLOSEWORKERTHREADTERMINATED), static_cast<WPARAM>(bIsFontListChanged), MAKELPARAM(bIsUnloadingInterrupted, bIsUnloadingSuccessful));

	return 0;
}

// Unload and close selected fonts worker thread
unsigned int __stdcall ButtonCloseWorkerThreadProc(void* lpParameter)
{
	HWND hWndListViewFontList{ GetDlgItem(hWndMain, PtrToInt(lpParameter)) };

	bool bIsInterrupted{ false };
	bool bIsFontListChanged{ false };

	FontList.reverse();
	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{};
	for (flcs.iItem = static_cast<int>(FontList.size()) - 1; flcs.iItem >= 0; flcs.iItem--)
	{
		// If target process terminated, wait for watch thread to terminate first
		if (IsTargetProcessTerminated.test_and_set())
		{
			SetEvent(hEventWorkerThreadReadyToTerminate);
			WaitForSingleObject(hThreadWatch, INFINITE);

			bIsInterrupted = true;

			break;
		}
		// Else do as usual
		else
		{
			IsTargetProcessTerminated.clear();
			if (ListView_GetItemState(hWndListViewFontList, flcs.iItem, LVIS_SELECTED) & LVIS_SELECTED)
			{
				std::wstring strTemp{ iter->GetFontName() };
				flcs.lpszFontName = strTemp.c_str();
				if (iter->IsLoaded())
				{
					if (iter->Unload())
					{
						bIsFontListChanged = true;

						SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::UNLOADED_CLOSED), reinterpret_cast<LPARAM>(&flcs));

						iter = FontList.erase(iter);
					}
					else
					{
						SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::NOTUNLOADED), reinterpret_cast<LPARAM>(&flcs));

						iter++;
					}
				}
				else
				{
					SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::CLOSED), reinterpret_cast<LPARAM>(&flcs));

					iter = FontList.erase(iter);
				}
			}
			else
			{
				iter++;
			}
		}
	}
	FontList.reverse();

	PostMessage(hWndMain, static_cast<UINT>(USERMESSAGE::BUTTONCLOSEWORKERTHREADTERMINATED), static_cast<WPARAM>(bIsFontListChanged), static_cast<LPARAM>(bIsInterrupted));

	return 0;
}

// Load selected fonts worker thread
unsigned int __stdcall ButtonLoadWorkerThreadProc(void* lpParameter)
{
	HWND hWndListViewFontList{ GetDlgItem(hWndMain, PtrToInt(lpParameter)) };

	bool bIsInterrupted{ false };
	bool bIsFontListChanged{ false };

	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{};
	for (flcs.iItem = 0; flcs.iItem < static_cast<int>(FontList.size()); flcs.iItem++)
	{
		// If target process terminated, wait for watch thread to terminate first
		if (IsTargetProcessTerminated.test_and_set())
		{
			SetEvent(hEventWorkerThreadReadyToTerminate);
			WaitForSingleObject(hThreadWatch, INFINITE);

			bIsInterrupted = true;

			break;
		}
		// Else do as usual
		else
		{
			IsTargetProcessTerminated.clear();
			if (ListView_GetItemState(hWndListViewFontList, flcs.iItem, LVIS_SELECTED) & LVIS_SELECTED)
			{
				if (iter->IsLoaded())
				{
					SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::UNTOUCHED), 0);
				}
				else
				{
					std::wstring strTemp{ iter->GetFontName() };
					flcs.lpszFontName = strTemp.c_str();
					if (iter->Load())
					{
						bIsFontListChanged = true;

						SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::LOADED), reinterpret_cast<LPARAM>(&flcs));
					}
					else
					{
						SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::NOTLOADED), reinterpret_cast<LPARAM>(&flcs));
					}
				}
			}
			iter++;
		}
	}

	PostMessage(hWndMain, static_cast<UINT>(USERMESSAGE::BUTTONCLOSEWORKERTHREADTERMINATED), static_cast<WPARAM>(bIsFontListChanged), static_cast<LPARAM>(bIsInterrupted));

	return 0;
}

// Unload selected fonts worker thread
unsigned int __stdcall ButtonUnloadWorkerThreadProc(void* lpParameter)
{
	HWND hWndListViewFontList{ GetDlgItem(hWndMain, PtrToInt(lpParameter)) };

	bool bIsInterrupted{ false };
	bool bIsFontListChanged{ false };

	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{};
	for (flcs.iItem = 0; flcs.iItem < static_cast<int>(FontList.size()); flcs.iItem++)
	{
		// If target process terminated, wait for watch thread to terminate first
		if (IsTargetProcessTerminated.test_and_set())
		{
			SetEvent(hEventWorkerThreadReadyToTerminate);
			WaitForSingleObject(hThreadWatch, INFINITE);

			bIsInterrupted = true;

			break;
		}
		// Else do as usual
		else
		{
			IsTargetProcessTerminated.clear();
			if (ListView_GetItemState(hWndListViewFontList, flcs.iItem, LVIS_SELECTED) & LVIS_SELECTED)
			{
				if (iter->IsLoaded())
				{
					std::wstring strTemp{ iter->GetFontName() };
					flcs.lpszFontName = strTemp.c_str();
					if (iter->Unload())
					{
						bIsFontListChanged = true;

						SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::UNLOADED), reinterpret_cast<LPARAM>(&flcs));
					}
					else
					{
						SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::NOTUNLOADED), reinterpret_cast<LPARAM>(&flcs));
					}
				}
				else
				{
					SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::FONTLISTCHANGED), static_cast<WPARAM>(FONTLISTCHANGED::UNTOUCHED), 0);
				}
			}
			iter++;
		}
	}

	PostMessage(hWndMain, static_cast<UINT>(USERMESSAGE::BUTTONCLOSEWORKERTHREADTERMINATED), static_cast<WPARAM>(bIsFontListChanged), static_cast<LPARAM>(bIsInterrupted));

	return 0;
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
		{
			assert(0 && "WaitForSingleObject() failed");
		}
		break;
	}

	// Singal worker thread and wait for worker thread to ready to exit
	// Because only one worker thread runs at a time, so use bitwise-or to get the handle to running worker thread
	bool bIsWorkerThreadRunning{ false };
	switch (WaitForSingleObject(reinterpret_cast<HANDLE>(reinterpret_cast<UINT_PTR>(hThreadCloseWorkerThreadProc) | reinterpret_cast<UINT_PTR>(hThreadButtonCloseWorkerThreadProc) | reinterpret_cast<UINT_PTR>(hThreadButtonLoadWorkerThreadProc) | reinterpret_cast<UINT_PTR>(hThreadButtonUnloadWorkerThreadProc)), 0))
	{
	case WAIT_TIMEOUT:
		{
			bIsWorkerThreadRunning = true;
			IsTargetProcessTerminated.test_and_set();
			WaitForSingleObject(hEventWorkerThreadReadyToTerminate, INFINITE);
		}
		break;
	default:
		break;
	}

	SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::WATCHTHREADTERMINATING), static_cast<WPARAM>(TERMINATION::TARGET), static_cast<LPARAM>(bIsWorkerThreadRunning));

	// Clear FontList
	FontResource::RegisterAddRemoveFontProc(NullAddFontProc, NullRemoveFontProc);
	FontList.clear();

	// Register global AddFont() and RemoveFont() procedures
	FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

	// Close HANDLE to proxy process and target process, duplicated handles and synchronization objects
	CloseHandle(TargetProcessInfo.hProcess);
	TargetProcessInfo.hProcess = NULL;
	CloseHandle(hProcessCurrentDuplicated);
	CloseHandle(hProcessTargetDuplicated);

	PostMessage(hWndMain, static_cast<UINT>(USERMESSAGE::WATCHTHREADTERMINATED), 0, static_cast<LPARAM>(bIsWorkerThreadRunning));

	return 0;
}

// Proxy process and target process watch thread
unsigned int __stdcall ProxyAndTargetProcessWatchThreadProc(void* lpParameter)
{
	// Wait for proxy process or target process or termination event
	TERMINATION t{};

	HANDLE handles[]{ ProxyProcessInfo.hProcess, TargetProcessInfo.hProcess, hEventTerminateWatchThread };
	switch (WaitForMultipleObjects(3, handles, FALSE, INFINITE))
	{
	case WAIT_OBJECT_0:
		{
			t = TERMINATION::PROXY;
		}
		break;
	case WAIT_OBJECT_0 + 1:
		{
			t = TERMINATION::TARGET;
		}
		break;
	case WAIT_OBJECT_0 + 2:
		{
			return 0;
		}
		break;
	default:
		{
			assert(0 && "WaitForSingleObject() failed");
		}
		break;
	}

	// Singal worker thread and wait for worker thread to ready to exit
	// Because only one worker thread runs at a time, so use bitwise-or to get the handle to running worker thread
	bool bIsWorkerThreadRunning{ false };
	switch (WaitForSingleObject(reinterpret_cast<HANDLE>(reinterpret_cast<UINT_PTR>(hThreadCloseWorkerThreadProc) | reinterpret_cast<UINT_PTR>(hThreadButtonCloseWorkerThreadProc) | reinterpret_cast<UINT_PTR>(hThreadButtonLoadWorkerThreadProc) | reinterpret_cast<UINT_PTR>(hThreadButtonUnloadWorkerThreadProc)), 0))
	{
	case WAIT_TIMEOUT:
		{
			bIsWorkerThreadRunning = true;
			IsTargetProcessTerminated.test_and_set();
			WaitForSingleObject(hEventWorkerThreadReadyToTerminate, INFINITE);
		}
		break;
	default:
		break;
	}

	SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::WATCHTHREADTERMINATING), static_cast<WPARAM>(t), static_cast<LPARAM>(bIsWorkerThreadRunning));

	// Clear FontList
	FontResource::RegisterAddRemoveFontProc(NullAddFontProc, NullRemoveFontProc);
	FontList.clear();

	// Register global AddFont() and RemoveFont() procedures
	FontResource::RegisterAddRemoveFontProc(GlobalAddFontProc, GlobalRemoveFontProc);

	// Terminate message thread
	SendMessage(hWndMessage, WM_CLOSE, 0, 0);
	WaitForSingleObject(hThreadMessage, INFINITE);

	// Close HANDLE to proxy process and target process, duplicated handles and synchronization objects
	CloseHandle(TargetProcessInfo.hProcess);
	TargetProcessInfo.hProcess = NULL;
	CloseHandle(hProcessCurrentDuplicated);
	CloseHandle(hProcessTargetDuplicated);
	CloseHandle(ProxyProcessInfo.hProcess);
	ProxyProcessInfo.hProcess = NULL;
	CloseHandle(hEventProxyAddFontFinished);
	CloseHandle(hEventProxyRemoveFontFinished);

	PostMessage(hWndMain, static_cast<UINT>(USERMESSAGE::WATCHTHREADTERMINATED), 0, static_cast<LPARAM>(bIsWorkerThreadRunning));

	return 0;
}

LRESULT CALLBACK MessageWndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);

HWND hWndMessage{};

// Message thread
unsigned int __stdcall MessageThreadProc(void* lpParameter)
{
	// Create message-only window
	WNDCLASS wc{ 0, MessageWndProc, 0, 0, static_cast<HINSTANCE>(GetModuleHandle(NULL)), NULL, NULL, NULL, NULL, L"FontLoaderExMessage" };
	if (!RegisterClass(&wc))
	{
		SetEvent(hEventMessageThreadNotReady);

		return 0xFFFFFFFFu;
	}
	if (!(hWndMessage = CreateWindow(L"FontLoaderExMessage", L"FontLoaderExMessage", NULL, 0, 0, 0, 0, HWND_MESSAGE, NULL, static_cast<HINSTANCE>(GetModuleHandle(NULL)), NULL)))
	{
		SetEvent(hEventMessageThreadNotReady);

		return 0xFFFFFFFFu;
	}

	SetEvent(hEventMessageThreadReady);

	MSG Message{};
	unsigned int uiRet{};
	BOOL bRet{};
	do
	{
		switch (bRet = GetMessage(&Message, NULL, 0, 0))
		{
		case -1:
			{
				uiRet = static_cast<unsigned int>(GetLastError());

				DestroyWindow(hWndMessage);
				BOOL bRetUnregisterClass{ UnregisterClass(L"FontLoaderExMessage", (HINSTANCE)GetModuleHandle(NULL)) };
				assert(bRetUnregisterClass);
			}
			break;
		case 0:
			{
				BOOL bRetUnregisterClass{ UnregisterClass(L"FontLoaderExMessage", (HINSTANCE)GetModuleHandle(NULL)) };
				assert(bRetUnregisterClass);

				uiRet = static_cast<unsigned int>(Message.wParam);
			}
			break;
		default:
			{
				DispatchMessage(&Message);
			}
			break;
		}
	} while (bRet);

	return uiRet;
}

LRESULT CALLBACK MessageWndProc(HWND hWndMessage, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	switch (Message)
	{
		// COPYDATASTRUCT::dwData = Command : enum COPYDATA
		// COPYDATASTRUCT::lpData = Data
	case WM_COPYDATA:
		{
			switch (static_cast<COPYDATA>(reinterpret_cast<PCOPYDATASTRUCT>(lParam)->dwData))
			{
				// Get proxy SeDebugPrivilege enabling result
			case COPYDATA::PROXYPROCESSDEBUGPRIVILEGEENABLINGFINISHED:
				{
					ProxyDebugPrivilegeEnablingResult = *static_cast<PROXYPROCESSDEBUGPRIVILEGEENABLING*>(reinterpret_cast<PCOPYDATASTRUCT>(lParam)->lpData);
					SetEvent(hEventProxyProcessDebugPrivilegeEnablingFinished);
				}
				break;
				// receive HWND to proxy process
			case COPYDATA::PROXYPROCESSHWNDSENT:
				{
					hWndProxy = *static_cast<HWND*>(reinterpret_cast<PCOPYDATASTRUCT>(lParam)->lpData);
					SetEvent(hEventProxyProcessHWNDRevieved);
				}
				break;
				// Get proxy dll injection result
			case COPYDATA::DLLINJECTIONFINISHED:
				{
					ProxyDllInjectionResult = *static_cast<PROXYDLLINJECTION*>(reinterpret_cast<PCOPYDATASTRUCT>(lParam)->lpData);
					SetEvent(hEventProxyDllInjectionFinished);
				}
				break;
				// Get proxy dll pull result
			case COPYDATA::DLLPULLINGFINISHED:
				{
					ProxyDllPullingResult = *static_cast<PROXYDLLPULL*>(reinterpret_cast<PCOPYDATASTRUCT>(lParam)->lpData);
					SetEvent(hEventProxyDllPullingFinished);
				}
				break;
				// Get add font result
			case COPYDATA::ADDFONTFINISHED:
				{
					ProxyAddFontResult = *static_cast<ADDFONT*>(reinterpret_cast<PCOPYDATASTRUCT>(lParam)->lpData);
					SetEvent(hEventProxyAddFontFinished);
				}
				break;
				// Get remove font result
			case COPYDATA::REMOVEFONTFINISHED:
				{
					ProxyRemoveFontResult = *static_cast<REMOVEFONT*>(reinterpret_cast<PCOPYDATASTRUCT>(lParam)->lpData);
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
			ret = DefWindowProc(hWndMessage, Message, wParam, lParam);
		}
		break;
	}

	return ret;
}