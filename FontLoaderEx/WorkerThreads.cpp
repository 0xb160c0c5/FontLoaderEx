#include <windows.h>
#include <windowsx.h>
#include <CommCtrl.h>
#include <list>
#include <sstream>
#include <atomic>
#include <climits>
#include "FontResource.h"
#include "Globals.h"
#include "resource.h"

std::atomic<bool> bIsWorkerThreadRunning{ false };
std::atomic<bool> bIsTargetProcessTerminated{ false };
HANDLE hEventWorkerThreadReadyToTerminate{};

// Process drag-drop font files onto the application icon stage II worker thread
void DragDropWorkerThreadProc(void* lpParameter)
{
	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{ 0, iter->GetFontName().c_str() };
	for (flcs.iItem = 0; flcs.iItem < static_cast<int>(FontList.size()); flcs.iItem++)
	{
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
void CloseWorkerThreadProc(void* lpParameter)
{
	bIsWorkerThreadRunning = true;
	hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);

	bool bIsUnloadingSuccessful{ true };
	bool bIsUnloadingInterrupted{ false };
	bool bIsFontListChanged{ false };

	FontList.reverse();
	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{};
	for (flcs.iItem = static_cast<int>(FontList.size()) - 1; flcs.iItem >= 0; flcs.iItem--)
	{
		// If target process terminated, wait for watch thread to terminate first
		if (bIsTargetProcessTerminated)
		{
			SetEvent(hEventWorkerThreadReadyToTerminate);
			WaitForSingleObject(hThreadWatch, INFINITE);

			bIsUnloadingInterrupted = true;
			break;
		}
		// Else do as usual
		else
		{
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

	bIsWorkerThreadRunning = false;
	SetEvent(hEventWorkerThreadReadyToTerminate);
	CloseHandle(hEventWorkerThreadReadyToTerminate);
}

// Unload and close selected fonts worker thread
void ButtonCloseWorkerThreadProc(void* lpParameter)
{
	bIsWorkerThreadRunning = true;
	hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);

	HWND hWndListViewFontList{ GetDlgItem(hWndMain, PtrToInt(lpParameter)) };

	bool bIsFontListChanged{ false };

	FontList.reverse();
	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{};
	for (flcs.iItem = static_cast<int>(FontList.size()) - 1; flcs.iItem >= 0; flcs.iItem--)
	{
		// If target process terminated, wait for watch thread to terminate first
		if (bIsTargetProcessTerminated)
		{
			SetEvent(hEventWorkerThreadReadyToTerminate);
			WaitForSingleObject(hThreadWatch, INFINITE);
			break;
		}
		// Else do as usual
		else
		{
			if ((ListView_GetItemState(hWndListViewFontList, flcs.iItem, LVIS_SELECTED) & LVIS_SELECTED))
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

	PostMessage(hWndMain, static_cast<UINT>(USERMESSAGE::BUTTONCLOSEWORKERTHREADTERMINATED), static_cast<WPARAM>(bIsFontListChanged), 0);

	bIsWorkerThreadRunning = false;
	SetEvent(hEventWorkerThreadReadyToTerminate);
	CloseHandle(hEventWorkerThreadReadyToTerminate);
}

// Load selected fonts worker thread
void ButtonLoadWorkerThreadProc(void* lpParameter)
{
	bIsWorkerThreadRunning = true;
	hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);

	HWND hWndListViewFontList{ GetDlgItem(hWndMain, PtrToInt(lpParameter)) };

	bool bIsFontListChanged{ false };

	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{};
	for (flcs.iItem = 0; flcs.iItem < static_cast<int>(FontList.size()); flcs.iItem++)
	{
		// If target process terminated, wait for watch thread to terminate first
		if (bIsTargetProcessTerminated)
		{
			SetEvent(hEventWorkerThreadReadyToTerminate);
			WaitForSingleObject(hThreadWatch, INFINITE);
			break;
		}
		// Else do as usual
		else
		{
			if ((ListView_GetItemState(hWndListViewFontList, flcs.iItem, LVIS_SELECTED) & LVIS_SELECTED) && (!(iter->IsLoaded())))
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
			iter++;
		}
	}

	PostMessage(hWndMain, (UINT)USERMESSAGE::BUTTONLOADWORKERTHREADTERMINATED, (WPARAM)bIsFontListChanged, 0);

	bIsWorkerThreadRunning = false;
	SetEvent(hEventWorkerThreadReadyToTerminate);
	CloseHandle(hEventWorkerThreadReadyToTerminate);
}

// Unload selected fonts worker thread
void ButtonUnloadWorkerThreadProc(void* lpParameter)
{
	bIsWorkerThreadRunning = true;
	hEventWorkerThreadReadyToTerminate = CreateEvent(NULL, TRUE, FALSE, NULL);

	HWND hWndListViewFontList{ GetDlgItem(hWndMain, PtrToInt(lpParameter)) };

	bool bIsFontListChanged{ false };

	std::list<FontResource>::iterator iter{ FontList.begin() };
	FONTLISTCHANGEDSTRUCT flcs{};
	for (flcs.iItem = 0; flcs.iItem < static_cast<int>(FontList.size()); flcs.iItem++)
	{
		// If target process terminated, wait for watch thread to terminate first
		if (bIsTargetProcessTerminated)
		{
			SetEvent(hEventWorkerThreadReadyToTerminate);
			WaitForSingleObject(hThreadWatch, INFINITE);
			break;
		}
		// Else do as usual
		else
		{
			if ((ListView_GetItemState(hWndListViewFontList, flcs.iItem, LVIS_SELECTED) & LVIS_SELECTED) && (iter->IsLoaded()))
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
			iter++;
		}
	}

	PostMessage(hWndMain, static_cast<UINT>(USERMESSAGE::BUTTONUNLOADWORKERTHREADTERMINATED), static_cast<WPARAM>(bIsFontListChanged), 0);

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

	SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::WATCHTHREADTERMINATED), MAKEWPARAM(WATCHTHREADTERMINATED::TARGET, TERMINATION::TARGET), static_cast<LPARAM>(bIsWorkerThreadRunning));

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
		break;
	}

	// Singal worker thread and wait for worker thread to ready to exit
	if (bIsWorkerThreadRunning)
	{
		bIsTargetProcessTerminated = true;
		WaitForSingleObject(hEventWorkerThreadReadyToTerminate, INFINITE);
	}

	SendMessage(hWndMain, static_cast<UINT>(USERMESSAGE::WATCHTHREADTERMINATED), MAKEWPARAM(WATCHTHREADTERMINATED::PROXY, t), static_cast<LPARAM>(bIsWorkerThreadRunning));

	if (bIsWorkerThreadRunning)
	{
		bIsTargetProcessTerminated = false;
	}

	return 0;
}

LRESULT CALLBACK MessageWndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);

HWND hWndMessage{};

// Message thread
unsigned int __stdcall MessageThreadProc(void* lpParameter)
{
	// Force Windows to create message queue for current thread
	MSG Message{};
	PeekMessage(&Message, NULL, 0, 0, PM_NOREMOVE);

	// Create message-only window
	WNDCLASS wc{ 0, MessageWndProc, 0, 0, static_cast<HINSTANCE>(GetModuleHandle(NULL)), NULL, NULL, NULL, NULL, L"FontLoaderExMessage" };
	if (!RegisterClass(&wc))
	{
		SetEvent(hEventMessageThreadNotReady);

		return 0xFFFFFFFF;
	}
	if (!(hWndMessage = CreateWindow(L"FontLoaderExMessage", L"FontLoaderExMessage", NULL, 0, 0, 0, 0, HWND_MESSAGE, NULL, static_cast<HINSTANCE>(GetModuleHandle(NULL)), NULL)))
	{
		SetEvent(hEventMessageThreadNotReady);

		return 0xFFFFFFFF;
	}

	SetEvent(hEventMessageThreadReady);

	BOOL bRet{};
	while ((bRet = GetMessage(&Message, NULL, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			unsigned int uiLastError{ static_cast<unsigned int>(GetLastError()) };
			DestroyWindow(hWndMessage);
			UnregisterClass(L"FontLoaderExMessage", (HINSTANCE)GetModuleHandle(NULL));

			return uiLastError;
		}
		else
		{
			DispatchMessage(&Message);
		}
	}
	UnregisterClass(L"FontLoaderExMessage", (HINSTANCE)GetModuleHandle(NULL));

	return static_cast<unsigned int>(Message.wParam);
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
				// Recieve HWND to proxy process
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