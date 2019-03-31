#if !defined(UNICODE) || !defined(_UNICODE)
#error Unicode character set required
#endif // UNICODE && _UNICODE

#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")
#pragma comment(lib, "Shlwapi.lib")

#include <windows.h>
#include <windowsx.h>
#include <tlhelp32.h>
#include <shlwapi.h>
#include <process.h>
#include <cwchar>

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
#ifdef _DEBUG
	// Wait for debugger to attach
	MessageBox(NULL, L"FontLoaderExProxy launched!", L"FontLoaderExProxy", NULL);
#endif // _DEBUG

	// Detect whether FontLoaderEx is running. If not running, launch it.
	HANDLE hEventParentProcessRunning{ OpenEvent(EVENT_ALL_ACCESS, FALSE, L"FontLoaderEx_EventParentProcessRunning_B980D8A4-C487-4306-9D17-3BA6A2CCA4A4") };
	if (!hEventParentProcessRunning)
	{
		switch (MessageBox(NULL, L"Never run FontLoaderExProxy directly, run FontLoaderEx instead.\r\n\r\nDo you want to launch FontloaderEx now?", L"FontLoaderExProxy", MB_ICONINFORMATION | MB_YESNO))
		{
		case IDYES:
			{
				STARTUPINFO sa{ sizeof(STARTUPINFO) };
				PROCESS_INFORMATION pi{};
				CreateProcess(L"FontLoaderEx.exe", NULL, NULL, NULL, FALSE, 0, NULL, NULL, &sa, &pi);
				CloseHandle(pi.hThread);
				CloseHandle(pi.hProcess);

				return 0;
			}
			break;
		case IDNO:
			{
				return 0;
			}
			break;
		default:
			{
				return 0;
			}
			break;
		}
	}
	else
	{
		CloseHandle(hEventParentProcessRunning);
	}

	// Create message-only window
	WNDCLASS wc{ 0, WndProc, 0, 0, hInstance, NULL, NULL, NULL, NULL, L"FontLoaderExProxy" };
	if (!RegisterClass(&wc))
	{
		return -1;
	}
	HWND hWndMain{};
	if (!(hWndMain = CreateWindow(L"FontLoaderExProxy", L"FontLoaderExProxy", NULL, 0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL)))
	{
		return -1;
	}

	MSG Msg{};
	BOOL bRet{};
	while ((bRet = GetMessage(&Msg, NULL, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			return (int)GetLastError();
		}
		else
		{
			DispatchMessage(&Msg);
		}
	}
	return (int)Msg.wParam;
}

HANDLE hParentProcess{};
HANDLE hTargetProcess{};
HWND hWndParentProcessMessage{};

HANDLE hEventMessageThreadReady{};
HANDLE hEventProxyProcessReady{};

DWORD dwTimeout{};

enum class USERMESSAGE : UINT { TERMINATE = WM_USER + 0x100, WATCHTHREADTERMINATED };
enum class COPYDATA : ULONG_PTR { PROXYPROCESSHWNDSENT, PROXYPROCESSDEBUGPRIVILEGEENABLEFINISHED, INJECTDLL, DLLINJECTIONFINISHED, PULLDLL, DLLPULLINGFINISHED, ADDFONT, ADDFONTFINISHED, REMOVEFONT, REMOVEFONTFINISHED, TERMINATE };
enum class PROXYPROCESSDEBUGPRIVILEGEENABLING { SUCCESSFUL, FAILED };
enum class PROXYDLLINJECTION { SUCCESSFUL, FAILED, FAILEDTOENUMERATEMODULES, GDI32NOTLOADED, MODULENOTFOUND };
enum class PROXYDLLPULL { SUCCESSFUL, FAILED };
enum class ADDFONT { SUCCESSFUL, FAILED };
enum class REMOVEFONT { SUCCESSFUL, FAILED };

void* pfnRemoteAddFontProc{};
void* pfnRemoteRemoveFontProc{};

#ifdef _WIN64
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll64.dll" };
#else
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll.dll" };
#endif // _WIN64

bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout);
bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout);
DWORD CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, std::size_t nParamSize);
bool EnableDebugPrivilege();

HANDLE hEventTerminateWatchThread{};
HANDLE hWatchThread{};

// Parent process watch thread
unsigned int __stdcall ParentProcessWatchThreadProc(void* lpParameter)
{
	// Wait for parent process or termination event
	HANDLE handles[]{ hParentProcess, hEventTerminateWatchThread };
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

	PostMessage((HWND)lpParameter, (UINT)USERMESSAGE::WATCHTHREADTERMINATED, NULL, NULL);

	return 0;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	switch ((USERMESSAGE)Msg)
	{
	case USERMESSAGE::WATCHTHREADTERMINATED:
		{
			// Destroy message-only window
			DestroyWindow(hWnd);
		}
		break;
	case USERMESSAGE::TERMINATE:
		{
			// Terminate watch thread
			SetEvent(hEventTerminateWatchThread);
			WaitForSingleObject(hWatchThread, INFINITE);
			CloseHandle(hEventTerminateWatchThread);
			CloseHandle(hWatchThread);

			// Destroy message-only window
			DestroyWindow(hWnd);
		}
		break;
	default:
		break;
	}
	switch (Msg)
	{
	case WM_CREATE:
		{
			// Get information from command line
			int argc{};
			LPWSTR* argv = CommandLineToArgvW(GetCommandLine(), &argc);

#pragma warning(push)
#pragma warning(disable:4312)
			// Get handles to parent process and message window
			hParentProcess = (HANDLE)std::wcstoul(argv[0], nullptr, 10);
			hTargetProcess = (HANDLE)std::wcstoul(argv[1], nullptr, 10);
			hWndParentProcessMessage = (HWND)std::wcstoul(argv[2], nullptr, 10);

			// Get handles to synchronization objects
			hEventMessageThreadReady = (HANDLE)std::wcstoul(argv[3], nullptr, 10);
			hEventProxyProcessReady = (HANDLE)std::wcstoul(argv[4], nullptr, 10);

			// Get timeout
			dwTimeout = (DWORD)std::wcstoul(argv[5], nullptr, 10);
#pragma warning(pop)

			// Wait for message thread to ready
			WaitForSingleObject(hEventMessageThreadReady, INFINITE);

			// Notify parent process ready
			SetEvent(hEventProxyProcessReady);

			// Send HWND to message window to parent process
			COPYDATASTRUCT cds2{ (ULONG_PTR)COPYDATA::PROXYPROCESSHWNDSENT, sizeof(HWND), &hWnd };
			FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds2, SendMessage);

			// Enable SeDebugPrivilege
			PROXYPROCESSDEBUGPRIVILEGEENABLING i{};
			COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::PROXYPROCESSDEBUGPRIVILEGEENABLEFINISHED, sizeof(PROXYPROCESSDEBUGPRIVILEGEENABLING), (void*)&i };
			if (EnableDebugPrivilege())
			{
				i = PROXYPROCESSDEBUGPRIVILEGEENABLING::SUCCESSFUL;
				FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
			}
			else
			{
				i = PROXYPROCESSDEBUGPRIVILEGEENABLING::FAILED;
				FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
			}

			// Start watch thread and create synchronization object
			hEventTerminateWatchThread = CreateEvent(NULL, TRUE, FALSE, NULL);
			hWatchThread = (HANDLE)_beginthreadex(nullptr, 0, ParentProcessWatchThreadProc, (void*)hWnd, 0, nullptr);
		}
		break;
	case WM_COPYDATA:
		{
			switch ((COPYDATA)((PCOPYDATASTRUCT)lParam)->dwData)
			{
				// Inject dll
			case COPYDATA::INJECTDLL:
				{
					// Check whether target process loads gdi32.dll as AddFontResourceEx() and RemoveFontResourceEx() are in it
					HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetThreadId(hTargetProcess)) };
					MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
					bool bIsGDI32Loaded{ false };
					if (!Module32First(hModuleSnapshot, &me32))
					{
						CloseHandle(hModuleSnapshot);

						PROXYDLLINJECTION i{ PROXYDLLINJECTION::FAILEDTOENUMERATEMODULES };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(PROXYDLLINJECTION), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}
					do
					{
						if (!lstrcmpi(me32.szModule, L"gdi32.dll"))
						{
							bIsGDI32Loaded = true;
							break;
						}
					} while (Module32Next(hModuleSnapshot, &me32));
					if (!bIsGDI32Loaded)
					{
						CloseHandle(hModuleSnapshot);

						PROXYDLLINJECTION i{ PROXYDLLINJECTION::GDI32NOTLOADED };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(PROXYDLLINJECTION), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}
					CloseHandle(hModuleSnapshot);

					// Inject FontLoaderExInjectionDll(64).dll into target process
					if (!InjectModule(hTargetProcess, szInjectionDllName, dwTimeout))
					{
						PROXYDLLINJECTION i{ PROXYDLLINJECTION::FAILED };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(PROXYDLLINJECTION), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}

					// Get base address of FontLoaderExInjectionDll(64).dll in target process
					HANDLE hModuleSnapshot2{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetProcessId(hTargetProcess)) };
					MODULEENTRY32 me322{ sizeof(MODULEENTRY32) };
					BYTE* pModBaseAddr{};
					if (!Module32First(hModuleSnapshot2, &me322))
					{
						CloseHandle(hModuleSnapshot2);

						PROXYDLLINJECTION i{ PROXYDLLINJECTION::MODULENOTFOUND };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(PROXYDLLINJECTION), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}
					do
					{
						if (!lstrcmpi(me322.szModule, szInjectionDllName))
						{
							pModBaseAddr = me322.modBaseAddr;
							break;
						}
					} while (Module32Next(hModuleSnapshot2, &me322));
					if (!pModBaseAddr)
					{
						CloseHandle(hModuleSnapshot2);

						PROXYDLLINJECTION i{ PROXYDLLINJECTION::MODULENOTFOUND };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(PROXYDLLINJECTION), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}
					CloseHandle(hModuleSnapshot2);

					// Calculate addresses of AddFont() and RemoveFont() in target process
					HMODULE hModInjectionDll{ LoadLibrary(szInjectionDllName) };
					void* pLocalAddFontProcAddr{ GetProcAddress(hModInjectionDll, "AddFont") };
					void* pLocalRemoveFontProcAddr{ GetProcAddress(hModInjectionDll, "RemoveFont") };
					FreeLibrary(hModInjectionDll);
					INT_PTR AddFontProcOffset{ (INT_PTR)pLocalAddFontProcAddr - (INT_PTR)hModInjectionDll };
					INT_PTR RemoveFontProcOffset{ (INT_PTR)pLocalRemoveFontProcAddr - (INT_PTR)hModInjectionDll };
					pfnRemoteAddFontProc = pModBaseAddr + AddFontProcOffset;
					pfnRemoteRemoveFontProc = pModBaseAddr + RemoveFontProcOffset;

					// Send success messsage to parent process
					PROXYDLLINJECTION i{ PROXYDLLINJECTION::SUCCESSFUL };
					COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(PROXYDLLINJECTION), (void*)&i };
					FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
				}
				break;
				// Pull dll
			case COPYDATA::PULLDLL:
				{
					PROXYDLLPULL i{};
					if (PullModule(hTargetProcess, szInjectionDllName, dwTimeout))
					{
						i = PROXYDLLPULL::SUCCESSFUL;
					}
					else
					{
						i = PROXYDLLPULL::FAILED;
					}
					COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLPULLINGFINISHED, sizeof(PROXYDLLPULL), (void*)&i };
					FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
				}
				break;
				// Load font
			case COPYDATA::ADDFONT:
				{
					ADDFONT i{};
					if (CallRemoteProc(hTargetProcess, pfnRemoteAddFontProc, ((PCOPYDATASTRUCT)lParam)->lpData, ((PCOPYDATASTRUCT)lParam)->cbData))
					{
						i = ADDFONT::SUCCESSFUL;
					}
					else
					{
						i = ADDFONT::FAILED;
					}
					COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::ADDFONTFINISHED, sizeof(ADDFONT), (void*)&i };
					FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
				}
				break;
				// Unload font
			case COPYDATA::REMOVEFONT:
				{
					REMOVEFONT i{};
					if (CallRemoteProc(hTargetProcess, pfnRemoteRemoveFontProc, ((PCOPYDATASTRUCT)lParam)->lpData, ((PCOPYDATASTRUCT)lParam)->cbData))
					{
						i = REMOVEFONT::SUCCESSFUL;
					}
					else
					{
						i = REMOVEFONT::FAILED;
					}
					COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::REMOVEFONTFINISHED, sizeof(REMOVEFONT), (void*)&i };
					FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
				}
				break;
				// Terminate self
			case COPYDATA::TERMINATE:
				{
					PostMessage(hWnd, (UINT)USERMESSAGE::TERMINATE, NULL, NULL);
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_DESTROY:
		{
			CloseHandle(hParentProcess);
			CloseHandle(hTargetProcess);

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

DWORD CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, std::size_t nParamSize)
{
	DWORD dwRet{};

	do
	{
		LPVOID lpRemoteBuffer{ VirtualAllocEx(hProcess, NULL, nParamSize, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE) };
		if (!lpRemoteBuffer)
		{
			dwRet = 0;
			break;
		}

		if (!WriteProcessMemory(hProcess, lpRemoteBuffer, lpParameter, nParamSize, NULL))
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			dwRet = 0;
			break;
		}

		HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, (LPTHREAD_START_ROUTINE)lpRemoteProcAddr, lpRemoteBuffer, 0, NULL) };
		if (!hRemoteThread)
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			dwRet = 0;
			break;
		}
		WaitForSingleObject(hRemoteThread, INFINITE);
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

		DWORD dwRemoteThreadExitCode{};
		if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
		{
			CloseHandle(hRemoteThread);

			dwRet = 0;
			break;
		}
		CloseHandle(hRemoteThread);

		dwRet = dwRemoteThreadExitCode;
	} while (false);

	return dwRet;
}

bool EnableDebugPrivilege()
{
	bool bRet{};

	// Enable SeDebugPrivilege
	do
	{
		HANDLE hToken{};
		LUID luid{};
		if (!OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, &hToken))
		{
			bRet = false;
			break;
		}

		if (!LookupPrivilegeValue(NULL, SE_DEBUG_NAME, &luid))
		{
			CloseHandle(hToken);

			bRet = false;
			break;
		}

		TOKEN_PRIVILEGES tp{ 1 , {luid, SE_PRIVILEGE_ENABLED} };
		if (!AdjustTokenPrivileges(hToken, FALSE, &tp, sizeof(tp), NULL, NULL))
		{
			CloseHandle(hToken);

			bRet = false;
			break;
		}
		CloseHandle(hToken);

		bRet = true;
	} while (false);

	return bRet;
}

bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout)
{
	bool bRet{};

	// Inject dll into target process
	do
	{
		WCHAR szDllPath[MAX_PATH]{};
		GetModuleFileName(NULL, szDllPath, MAX_PATH);
		PathRemoveFileSpec(szDllPath);
		PathAppend(szDllPath, szModuleName);

		// Allocate buffer in target process
		LPVOID lpRemoteBuffer{ VirtualAllocEx(hProcess, NULL, (std::wcslen(szDllPath) + 1) * sizeof(WCHAR), MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE) };
		if (!lpRemoteBuffer)
		{
			bRet = false;
			break;
		}

		// Write dll name to remote buffer
		if (!WriteProcessMemory(hProcess, lpRemoteBuffer, (LPVOID)szDllPath, (std::wcslen(szDllPath) + 1) * sizeof(WCHAR), NULL))
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;
			break;
		}

		HMODULE hModule{ GetModuleHandle(L"Kernel32") };
		if (!hModule)
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;
			break;
		}

		// Create remote thread to inject dll
		LPTHREAD_START_ROUTINE addr{ (LPTHREAD_START_ROUTINE)GetProcAddress(hModule, "LoadLibraryW") };
		HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, addr, lpRemoteBuffer, 0, NULL) };
		if (!hRemoteThread)
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;
			break;
		}
		if (WaitForSingleObject(hRemoteThread, Timeout) == WAIT_TIMEOUT)
		{
			CloseHandle(hRemoteThread);
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			bRet = false;
			break;
		}
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

		// Get exit code of remote thread
		DWORD dwRemoteThreadExitCode{};
		if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}

		if (!dwRemoteThreadExitCode)
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}
		CloseHandle(hRemoteThread);

		bRet = true;
	} while (false);

	return bRet;
}

bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout)
{
	bool bRet{};

	// Unload dll from target process
	do
	{
		// Find HMODULE of szModuleName in target process
		HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetProcessId(hProcess)) };
		MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
		HMODULE hModInjectionDll{};
		if (!Module32First(hModuleSnapshot, &me32))
		{
			CloseHandle(hModuleSnapshot);

			bRet = false;
			break;
		}
		do
		{
			if (!lstrcmpi(me32.szModule, szModuleName))
			{
				hModInjectionDll = me32.hModule;
				break;
			}
		} while (Module32Next(hModuleSnapshot, &me32));
		if (!hModInjectionDll)
		{
			CloseHandle(hModuleSnapshot);

			bRet = false;
			break;
		}
		CloseHandle(hModuleSnapshot);

		HMODULE hModule{ GetModuleHandle(L"Kernel32") };
		if (!hModule)
		{
			bRet = false;
			break;
		}

		// Create remote thread to unload dll
		LPTHREAD_START_ROUTINE addr{ (LPTHREAD_START_ROUTINE)GetProcAddress(hModule, "FreeLibrary") };
		HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, addr, (LPVOID)hModInjectionDll, 0, NULL) };
		if (!hRemoteThread)
		{
			bRet = false;
			break;
		}
		if (WaitForSingleObject(hRemoteThread, Timeout) == WAIT_TIMEOUT)
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}

		// Get exit code of remote thread
		DWORD dwRemoteThreadExitCode{};
		if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}
		if (!dwRemoteThreadExitCode)
		{
			CloseHandle(hRemoteThread);

			bRet = false;
			break;
		}
		CloseHandle(hRemoteThread);

		bRet = true;
	} while (false);

	return bRet;
}