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
#include <string>
#include <sstream>

constexpr WCHAR szAppName[]{ L"FontLoaderExProxy" };
constexpr WCHAR szParentAppName[]{ L"FontLoaderEx" };

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
	std::wstringstream ssMessage{};
	std::wstring strMessage{};

#ifdef _DEBUG
	// Wait for debugger to attach
	ssMessage << szAppName << L" launched!";
	strMessage = ssMessage.str();
	MessageBox(NULL, strMessage.c_str(), szAppName, NULL);
	ssMessage.str(L"");
#endif // _DEBUG

	// Detect whether FontLoaderEx is running. If not running, launch it.
	ssMessage << L"Never run " << szAppName << L" directly, run " << szParentAppName << " instead.\r\n\r\nDo you want to launch FontloaderEx now?";
	strMessage = ssMessage.str();
	HANDLE hEventParentProcessRunning{ OpenEvent(EVENT_ALL_ACCESS, FALSE, L"FontLoaderEx_EventParentProcessRunning_B980D8A4-C487-4306-9D17-3BA6A2CCA4A4") };
	if (!hEventParentProcessRunning)
	{
		switch (MessageBox(NULL, strMessage.c_str(), szAppName, MB_ICONINFORMATION | MB_YESNO))
		{
		case IDYES:
			{
				STARTUPINFO sa{ sizeof(STARTUPINFO) };
				PROCESS_INFORMATION pi{};
				CreateProcess(L"FontLoaderEx.exe", NULL, NULL, NULL, FALSE, 0, NULL, NULL, &sa, &pi);
				CloseHandle(pi.hThread);
				CloseHandle(pi.hProcess);
			}
			break;
		case IDNO:
			break;
		default:
			break;
		}

		return 0;
	}
	else
	{
		CloseHandle(hEventParentProcessRunning);
	}

	// Create message-only window
	WNDCLASS wc{ 0, WndProc, 0, 0, hInstance, NULL, NULL, NULL, NULL, szAppName };
	if (!RegisterClass(&wc))
	{
		return -1;
	}
	HWND hWndMain{};
	if (!(hWndMain = CreateWindow(szAppName, szAppName, NULL, 0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL)))
	{
		return -1;
	}

	MSG Message{};
	BOOL bRet{};
	while ((bRet = GetMessage(&Message, NULL, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			return (int)GetLastError();
		}
		else
		{
			DispatchMessage(&Message);
		}
	}
	return (int)Message.wParam;
}

HANDLE hProcessParent{};
HANDLE hProcessTarget{};
HWND hWndParentMessage{};

HANDLE hEventMessageThreadReady{};
HANDLE hEventProxyProcessReady{};

DWORD dwTimeout{};

enum class USERMESSAGE : UINT { TERMINATE = WM_USER + 0x100, WATCHTHREADTERMINATED };
enum class COPYDATA : ULONG_PTR { PROXYPROCESSHWNDSENT, PROXYPROCESSDEBUGPRIVILEGEENABLEFINISHED, INJECTDLL, DLLINJECTIONFINISHED, PULLDLL, DLLPULLINGFINISHED, ADDFONT, ADDFONTFINISHED, REMOVEFONT, REMOVEFONTFINISHED, TERMINATE };
enum class PROXYPROCESSDEBUGPRIVILEGEENABLING : UINT { SUCCESSFUL, FAILED };
enum class PROXYDLLINJECTION : UINT { SUCCESSFUL, FAILED, FAILEDTOENUMERATEMODULES, GDI32NOTLOADED, MODULENOTFOUND };
enum class PROXYDLLPULL : UINT { SUCCESSFUL, FAILED };
enum class ADDFONT : UINT { SUCCESSFUL, FAILED };
enum class REMOVEFONT : UINT { SUCCESSFUL, FAILED };

void* lpRemoteAddFontProcAddr{};
void* lpRemoteRemoveFontProcAddr{};

#ifdef _WIN64
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll64.dll" };
#else
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll.dll" };
#endif // _WIN64

bool EnableDebugPrivilege();
bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout);
bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout);
DWORD CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, std::size_t cbParamSize, DWORD dwTimeout);

HANDLE hEventTerminateWatchThread{};
HANDLE hThreadWatch{};

// Parent process watch thread
unsigned int __stdcall ParentProcessWatchThreadProc(void* lpParameter)
{
	// Wait for parent process or termination event
	HANDLE handles[]{ hProcessParent, hEventTerminateWatchThread };
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

	PostMessage(reinterpret_cast<HWND>(lpParameter), static_cast<UINT>(USERMESSAGE::WATCHTHREADTERMINATED), 0, 0);

	return 0;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	switch (static_cast<USERMESSAGE>(Message))
	{
	case USERMESSAGE::WATCHTHREADTERMINATED:
		{
			// Destroy message-only window
			PostMessage(hWnd, WM_CLOSE, 0, 0);
		}
		break;
	case USERMESSAGE::TERMINATE:
		{
			// Terminate watch thread
			SetEvent(hEventTerminateWatchThread);
			WaitForSingleObject(hThreadWatch, INFINITE);
			CloseHandle(hEventTerminateWatchThread);
			CloseHandle(hThreadWatch);

			// Destroy message-only window
			PostMessage(hWnd, WM_CLOSE, 0, 0);
		}
		break;
	default:
		break;
	}
	switch (Message)
	{
	case WM_CREATE:
		{
			// Get information from command line
			int argc{};
			LPWSTR* argv = CommandLineToArgvW(GetCommandLine(), &argc);
#ifdef _WIN64
			// Get handles to parent process and message window
			hProcessParent = reinterpret_cast<HANDLE>(std::wcstoull(argv[0], nullptr, 16));
			hProcessTarget = reinterpret_cast<HANDLE>(std::wcstoull(argv[1], nullptr, 16));
			hWndParentMessage = reinterpret_cast<HWND>(std::wcstoull(argv[2], nullptr, 16));

			// Get handles to synchronization objects
			hEventMessageThreadReady = reinterpret_cast<HANDLE>(std::wcstoull(argv[3], nullptr, 16));
			hEventProxyProcessReady = reinterpret_cast<HANDLE>(std::wcstoull(argv[4], nullptr, 16));
#else
			// Get handles to parent process and message window
			hProcessParent = reinterpret_cast<HANDLE>(std::wcstoul(argv[0], nullptr, 16));
			hProcessTarget = reinterpret_cast<HANDLE>(std::wcstoul(argv[1], nullptr, 16));
			hWndParentMessage = reinterpret_cast<HWND>(std::wcstoul(argv[2], nullptr, 16));

			// Get handles to synchronization objects
			hEventMessageThreadReady = reinterpret_cast<HANDLE>(std::wcstoul(argv[3], nullptr, 16));
			hEventProxyProcessReady = reinterpret_cast<HANDLE>(std::wcstoul(argv[4], nullptr, 16));
#endif	// _WIN64

			// Get timeout
			dwTimeout = static_cast<DWORD>(std::wcstoul(argv[5], nullptr, 10));

			// Wait for message thread to ready
			WaitForSingleObject(hEventMessageThreadReady, INFINITE);

			// Notify parent process ready
			SetEvent(hEventProxyProcessReady);

			// Send HWND to message window to parent process
			COPYDATASTRUCT cds2{ static_cast<ULONG_PTR>(COPYDATA::PROXYPROCESSHWNDSENT), sizeof(HWND), &hWnd };
			FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds2, SendMessage);

			// Enable SeDebugPrivilege
			PROXYPROCESSDEBUGPRIVILEGEENABLING i{};
			COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::PROXYPROCESSDEBUGPRIVILEGEENABLEFINISHED), sizeof(PROXYPROCESSDEBUGPRIVILEGEENABLING), &i };
			if (EnableDebugPrivilege())
			{
				i = PROXYPROCESSDEBUGPRIVILEGEENABLING::SUCCESSFUL;
				FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);
			}
			else
			{
				i = PROXYPROCESSDEBUGPRIVILEGEENABLING::FAILED;
				FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);
			}

			// Start watch thread and create synchronization object
			hEventTerminateWatchThread = CreateEvent(NULL, TRUE, FALSE, NULL);
			hThreadWatch = reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, ParentProcessWatchThreadProc, hWnd, 0, nullptr));
		}
		break;
	case WM_COPYDATA:
		{
			// COPYDATASTRUCT::dwData = Command : enum COPYDATA
			// COPYDATASTRUCT::lpData = Data
			switch (static_cast<COPYDATA>(reinterpret_cast<PCOPYDATASTRUCT>(lParam)->dwData))
			{
				// Inject dll
			case COPYDATA::INJECTDLL:
				{
					PROXYDLLINJECTION i{};

					// Check whether target process loads gdi32.dll as AddFontResourceEx() and RemoveFontResourceEx() are in it
					HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetThreadId(hProcessTarget)) };
					MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
					bool bIsGDI32Loaded{ false };
					if (!Module32First(hModuleSnapshot, &me32))
					{
						CloseHandle(hModuleSnapshot);

						i = PROXYDLLINJECTION::FAILEDTOENUMERATEMODULES;
						COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::DLLINJECTIONFINISHED), sizeof(PROXYDLLINJECTION), &i };
						FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);

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

						i = PROXYDLLINJECTION::GDI32NOTLOADED;
						COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::DLLINJECTIONFINISHED), sizeof(PROXYDLLINJECTION), &i };
						FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);

						break;
					}
					CloseHandle(hModuleSnapshot);

					// Inject FontLoaderExInjectionDll(64).dll into target process
					if (!InjectModule(hProcessTarget, szInjectionDllName, dwTimeout))
					{
						i = PROXYDLLINJECTION::FAILED;
						COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::DLLINJECTIONFINISHED), sizeof(PROXYDLLINJECTION), &i };
						FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);

						break;
					}

					// Get base address of FontLoaderExInjectionDll(64).dll in target process
					HANDLE hModuleSnapshot2{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetProcessId(hProcessTarget)) };
					MODULEENTRY32 me322{ sizeof(MODULEENTRY32) };
					BYTE* lpModBaseAddr{};
					if (!Module32First(hModuleSnapshot2, &me322))
					{
						CloseHandle(hModuleSnapshot2);

						i = PROXYDLLINJECTION::MODULENOTFOUND;
						COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::DLLINJECTIONFINISHED), sizeof(PROXYDLLINJECTION), &i };
						FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);

						break;
					}
					do
					{
						if (!lstrcmpi(me322.szModule, szInjectionDllName))
						{
							lpModBaseAddr = me322.modBaseAddr;

							break;
						}
					} while (Module32Next(hModuleSnapshot2, &me322));
					if (!lpModBaseAddr)
					{
						CloseHandle(hModuleSnapshot2);

						i = PROXYDLLINJECTION::MODULENOTFOUND;
						COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::DLLINJECTIONFINISHED), sizeof(PROXYDLLINJECTION), &i };
						FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);

						break;
					}
					CloseHandle(hModuleSnapshot2);

					// Calculate addresses of AddFont() and RemoveFont() in target process
					HMODULE hModInjectionDll{ LoadLibrary(szInjectionDllName) };
					void* lpLocalAddFontProcAddr{ GetProcAddress(hModInjectionDll, "AddFont") };
					void* lpLocalRemoveFontProcAddr{ GetProcAddress(hModInjectionDll, "RemoveFont") };
					FreeLibrary(hModInjectionDll);
					lpRemoteAddFontProcAddr = reinterpret_cast<void*>(reinterpret_cast<UINT_PTR>(lpModBaseAddr) + (reinterpret_cast<UINT_PTR>(lpLocalAddFontProcAddr) - reinterpret_cast<UINT_PTR>(hModInjectionDll)));
					lpRemoteRemoveFontProcAddr = reinterpret_cast<void*>(reinterpret_cast<UINT_PTR>(lpModBaseAddr) + (reinterpret_cast<UINT_PTR>(lpLocalRemoveFontProcAddr) - reinterpret_cast<UINT_PTR>(hModInjectionDll)));

					// Send success messsage to parent process
					i = PROXYDLLINJECTION::SUCCESSFUL;
					COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::DLLINJECTIONFINISHED), sizeof(PROXYDLLINJECTION), &i };
					FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);
				}
				break;
				// Pull dll
			case COPYDATA::PULLDLL:
				{
					PROXYDLLPULL i{};
					if (PullModule(hProcessTarget, szInjectionDllName, dwTimeout))
					{
						i = PROXYDLLPULL::SUCCESSFUL;
					}
					else
					{
						i = PROXYDLLPULL::FAILED;
					}
					COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::DLLPULLINGFINISHED), sizeof(PROXYDLLPULL), &i };
					FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);
				}
				break;
				// Load font
			case COPYDATA::ADDFONT:
				{
					ADDFONT i{};
					if (CallRemoteProc(hProcessTarget, lpRemoteAddFontProcAddr, reinterpret_cast<PCOPYDATASTRUCT>(lParam)->lpData, reinterpret_cast<PCOPYDATASTRUCT>(lParam)->cbData, INFINITE))
					{
						i = ADDFONT::SUCCESSFUL;
					}
					else
					{
						i = ADDFONT::FAILED;
					}
					COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::ADDFONTFINISHED), sizeof(ADDFONT), &i };
					FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);
				}
				break;
				// Unload font
			case COPYDATA::REMOVEFONT:
				{
					REMOVEFONT i{};
					if (CallRemoteProc(hProcessTarget, lpRemoteRemoveFontProcAddr, reinterpret_cast<PCOPYDATASTRUCT>(lParam)->lpData, reinterpret_cast<PCOPYDATASTRUCT>(lParam)->cbData, INFINITE))
					{
						i = REMOVEFONT::SUCCESSFUL;
					}
					else
					{
						i = REMOVEFONT::FAILED;
					}
					COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::REMOVEFONTFINISHED), sizeof(REMOVEFONT), &i };
					FORWARD_WM_COPYDATA(hWndParentMessage, hWnd, &cds, SendMessage);
				}
				break;
				// Terminate self
			case COPYDATA::TERMINATE:
				{
					PostMessage(hWnd, static_cast<UINT>(USERMESSAGE::TERMINATE), 0, 0);
				}
				break;
			default:
				break;
			}
		}
		break;
	case WM_DESTROY:
		{
			CloseHandle(hProcessParent);
			CloseHandle(hProcessTarget);
			CloseHandle(hEventMessageThreadReady);
			CloseHandle(hEventProxyProcessReady);

			PostQuitMessage(0);
		}
		break;
	default:
		{
			ret = DefWindowProc(hWnd, Message, wParam, lParam);
		}
		break;
	}

	return ret;
}

bool EnableDebugPrivilege()
{
	// Enable SeDebugPrivilege

	bool bRet{};

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

bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout)
{
	// Inject dll into target process
	bool bRet{};

	// Make full path to module
	WCHAR szDllPath[MAX_PATH]{};
	GetModuleFileName(NULL, szDllPath, MAX_PATH);
	PathRemoveFileSpec(szDllPath);
	PathAppend(szDllPath, szModuleName);

	// Call LoadLibraryW with module full path to inject dll into hProcess
	bRet = CallRemoteProc(hProcess, GetProcAddress(GetModuleHandle(L"Kernel32"), "LoadLibraryW"), szDllPath, (std::wcslen(szDllPath) + 1) * sizeof(WCHAR), dwTimeout);

	return bRet;
}

bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD dwTimeout)
{
	// Unload dll from target process
	bool bRet{};

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

		// Call FreeLibrary with HMODULE to unload dll from hProcess
		bRet = CallRemoteProc(hProcess, GetProcAddress(GetModuleHandle(L"Kernel32"), "FreeLibrary"), hModInjectionDll, 0, dwTimeout);
	} while (false);

	return bRet;
}

DWORD CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, std::size_t cbParamSize, DWORD dwTimeout)
{
	DWORD dwRet{};

	do
	{
		LPVOID lpRemoteBuffer{};
		// If cbParamSize == 0, directly copy lpParameter to lpRemoteBuffer
		if (cbParamSize == 0)
		{
			lpRemoteBuffer = lpParameter;
		}
		// Else as normal
		else
		{
			// Allocate buffer in target process
			lpRemoteBuffer = VirtualAllocEx(hProcess, NULL, cbParamSize, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
			if (!lpRemoteBuffer)
			{
				dwRet = 0;

				break;
			}

			// Write parameter to remote buffer
			if (!WriteProcessMemory(hProcess, lpRemoteBuffer, lpParameter, cbParamSize, NULL))
			{
				VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

				dwRet = 0;

				break;
			}
		}

		// Create remote thread to call function
		HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, static_cast<LPTHREAD_START_ROUTINE>(lpRemoteProcAddr), lpRemoteBuffer, 0, NULL) };
		if (!hRemoteThread)
		{
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			dwRet = 0;

			break;
		}

		// Wait for remote thread to terminate with timeout
		if (WaitForSingleObject(hRemoteThread, dwTimeout) == WAIT_TIMEOUT)
		{
			CloseHandle(hRemoteThread);
			VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

			dwRet = 0;

			break;
		}
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

		// Get exit code of remote thread
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