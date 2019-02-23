#pragma comment(lib, "Shlwapi.lib")

#include <windows.h>
#include <windowsx.h>
#include <tlhelp32.h>
#include <shlwapi.h>
#include <cwchar>

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam);

int WINAPI wWinMain(_In_ HINSTANCE hInstance, _In_opt_ HINSTANCE hPrevInstance, _In_ LPWSTR lpCmdLine, _In_ int nShowCmd)
{
#ifdef _DEBUG
	//Wait for debugger to attach
	MessageBox(NULL, L"", L"", NULL);
#endif

	//Detect whether FontLoaderEx launchs FontloaderExProxy
	HANDLE hEventParentProcessRunning{ OpenEvent(EVENT_ALL_ACCESS, FALSE, L"FontLoaderEx_EventParentProcessRunning") };
	if (!hEventParentProcessRunning)
	{
		return 0;
	}

	//Create window
	WNDCLASS wndclass{ 0, WndProc, 0, 0, hInstance, NULL, NULL, NULL, NULL, L"FontLoaderExProxy" };
	if (!RegisterClass(&wndclass))
	{
		return -1;
	}

	HWND hWnd{};
	if (!(hWnd = CreateWindow(L"FontLoaderExProxy", L"FontLoaderExProxy", NULL, 0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL)))
	{
		return -1;
	}

	MSG Msg{};
	BOOL bRet{};
	while ((bRet = GetMessage(&Msg, hWnd, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			return (int)GetLastError();
			break;
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

enum class USERMESSAGE : UINT { TERMINATE = WM_USER + 0x100 };

enum class COPYDATA : ULONG_PTR { PROXYPROCESSHWNDSENT, INJECTDLL, DLLINJECTIONFINISHED, PULLDLL, DLLPULLFINISHED, ADDFONT, ADDFONTFINISHED, REMOVEFONT, REMOVEFONTFINISHED, TERMINATE };
enum class PROXYDLLINJECTION : int { SUCCESSFUL, FAILED, FAILEDTOENUMERATEMODULES, GDI32NOTLOADED, MODULENOTFOUND };
enum class PROXYDLLPULL : int { SUCCESSFUL, FAILED };

void* pfnRemoteAddFontProc{};
void* pfnRemoteRemoveFontProc{};

#ifdef _WIN64
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll64.dll" };
#else
const WCHAR szInjectionDllName[]{ L"FontLoaderExInjectionDll.dll" };
#endif

bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout);
bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout);
DWORD CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, size_t nParamSize, DWORD Timeout);

LRESULT CALLBACK WndProc(HWND hWnd, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};
	switch (Message)
	{
	case WM_CREATE:
		{
			//Get handles to parent process and target process and message window from command line
			int argc{};
			LPWSTR* argv = CommandLineToArgvW(GetCommandLine(), &argc);

#pragma warning(push)
#pragma warning(disable:4312)
			hParentProcess = (HANDLE)std::wcstoul(argv[0], nullptr, 10);
			hTargetProcess = (HANDLE)std::wcstoul(argv[1], nullptr, 10);
			hWndParentProcessMessage = (HWND)std::wcstoul(argv[2], nullptr, 10);

			//Get handles to synchronization objects
			hEventMessageThreadReady = (HANDLE)std::wcstoul(argv[3], nullptr, 10);
			hEventProxyProcessReady = (HANDLE)std::wcstoul(argv[4], nullptr, 10);
#pragma warning(pop)

			//Wait for message thread to ready
			WaitForSingleObject(hEventMessageThreadReady, INFINITE);

			//Notify parent process ready
			SetEvent(hEventProxyProcessReady);

			//Send HWND to parent process
			COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::PROXYPROCESSHWNDSENT, sizeof(HWND), &hWnd };
			FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
		}
		break;
	case WM_COPYDATA:
		{
			COPYDATASTRUCT* pcds{ (PCOPYDATASTRUCT)lParam };
			switch (pcds->dwData)
			{
				//Inject dll
			case (ULONG_PTR)COPYDATA::INJECTDLL:
				{
					//Check whether target process loads gdi32.dll as AddFontResourceEx() and RemoveFontResourceEx() are in it
					HANDLE hModuleSnapshot1{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetThreadId(hTargetProcess)) };
					MODULEENTRY32 me321{ sizeof(MODULEENTRY32) };
					bool bIsGDI32Loaded{ false };
					if (!Module32First(hModuleSnapshot1, &me321))
					{
						CloseHandle(hModuleSnapshot1);
						int i{ (int)PROXYDLLINJECTION::FAILEDTOENUMERATEMODULES };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(int), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}
					do
					{
						if (!lstrcmpi(me321.szModule, L"gdi32.dll"))
						{
							bIsGDI32Loaded = true;
							break;
						}
					} while (Module32Next(hModuleSnapshot1, &me321));
					if (!bIsGDI32Loaded)
					{
						CloseHandle(hModuleSnapshot1);
						int i{ (int)PROXYDLLINJECTION::GDI32NOTLOADED };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(int), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}
					CloseHandle(hModuleSnapshot1);

					//Inject FontLoaderExInjectionDll(64).dll into target process
					if (!InjectModule(hTargetProcess, szInjectionDllName, 5000))
					{
						int i{ (int)PROXYDLLINJECTION::FAILED };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(int), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}

					//Get base address of FontLoaderExInjectionDll(64).dll in target process
					HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetProcessId(hTargetProcess)) };
					MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
					BYTE* pModBaseAddr{};
					if (!Module32First(hModuleSnapshot, &me32))
					{
						CloseHandle(hModuleSnapshot);
						int i{ (int)PROXYDLLINJECTION::MODULENOTFOUND };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(int), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}
					do
					{
						if (!lstrcmpi(me32.szModule, szInjectionDllName))
						{
							pModBaseAddr = me32.modBaseAddr;
							break;
						}
					} while (Module32Next(hModuleSnapshot, &me32));
					if (!pModBaseAddr)
					{
						CloseHandle(hModuleSnapshot);
						int i{ (int)PROXYDLLINJECTION::MODULENOTFOUND };
						COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(int), (void*)&i };
						FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
						break;
					}
					CloseHandle(hModuleSnapshot);

					//Calculate addresses of AddFont() and RemoveFont() in target process
					HMODULE hModInjectionDll{ LoadLibrary(szInjectionDllName) };
					void* pLocalAddFontProcAddr{ GetProcAddress(hModInjectionDll, "AddFont") };
					void* pLocalRemoveFontProcAddr{ GetProcAddress(hModInjectionDll, "RemoveFont") };
					INT_PTR AddFontProcOffset{ (INT_PTR)pLocalAddFontProcAddr - (INT_PTR)hModInjectionDll };
					INT_PTR RemoveFontProcOffset{ (INT_PTR)pLocalRemoveFontProcAddr - (INT_PTR)hModInjectionDll };
					FreeLibrary(hModInjectionDll);
					pfnRemoteAddFontProc = pModBaseAddr + AddFontProcOffset;
					pfnRemoteRemoveFontProc = pModBaseAddr + RemoveFontProcOffset;

					//Send success messsage to parent process
					int i{ (int)PROXYDLLINJECTION::SUCCESSFUL };
					COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLINJECTIONFINISHED, sizeof(int), (void*)&i };
					FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
				}
				break;
				//Pull dll
			case (ULONG_PTR)COPYDATA::PULLDLL:
				{
					int i{};
					if (!PullModule(hTargetProcess, szInjectionDllName, 5000))
					{
						i = (int)PROXYDLLPULL::FAILED;
					}
					else
					{
						i = (int)PROXYDLLPULL::SUCCESSFUL;
					}
					COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::DLLPULLFINISHED, sizeof(int), (void*)&i };
					FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
				}
				break;
				//Load font
			case (ULONG_PTR)COPYDATA::ADDFONT:
				{
					bool b{};
					if (!CallRemoteProc(hTargetProcess, pfnRemoteAddFontProc, pcds->lpData, (std::wcslen((LPWSTR)pcds->lpData) + 1) * sizeof(wchar_t), 5000))
					{
						b = false;
					}
					else
					{
						b = true;
					}
					COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::ADDFONTFINISHED, sizeof(int), (void*)&b };
					FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
				}
				break;
				//Unload font
			case (ULONG_PTR)COPYDATA::REMOVEFONT:
				{
					bool b{};
					if (!CallRemoteProc(hTargetProcess, pfnRemoteRemoveFontProc, pcds->lpData, (std::wcslen((LPWSTR)pcds->lpData) + 1) * sizeof(wchar_t), 5000))
					{
						b = false;
					}
					else
					{
						b = true;
					}
					COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::REMOVEFONTFINISHED, sizeof(int), (void*)&b };
					FORWARD_WM_COPYDATA(hWndParentProcessMessage, hWnd, &cds, SendMessage);
				}
				break;
				//Terminate self
			case (ULONG_PTR)COPYDATA::TERMINATE:
				{
					PostMessage(hWnd, (UINT)USERMESSAGE::TERMINATE, NULL, NULL);
				}
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
	case (UINT)USERMESSAGE::TERMINATE:
		{
			DestroyWindow(hWnd);
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

bool InjectModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout)
{
	//Inject dll to target process
	WCHAR szDllPath[MAX_PATH]{};
	GetModuleFileName(NULL, szDllPath, MAX_PATH);
	PathRemoveFileSpec(szDllPath);
	PathAppend(szDllPath, szModuleName);
	LPVOID lpRemoteBuffer{ VirtualAllocEx(hProcess, NULL, (std::wcslen(szDllPath) + 1) * sizeof(WCHAR), MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE) };
	if (!lpRemoteBuffer)
	{
		DWORD dw = GetLastError();
		return false;
	}
	if (!WriteProcessMemory(hProcess, lpRemoteBuffer, (LPVOID)szDllPath, (std::wcslen(szDllPath) + 1) * sizeof(WCHAR), NULL))
	{
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);
		return false;
	}
	HMODULE hModule{ GetModuleHandle(L"Kernel32") };
	if (!hModule)
	{
		return false;
	}
	LPTHREAD_START_ROUTINE addr{ (LPTHREAD_START_ROUTINE)GetProcAddress(hModule, "LoadLibraryW") };
	HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, addr, lpRemoteBuffer, 0, NULL) };
	if (!hRemoteThread)
	{
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);
		return false;
	}
	if (WaitForSingleObject(hRemoteThread, Timeout) == WAIT_TIMEOUT)
	{
		CloseHandle(hRemoteThread);
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);
		return false;
	}
	VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

	//Get exit code of remote thread
	DWORD dwRemoteThreadExitCode{};
	if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
	{
		CloseHandle(hRemoteThread);
		return false;
	}
	if (!dwRemoteThreadExitCode)
	{
		CloseHandle(hRemoteThread);
		return false;
	}
	CloseHandle(hRemoteThread);

	return true;
}

bool PullModule(HANDLE hProcess, LPCWSTR szModuleName, DWORD Timeout)
{
	//Find HMODULE of szModuleName in target process
	HANDLE hModuleSnapshot{ CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetProcessId(hProcess)) };
	MODULEENTRY32 me32{ sizeof(MODULEENTRY32) };
	HMODULE hModInjectionDll{};
	if (!Module32First(hModuleSnapshot, &me32))
	{
		CloseHandle(hModuleSnapshot);
		return false;
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
		return false;
	}
	CloseHandle(hModuleSnapshot);

	//Unload FontLoaderExInjectionDll(64).dll from target process
	HMODULE hModule{ GetModuleHandle(L"Kernel32") };
	LPTHREAD_START_ROUTINE addr{ (LPTHREAD_START_ROUTINE)GetProcAddress(hModule, "FreeLibrary") };
	HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, addr, (LPVOID)hModInjectionDll, 0, NULL) };
	if (!hRemoteThread)
	{
		return false;
	}
	if (WaitForSingleObject(hRemoteThread, Timeout) == WAIT_TIMEOUT)
	{
		CloseHandle(hRemoteThread);
		return false;
	}

	//Get exit code of remote thread
	DWORD dwRemoteThreadExitCode{};
	if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
	{
		CloseHandle(hRemoteThread);
		return false;
	}
	if (!dwRemoteThreadExitCode)
	{
		CloseHandle(hRemoteThread);
		return false;
	}
	CloseHandle(hRemoteThread);

	return true;
}

DWORD CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, size_t nParamSize, DWORD Timeout)
{
	LPVOID lpRemoteBuffer{ VirtualAllocEx(hProcess, NULL, nParamSize, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE) };
	if (!lpRemoteBuffer)
	{
		return false;
	}

	if (!WriteProcessMemory(hProcess, lpRemoteBuffer, lpParameter, nParamSize, NULL))
	{
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);
		return false;
	}

	HANDLE hRemoteThread{ CreateRemoteThread(hProcess, NULL, 0, (LPTHREAD_START_ROUTINE)lpRemoteProcAddr, lpRemoteBuffer, 0, NULL) };
	if (!hRemoteThread)
	{
		VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);
		return false;
	}
	if (WaitForSingleObject(hRemoteThread, Timeout) == WAIT_TIMEOUT)
	{
		return false;
	}
	VirtualFreeEx(hProcess, lpRemoteBuffer, 0, MEM_RELEASE);

	DWORD dwRemoteThreadExitCode{};
	if (!GetExitCodeThread(hRemoteThread, &dwRemoteThreadExitCode))
	{
		CloseHandle(hRemoteThread);
		return false;
	}

	CloseHandle(hRemoteThread);
	return dwRemoteThreadExitCode;
}