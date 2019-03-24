#include <Windows.h>
#include <windowsx.h>
#include <cstddef>
#include "FontResource.h"
#include "Globals.h"

FontResource::pfnAddFontProc FontResource::AddFontProc_{};
FontResource::pfnRemoveFontProc FontResource::RemoveFontProc_{};

HANDLE hEventProxyAddFontFinished{};
HANDLE hEventProxyRemoveFontFinished{};

ADDFONT ProxyAddFontResult{};
REMOVEFONT ProxyRemoveFontResult{};

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

#ifdef _DEBUG
bool DefaultAddFontProc(const wchar_t* lpszFontName)
{
	Sleep(ADDFONT_WAIT_MILLISEC);

	return true;
}

bool DefaultRemoveFontProc(const wchar_t* lpszFontName)
{
	Sleep(REMOVEFONT_WAIT_MILLISEC);

	return true;
}
#else
bool DefaultAddFontProc(const wchar_t* lpszFontName)
{
	return AddFontResourceEx(lpszFontName, 0, NULL);
}

bool DefaultRemoveFontProc(const wchar_t* lpszFontName)
{
	return RemoveFontResourceEx(lpszFontName, 0, NULL);
}
#endif // _DEBUG

bool RemoteAddFontProc(const wchar_t* lpszFontName)
{
	return CallRemoteProc(TargetProcessInfo.hProcess, pfnRemoteAddFontProc, (void*)lpszFontName, (std::wcslen(lpszFontName) + 1) * sizeof(wchar_t));
}

bool RemoteRemoveFontProc(const wchar_t* lpszFontName)
{
	return CallRemoteProc(TargetProcessInfo.hProcess, pfnRemoteRemoveFontProc, (void*)lpszFontName, (std::wcslen(lpszFontName) + 1) * sizeof(wchar_t));
}

bool ProxyAddFontProc(const wchar_t* lpszFontName)
{
	bool bRet{};

	COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::ADDFONT, (DWORD)(std::wcslen(lpszFontName) + 1) * sizeof(wchar_t), (void*)lpszFontName };
	FORWARD_WM_COPYDATA(hWndProxy, hWndMain, &cds, SendMessage);
	WaitForSingleObject(hEventProxyAddFontFinished, INFINITE);
	ResetEvent(hEventProxyAddFontFinished);
	switch (ProxyAddFontResult)
	{
	case ADDFONT::SUCCESSFUL:
		{
			bRet = true;
		}
		break;
	case ADDFONT::FAILED:
		{
			bRet = false;
		}
	default:
		break;
	}

	return bRet;
}

bool ProxyRemoveFontProc(const wchar_t* lpszFontName)
{
	bool bRet{};

	COPYDATASTRUCT cds{ (ULONG_PTR)COPYDATA::REMOVEFONT, (DWORD)(std::wcslen(lpszFontName) + 1) * sizeof(wchar_t), (void*)lpszFontName };
	FORWARD_WM_COPYDATA(hWndProxy, hWndMain, &cds, SendMessage);
	WaitForSingleObject(hEventProxyRemoveFontFinished, INFINITE);
	ResetEvent(hEventProxyRemoveFontFinished);
	switch (ProxyRemoveFontResult)
	{
	case REMOVEFONT::SUCCESSFUL:
		{
			bRet = true;
		}
		break;
	case REMOVEFONT::FAILED:
		{
			bRet = false;
		}
	default:
		break;
	}

	return bRet;
}

bool NullAddFontProc(const wchar_t* lpszFontName)
{
	return true;
}

bool NullRemoveFontProc(const wchar_t* lpszFontName)
{
	return true;
}

FontResource::FontResource(const std::wstring& FontName) : strFontName_(FontName)
{
	return;
}

FontResource::FontResource(const wchar_t* FontName) : strFontName_(FontName)
{
	return;
}

FontResource::~FontResource()
{
	if (bIsLoaded_)
	{
		RemoveFontProc_(strFontName_.c_str());
	}
}

void FontResource::RegisterAddRemoveFontProc(pfnAddFontProc AddFontProc, pfnRemoveFontProc RemoveFontProc)
{
	AddFontProc_ = AddFontProc;
	RemoveFontProc_ = RemoveFontProc;
}

bool FontResource::Load()
{
	bool bRet;

	if (!bIsLoaded_)
	{
		if (AddFontProc_(strFontName_.c_str()))
		{
			bIsLoaded_ = true;
			bRet = true;
		}
		else
		{
			bRet = false;
		}
	}
	else
	{
		bRet = true;
	}

	return bRet;
}

bool FontResource::Unload()
{
	bool bRet;

	if (bIsLoaded_)
	{
		if (RemoveFontProc_(strFontName_.c_str()))
		{
			bIsLoaded_ = false;
			bRet = true;
		}
		else
		{
			bRet = false;
		}
	}
	else
	{
		bRet = true;
	}

	return bRet;
}

const std::wstring & FontResource::GetFontName()
{
	return strFontName_;
}

bool FontResource::IsLoaded()
{
	return bIsLoaded_;
}