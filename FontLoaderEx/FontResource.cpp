#include <cstring>
#include "FontResource.h"
#include "Globals.h"

FontResource::pfnAddFontProc FontResource::AddFontProc_{};
FontResource::pfnRemoveFontProc FontResource::RemoveFontProc_{};

DWORD CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, size_t nParamSize)
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
	if (WaitForSingleObject(hRemoteThread, 5000) == WAIT_TIMEOUT)
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

#ifdef _DEBUG
bool DefaultAddFontProc(const wchar_t* lpFontName)
{
	Sleep(300);
	return true;
}

bool DefaultRemoveFontProc(const wchar_t* lpFontName)
{
	Sleep(300);
	return true;
}
#else
bool DefaultAddFontProc(const wchar_t* lpFontName)
{
	return AddFontResourceEx(lpFontName, 0, NULL);
}

bool DefaultRemoveFontProc(const wchar_t* lpFontName)
{
	return RemoveFontResourceEx(lpFontName, 0, NULL);
}
#endif // _DEBUG

bool RemoteAddFontProc(const wchar_t* lpFontName)
{
	return CallRemoteProc(TargetProcessInfo.hProcess, lpRemoteAddFontProc, (void*)lpFontName, (std::wcslen(lpFontName) + 1) * sizeof(wchar_t));
}

bool RemoteRemoveFontProc(const wchar_t* lpFontName)
{
	return CallRemoteProc(TargetProcessInfo.hProcess, lpRemoteRemoveFontProc, (void*)lpFontName, (std::wcslen(lpFontName) + 1) * sizeof(wchar_t));
}

bool NullAddFontProc(const wchar_t* lpFontName)
{
	return true;
}

bool NullRemoveFontProc(const wchar_t* lpFontName)
{
	return true;
}

FontResource::FontResource(const std::wstring& FontName) : strFontPath_(FontName)
{
}

FontResource::FontResource(const wchar_t* FontName) : strFontPath_(FontName)
{
}

FontResource::~FontResource()
{
	if (bIsLoaded_)
	{
		RemoveFontProc_(strFontPath_.c_str());
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
		if (AddFontProc_(strFontPath_.c_str()))
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
		if (RemoveFontProc_(strFontPath_.c_str()))
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

const std::wstring & FontResource::GetFontPath()
{
	return strFontPath_;
}

bool FontResource::IsLoaded()
{
	return bIsLoaded_;
}