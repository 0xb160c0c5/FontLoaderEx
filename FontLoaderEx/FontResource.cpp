﻿#include <Windows.h>
#include <windowsx.h>
#include <cwchar>
#include <cassert>
#include "FontResource.h"
#include "Globals.h"

FontResource::AddFontProc FontResource::AddFontProc_{};
FontResource::RemoveFontProc FontResource::RemoveFontProc_{};

ADDFONT ProxyAddFontResult{};
REMOVEFONT ProxyRemoveFontResult{};

#ifdef _DEBUG
bool GlobalAddFontProc(const wchar_t* lpszFontName)
{
	Sleep(ADDFONT_WAIT_MILLISEC);

	return true;
}

bool GlobalRemoveFontProc(const wchar_t* lpszFontName)
{
	Sleep(REMOVEFONT_WAIT_MILLISEC);

	return true;
}
#else
bool GlobalAddFontProc(const wchar_t* lpszFontName)
{
	return AddFontResourceEx(lpszFontName, 0, NULL);
}

bool GlobalRemoveFontProc(const wchar_t* lpszFontName)
{
	return RemoveFontResourceEx(lpszFontName, 0, NULL);
}
#endif // _DEBUG

bool RemoteAddFontProc(const wchar_t* lpszFontName)
{
	DWORD dwRemoteThreadExitCode{};
	if (CallRemoteProc(TargetProcessInfo.hProcess, lpRemoteAddFontProcAddr, const_cast<wchar_t*>(lpszFontName), (std::wcslen(lpszFontName) + 1) * sizeof(wchar_t), INFINITE, &dwRemoteThreadExitCode))
	{
		return static_cast<bool>(dwRemoteThreadExitCode);
	}
	else
	{
		return false;
	}
}

bool RemoteRemoveFontProc(const wchar_t* lpszFontName)
{
	DWORD dwRemoteThreadExitCode{};
	if (CallRemoteProc(TargetProcessInfo.hProcess, lpRemoteRemoveFontProcAddr, const_cast<wchar_t*>(lpszFontName), (std::wcslen(lpszFontName) + 1) * sizeof(wchar_t), INFINITE, &dwRemoteThreadExitCode))
	{
		return static_cast<bool>(dwRemoteThreadExitCode);
	}
	else
	{
		return false;
	}
}

bool ProxyAddFontProc(const wchar_t* lpszFontName)
{
	bool bRet{};

	COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::ADDFONT), static_cast<DWORD>((std::wcslen(lpszFontName) + 1) * sizeof(wchar_t)), const_cast<wchar_t*>(lpszFontName) };
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
		break;
	default:
		{
			assert(0 && "invalid ProxyAddFontResult");
		}
		break;
	}

	return bRet;
}

bool ProxyRemoveFontProc(const wchar_t* lpszFontName)
{
	bool bRet{};

	COPYDATASTRUCT cds{ static_cast<ULONG_PTR>(COPYDATA::REMOVEFONT), static_cast<DWORD>((std::wcslen(lpszFontName) + 1) * sizeof(wchar_t)), const_cast<wchar_t*>(lpszFontName) };
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
		break;
	default:
		{
			assert(0 && "invalid ProxyRemoveFontResult");
		}
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

void FontResource::RegisterAddRemoveFontProc(AddFontProc AddFontProc, RemoveFontProc RemoveFontProc)
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

const std::wstring & FontResource::GetFontName() const
{
	return strFontName_;
}

bool FontResource::IsLoaded() const
{
	return bIsLoaded_;
}