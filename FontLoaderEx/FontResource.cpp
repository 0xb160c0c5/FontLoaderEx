#include "FontResource.h"

FontResource::FontResource(const std::wstring FontPath) : strFontPath_(FontPath)
{
}

FontResource::FontResource(const wchar_t * FontPath) : strFontPath_(FontPath)
{
}

FontResource::~FontResource()
{
	if (bIsLoaded_)
	{
		RemoveFontResourceEx(strFontPath_.c_str(), 0, NULL);
	}
}

bool FontResource::Load()
{
#ifndef _DEBUG
	bool bRet;
	if (!bIsLoaded_)
	{
		if (AddFontResourceEx(strFontPath_.c_str(), 0, NULL))
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
#else
	Sleep(500);
	bIsLoaded_ = true;
	return true;
#endif
}

bool FontResource::Unload()
{
#ifndef _DEBUG
	bool bRet;
	if (bIsLoaded_)
	{
		if (RemoveFontResourceEx(strFontPath_.c_str(), 0, NULL))
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
#else
	Sleep(500);
	bIsLoaded_ = false;
	return true;
#endif
}

const std::wstring & FontResource::GetFontPath()
{
	return strFontPath_;
}

bool FontResource::IsLoaded()
{
	return bIsLoaded_;
}