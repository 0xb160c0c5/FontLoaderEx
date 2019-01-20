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
		RemoveFontResource(strFontPath_.c_str());
	}
}

bool FontResource::Load()
{
#ifndef _DEBUG
	bool bRet;
	if (!bIsLoaded_)
	{
		if (AddFontResource(strFontPath_.c_str()))
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
		if (RemoveFontResource(strFontPath_.c_str()))
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