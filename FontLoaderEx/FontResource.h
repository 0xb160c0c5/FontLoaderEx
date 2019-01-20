#pragma once

#include <string>
#include <Windows.h>

class FontResource
{
public:
	FontResource(const std::wstring FontPath);
	FontResource(const wchar_t* FontPath);
	bool Load();
	bool Unload();
	const std::wstring& GetFontPath();
	bool IsLoaded();
	~FontResource();

private:
	std::wstring strFontPath_;
	bool bIsLoaded_ = FALSE;
};

