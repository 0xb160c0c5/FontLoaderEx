#pragma once

#include <string>
#include <Windows.h>

class FontResource
{
public:
	typedef bool (*pfnAddFontProc)(const wchar_t* lpFontName);
	typedef bool (*pfnRemoveFontProc)(const wchar_t* lpFontName);
	static void RegisterAddRemoveFontProc(pfnAddFontProc AddFontProc, pfnRemoveFontProc RemoveFontProc);
	FontResource(const std::wstring& FontName);
	FontResource(const wchar_t* FontName);
	bool Load();
	bool Unload();
	const std::wstring& GetFontPath();
	bool IsLoaded();
	~FontResource();
private:
	static pfnAddFontProc AddFontProc_;
	static pfnRemoveFontProc RemoveFontProc_;
	std::wstring strFontPath_;
	bool bIsLoaded_ = FALSE;
};
