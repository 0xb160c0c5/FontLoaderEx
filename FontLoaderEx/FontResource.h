#pragma once

#include <Windows.h>
#include <string>

#ifdef _DEBUG
	#define ADDFONT_WAIT_MILLISEC 300
	#define REMOVEFONT_WAIT_MILLISEC 300
#endif // _DEBUG

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
	const std::wstring& GetFontName();
	bool IsLoaded();
	~FontResource();
private:
	static pfnAddFontProc AddFontProc_;
	static pfnRemoveFontProc RemoveFontProc_;
	std::wstring strFontName_;
	bool bIsLoaded_ = FALSE;
};