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
	typedef bool(*AddFontProc)(const wchar_t* lpFontName);
	typedef bool(*RemoveFontProc)(const wchar_t* lpFontName);
	static void RegisterAddRemoveFontProc(AddFontProc AddFontProc, RemoveFontProc RemoveFontProc);
	FontResource(const std::wstring& FontName);
	FontResource(const wchar_t* FontName);
	bool Load();
	bool Unload();
	const std::wstring& GetFontName();
	bool IsLoaded();
	~FontResource();
private:
	static AddFontProc AddFontProc_;
	static RemoveFontProc RemoveFontProc_;
	std::wstring strFontName_;
	bool bIsLoaded_ = FALSE;
};