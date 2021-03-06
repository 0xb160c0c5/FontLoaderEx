﻿#if !defined(UNICODE) || !defined(_UNICODE)
#error Unicode character set required
#endif // UNICODE && _UNICODE

#include <Windows.h>

#define EXPORT_FUNCTION comment(linker, "/EXPORT:" __FUNCTION__ "=" __FUNCDNAME__)	// Export function and avoid name mangling with __stdcall

DWORD WINAPI AddFont(_In_ LPVOID lpParameter)
{
#pragma EXPORT_FUNCTION
#ifdef _DEBUG
	WCHAR szMessage[512]{ L"AddFont() called!\r\nFont Name: " };
	wcscat_s(szMessage, static_cast<wchar_t*>(lpParameter));
	MessageBox(NULL, szMessage, L"FontLoaderEx", NULL);

	return 1;
#else
	return AddFontResourceEx(static_cast<LPWSTR>(lpParameter), FR_PRIVATE, NULL);
#endif // _DEBUG
}

DWORD WINAPI RemoveFont(_In_ LPVOID lpParameter)
{
#pragma EXPORT_FUNCTION
#ifdef _DEBUG
	WCHAR szMessage[512]{ L"RemoveFont() called!\r\nFont Name: " };
	wcscat_s(szMessage, static_cast<wchar_t*>(lpParameter));
	MessageBox(NULL, szMessage, L"FontLoaderEx", NULL);

	return 1;
#else
	return RemoveFontResourceEx(static_cast<LPWSTR>(lpParameter), FR_PRIVATE, NULL);
#endif // _DEBUG
}