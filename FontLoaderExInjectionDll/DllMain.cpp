#include <Windows.h>

#define EXPORT_FUNCTION comment(linker, "/EXPORT:" __FUNCTION__ "=" __FUNCDNAME__)	// Export function and avoid name mangling with __stdcall

DWORD WINAPI AddFont(_In_ LPVOID lpParameter)
{
#pragma EXPORT_FUNCTION
#ifdef _DEBUG
	WCHAR szMessage[512]{ L"AddFont() called!\r\nFontName: " };
	WCHAR* lpszMessage = &szMessage[29];
	WCHAR* lpszFontName = (WCHAR*)lpParameter;
	while (*lpszFontName)
	{
		*lpszMessage++ = *lpszFontName++;
	};
	MessageBox(NULL, szMessage, L"Test", NULL);
	return 1;
#else
	return AddFontResourceEx((LPCWSTR)lpParameter, FR_PRIVATE, NULL);
#endif // _DEBUG
}

DWORD WINAPI RemoveFont(_In_ LPVOID lpParameter)
{
#pragma EXPORT_FUNCTION
#ifdef _DEBUG
	WCHAR szMessage[512]{ L"RemoveFont() called!\r\nFontName: " };
	WCHAR* lpszMessage = &szMessage[32];
	WCHAR* lpszFontName = (WCHAR*)lpParameter;
	while (*lpszFontName)
	{
		*lpszMessage++ = *lpszFontName++;
	};
	MessageBox(NULL, szMessage, L"Test", NULL);
	return 1;
#else
	return RemoveFontResourceEx((LPCWSTR)lpParameter, FR_PRIVATE, NULL);
#endif // _DEBUG
}