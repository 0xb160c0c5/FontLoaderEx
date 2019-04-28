#include <Windows.h>
#include <cwchar>

#define EXPORT_FUNCTION comment(linker, "/EXPORT:" __FUNCTION__ "=" __FUNCDNAME__)	// Export function and avoid name mangling with __stdcall

#pragma warning(disable: 4996)

DWORD WINAPI AddFont(_In_ LPVOID lpParameter)
{
#pragma EXPORT_FUNCTION
#ifdef _DEBUG
	WCHAR szMessage[512]{ L"AddFont() called!\r\nFont Name: " };
	std::wcscpy(&szMessage[30], (wchar_t*)lpParameter);
	MessageBox(NULL, szMessage, L"FontLoaderEx", NULL);

	return 1;
#else
	return AddFontResourceEx((LPCWSTR)lpParameter, FR_PRIVATE, NULL);
#endif // _DEBUG
}

DWORD WINAPI RemoveFont(_In_ LPVOID lpParameter)
{
#pragma EXPORT_FUNCTION
#ifdef _DEBUG
	WCHAR szMessage[512]{ L"RemoveFont() called!\r\nFont Name: " };
	std::wcscpy(&szMessage[33], (wchar_t*)lpParameter);
	MessageBox(NULL, szMessage, L"FontLoaderEx", NULL);

	return 1;
#else
	return RemoveFontResourceEx((LPCWSTR)lpParameter, FR_PRIVATE, NULL);
#endif // _DEBUG
}