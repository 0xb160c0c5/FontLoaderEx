#include <Windows.h>

#define EXPORT_FUNCTION comment(linker, "/EXPORT:" __FUNCTION__ "=" __FUNCDNAME__)	//Export function and avoid name mangling with __stdcall

DWORD WINAPI AddFont(_In_ LPVOID lpParameter)
{
#pragma EXPORT_FUNCTION
#ifdef _DEBUG
	MessageBox(NULL, L"AddFont() called!", L"Test", NULL);
	return 1;
#else
	return AddFontResourceEx((LPCWSTR)lpParameter, FR_PRIVATE, NULL);
#endif // _DEBUG
}

DWORD WINAPI RemoveFont(_In_ LPVOID lpParameter)
{
#pragma EXPORT_FUNCTION
#ifdef _DEBUG
	MessageBox(NULL, L"RemoveFont() called!", L"Test", NULL);
	return 1;
#else
	return RemoveFontResourceEx((LPCWSTR)lpParameter, FR_PRIVATE, NULL);
#endif // _DEBUG
}