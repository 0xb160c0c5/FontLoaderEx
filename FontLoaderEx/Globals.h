#pragma once

#include <windows.h>
#include <list>
#include <string>
#include "FontResource.h"

extern std::list<FontResource> FontList;

extern HWND hWndMain;

enum class USERMESSAGE : UINT { FONTLISTCHANGED = WM_USER + 0x100, DRAGDROPWORKERTHREADTERMINATED, BUTTONCLOSEWORKERTHREADTERMINATED, BUTTONLOADWORKERTHREADTERMINATED, BUTTONUNLOADWORKERTHREADTERMINATED, CLOSEWORKERTHREADTERMINATED, WATCHTHREADTERMINATING, WATCHTHREADTERMINATED, CHILDWINDOWPOSCHANGED, TRAYNOTIFYICON };

enum class FONTLISTCHANGED : UINT { OPENED = 1, OPENED_LOADED, OPENED_NOTLOADED, LOADED, NOTLOADED, UNLOADED, NOTUNLOADED, UNLOADED_CLOSED, CLOSED, UNTOUCHED };

struct FONTLISTCHANGEDSTRUCT
{
	int iItem;
	LPCWSTR lpszFontName;
};

enum class TERMINATION : DWORD { DIRECT = 1 , SURROGATE, TARGET };

extern void DragDropWorkerThreadProc(void* lpParameter);
extern unsigned int __stdcall CloseWorkerThreadProc(void* lpParameter);
extern unsigned int __stdcall ButtonCloseWorkerThreadProc(void* lpParameter);
extern unsigned int __stdcall ButtonLoadWorkerThreadProc(void* lpParameter);
extern unsigned int __stdcall ButtonUnloadWorkerThreadProc(void* lpParameter);
extern unsigned int __stdcall TargetProcessWatchThreadProc(void* lpParameter);
extern unsigned int __stdcall SurrogateAndTargetProcessWatchThreadProc(void* lpParameter);
extern unsigned int __stdcall MessageThreadProc(void* lpParameter);

extern HANDLE hThreadCloseWorkerThreadProc;
extern HANDLE hThreadButtonCloseWorkerThreadProc;
extern HANDLE hThreadButtonLoadWorkerThreadProc;
extern HANDLE hThreadButtonUnloadWorkerThreadProc;
extern HANDLE hThreadWatch;
extern HANDLE hThreadMessage;

extern HANDLE hEventWorkerThreadReadyToTerminate;

struct ProcessInfo
{
	HANDLE hProcess;
	std::wstring strProcessName;
	DWORD dwProcessID;
};

extern ProcessInfo TargetProcessInfo;
extern void* lpRemoteAddFontProcAddr;
extern void* lpRemoteRemoveFontProcAddr;

extern ProcessInfo SurrogateProcessInfo;
extern HWND hWndSurrogate;
extern HWND hWndMessage;

extern HANDLE hProcessCurrentDuplicated;
extern HANDLE hProcessTargetDuplicated;

extern HANDLE hEventParentProcessRunning;
extern HANDLE hEventMessageThreadNotReady;
extern HANDLE hEventMessageThreadReady;
extern HANDLE hEventTerminateWatchThread;
extern HANDLE hEventSurrogateProcessReady;
extern HANDLE hEventSurrogateProcessDebugPrivilegeEnablingFinished;
extern HANDLE hEventSurrogateProcessHWNDRevieved;
extern HANDLE hEventSurrogateDllInjectionFinished;
extern HANDLE hEventSurrogateDllPullingFinished;
extern HANDLE hEventSurrogateAddFontFinished;
extern HANDLE hEventSurrogateRemoveFontFinished;

enum class COPYDATA : ULONG_PTR { SURROGATEPROCESSHWNDSENT = 1, SURROGATEPROCESSDEBUGPRIVILEGEENABLINGFINISHED, INJECTDLL, DLLINJECTIONFINISHED, PULLDLL, DLLPULLINGFINISHED, ADDFONT, ADDFONTFINISHED, REMOVEFONT, REMOVEFONTFINISHED, TERMINATE };
enum class SURROGATEPROCESSDEBUGPRIVILEGEENABLING : UINT { SUCCESSFUL = 1, FAILED };
enum class SURROGATEDLLINJECTION : UINT { SUCCESSFUL = 1, FAILED, FAILEDTOENUMERATEMODULES, GDI32NOTLOADED, MODULENOTFOUND };
enum class SURROGATEDLLPULL : UINT { SUCCESSFUL = 1, FAILED };
enum class ADDFONT : UINT { SUCCESSFUL = 1, FAILED };
enum class REMOVEFONT : UINT { SUCCESSFUL = 1, FAILED };

extern SURROGATEPROCESSDEBUGPRIVILEGEENABLING SurrogateDebugPrivilegeEnablingResult;
extern SURROGATEDLLINJECTION SurrogateDllInjectionResult;
extern SURROGATEDLLPULL SurrogateDllPullingResult;
extern ADDFONT SurrogateAddFontResult;
extern REMOVEFONT SurrogateRemoveFontResult;

extern bool GlobalAddFontProc(const wchar_t* lpszFontName);
extern bool GlobalRemoveFontProc(const wchar_t* lpszFontName);
extern bool RemoteAddFontProc(const wchar_t* lpszFontName);
extern bool RemoteRemoveFontProc(const wchar_t* lpszFontName);
extern bool SurrogateAddFontProc(const wchar_t* lpszFontName);
extern bool SurrogateRemoveFontProc(const wchar_t* lpszFontName);
extern bool NullAddFontProc(const wchar_t* lpszFontName);
extern bool NullRemoveFontProc(const wchar_t* lpszFontName);

extern bool CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, std::size_t cbParamSize, DWORD dwTimeout, LPDWORD lpdwRemoteThreadExitCode);