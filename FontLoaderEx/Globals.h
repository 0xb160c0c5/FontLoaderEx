#pragma once

#include <windows.h>
#include <list>
#include "FontResource.h"

extern std::list<FontResource> FontList;

extern HWND hWndMainWindow;
extern HWND hWndButtonOpen;
extern HWND hWndButtonClose;
extern HWND hWndButtonLoad;
extern HWND hWndButtonUnload;
extern HWND hWndButtonBroadcastWM_FONTCHANGE;
extern HWND hWndButtonSelectProcess;
extern HWND hWndListViewFontList;
extern HWND hWndEditMessage;

enum class USERMESSAGE : UINT { WORKINGTHREADTERMINATED = WM_USER + 0x100, CLOSEWORKINGTHREADTERMINATED, TERMINATEWATCHTHREAD, WATCHTHREADTERMINATED, TERMINATEMESSAGETHREAD };

extern bool bDragDropHasFonts;

extern void DragDropWorkingThreadProc(void* lpParameter);
extern void CloseWorkingThreadProc(void* lpParameter);
extern void ButtonCloseWorkingThreadProc(void* lpParameter);
extern void ButtonLoadWorkingThreadProc(void* lpParameter);
extern void ButtonUnloadWorkingThreadProc(void* lpParameter);
extern void TargetProcessWatchThreadProc(void* lpParameter);
extern void ProxyAndTargetProcessWatchThreadProc(void* lpParameter);
extern void MessageThreadProc(void* lpParameter);

extern HANDLE hWatchThread;
extern HANDLE hMessageThread;

struct ProcessInfo
{
	HANDLE hProcess;
	std::wstring ProcessName;
	std::uint32_t ProcessID;
};

extern ProcessInfo TargetProcessInfo;
extern void* pfnRemoteAddFontProc;
extern void* pfnRemoteRemoveFontProc;

extern PROCESS_INFORMATION piProxyProcess;
extern HWND hWndProxy;
extern HWND hWndMessage;

extern HANDLE hEventParentProcessRunning;
extern HANDLE hEventMessageThreadReady;
extern HANDLE hEventTerminateWatchThread;
extern HANDLE hEventProxyProcessReady;
extern HANDLE hEventProxyProcessHWNDRevieved;
extern HANDLE hEventProxyDllInjectionFinished;
extern HANDLE hEventProxyDllPullFinished;
extern HANDLE hEventProxyAddFontFinished;
extern HANDLE hEventProxyRemoveFontFinished;

extern HANDLE hCurrentProcessDuplicated;
extern HANDLE hTargetProcessDuplicated;

enum class COPYDATA : ULONG_PTR { PROXYPROCESSHWNDSENT, INJECTDLL, DLLINJECTIONFINISHED, PULLDLL, DLLPULLFINISHED, ADDFONT, ADDFONTFINISHED, REMOVEFONT, REMOVEFONTFINISHED, TERMINATE };

extern int ProxyDllInjectionResult;
extern int ProxyDllPullResult;
extern bool ProxyAddFontResult;
extern bool ProxyRemoveFontResult;
enum class PROXYDLLINJECTION : int { SUCCESSFUL, FAILED, FAILEDTOENUMERATEMODULES, GDI32NOTLOADED, MODULENOTFOUND };
enum class PROXYDLLPULL : int { SUCCESSFUL, FAILED };

extern bool DefaultAddFontProc(const wchar_t* lpFontName);
extern bool DefaultRemoveFontProc(const wchar_t* lpFontName);
extern bool RemoteAddFontProc(const wchar_t* lpFontName);
extern bool RemoteRemoveFontProc(const wchar_t* lpFontName);
extern bool ProxyAddFontProc(const wchar_t* lpFontName);
extern bool ProxyRemoveFontProc(const wchar_t* lpFontName);
extern bool NullAddFontProc(const wchar_t* lpFontName);
extern bool NullRemoveFontProc(const wchar_t* lpFontName);

extern void EnableAllButtons();
extern void DisableAllButtons();