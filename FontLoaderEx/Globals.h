#pragma once

#include <windows.h>
#include <list>
#include "FontResource.h"

extern std::list<FontResource> FontList;

extern HWND hWndMain;
extern HMENU hMenuContextListViewFontList;

enum class ID : WORD { ButtonOpen = 20, ButtonClose, ButtonLoad, ButtonUnload, ButtonBroadcastWM_FONTCHANGE, StaticTimeout, EditTimeout, ButtonSelectProcess, ListViewFontList, Splitter, EditMessage };

enum class USERMESSAGE : UINT {DRAGDROPWORKINGTHREADTERMINATED = WM_USER + 0x100, BUTTONCLOSEWORKINGTHREADTERMINATED, BUTTONLOADWORKINGTHREADTERMINATED, BUTTONUNLOADWORKINGTHREADTERMINATED, CLOSEWORKINGTHREADTERMINATED,TERMINATEWATCHTHREAD, WATCHTHREADTERMINATED, TERMINATEMESSAGETHREAD };

extern bool bDragDropHasFonts;

extern void DragDropWorkingThreadProc(void* lpParameter);
extern void CloseWorkingThreadProc(void* lpParameter);
extern void ButtonCloseWorkingThreadProc(void* lpParameter);
extern void ButtonLoadWorkingThreadProc(void* lpParameter);
extern void ButtonUnloadWorkingThreadProc(void* lpParameter);
extern unsigned int __stdcall TargetProcessWatchThreadProc(void* lpParameter);
extern unsigned int __stdcall ProxyAndTargetProcessWatchThreadProc(void* lpParameter);
extern unsigned int __stdcall MessageThreadProc(void* lpParameter);

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
extern HANDLE hEventMessageThreadNotReady;
extern HANDLE hEventMessageThreadReady;
extern HANDLE hEventTerminateWatchThread;
extern HANDLE hEventProxyProcessReady;
extern HANDLE hEventProxyProcessDebugPrivilegeEnablingFinished;
extern HANDLE hEventProxyProcessHWNDRevieved;
extern HANDLE hEventProxyDllInjectionFinished;
extern HANDLE hEventProxyDllPullingFinished;
extern HANDLE hEventProxyAddFontFinished;
extern HANDLE hEventProxyRemoveFontFinished;

enum class COPYDATA : ULONG_PTR { PROXYPROCESSHWNDSENT, PROXYPROCESSDEBUGPRIVILEGEENABLINGFINISHED, INJECTDLL, DLLINJECTIONFINISHED, PULLDLL, DLLPULLINGFINISHED, ADDFONT, ADDFONTFINISHED, REMOVEFONT, REMOVEFONTFINISHED, TERMINATE };
enum class PROXYPROCESSDEBUGPRIVILEGEENABLING { SUCCESSFUL, FAILED };
enum class PROXYDLLINJECTION { SUCCESSFUL, FAILED, FAILEDTOENUMERATEMODULES, GDI32NOTLOADED, MODULENOTFOUND };
enum class PROXYDLLPULL { SUCCESSFUL, FAILED };
enum class ADDFONT { SUCCESSFUL, FAILED };
enum class REMOVEFONT { SUCCESSFUL, FAILED };

extern PROXYPROCESSDEBUGPRIVILEGEENABLING ProxyDebugPrivilegeEnablingResult;
extern PROXYDLLINJECTION ProxyDllInjectionResult;
extern PROXYDLLPULL ProxyDllPullingResult;
extern ADDFONT ProxyAddFontResult;
extern REMOVEFONT ProxyRemoveFontResult;

extern bool DefaultAddFontProc(const wchar_t* lpFontName);
extern bool DefaultRemoveFontProc(const wchar_t* lpFontName);
extern bool RemoteAddFontProc(const wchar_t* lpFontName);
extern bool RemoteRemoveFontProc(const wchar_t* lpFontName);
extern bool ProxyAddFontProc(const wchar_t* lpFontName);
extern bool ProxyRemoveFontProc(const wchar_t* lpFontName);
extern bool NullAddFontProc(const wchar_t* lpFontName);
extern bool NullRemoveFontProc(const wchar_t* lpFontName);