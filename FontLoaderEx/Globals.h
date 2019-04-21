#pragma once

#include <windows.h>
#include <list>
#include <string>
#include "FontResource.h"

extern std::list<FontResource> FontList;

extern HWND hWndMain;
extern HMENU hMenuContextListViewFontList;

enum class ID : WORD { ButtonOpen = 20, ButtonClose, ButtonLoad, ButtonUnload, ButtonBroadcastWM_FONTCHANGE, StaticTimeout, EditTimeout, ButtonSelectProcess, ListViewFontList, Splitter, EditMessage };

enum class USERMESSAGE : UINT { DRAGDROPWORKERTHREADTERMINATED = WM_USER + 0x100, BUTTONCLOSEWORKERTHREADTERMINATED, BUTTONLOADWORKERTHREADTERMINATED, BUTTONUNLOADWORKERTHREADTERMINATED, CLOSEWORKERTHREADTERMINATED, TERMINATEWATCHTHREAD, WATCHTHREADTERMINATED, TERMINATEMESSAGETHREAD, CHILDWINDOWPOSCHANGED };

extern void DragDropWorkerThreadProc(void* lpParameter);
extern void CloseWorkerThreadProc(void* lpParameter);
extern void ButtonCloseWorkerThreadProc(void* lpParameter);
extern void ButtonLoadWorkerThreadProc(void* lpParameter);
extern void ButtonUnloadWorkerThreadProc(void* lpParameter);
extern unsigned int __stdcall TargetProcessWatchThreadProc(void* lpParameter);
extern unsigned int __stdcall ProxyAndTargetProcessWatchThreadProc(void* lpParameter);
extern unsigned int __stdcall MessageThreadProc(void* lpParameter);

extern HANDLE hThreadWatch;
extern HANDLE hThreadMessage;

struct ProcessInfo
{
	HANDLE hProcess;
	std::wstring strProcessName;
	DWORD dwProcessID;
};

extern ProcessInfo TargetProcessInfo;
extern void* lpRemoteAddFontProcAddr;
extern void* lpRemoteRemoveFontProcAddr;

extern ProcessInfo ProxyProcessInfo;
extern HWND hWndProxy;
extern HWND hWndMessage;

extern HANDLE hProcessCurrentDuplicated;
extern HANDLE hProcessTargetDuplicated;

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

extern bool GlobalAddFontProc(const wchar_t* lpszFontName);
extern bool GlobalRemoveFontProc(const wchar_t* lpszFontName);
extern bool RemoteAddFontProc(const wchar_t* lpszFontName);
extern bool RemoteRemoveFontProc(const wchar_t* lpszFontName);
extern bool ProxyAddFontProc(const wchar_t* lpszFontName);
extern bool ProxyRemoveFontProc(const wchar_t* lpszFontName);
extern bool NullAddFontProc(const wchar_t* lpszFontName);
extern bool NullRemoveFontProc(const wchar_t* lpszFontName);

extern DWORD CallRemoteProc(HANDLE hProcess, void* lpRemoteProcAddr, void* lpParameter, std::size_t cbParamSize, DWORD dwTimeout);