#pragma once

#include <windows.h>
#include <list>
#include "FontResource.h"

extern std::list<FontResource> FontList;

extern HWND hWndMainWindow;
extern HWND hWndButtonOpen;
extern HWND hWndButtonClose;
extern HWND hWndButtonCloseAll;
extern HWND hWndButtonLoad;
extern HWND hWndButtonLoadAll;
extern HWND hWndButtonUnload;
extern HWND hWndButtonUnloadAll;
extern HWND hWndButtonBroadcastWM_FONTCHANGE;
extern HWND hWndButtonSelectProcess;
extern HWND hWndListViewFontList;
extern HWND hWndEditMessage;

const UINT UM_WORKINGTHREADTERMINATED{ WM_USER + 0x100 };
const UINT UM_CLOSEWORKINGTHREADTERMINATED{ WM_USER + 0x101 };
const UINT UM_TERMINATEWATCHPROCESS{ WM_USER + 0x102 };
const UINT UM_WATCHPROCESSTERMINATED{ WM_USER + 0x103 };

extern bool DragDropHasFonts;

extern void DragDropWorkingThreadProc(void* lpParameter);
extern void CloseWorkingThreadProc(void* lpParameter);
extern void ButtonCloseWorkingThreadProc(void* lpParameter);
extern void ButtonCloseAllWorkingThreadProc(void* lpParameter);
extern void ButtonLoadWorkingThreadProc(void* lpParameter);
extern void ButtonLoadAllWorkingThreadProc(void* lpParameter);
extern void ButtonUnloadWorkingThreadProc(void* lpParameter);
extern void ButtonUnloadAllWorkingThreadProc(void* lpParameter);
extern void TargetProcessWatchThreadProc(void* lpParameter);

extern void EnableAllButtons();
extern void DisableAllButtons();

struct ProcessInfo
{
	HANDLE hProcess;
	std::wstring ProcessName;
	std::uint32_t ProcessID;
};

extern ProcessInfo TargetProcessInfo;
extern void* lpRemoteAddFontProc;
extern void* lpRemoteRemoveFontProc;
extern HANDLE hTargetProcessWatchThread;

extern bool DefaultAddFontProc(const wchar_t* lpFontName);
extern bool DefaultRemoveFontProc(const wchar_t* lpFontName);
extern bool RemoteAddFontProc(const wchar_t* lpFontName);
extern bool RemoteRemoveFontProc(const wchar_t* lpFontName);
extern bool NullAddFontProc(const wchar_t* lpFontName);
extern bool NullRemoveFontProc(const wchar_t* lpFontName);