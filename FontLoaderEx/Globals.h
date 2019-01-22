#pragma once

#include <windows.h>
#include <list>
#include "FontResource.h"

extern CRITICAL_SECTION CriticalSection;

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
extern HWND hWndListViewFontList;
extern HWND hWndEditMessage;

extern const UINT UM_WORKINGTHREADTERMINATED;
extern const UINT UM_CLOSEWORKINGTHREADTERMINATED;

extern bool DragDropHasFonts;

extern DWORD DragDropWorkingThreadProc(void* lpParameter);
extern DWORD CloseWorkingThreadProc(void* lpParameter);
extern DWORD ButtonCloseWorkingThreadProc(void* lpParameter);
extern DWORD ButtonCloseAllWorkingThreadProc(void* lpParameter);
extern DWORD ButtonLoadWorkingThreadProc(void* lpParameter);
extern DWORD ButtonLoadAllWorkingThreadProc(void* lpParameter);
extern DWORD ButtonUnloadWorkingThreadProc(void* lpParameter);
extern DWORD ButtonUnloadAllWorkingThreadProc(void* lpParameter);