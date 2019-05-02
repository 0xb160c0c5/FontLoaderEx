// Splitter user control

#pragma once

#include <windows.h>

#if !defined(UNICODE) || !defined(_UNICODE)
#error Unicode character set required
#endif

// Splitter class name
constexpr WCHAR UC_SPLITTER[]{ L"UserControl_Splitter" };

// Splitter styles
constexpr DWORD SPS_HORZ = 0b1u;	// Horizontal splitter;
constexpr DWORD SPS_VERT = 0b10u;	// Vertical splitter;
constexpr DWORD SPS_NOCAPTURE = 0b100u;	// Do not capture the mouse while dragging;

typedef struct tagSPLITTER	// lParam in WM_NOTIFY
{
	NMHDR hdr;
	POINT ptCursor;	// Cursor point relative to parent window client
	LONG CursorOffset;	// Cursor offset relative to splitter top/left
} SPLITTERSTRUCT, *LPSPLITTERSTRUCT;

enum class SPLITTERNOTIFICATION : UINT { DRAGBEGIN, DRAGGING, DRAGEND };

ATOM InitSplitter();