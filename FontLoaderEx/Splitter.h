// Splitter user control

#pragma once

#include <windows.h>

#if !defined(UNICODE) || !defined(_UNICODE)
	#error Unicode character set required
#endif

const WCHAR UC_SPLITTER[]{ L"UserControl_Splitter" };

typedef struct tagSPLITTER
{
	NMHDR nmhdr;
	POINT ptCursor;
} SPLITTERSTRUCT, *LPSPLITTERSTRUCT;

enum class SPLITTERNOTIFICATION : UINT{ DRAGBEGIN, DRAGGING, DRAGEND };

ATOM InitSplitter();