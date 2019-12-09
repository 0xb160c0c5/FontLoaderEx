// Splitter user control
#pragma once

#if !defined(UNICODE) || !defined(_UNICODE)
#error Unicode character set required
#endif

#include <windows.h>

// Splitter class name
constexpr WCHAR UC_SPLITTER[]{ L"UserControl_Splitter" };

// ===================Splitter styles===================
constexpr DWORD SPS_HORZ{ 0b1u };	// Horizontal splitter;
constexpr DWORD SPS_VERT{ 0b10u };	// Vertical splitter;
constexpr DWORD SPS_PARENTWIDTH{ 0b100u };	// Splitter is as wide as parent window's client, ignoring the x and cx parameter in CreateWindow()
constexpr DWORD SPS_PARENTHEIGHT{ 0b1000u };	// Splitter is as tall as parent window's client, ignoring the y and cy parameter in CreateWindow()
constexpr DWORD SPS_AUTODRAG{ 0b10000u };	// Automatically move the splitter when holding and moving left mouse button
constexpr DWORD SPS_NOCAPTURE{ 0b100000u };	// Do not capture the cursor while dragging
constexpr DWORD SPS_NONOTIFY{ 0b1000000u };	// Do not send WM_NOTIFY to parent window
/*
	Remarks:

	One of SPS_VERT and SPS_HORZ must be specified and only one of them can be specified at a time.
	SPS_PARENTWIDTH and SPS_PARENTHEIGHT can't be both specified.
	SPS_VERT can't coexist with SPS_PARENTWIDTH, so does SPS_HORZ and SPS_PARENTHEIGHT.
*/

// ===================Splitter messages===================
enum SPLITTERMESSAGE : UINT { SPM_ROTATE = WM_USER + 1, SPM_SETRANGE, SPM_GETRANGE, SPM_SETMARGIN, SPM_GETMARGIN, SPM_SETLINKEDCTL, SPM_GETLINKEDCTL, SPM_ADDLINKEDCTL, SPM_REMOVELINKEDCTL };
enum SETLINKEDCONTROL : WORD { SLC_TOP = 1, SLC_BOTTOM, SLC_LEFT, SLC_RIGHT };
/*
	SPM_ROTATE

	Rotate the splitter by 90 degree.

	wParam, lParam:
		Not used.

	Return value:
		Not used.

	Remarks:

	After rotation, the splitter is redrawed, but the size and position remain unchanged.
	If the splitter has SPS_PARENTWIDTH style, it will be cleared and SPS_PARENTHEIGHT will be set, and vice versa.
	Other styles are retained.
*/

/*
	SPM_SETRANGE

	Set the movable area of splitter in the parent window's client area.

	wParam:
		Low word is the left/top border of the movable area of splitter, high word is the right/bottom border of the movable area of splitter.
		Set wParam to 0 will clear the movable range.

	lParam:
		Not used.

	Return value:
		Non-zero for success and zero for failure.

	Remarks:

	The movable range must be larger than the height/width of splitter or this message will return zero and has no effect.
*/

/*
	SPM_GETRANGE

	Get the movable area of splitter in the parent window's client area.

	wParam, lParam:
		Not used.

	Return value:
		Low word is the left/top border of the movable area, high word is the right/bottom border of the movable area.
		If return value is 0, the movable range is not set.
*/

/*
	SPM_SETMARGIN

	Set the margin of the line to the left and right/top and bottom of the splitter.

	wParam:
		The margin width/height.

	lParam:
		Not used.

	Return value:
		Non-zero for success and zero for failure.

	Remarks:
		The margin value must be lesser than half of the height/width of the splitter or this message will return zero and has no effect.
		If not set, the default margin is 0.
*/

/*
	SPM_GETMARGIN

	Get the margin of the line to the left and right/top and bottom of the splitter.

	wParam, lParam:
		Not used.

	Return value:
		the margin.

	Remarks:
		If return value is 0, the margin is not set;
*/

/*
	SPM_SETLINKEDCTL

	Set the linked controls.

	wParam:
		Low word is the number of controls that is about to links to the splitter.
		High word is a flag that determines which edge of the splitter the controls link to.
			SLC_TOP		The top edge of the splitter is linked to the bottom edge of the controls
			SLC_BOTTOM	The bottom edge of the splitter is linked to the top edge of the controls
			SLC_LEFT	The left edge of the splitter is linked to the right edge of the controls
			SLC_RIGHT	The right edge of the splitter is linked to the left edge of the controls

	lParam:
		A pointer to an array that contains the HWND controls.

	Return value:
		The number of controls that linked to the splitter, or 0 if failed.

	Remarks:
		When dragging the splitter, the edges that linked controls linked to the splitter are moved with the splitter.
		The linked controls must be the sibling windows of the splitter.
		Duplicated HWND is only added once.
		SLC_TOP/SLC_BOTTOM is not compatible with a splitter that has SPS_VERT style, so does SLC_LEFT/SLC_RIGHT with a splitter that has SPS_HORZ style.
*/

/*
	SPM_GETLINKEDCTL

	Get the linked controls.

	wParam:
		Low word is not used.
		High word is a flag that determines which edge of the splitter the controls link to.
			SLC_TOP		The top edge of the splitter is linked to the bottom edge of the controls
			SLC_BOTTOM	The bottom edge of the splitter is linked to the top edge of the controls
			SLC_LEFT	The left edge of the splitter is linked to the right edge of the controls
			SLC_RIGHT	The right edge of the splitter is linked to the left edge of the controls

	lParam:
		A pointer to an array that receives the HWND to controls.
		If lParam is NULL, the HWND to linked controls are not returned.

	Return value:
		The number of controls that linked to the splitter depending on SLC_*.
*/

/*
	SPM_ADDLINKEDCTL

	Add linked controls to the splitter.

	wParam:
		Low word is the number of controls that is about to add to the linked controls of the splitter.
		High word is a flag that determines which edge of the splitter the controls link to.
			SLC_TOP		The top edge of the splitter is linked to the bottom edge of the controls
			SLC_BOTTOM	The bottom edge of the splitter is linked to the top edge of the controls
			SLC_LEFT	The left edge of the splitter is linked to the right edge of the controls
			SLC_RIGHT	The right edge of the splitter is linked to the left edge of the controls

	lParam:
		A pointer to an array that contains the HWND to controls that are about to add.

	Return value:
		The number of controls added that linked to the splitter, or 0 if failed.

	Remarks:
		When dragging the splitter, The edges that linked controls linked to the splitter are moved with the splitter.
		The linked controls must be the sibling windows of the splitter.
		Duplicated HWND is only added once.
		SLC_TOP/SLC_BOTTOM is not compatible with a splitter that has SPS_VERT style, so does SLC_LEFT/SLC_RIGHT with a splitter that has SPS_HORZ style.
*/

/*
	SPM_REMOVELINKEDCTL

	Remove linked controls from the splitter.

	wParam:
		Low word is the number of controls that is about to remove from the linked controls of the splitter.
		High word is a flag that determines which edge of the splitter the controls link to.
			SLC_TOP		The top edge of the splitter is linked to the bottom edge of the controls
			SLC_BOTTOM	The bottom edge of the splitter is linked to the top edge of the controls
			SLC_LEFT	The left edge of the splitter is linked to the right edge of the controls
			SLC_RIGHT	The right edge of the splitter is linked to the left edge of the controls

	lParam:
		A pointer to an array that receives the HWND to controls that are about to remove.

	Return value:
		The number of controls removed that linked to the splitter depending on SLC_*.
*/

/*
	The splitter also accepts WM_MOVE and WM_SIZE.

	Sending the splitter an WM_MOVE will move it, and send an WM_SIZE will resize it,
	ignoring the x/y in WM_MOVE and cx/cy in WM_SIZE depending on SPS_PARENTWIDTH or SPS_PARENTHEIGHT.
*/

// ===================Splitter notifications===================
enum SPLITTERNOTIFICATION : UINT { SPN_DRAGBEGIN = 1, SPN_DRAGGING, SPN_DRAGEND };
/*
	DRAGBEGIN is sent when the cursor is in the client area and left mouse button is pressed.

	DRAGGING is sent when left mouse button is being hold and cursor is moving.
	If the splitter has SPS_NOCAPTURE style, it is only sent when the cursor is in the client area of the splitter.
	If the splitter has SPS_AUTODRAG style, this notification is sent after the splitter is moved.

	DARGEND is sent when left mouse button is released.
	If SPS_NOCAPTURE style is specified, it is only sent when the cursor is in the client area of the splitter.

	These notifications are sent via WM_NOTIFY, whose lParam points to a NMSPLITTER structure.
	If the movable range of splitter is set and the cursor is moved outside the movable range when holding left mouse button,
	the NMSPLITTER::ptCursorOffset will be adjusted to prevent further moving the splitter.
*/

/*
	The splitter also sends WM_CTLCORLORSTATIC to parent window.
	User can use this message to change the background color and line style of the splitter.
	The caller must take responsibility to manage the life cycle of GDI objects.
*/

// ===================Splitter structure===================
typedef struct tagNMSPLITTER
{
	NMHDR hdr;
	POINT ptCursor;	// Cursor position relatives to parent window's client area
	POINT ptCursorOffset;	// Cursor position relatives to splitter window when holding down the left button
} NMSPLITTER, *PNMSPLITTER, *LPNMSPLITTER;

// ===================Splitter function===================
ATOM InitSplitter();	// Initialize the splitter class
/*
	Return value:
		The ATOM of the splitter class if successful, or NULL if failed.
*/