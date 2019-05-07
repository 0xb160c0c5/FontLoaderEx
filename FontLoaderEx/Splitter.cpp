#include <windows.h>
#include <windowsx.h>
#include "Splitter.h"

LRESULT CALLBACK SplitterProc(HWND hWndSplitter, UINT Message, WPARAM wParam, LPARAM lParam);

ATOM InitSplitter()
{
	WNDCLASS wc{ 0, SplitterProc, 0, 0, static_cast<HINSTANCE>(GetModuleHandle(NULL)), NULL, NULL, GetSysColorBrush(COLOR_BTNFACE), NULL, UC_SPLITTER };

	return RegisterClass(&wc);
}

LRESULT CALLBACK SplitterProc(HWND hWndSplitter, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	static DWORD dwSplitterStyle;

	static LONG cyCursorOffset{};

	switch (Message)
	{
	case WM_CREATE:
		{
			dwSplitterStyle = static_cast<DWORD>(((LPCREATESTRUCT)lParam)->style);
			if ((dwSplitterStyle & SPS_HORZ) && (dwSplitterStyle & SPS_VERT))
			{
				ret = static_cast<LRESULT>(-1);
			}
		}
		break;
	case WM_PAINT:
		{
			// Draw a horizontal line in the middle
			PAINTSTRUCT ps{};
			HDC hDCSplitter{ BeginPaint(hWndSplitter, &ps) };

			RECT rcSplitterClient{};
			GetClientRect(hWndSplitter, &rcSplitterClient);
			HPEN hPenSplitter{ CreatePen(PS_SOLID, (rcSplitterClient.bottom - rcSplitterClient.top) / 5, static_cast<COLORREF>(GetSysColor(COLOR_GRAYTEXT))) };
			SelectPen(hDCSplitter, hPenSplitter);

			if (dwSplitterStyle & SPS_HORZ)
			{
				MoveToEx(hDCSplitter, rcSplitterClient.left + (rcSplitterClient.bottom - rcSplitterClient.top) / 2, (rcSplitterClient.bottom - rcSplitterClient.top) / 2, NULL);
				LineTo(hDCSplitter, (rcSplitterClient.right - rcSplitterClient.left) - (rcSplitterClient.bottom - rcSplitterClient.top) / 2, (rcSplitterClient.bottom - rcSplitterClient.top) / 2);
			}
			else if (dwSplitterStyle & SPS_VERT)
			{
				MoveToEx(hDCSplitter, rcSplitterClient.top + (rcSplitterClient.right - rcSplitterClient.left) / 2, (rcSplitterClient.right - rcSplitterClient.left) / 2, NULL);
				LineTo(hDCSplitter, (rcSplitterClient.bottom - rcSplitterClient.top) - (rcSplitterClient.right - rcSplitterClient.left) / 2, (rcSplitterClient.right - rcSplitterClient.left) / 2);
			}

			EndPaint(hWndSplitter, &ps);
		}
		break;
	case WM_LBUTTONDOWN:
		{
			// Capture cursor and send WM_NOTIFY with SPLITTERNOTIFICATION::DRAGBEGIN to parent window
			if (!(dwSplitterStyle & SPS_NOCAPTURE))
			{
				SetCapture(hWndSplitter);
			}

			RECT rcSplitter{}, rcSplitterClient{};
			GetWindowRect(hWndSplitter, &rcSplitter);
			GetClientRect(hWndSplitter, &rcSplitterClient);
			MapWindowRect(hWndSplitter, HWND_DESKTOP, &rcSplitterClient);
			POINT ptCursor{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
			if (dwSplitterStyle & SPS_HORZ)
			{
				cyCursorOffset = ptCursor.y + (rcSplitterClient.top - rcSplitter.top);
			}
			else if (dwSplitterStyle & SPS_VERT)
			{
				cyCursorOffset = ptCursor.x + (rcSplitterClient.left - rcSplitter.left);
			}
			ClientToScreen(hWndSplitter, &ptCursor);
			ScreenToClient(GetParent(hWndSplitter), &ptCursor);
			SPLITTERSTRUCT ss{ { hWndSplitter, static_cast<UINT_PTR>(GetDlgCtrlID(hWndSplitter)), static_cast<UINT>(SPLITTERNOTIFICATION::DRAGBEGIN) }, ptCursor, cyCursorOffset };
			SendMessage(GetParent(hWndSplitter), WM_NOTIFY, static_cast<WPARAM>(GetDlgCtrlID(hWndSplitter)), reinterpret_cast<LPARAM>(&ss));
		}
		break;
	case WM_MOUSEMOVE:
		{
			// If left button is being hold, send WM_NOTIFY with SPLITTERNOTIFICATION::DRAGGING to parent window
			if ((wParam == MK_LBUTTON))
			{
				POINT ptCursor{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
				ClientToScreen(hWndSplitter, &ptCursor);
				ScreenToClient(GetParent(hWndSplitter), &ptCursor);
				SPLITTERSTRUCT ss{ { hWndSplitter, static_cast<UINT_PTR>(GetDlgCtrlID(hWndSplitter)), (UINT)SPLITTERNOTIFICATION::DRAGGING }, ptCursor, cyCursorOffset };
				SendMessage(GetParent(hWndSplitter), WM_NOTIFY, static_cast<WPARAM>(GetDlgCtrlID(hWndSplitter)), reinterpret_cast<LPARAM>(&ss));
			}
		}
		break;
	case WM_LBUTTONUP:
		{
			// Release cursor and send WM_NOTIFY with SPLITTERNOTIFICATION::DRAGBEGIN to parent window
			if (!(dwSplitterStyle & SPS_NOCAPTURE))
			{
				ReleaseCapture();
			}

			POINT ptCursor{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
			ClientToScreen(hWndSplitter, &ptCursor);
			ScreenToClient(GetParent(hWndSplitter), &ptCursor);
			SPLITTERSTRUCT ss{ { hWndSplitter, static_cast<UINT_PTR>(GetDlgCtrlID(hWndSplitter)), (UINT)SPLITTERNOTIFICATION::DRAGEND }, ptCursor, cyCursorOffset };
			SendMessage(GetParent(hWndSplitter), WM_NOTIFY, static_cast<WPARAM>(GetDlgCtrlID(hWndSplitter)), reinterpret_cast<LPARAM>(&ss));
		}
		break;
	case WM_SETCURSOR:
		{
			if (reinterpret_cast<HWND>(wParam) == hWndSplitter)
			{
				if (dwSplitterStyle & SPS_HORZ)
				{
					SetCursor(LoadCursor(NULL, IDC_SIZENS));
				}
				else if (dwSplitterStyle & SPS_VERT)
				{
					SetCursor(LoadCursor(NULL, IDC_SIZEWE));
				}

				ret = static_cast<LRESULT>(TRUE);
			}
		}
		break;
	case WM_STYLECHANGED:
		{
			if (wParam == GWL_STYLE)
			{
				dwSplitterStyle = reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew;
			}
		}
		break;
	default:
		{
			ret = DefWindowProc(hWndSplitter, Message, wParam, lParam);
		}
		break;
	}

	return ret;
}