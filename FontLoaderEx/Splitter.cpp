#include <windows.h>
#include <windowsx.h>
#include "Splitter.h"

LRESULT CALLBACK SplitterProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);

ATOM InitSplitter()
{
	WNDCLASS wc{ CS_HREDRAW | CS_VREDRAW, SplitterProc, 0, 0, GetModuleHandle(NULL), NULL, LoadCursor(NULL, IDC_SIZENS), GetSysColorBrush(COLOR_WINDOW), NULL, UC_SPLITTER };

	return RegisterClass(&wc);
}

LRESULT CALLBACK SplitterProc(HWND hWndSplitter, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	switch (Msg)
	{
		//Draw a horizontal line in the middle
	case WM_PAINT:
		{
			PAINTSTRUCT ps{};
			HDC hDC{ BeginPaint(hWndSplitter, &ps) };

			SelectPen(hDC, GetStockPen(BLACK_PEN));
			RECT rectSplitterClient{};
			GetClientRect(hWndSplitter, &rectSplitterClient);
			MoveToEx(hDC, rectSplitterClient.left + 2, (rectSplitterClient.bottom - rectSplitterClient.top) / 2, NULL);
			LineTo(hDC, rectSplitterClient.right - rectSplitterClient.left - 2, (rectSplitterClient.bottom - rectSplitterClient.top) / 2);

			EndPaint(hWndSplitter, &ps);
		}
		break;
	case WM_LBUTTONDOWN:
		{
			SetCapture(hWndSplitter);

			SPLITTERSTRUCT ss{ hWndSplitter, (UINT_PTR)GetDlgCtrlID(hWndSplitter), (UINT)SPLITTERNOTIFICATION::DRAGBEGIN, { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) } };
			SendMessage(GetParent(hWndSplitter), WM_NOTIFY, (WPARAM)GetDlgCtrlID(hWndSplitter), (LPARAM)&ss);
		}
		break;
	case WM_MOUSEMOVE:
		{
			if ((wParam == MK_LBUTTON))
			{
				SPLITTERSTRUCT ss{ hWndSplitter, (UINT_PTR)GetDlgCtrlID(hWndSplitter), (UINT)SPLITTERNOTIFICATION::DRAGGING, { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) } };
				SendMessage(GetParent(hWndSplitter), WM_NOTIFY, (WPARAM)GetDlgCtrlID(hWndSplitter), (LPARAM)&ss);
			}
		}
		break;
	case WM_LBUTTONUP:
		{
			ReleaseCapture();

			SPLITTERSTRUCT ss{ hWndSplitter, (UINT_PTR)GetDlgCtrlID(hWndSplitter), (UINT)SPLITTERNOTIFICATION::DRAGEND, { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) } };
			SendMessage(GetParent(hWndSplitter), WM_NOTIFY, (WPARAM)GetDlgCtrlID(hWndSplitter), (LPARAM)&ss);
		}
		break;
	default:
		{
			ret = DefWindowProc(hWndSplitter, Msg, wParam, lParam);
		}
		break;
	}

	return ret;
}