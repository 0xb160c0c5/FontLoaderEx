#include <windows.h>
#include <windowsx.h>
#include "Splitter.h"

LRESULT CALLBACK SplitterProc(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam);

ATOM InitSplitter()
{
	WNDCLASS wc{ CS_HREDRAW | CS_VREDRAW, SplitterProc, 0, 0, (HINSTANCE)GetModuleHandle(NULL), NULL, LoadCursor(NULL, IDC_SIZENS), GetSysColorBrush(COLOR_BTNFACE), NULL, UC_SPLITTER };

	return RegisterClass(&wc);
}

LRESULT CALLBACK SplitterProc(HWND hWndSplitter, UINT Msg, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	static HPEN hPen{};

	switch (Msg)
	{
	case WM_CREATE:
		{
			hPen = CreatePen(PS_SOLID, 0, (COLORREF)GetSysColor(COLOR_GRAYTEXT));
		}
		break;
	case WM_PAINT:
		{
			// Draw a horizontal line in the middle
			PAINTSTRUCT ps{};
			HDC hDCSplitter{ BeginPaint(hWndSplitter, &ps) };

			SelectPen(hDCSplitter, hPen);
			RECT rectSplitterClient{};
			GetClientRect(hWndSplitter, &rectSplitterClient);
			MoveToEx(hDCSplitter, rectSplitterClient.left + (rectSplitterClient.bottom - rectSplitterClient.top) / 2, (rectSplitterClient.bottom - rectSplitterClient.top) / 2, NULL);
			LineTo(hDCSplitter, (rectSplitterClient.right - rectSplitterClient.left) - (rectSplitterClient.bottom - rectSplitterClient.top) / 2, (rectSplitterClient.bottom - rectSplitterClient.top) / 2);

			EndPaint(hWndSplitter, &ps);
		}
		break;
	case WM_LBUTTONDOWN:
		{
			// Capture mouse and send SPLITTERNOTIFICATION::DRAGBEGIN to parent window
			SetCapture(hWndSplitter);

			SPLITTERSTRUCT ss{ hWndSplitter, (UINT_PTR)GetDlgCtrlID(hWndSplitter), (UINT)SPLITTERNOTIFICATION::DRAGBEGIN, { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) } };
			SendMessage(GetParent(hWndSplitter), WM_NOTIFY, (WPARAM)GetDlgCtrlID(hWndSplitter), (LPARAM)&ss);
		}
		break;
	case WM_MOUSEMOVE:
		{
			// If holding left button, send SPLITTERNOTIFICATION::DRAGGING to parent window
			if ((wParam == MK_LBUTTON))
			{
				SPLITTERSTRUCT ss{ hWndSplitter, (UINT_PTR)GetDlgCtrlID(hWndSplitter), (UINT)SPLITTERNOTIFICATION::DRAGGING, { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) } };
				SendMessage(GetParent(hWndSplitter), WM_NOTIFY, (WPARAM)GetDlgCtrlID(hWndSplitter), (LPARAM)&ss);
			}
		}
		break;
	case WM_LBUTTONUP:
		{
			// Release mouse and send SPLITTERNOTIFICATION::DRAGBEGIN to parent window
			ReleaseCapture();

			SPLITTERSTRUCT ss{ hWndSplitter, (UINT_PTR)GetDlgCtrlID(hWndSplitter), (UINT)SPLITTERNOTIFICATION::DRAGEND, { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) } };
			SendMessage(GetParent(hWndSplitter), WM_NOTIFY, (WPARAM)GetDlgCtrlID(hWndSplitter), (LPARAM)&ss);
		}
		break;
	case WM_DESTROY:
		{
			DeletePen(hPen);
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