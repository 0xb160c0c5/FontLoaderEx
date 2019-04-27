#include <windows.h>
#include <windowsx.h>
#include "Splitter.h"

LRESULT CALLBACK SplitterProc(HWND hWndSplitter, UINT Message, WPARAM wParam, LPARAM lParam);

ATOM InitSplitter()
{
	WNDCLASS wc{ CS_HREDRAW | CS_VREDRAW, SplitterProc, 0, 0, (HINSTANCE)GetModuleHandle(NULL), NULL, LoadCursor(NULL, IDC_SIZENS), GetSysColorBrush(COLOR_BTNFACE), NULL, UC_SPLITTER };

	return RegisterClass(&wc);
}

LRESULT CALLBACK SplitterProc(HWND hWndSplitter, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	static HPEN hPenSplitter{};

	switch (Message)
	{
	case WM_CREATE:
		{
			hPenSplitter = CreatePen(PS_SOLID, 0, (COLORREF)GetSysColor(COLOR_GRAYTEXT));
		}
		break;
	case WM_PAINT:
		{
			// Draw a horizontal line in the middle
			PAINTSTRUCT ps{};
			HDC hDCSplitter{ BeginPaint(hWndSplitter, &ps) };

			SelectPen(hDCSplitter, hPenSplitter);

			RECT rcSplitterClient{};
			GetClientRect(hWndSplitter, &rcSplitterClient);
			MoveToEx(hDCSplitter, rcSplitterClient.left + (rcSplitterClient.bottom - rcSplitterClient.top) / 2, (rcSplitterClient.bottom - rcSplitterClient.top) / 2, NULL);
			LineTo(hDCSplitter, (rcSplitterClient.right - rcSplitterClient.left) - (rcSplitterClient.bottom - rcSplitterClient.top) / 2, (rcSplitterClient.bottom - rcSplitterClient.top) / 2);

			EndPaint(hWndSplitter, &ps);
		}
		break;
	case WM_LBUTTONDOWN:
		{
			// Capture mouse and send WM_NOTIFY with SPLITTERNOTIFICATION::DRAGBEGIN to parent window
			SetCapture(hWndSplitter);

			SPLITTERSTRUCT ss{ hWndSplitter, (UINT_PTR)GetDlgCtrlID(hWndSplitter), (UINT)SPLITTERNOTIFICATION::DRAGBEGIN, { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) } };
			SendMessage(GetParent(hWndSplitter), WM_NOTIFY, (WPARAM)GetDlgCtrlID(hWndSplitter), (LPARAM)&ss);
		}
		break;
	case WM_MOUSEMOVE:
		{
			// If left button is being hold, send WM_NOTIFY with SPLITTERNOTIFICATION::DRAGGING to parent window
			if ((wParam == MK_LBUTTON))
			{
				SPLITTERSTRUCT ss{ hWndSplitter, (UINT_PTR)GetDlgCtrlID(hWndSplitter), (UINT)SPLITTERNOTIFICATION::DRAGGING, { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) } };
				SendMessage(GetParent(hWndSplitter), WM_NOTIFY, (WPARAM)GetDlgCtrlID(hWndSplitter), (LPARAM)&ss);
			}
		}
		break;
	case WM_LBUTTONUP:
		{
			// Release mouse and send WM_NOTIFY with SPLITTERNOTIFICATION::DRAGBEGIN to parent window
			ReleaseCapture();

			SPLITTERSTRUCT ss{ hWndSplitter, (UINT_PTR)GetDlgCtrlID(hWndSplitter), (UINT)SPLITTERNOTIFICATION::DRAGEND, { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) } };
			SendMessage(GetParent(hWndSplitter), WM_NOTIFY, (WPARAM)GetDlgCtrlID(hWndSplitter), (LPARAM)&ss);
		}
		break;
	case WM_DESTROY:
		{
			DeletePen(hPenSplitter);
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