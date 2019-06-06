#include <windows.h>
#include <windowsx.h>
#include <vector>
#include <array>
#include <algorithm>
#include <cassert>
#include "Splitter.h"

LRESULT CALLBACK SplitterProc(HWND hWndSplitter, UINT Message, WPARAM wParam, LPARAM lParam);

ATOM InitSplitter()
{
	WNDCLASS wc{ 0, SplitterProc, 0, 0, static_cast<HINSTANCE>(GetModuleHandle(NULL)), NULL, NULL, NULL, NULL, UC_SPLITTER };

	return RegisterClass(&wc);
}

LRESULT CALLBACK SplitterProc(HWND hWndSplitter, UINT Message, WPARAM wParam, LPARAM lParam)
{
	LRESULT ret{};

	static DWORD dwSplitterStyle{};
	static WORD idSplitter{};
	static POINT ptSplitterRange{};
	static DWORD dwLineMargin{};
	static POINT ptCursorOffset{};
	static std::array<std::vector<HWND>, 2> LinkedControl;

	switch (Message)
	{
	case SPM_ROTATE:
		{
			// Assign new styles
			DWORD dwSplitterStyleNew{ (dwSplitterStyle & (~(SPS_HORZ | SPS_VERT))) | ((dwSplitterStyle & (SPS_HORZ | SPS_VERT)) ^ (SPS_HORZ | SPS_VERT)) };
			if (dwSplitterStyleNew & SPS_PARENTWIDTH)
			{
				dwSplitterStyle = (dwSplitterStyleNew & (~SPS_PARENTWIDTH)) | SPS_PARENTHEIGHT;
			}
			if (dwSplitterStyleNew & SPS_PARENTHEIGHT)
			{
				dwSplitterStyle = (dwSplitterStyleNew & (~SPS_PARENTHEIGHT)) | SPS_PARENTWIDTH;
			}
			SetWindowLongPtr(hWndSplitter, GWL_STYLE, static_cast<LONG>(dwSplitterStyleNew));

			// Redraw splitter
			InvalidateRect(hWndSplitter, NULL, FALSE);
		}
		break;
	case SPM_SETRANGE:
		{
			if (wParam)
			{
				HWND hWndSplitterParent{ GetAncestor(hWndSplitter, GA_PARENT) };
				RECT rcSplitter{};
				GetWindowRect(hWndSplitter, &rcSplitter);
				MapWindowRect(HWND_DESKTOP, hWndSplitterParent, &rcSplitter);
				if (dwSplitterStyle & SPS_HORZ)
				{
					ptSplitterRange = { LOWORD(wParam), HIWORD(wParam) - (rcSplitter.bottom - rcSplitter.top) };
				}
				else if (dwSplitterStyle & SPS_VERT)
				{
					ptSplitterRange = { LOWORD(wParam), HIWORD(wParam) - (rcSplitter.right - rcSplitter.left) };
				}
				if (ptSplitterRange.y >= ptSplitterRange.x)
				{
					ret = static_cast<LRESULT>(TRUE);
				}
				else
				{
					ptSplitterRange = {};

					ret = static_cast<LRESULT>(FALSE);
				}
			}
			else
			{
				ptSplitterRange = {};

				ret = static_cast<LRESULT>(TRUE);
			}
		}
		break;
	case SPM_GETRANGE:
		{
			ret = MAKELRESULT(ptSplitterRange.x, ptSplitterRange.y);
		}
		break;
	case SPM_SETMARGIN:
		{
			dwLineMargin = static_cast<DWORD>(wParam);
			RECT rcSplitterClient{};
			GetClientRect(hWndSplitter, &rcSplitterClient);
			if (dwSplitterStyle & SPS_HORZ)
			{
				POINT ptLineStart{ rcSplitterClient.left + static_cast<LONG>(dwLineMargin), rcSplitterClient.top + (rcSplitterClient.bottom - rcSplitterClient.top) / 2 };
				RECT rcSplitterClientLeftPart{ rcSplitterClient.left, rcSplitterClient.top, rcSplitterClient.left + (rcSplitterClient.right - rcSplitterClient.left) / 2, rcSplitterClient.bottom };
				if (!PtInRect(&rcSplitterClientLeftPart, ptLineStart))
				{
					dwLineMargin = 0;

					ret = static_cast<LRESULT>(FALSE);
					break;
				}
			}
			else if (dwSplitterStyle & SPS_VERT)
			{
				POINT ptLineStart{ rcSplitterClient.left + (rcSplitterClient.right - rcSplitterClient.left) / 2, rcSplitterClient.top + static_cast<LONG>(dwLineMargin) };
				RECT rcSplitterClientUpperPart{ rcSplitterClient.left, rcSplitterClient.top, rcSplitterClient.right, rcSplitterClient.top + (rcSplitterClient.bottom - rcSplitterClient.top) / 2 };
				if (!PtInRect(&rcSplitterClientUpperPart, ptLineStart))
				{
					dwLineMargin = 0;

					ret = static_cast<LRESULT>(FALSE);
					break;
				}
			}
			else
			{
				dwLineMargin = 0;

				ret = static_cast<LRESULT>(FALSE);
				break;
			}

			// Redraw splitter
			InvalidateRect(hWndSplitter, NULL, FALSE);

			ret = static_cast<LRESULT>(TRUE);
		}
		break;
	case SPM_GETMARGIN:
		{
			ret = static_cast<LRESULT>(dwLineMargin);
		}
		break;
	case SPM_SETLINKEDCTL:
		{
			switch (HIWORD(wParam))
			{
			case SLC_TOP:
				{
					if (dwSplitterStyle & SPS_HORZ)
					{
						LinkedControl[0].clear();
						try
						{
							for (WORD i = 0; i < LOWORD(wParam); i++)
							{
								if (IsWindow(reinterpret_cast<HWND*>(lParam)[i]) && (GetAncestor(reinterpret_cast<HWND*>(lParam)[i], GA_PARENT) == GetAncestor(hWndSplitter, GA_PARENT)))
								{
									LinkedControl[0].push_back(reinterpret_cast<HWND*>(lParam)[i]);
								}
							}
							LinkedControl[0].erase(std::unique(LinkedControl[0].begin(), LinkedControl[0].end()), LinkedControl[0].end());
						}
						catch (...)
						{
							LinkedControl[0].clear();

							ret = 0;
							break;
						}

						ret = static_cast<LRESULT>(LinkedControl[0].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_BOTTOM:
				{
					if (dwSplitterStyle & SPS_HORZ)
					{
						LinkedControl[1].clear();
						try
						{
							for (WORD i = 0; i < LOWORD(wParam); i++)
							{
								if (IsWindow(reinterpret_cast<HWND*>(lParam)[i]) && (GetAncestor(reinterpret_cast<HWND*>(lParam)[i], GA_PARENT) == GetAncestor(hWndSplitter, GA_PARENT)))
								{
									LinkedControl[1].push_back(reinterpret_cast<HWND*>(lParam)[i]);
								}
							}
							LinkedControl[1].erase(std::unique(LinkedControl[1].begin(), LinkedControl[1].end()), LinkedControl[1].end());
						}
						catch (...)
						{
							LinkedControl[1].clear();

							ret = 0;
							break;
						}

						ret = static_cast<LRESULT>(LinkedControl[1].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_LEFT:
				{
					if (dwSplitterStyle & SPS_VERT)
					{
						LinkedControl[0].clear();
						try
						{
							for (WORD i = 0; i < LOWORD(wParam); i++)
							{
								if (IsWindow(reinterpret_cast<HWND*>(lParam)[i]) && (GetAncestor(reinterpret_cast<HWND*>(lParam)[i], GA_PARENT) == GetAncestor(hWndSplitter, GA_PARENT)))
								{
									LinkedControl[0].push_back(reinterpret_cast<HWND*>(lParam)[i]);
								}
							}
							LinkedControl[0].erase(std::unique(LinkedControl[0].begin(), LinkedControl[0].end()), LinkedControl[0].end());
						}
						catch (...)
						{
							LinkedControl[0].clear();

							ret = 0;
							break;
						}

						ret = static_cast<LRESULT>(LinkedControl[0].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_RIGHT:
				{
					if (dwSplitterStyle & SPS_VERT)
					{
						LinkedControl[1].clear();
						try
						{
							for (WORD i = 0; i < LOWORD(wParam); i++)
							{
								if (IsWindow(reinterpret_cast<HWND*>(lParam)[i]) && (GetAncestor(reinterpret_cast<HWND*>(lParam)[i], GA_PARENT) == GetAncestor(hWndSplitter, GA_PARENT)))
								{
									LinkedControl[1].push_back(reinterpret_cast<HWND*>(lParam)[i]);
								}
							}
							LinkedControl[1].erase(std::unique(LinkedControl[1].begin(), LinkedControl[1].end()), LinkedControl[1].end());
						}
						catch (...)
						{
							LinkedControl[1].clear();

							ret = 0;
							break;
						}

						ret = static_cast<LRESULT>(LinkedControl[1].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			default:
				{
					ret = 0;
				}
				break;
			}
		}
		break;
	case SPM_GETLINKEDCTL:
		{
			switch (HIWORD(wParam))
			{
			case SLC_TOP:
				{
					if (dwSplitterStyle & SPS_HORZ)
					{
						if (lParam)
						{
							for (WORD i = 0; i < static_cast<WORD>(LinkedControl[0].size()); i++)
							{
								reinterpret_cast<HWND*>(lParam)[i] = LinkedControl[0][i];
							}
						}

						ret = static_cast<LRESULT>(LinkedControl[0].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_BOTTOM:
				{
					if (dwSplitterStyle & SPS_HORZ)
					{
						if (lParam)
						{
							for (WORD i = 0; i < static_cast<WORD>(LinkedControl[1].size()); i++)
							{
								reinterpret_cast<HWND*>(lParam)[i] = LinkedControl[1][i];
							}
						}

						ret = static_cast<LRESULT>(LinkedControl[1].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_LEFT:
				{
					if (dwSplitterStyle & SPS_VERT)
					{
						if (lParam)
						{
							for (WORD i = 0; i < static_cast<WORD>(LinkedControl[0].size()); i++)
							{
								reinterpret_cast<HWND*>(lParam)[i] = LinkedControl[0][i];
							}
						}

						ret = static_cast<LRESULT>(LinkedControl[0].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_RIGHT:
				{
					if (dwSplitterStyle & SPS_VERT)
					{
						if (lParam)
						{
							for (WORD i = 0; i < static_cast<WORD>(LinkedControl[1].size()); i++)
							{
								reinterpret_cast<HWND*>(lParam)[i] = LinkedControl[1][i];
							}
						}

						ret = static_cast<LRESULT>(LinkedControl[1].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			default:
				{
					ret = 0;
				}
				break;
			}
		}
		break;
	case SPM_ADDLINKEDCTL:
		{
			std::vector<HWND> LinkedControlTemp{};

			switch (HIWORD(wParam))
			{
			case SLC_TOP:
				{
					if (dwSplitterStyle & SPS_HORZ)
					{
						try
						{
							for (WORD i = 0; i < LOWORD(wParam); i++)
							{
								if (IsWindow(reinterpret_cast<HWND*>(lParam)[i]) && (GetAncestor(reinterpret_cast<HWND*>(lParam)[i], GA_PARENT) == GetAncestor(hWndSplitter, GA_PARENT)))
								{
									LinkedControlTemp.push_back(reinterpret_cast<HWND*>(lParam)[i]);
								}
							}
							LinkedControlTemp.erase(std::unique(LinkedControlTemp.begin(), LinkedControlTemp.end()), LinkedControlTemp.end());
							LinkedControl[0].reserve(LinkedControl[0].size() + LinkedControlTemp.size());
						}
						catch (...)
						{
							ret = 0;
							break;
						}
						LinkedControl[0].insert(LinkedControl[0].end(), LinkedControlTemp.begin(), LinkedControlTemp.end());

						ret = static_cast<LRESULT>(LinkedControlTemp.size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_BOTTOM:
				{
					if (dwSplitterStyle & SPS_HORZ)
					{
						try
						{
							for (WORD i = 0; i < LOWORD(wParam); i++)
							{
								if (IsWindow(reinterpret_cast<HWND*>(lParam)[i]) && (GetAncestor(reinterpret_cast<HWND*>(lParam)[i], GA_PARENT) == GetAncestor(hWndSplitter, GA_PARENT)))
								{
									LinkedControlTemp.push_back(reinterpret_cast<HWND*>(lParam)[i]);
								}
							}
							LinkedControlTemp.erase(std::unique(LinkedControlTemp.begin(), LinkedControlTemp.end()), LinkedControlTemp.end());
							LinkedControl[1].reserve(LinkedControl[1].size() + LinkedControlTemp.size());
						}
						catch (...)
						{
							ret = 0;
							break;
						}
						LinkedControl[1].insert(LinkedControl[1].end(), LinkedControlTemp.begin(), LinkedControlTemp.end());

						ret = static_cast<LRESULT>(LinkedControlTemp.size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_LEFT:
				{
					if (dwSplitterStyle & SPS_VERT)
					{
						try
						{
							for (WORD i = 0; i < LOWORD(wParam); i++)
							{
								if (IsWindow(reinterpret_cast<HWND*>(lParam)[i]) && (GetAncestor(reinterpret_cast<HWND*>(lParam)[i], GA_PARENT) == GetAncestor(hWndSplitter, GA_PARENT)))
								{
									LinkedControlTemp.push_back(reinterpret_cast<HWND*>(lParam)[i]);
								}
							}
							LinkedControlTemp.erase(std::unique(LinkedControlTemp.begin(), LinkedControlTemp.end()), LinkedControlTemp.end());
							LinkedControl[0].reserve(LinkedControl[0].size() + LinkedControlTemp.size());
						}
						catch (...)
						{
							ret = 0;
							break;
						}
						LinkedControl[0].insert(LinkedControl[0].end(), LinkedControlTemp.begin(), LinkedControlTemp.end());

						ret = static_cast<LRESULT>(LinkedControlTemp.size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_RIGHT:
				{
					if (dwSplitterStyle & SPS_VERT)
					{
						try
						{
							for (WORD i = 0; i < LOWORD(wParam); i++)
							{
								if (IsWindow(reinterpret_cast<HWND*>(lParam)[i]) && (GetAncestor(reinterpret_cast<HWND*>(lParam)[i], GA_PARENT) == GetAncestor(hWndSplitter, GA_PARENT)))
								{
									LinkedControlTemp.push_back(reinterpret_cast<HWND*>(lParam)[i]);
								}
							}
							LinkedControlTemp.erase(std::unique(LinkedControlTemp.begin(), LinkedControlTemp.end()), LinkedControlTemp.end());
							LinkedControl[1].reserve(LinkedControl[1].size() + LinkedControlTemp.size());
						}
						catch (...)
						{
							ret = 0;
							break;
						}
						LinkedControl[1].insert(LinkedControl[1].end(), LinkedControlTemp.begin(), LinkedControlTemp.end());

						ret = static_cast<LRESULT>(LinkedControlTemp.size());
					}
				}
				break;
			default:
				{
					ret = 0;
				}
				break;
			}
		}
		break;
	case SPM_REMOVELINKEDCTL:
		{
			switch (HIWORD(wParam))
			{
			case SLC_TOP:
				{
					if (dwSplitterStyle & SPS_HORZ)
					{
						std::size_t LinkedControlOriginalSize{ LinkedControl[0].size() };
						for (WORD i = 0; i < LOWORD(wParam); i++)
						{
							LinkedControl[0].erase(std::find(LinkedControl[0].begin(), LinkedControl[0].end(), reinterpret_cast<HWND*>(lParam)[i]));
						}

						ret = static_cast<LRESULT>(LinkedControlOriginalSize - LinkedControl[0].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_BOTTOM:
				{
					if (dwSplitterStyle & SPS_HORZ)
					{
						std::size_t LinkedControlOriginalSize{ LinkedControl[1].size() };
						for (WORD i = 0; i < LOWORD(wParam); i++)
						{
							LinkedControl[1].erase(std::find(LinkedControl[1].begin(), LinkedControl[1].end(), reinterpret_cast<HWND*>(lParam)[i]));
						}

						ret = static_cast<LRESULT>(LinkedControlOriginalSize - LinkedControl[1].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_LEFT:
				{
					if (dwSplitterStyle & SPS_VERT)
					{
						std::size_t LinkedControlOriginalSize{ LinkedControl[0].size() };
						for (WORD i = 0; i < LOWORD(wParam); i++)
						{
							LinkedControl[0].erase(std::find(LinkedControl[0].begin(), LinkedControl[0].end(), reinterpret_cast<HWND*>(lParam)[i]));
						}

						ret = static_cast<LRESULT>(LinkedControlOriginalSize - LinkedControl[0].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			case SLC_RIGHT:
				{
					if (dwSplitterStyle & SPS_VERT)
					{
						std::size_t LinkedControlOriginalSize{ LinkedControl[1].size() };
						for (WORD i = 0; i < LOWORD(wParam); i++)
						{
							LinkedControl[1].erase(std::find(LinkedControl[1].begin(), LinkedControl[1].end(), reinterpret_cast<HWND*>(lParam)[i]));
						}

						ret = static_cast<LRESULT>(LinkedControlOriginalSize - LinkedControl[1].size());
					}
					else
					{
						ret = 0;
					}
				}
				break;
			default:
				{
					ret = 0;
				}
				break;
			}
		}
		break;
	case WM_CREATE:
		{
			// Get splitter styles and ID
			dwSplitterStyle = static_cast<DWORD>(reinterpret_cast<LPCREATESTRUCT>(lParam)->style);
			idSplitter = static_cast<WORD>(reinterpret_cast<UINT_PTR>(reinterpret_cast<LPCREATESTRUCT>(lParam)->hMenu) & 0xFFFF);

			// Validate styles
			// One of SPS_VERT and SPS_HORZ must be specified and only one of them can be specified at a time
			if (static_cast<bool>(dwSplitterStyle & SPS_HORZ) == static_cast<bool>(dwSplitterStyle & SPS_VERT))
			{
				dwSplitterStyle = dwSplitterStyle & (~(SPS_HORZ | SPS_VERT));
			}
			// SPS_PARENTWIDTH and SPS_PARENTHEIGHT can't be both specified
			if ((dwSplitterStyle & SPS_PARENTWIDTH) && (dwSplitterStyle & SPS_PARENTHEIGHT))
			{
				dwSplitterStyle = dwSplitterStyle & (~(SPS_PARENTWIDTH | SPS_PARENTHEIGHT));
			}
			// SPS_HORZ can't coexist with SPS_PARENTHEIGHT, so does SPS_VERT and SPS_PARENTWIDTH
			if ((dwSplitterStyle & SPS_HORZ) && (dwSplitterStyle & SPS_PARENTHEIGHT))
			{
				dwSplitterStyle = dwSplitterStyle & (~(SPS_PARENTHEIGHT));
			}
			if ((dwSplitterStyle & SPS_VERT) && (dwSplitterStyle & SPS_PARENTWIDTH))
			{
				dwSplitterStyle = dwSplitterStyle & (~(SPS_PARENTWIDTH));
			}
			SetWindowLongPtr(hWndSplitter, GWL_STYLE, static_cast<LONG>(dwSplitterStyle));

			ret = 0;
		}
		break;
	case WM_ERASEBKGND:
		{
			// Prevent Windows from drawing the background
			ret = static_cast<LRESULT>(TRUE);
		}
		break;
	case WM_PAINT:
		{
			PAINTSTRUCT ps{};
			HDC hDCSplitter{ BeginPaint(hWndSplitter, &ps) };

			// Send WM_CTLCOLORSTATIC to parent window
			HWND hWndSplitterParent{ GetAncestor(hWndSplitter, GA_PARENT) };
			HBRUSH hBrSplitterBackground{ FORWARD_WM_CTLCOLORSTATIC(hWndSplitterParent, hDCSplitter, hWndSplitter, SendMessage) };

			// Draw the background
			RECT rcSplitterClient{};
			GetClientRect(hWndSplitter, &rcSplitterClient);
			if (!hBrSplitterBackground)
			{
				hBrSplitterBackground = GetSysColorBrush(COLOR_3DFACE);
			}
			FillRect(hDCSplitter, &rcSplitterClient, hBrSplitterBackground);

			// Draw a line in the middle
			if (dwSplitterStyle & SPS_HORZ)
			{
				MoveToEx(hDCSplitter, rcSplitterClient.left + dwLineMargin, rcSplitterClient.top + (rcSplitterClient.bottom - rcSplitterClient.top) / 2, NULL);
				LineTo(hDCSplitter, rcSplitterClient.right - dwLineMargin, rcSplitterClient.top + (rcSplitterClient.bottom - rcSplitterClient.top) / 2);
			}
			else if (dwSplitterStyle & SPS_VERT)
			{
				MoveToEx(hDCSplitter, rcSplitterClient.left + (rcSplitterClient.right - rcSplitterClient.left) / 2, rcSplitterClient.top + dwLineMargin, NULL);
				LineTo(hDCSplitter, rcSplitterClient.left + (rcSplitterClient.right - rcSplitterClient.left) / 2, rcSplitterClient.bottom - dwLineMargin);
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

			if (!(dwSplitterStyle & SPS_NONOTIFY))
			{
				HWND hWndSplitterParent{ GetAncestor(hWndSplitter, GA_PARENT) };
				RECT rcSplitter{}, rcSplitterClient{};
				GetWindowRect(hWndSplitter, &rcSplitter);
				GetClientRect(hWndSplitter, &rcSplitterClient);
				MapWindowRect(hWndSplitter, HWND_DESKTOP, &rcSplitterClient);
				POINT ptCursor{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
				ptCursorOffset = { ptCursor.x + (rcSplitterClient.left - rcSplitter.left), ptCursor.y + (rcSplitterClient.top - rcSplitter.top) };
				MapWindowPoints(hWndSplitter, hWndSplitterParent, &ptCursor, 1);
				NMSPLITTER nms{ { hWndSplitter, static_cast<UINT_PTR>(idSplitter), static_cast<UINT>(SPN_DRAGBEGIN) }, ptCursor, ptCursorOffset };
				SendMessage(hWndSplitterParent, WM_NOTIFY, static_cast<WPARAM>(idSplitter), reinterpret_cast<LPARAM>(&nms));
			}
		}
		break;
	case WM_MOUSEMOVE:
		{
			// If left button is being hold, send WM_NOTIFY with SPLITTERNOTIFICATION::DRAGGING to parent window
			if ((wParam == MK_LBUTTON))
			{
				HWND hWndSplitterParent{ GetAncestor(hWndSplitter, GA_PARENT) };
				POINT ptCursor{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) }, ptSplitter{}, ptCursorOffsetNew{ ptCursorOffset };
				MapWindowPoints(hWndSplitter, hWndSplitterParent, &ptCursor, 1);
				ptSplitter = { ptCursor.x - ptCursorOffsetNew.x, ptCursor.y - ptCursorOffsetNew.y };
				if ((ptSplitterRange.x != 0) || (ptSplitterRange.y != 0))
				{
					if (dwSplitterStyle & SPS_HORZ)
					{
						if (ptSplitter.y < ptSplitterRange.x)
						{
							ptSplitter.y = ptSplitterRange.x;
							ptCursorOffsetNew.y = ptCursor.y - ptSplitterRange.x;
						}
						if (ptSplitter.y > ptSplitterRange.y)
						{
							ptSplitter.y = ptSplitterRange.y;
							ptCursorOffsetNew.y = ptCursor.y - ptSplitterRange.y;
						}
					}
					if (dwSplitterStyle & SPS_VERT)
					{
						if (ptSplitter.x < ptSplitterRange.x)
						{
							ptSplitter.x = ptSplitterRange.x;
							ptCursorOffsetNew.x = ptCursor.x - ptSplitterRange.x;
						}
						if (ptSplitter.x > ptSplitterRange.y)
						{
							ptSplitter.x = ptSplitterRange.y;
							ptCursorOffsetNew.x = ptCursor.x - ptSplitterRange.y;
						}
					}
				}

				if (dwSplitterStyle & SPS_AUTODRAG)
				{
					FORWARD_WM_MOVE(hWndSplitter, ptSplitter.x, ptSplitter.y, SendMessage);
				}

				if (!(dwSplitterStyle & SPS_NONOTIFY))
				{
					NMSPLITTER nms{ { hWndSplitter, static_cast<UINT_PTR>(idSplitter), static_cast<UINT>(SPN_DRAGGING) }, ptCursor, ptCursorOffsetNew };
					SendMessage(hWndSplitterParent, WM_NOTIFY, static_cast<WPARAM>(idSplitter), reinterpret_cast<LPARAM>(&nms));
				}
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

			if (!(dwSplitterStyle & SPS_NONOTIFY))
			{
				HWND hWndSplitterParent{ GetAncestor(hWndSplitter, GA_PARENT) };
				POINT ptCursor{ GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
				MapWindowPoints(hWndSplitter, hWndSplitterParent, &ptCursor, 1);
				NMSPLITTER nms{ { hWndSplitter, static_cast<UINT_PTR>(idSplitter), static_cast<UINT>(SPN_DRAGEND) }, ptCursor, ptCursorOffset };
				SendMessage(hWndSplitterParent, WM_NOTIFY, static_cast<WPARAM>(idSplitter), reinterpret_cast<LPARAM>(&nms));
			}
		}
		break;
	case WM_SETCURSOR:
		{
			if (reinterpret_cast<HWND>(wParam) == hWndSplitter)
			{
				if (dwSplitterStyle & SPS_HORZ)
				{
					SetCursor(LoadCursor(NULL, IDC_SIZENS));

					ret = static_cast<LRESULT>(TRUE);
				}
				else if (dwSplitterStyle & SPS_VERT)
				{
					SetCursor(LoadCursor(NULL, IDC_SIZEWE));

					ret = static_cast<LRESULT>(TRUE);
				}
				else
				{
					ret = static_cast<LRESULT>(FALSE);
				}
			}
		}
		break;
	case WM_MOVE:
		{
			HWND hWndSplitterParent{ GetAncestor(hWndSplitter, GA_PARENT) };
			RECT rcSplitter{}, rcSplitterParentClient{};
			GetWindowRect(hWndSplitter, &rcSplitter);
			GetClientRect(hWndSplitterParent, &rcSplitterParentClient);
			int xPosSplitter{}, yPosSplitter{}, cxSplitter{}, cySplitter{};
			UINT uFlags{ SWP_NOZORDER | SWP_NOACTIVATE };
			if (dwSplitterStyle & SPS_PARENTWIDTH)
			{
				xPosSplitter = rcSplitterParentClient.left;
				yPosSplitter = GET_Y_LPARAM(lParam);
				cxSplitter = rcSplitterParentClient.right - rcSplitterParentClient.left;
				cySplitter = rcSplitter.bottom - rcSplitter.top;
			}
			else if (dwSplitterStyle & SPS_PARENTHEIGHT)
			{
				xPosSplitter = GET_X_LPARAM(lParam);
				yPosSplitter = rcSplitterParentClient.top;
				cxSplitter = rcSplitter.right - rcSplitter.left;
				cySplitter = rcSplitterParentClient.bottom - rcSplitterParentClient.top;
			}
			else
			{
				xPosSplitter = GET_X_LPARAM(lParam);
				yPosSplitter = GET_Y_LPARAM(lParam);
				uFlags |= SWP_NOSIZE;
			}
			SetWindowPos(hWndSplitter, NULL, xPosSplitter, yPosSplitter, cxSplitter, cySplitter, uFlags);

			MapWindowRect(HWND_DESKTOP, hWndSplitterParent, &rcSplitter);
			for (const auto& i : LinkedControl[0])
			{
				RECT rcLinkedWindow{};
				GetWindowRect(i, &rcLinkedWindow);
				MapWindowRect(HWND_DESKTOP, hWndSplitterParent, &rcLinkedWindow);
				if (dwSplitterStyle & SPS_HORZ)
				{
					SetWindowPos(i, NULL, 0, 0, rcLinkedWindow.right - rcLinkedWindow.left, yPosSplitter - rcLinkedWindow.top, SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
				}
				else if (dwSplitterStyle & SPS_VERT)
				{
					SetWindowPos(i, NULL, 0, 0, xPosSplitter - rcLinkedWindow.left, rcLinkedWindow.bottom - rcLinkedWindow.top, SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
				}
			}
			for (const auto& i : LinkedControl[1])
			{
				RECT rcLinkedWindow{};
				GetWindowRect(i, &rcLinkedWindow);
				MapWindowRect(HWND_DESKTOP, hWndSplitterParent, &rcLinkedWindow);
				if (dwSplitterStyle & SPS_HORZ)
				{
					SetWindowPos(i, NULL, rcLinkedWindow.left, yPosSplitter + cySplitter, rcLinkedWindow.right - rcLinkedWindow.left, rcLinkedWindow.bottom - (yPosSplitter + cySplitter), SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
				}
				else if (dwSplitterStyle & SPS_VERT)
				{
					SetWindowPos(i, NULL, xPosSplitter + cxSplitter, rcLinkedWindow.top, rcLinkedWindow.right - (xPosSplitter + cxSplitter), rcLinkedWindow.bottom - rcLinkedWindow.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
				}
			}
		}
		break;
	case WM_SIZE:
		{
			HWND hWndSplitterParent{ GetAncestor(hWndSplitter, GA_PARENT) };
			RECT rcSplitter{}, rcSplitterParentClient{};
			GetWindowRect(hWndSplitter, &rcSplitter);
			MapWindowRect(HWND_DESKTOP, hWndSplitterParent, &rcSplitter);
			GetClientRect(hWndSplitterParent, &rcSplitterParentClient);
			int xPosSplitter{}, yPosSplitter{}, cxSplitter{}, cySplitter{};
			UINT uFlags{ SWP_NOZORDER | SWP_NOACTIVATE };
			if (dwSplitterStyle & SPS_PARENTWIDTH)
			{
				xPosSplitter = rcSplitterParentClient.left;
				yPosSplitter = rcSplitter.top;
				cxSplitter = rcSplitterParentClient.right - rcSplitterParentClient.left;
				cySplitter = HIWORD(lParam);
			}
			else if (dwSplitterStyle & SPS_PARENTHEIGHT)
			{
				xPosSplitter = rcSplitter.left;
				yPosSplitter = rcSplitterParentClient.top;
				cxSplitter = LOWORD(lParam);
				cySplitter = rcSplitterParentClient.bottom - rcSplitterParentClient.top;
			}
			else
			{
				cxSplitter = LOWORD(lParam);
				cySplitter = HIWORD(lParam);
				uFlags |= SWP_NOMOVE;
			}
			SetWindowPos(hWndSplitter, NULL, xPosSplitter, yPosSplitter, cxSplitter, cySplitter, uFlags);

			MapWindowRect(HWND_DESKTOP, hWndSplitterParent, &rcSplitter);
			for (const auto& i : LinkedControl[0])
			{
				RECT rcLinkedWindow{};
				GetWindowRect(i, &rcLinkedWindow);
				MapWindowRect(HWND_DESKTOP, hWndSplitterParent, &rcLinkedWindow);
				if (dwSplitterStyle & SPS_HORZ)
				{
					SetWindowPos(i, NULL, 0, 0, rcLinkedWindow.right - rcLinkedWindow.left, yPosSplitter - rcLinkedWindow.top, SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
				}
				else if (dwSplitterStyle & SPS_VERT)
				{
					SetWindowPos(i, NULL, 0, 0, xPosSplitter - rcLinkedWindow.left, rcLinkedWindow.bottom - rcLinkedWindow.top, SWP_NOZORDER | SWP_NOMOVE | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
				}
			}
			for (const auto& i : LinkedControl[1])
			{
				RECT rcLinkedWindow{};
				GetWindowRect(i, &rcLinkedWindow);
				MapWindowRect(HWND_DESKTOP, hWndSplitterParent, &rcLinkedWindow);
				if (dwSplitterStyle & SPS_HORZ)
				{
					SetWindowPos(i, NULL, rcLinkedWindow.left, yPosSplitter + cySplitter, rcLinkedWindow.right - rcLinkedWindow.left, rcLinkedWindow.bottom - (yPosSplitter + cySplitter), SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
				}
				else if (dwSplitterStyle & SPS_VERT)
				{
					SetWindowPos(i, NULL, xPosSplitter + cxSplitter, rcLinkedWindow.top, rcLinkedWindow.right - (xPosSplitter + cxSplitter), rcLinkedWindow.bottom - rcLinkedWindow.top, SWP_NOZORDER | SWP_NOACTIVATE | SWP_ASYNCWINDOWPOS);
				}
			}
		}
		break;
	case WM_STYLECHANGING:
		{
			if (wParam == GWL_STYLE)
			{
				// Validate styles
				// One of SPS_VERT and SPS_HORZ must be specified and only one of them can be specified at a time
				if (static_cast<bool>(reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & SPS_HORZ) == static_cast<bool>(reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & SPS_VERT))
				{
					reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew = reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & (~(SPS_HORZ | SPS_VERT));
				}
				// SPS_PARENTWIDTH and SPS_PARENTHEIGHT can't be both specified
				if ((reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & SPS_PARENTWIDTH) && (reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & SPS_PARENTHEIGHT))
				{
					reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew = reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & (~(SPS_PARENTWIDTH | SPS_PARENTHEIGHT));
				}
				// SPS_HORZ can't coexist with SPS_PARENTHEIGHT, so does SPS_VERT and SPS_PARENTWIDTH
				if ((reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & SPS_HORZ) && (reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & SPS_PARENTHEIGHT))
				{
					reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew = reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & (~(SPS_PARENTHEIGHT));
				}
				if ((reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & SPS_VERT) && (reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & SPS_PARENTWIDTH))
				{
					reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew = reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew & (~(SPS_PARENTWIDTH));
				}

				// Redraw splitter
				if (reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleNew != reinterpret_cast<LPSTYLESTRUCT>(lParam)->styleOld)
				{
					InvalidateRect(hWndSplitter, NULL, FALSE);
				}
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