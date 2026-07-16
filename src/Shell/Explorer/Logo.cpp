#include "pch.h"

#include "Logo.h"

#include "cabinet.h"
#include "shundoc.h"

int g_bmpWidth = 0;
int g_bmpHeight = 0;
HWND g_hwndWatermark = nullptr;

BOOL DrawLogoBitmap(HDC hdc, HBITMAP hbm)
{
    BOOL fRet = FALSE;
    HDC hdcMem = CreateCompatibleDC(hdc);
    if (hdcMem)
    {
        HBITMAP hbmMem = (HBITMAP)SelectObject(hdcMem, hbm);
        fRet = BitBlt(hdc, 0, 0, g_bmpWidth, g_bmpHeight, hdcMem, 0, 0, SRCCOPY);
        SelectObject(hdcMem, hbmMem);
        DeleteDC(hdcMem);
    }
    return fRet;
}

void GetWatermarkOffset(POINT* pPt, RECT* pClientRect)
{
    ASSERT(pPt != NULL); // 419
    ASSERT(pClientRect != NULL); // 420

    pPt->x = IsBiDiLocalizedSystem() ? pClientRect->left + 10 : pClientRect->right - g_bmpWidth - 10;
    pPt->y = pClientRect->bottom - g_bmpHeight - 10;
}

void RePositionWatermark(HWND hWnd, BOOL bRepaint)
{
    RECT rcWorkArea = {};
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &rcWorkArea, 0);

    POINT ptOffset = { 0, 0 };
    GetWatermarkOffset(&ptOffset, &rcWorkArea);
    MoveWindow(hWnd, ptOffset.x, ptOffset.y, g_bmpWidth, g_bmpHeight, bRepaint);
}

LRESULT CALLBACK WatermarkWndProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    if (uMsg == 5)
    {
        if (wParam == 1)
        {
            ShowWindow(hWnd, 10);
            return 0;
        }
        return DefWindowProc(hWnd, 5u, wParam, lParam);
    }
    if (uMsg == 15)
    {
        COLORREF crPixel = 0;

        PAINTSTRUCT ps = { nullptr };
        HDC hdc = BeginPaint(hWnd, &ps);
        if (hdc)
        {
            HBITMAP h = LoadBitmap(g_hinstCabinet, (LPCWSTR)0xB8);
            if (h)
            {
                HDC CompatibleDC = CreateCompatibleDC(hdc);
                HDC v10 = CompatibleDC;
                if (CompatibleDC)
                {
                    HGDIOBJ v12 = SelectObject(CompatibleDC, h);
                    if (!v12 || (crPixel = GetPixel(v10, 0, 0), crPixel != -1))
                    {
                        int DeviceCaps = GetDeviceCaps(hdc, BITSPIXEL);
                        SetLayeredWindowAttributes(hWnd, crPixel, 0xA5u, 2 * (DeviceCaps >= 16) + 1);
                    }
                    SelectObject(v10, v12);
                    DeleteDC(v10);
                }
                DrawLogoBitmap(hdc, h);
                DeleteObject(h);
            }
            EndPaint(hWnd, &ps);
        }
    }
    else
    {
        if (uMsg != 16)
        {
            switch (uMsg)
            {
                case WM_SETTINGCHANGE:
                    if (wParam == 47)
                        RePositionWatermark(hWnd, TRUE);
                    break;
                case WM_WINDOWPOSCHANGING:
                    if (lParam)
                    {
                        WINDOWPOS* pwp = (WINDOWPOS*)lParam;
                        UINT flags = pwp->flags;
                        if ((flags & 0x80u) != 0)
                            pwp->flags = flags & 0xFFFFFF7F;
                        int v7 = pwp->flags;
                        if ((v7 & 4) != 0)
                        {
                            pwp->hwndInsertAfter = (HWND)-1;
                            pwp->flags = v7 & 0xFFFFFFFB;
                        }
                    }
                    break;
                case WM_DISPLAYCHANGE:
                    RePositionWatermark(hWnd, FALSE);
                    InvalidateRect(hWnd, nullptr, FALSE);
                    break;
                default:
                    return DefWindowProc(hWnd, uMsg, wParam, lParam);
            }
        }
    }
    return 0;
}

HWND CreateWatermark(HINSTANCE hInstance)
{
    COLORREF crKey = CLR_INVALID;
    HWND hwnd = NULL;
    HDC hDC = NULL;

    if (g_hwndWatermark)
        DeleteWatermark(hInstance);

    WNDCLASSEX wc = {};
    wc.hInstance = hInstance;
    wc.cbSize = sizeof(wc);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = WatermarkWndProc;
    wc.lpszClassName = L"CSWatermark";
    if (!RegisterClassEx(&wc))
        return NULL;

    HBITMAP hbmWatermark = LoadBitmap(g_hinstCabinet, MAKEINTRESOURCE(IDB_STARTER_WATERMARK));
    if (!hbmWatermark)
        return NULL;

    HWND hwndDesktop = GetDesktopWindow();

    BITMAP bm = {};
    if (GetObject(hbmWatermark, sizeof(BITMAP), &bm))
    {
        g_bmpWidth = bm.bmWidth;
        g_bmpHeight = bm.bmHeight;
        hDC = GetDC(hwndDesktop);
        if (hDC)
        {
            HDC hdcMem = CreateCompatibleDC(hDC);
            ReleaseDC(hwndDesktop, hDC);
            if (hdcMem)
            {
                HBITMAP hbmMem = (HBITMAP)SelectObject(hdcMem, hbmWatermark);
                if (hbmMem)
                {
                    crKey = GetPixel(hdcMem, 0, 0);
                }
                SelectObject(hdcMem, hbmMem);
                DeleteDC(hdcMem);

                RECT rcWorkArea;
                if (crKey != CLR_INVALID && SystemParametersInfo(SPI_GETWORKAREA, 0, &rcWorkArea, 0))
                {
                    POINT ptWatermark = { 0, 0 };
                    GetWatermarkOffset(&ptWatermark, &rcWorkArea);
                    hwnd = SHFusionCreateWindowEx(
                        WS_EX_TOPMOST | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE | WS_EX_LAYERED,
                        TEXT("CSWatermark"),
                        NULL,
                        WS_POPUP,
                        ptWatermark.x, ptWatermark.y,
                        g_bmpWidth, g_bmpHeight,
                        NULL,
                        NULL,
                        hInstance, NULL);
                }
            }
        }
    }

    DeleteObject(hbmWatermark);
    if (hwnd)
    {
        SetLayeredWindowAttributes(hwnd, crKey, 0xA5, (GetDeviceCaps(hDC, BITSPIXEL) >= 16) ? (LWA_COLORKEY | LWA_ALPHA) : LWA_COLORKEY);
        ShowWindow(hwnd, SW_SHOW);
        UpdateWindow(hwnd);
        g_hwndWatermark = hwnd;
    }

    return hwnd;
}

void DeleteWatermark(HINSTANCE hInstance)
{
    if (g_hwndWatermark)
    {
        DestroyWindow(g_hwndWatermark);
        UnregisterClassW(L"CSWatermark", hInstance);
        g_hwndWatermark = nullptr;
    }
}

void UpdateWatermark()
{
    if (g_hwndWatermark)
    {
        SetWindowPos(g_hwndWatermark, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOMOVE | SWP_NOOWNERZORDER);
        RePositionWatermark(g_hwndWatermark, TRUE);
    }
}