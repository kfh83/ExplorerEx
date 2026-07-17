#include "pch.h"

#include "DPIHelpers.h"

int g_iLPX;
int g_iLPY;
BOOL g_fHighDPI;
BOOL g_fHighDPIAware;

void InitDPI()
{
    BOOL fIsHighDpiAware = IsProcessDPIAware();
    if (g_iLPX == -1 || g_fHighDPIAware != fIsHighDpiAware)
    {
        g_fHighDPIAware = fIsHighDpiAware;
        HDC hdcScreen = GetDC(nullptr);
        if (hdcScreen)
        {
            g_iLPX = GetDeviceCaps(hdcScreen, LOGPIXELSX);
            g_iLPY = GetDeviceCaps(hdcScreen, LOGPIXELSY);
            g_fHighDPI = g_iLPX != USER_DEFAULT_SCREEN_DPI;
            ReleaseDC(nullptr, hdcScreen);
        }
    }
}

void SHLogicalToPhysicalDPI(SIZE* pSize)
{
    InitDPI();

    pSize->cx = MulDiv(pSize->cx, g_iLPX, USER_DEFAULT_SCREEN_DPI);
    pSize->cy = MulDiv(pSize->cy, g_iLPY, USER_DEFAULT_SCREEN_DPI);
}

void SHLogicalToPhysicalDPI(int* px, int* py)
{
    InitDPI();

    if (px)
    {
        *px = MulDiv(*px, g_iLPX, USER_DEFAULT_SCREEN_DPI);
    }
    if (py)
    {
        *py = MulDiv(*py, g_iLPY, USER_DEFAULT_SCREEN_DPI);
    }
}
