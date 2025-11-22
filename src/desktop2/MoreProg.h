// window class name of More Programs pane control
#define WC_MOREPROGRAMS TEXT("Desktop More Programs Pane")

#include "COWSite.h"

enum OPENHOSTVIEW
{
    OHVIEW_0 = 0x0,
    OHVIEW_1 = 0x1,
    OHVIEW_2 = 0x2,
    OHVIEW_3 = 0x3,
    OHVIEW_4 = 0x4,
    OHVIEW_5 = 0x5,
};

class CMorePrograms
    : public IDropTarget
    , public CAccessible
    , public IServiceProvider
    , public CObjectWithSite
    , public IOleCommandTarget
{
public:
    /*
     *  Interface stuff...
     */

    // *** IUnknown ***
    STDMETHODIMP QueryInterface(REFIID riid, LPVOID *ppvOut);
    STDMETHODIMP_(ULONG) AddRef();
    STDMETHODIMP_(ULONG) Release();

    // *** IDropTarget ***
    STDMETHODIMP DragEnter(IDataObject *pdto, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
    STDMETHODIMP DragOver(DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);
    STDMETHODIMP DragLeave();
    STDMETHODIMP Drop(IDataObject *pdto, DWORD grfKeyState, POINTL pt, DWORD *pdwEffect);

    // *** IAccessible overridden methods ***
    STDMETHODIMP get_accRole(VARIANT varChild, VARIANT *pvarRole);
    STDMETHODIMP get_accState(VARIANT varChild, VARIANT *pvarState);
    STDMETHODIMP get_accKeyboardShortcut(VARIANT varChild, BSTR *pszKeyboardShortcut);
    STDMETHODIMP get_accDefaultAction(VARIANT varChild, BSTR *pszDefAction);
    STDMETHODIMP accDoDefaultAction(VARIANT varChild);

	// *** IServiceProvider ***
	STDMETHODIMP QueryService(REFGUID guidService, REFIID riid, void **ppvObject);

    // *** IOleCommandTarget ***
    STDMETHODIMP QueryStatus(const GUID *pguidCmdGroup, ULONG cCmds, OLECMD prgCmds[], OLECMDTEXT *pCmdText);
    STDMETHODIMP Exec(const GUID *pguidCmdGroup, DWORD nCmdID, DWORD nCmdexecopt, VARIANT *pvarargIn, VARIANT *pvarargOut);

private:
    CMorePrograms(HWND hwnd);
    ~CMorePrograms();

    static LRESULT CALLBACK s_WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnNCCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnCreate(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnNCDestroy(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnCtlColorBtn(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnDrawItem(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnCommand(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnSysColorChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnDisplayChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnSettingChange(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnContextMenu(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnEraseBkgnd(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnNotify(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnSMNFindItem(PSMNDIALOGMESSAGE pdm);
    LRESULT _OnSMNShowNewAppsTip(PSMNMBOOL psmb);
    LRESULT _OnSMNDismiss();
    LRESULT _OnSize(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
	LRESULT _OnTimer(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    LRESULT _OnMouseLeave();

    void    _InitMetrics();
    HWND    _CreateTooltip();
    void    _TooltipAddTool();
    void    _PopBalloon();
    void    _BuildHoverRect(const LPPOINT ppt);
    void    _TrackShellMenu(DWORD dwFlags);
    HRESULT _GetCurView(OPENHOSTVIEW *pView);
    int     _OnSetCurView(OPENHOSTVIEW view);
    int     _Mark(SMNDIALOGMESSAGE *pdm, UINT a3);

    friend BOOL MorePrograms_RegisterClass();

    enum { IDC_ALL = 1,
           IDC_KEYPRESS = 2 };

private:
    HWND _hwnd;
    HWND _hwndButton;
    HWND _hwndTT;
    HWND _hwndBalloon;

    HTHEME _hTheme;

    HFONT _hf;
    HFONT _hfTTBold;                // Bold tooltip font
    HFONT _hfMarlett;
    HBRUSH _hbrBk;                  // Always a stock object

    IDropTargetHelper *_pdth;       // For friendly-looking drag/drop

    COLORREF _clrText;
    COLORREF _clrTextHot;
    COLORREF _clrBk;

    int      _colorHighlight;       // GetSysColor
    int      _colorHighlightText;   // GetSysColor

    DWORD    _tmHoverStart;         // When did the user start a drag/drop hover?

    RECT    field_64;

    // Assorted metrics for painting
    int     _tmAscent;              // Ascent of main font
    int     _tmAscentMarlett;       // Ascent of Marlett font
    int     _cxText;                // width of entire client text
    int     _cxText2;               // Vista - New
    int     _cxTextIndent;          // distance to beginning of text
    int     _cxArrow;               // width of the arrow image or glyph
    MARGINS _margins;               // margins for the proglist listview
    int     _iTextCenterVal;        // space added to top of text to center with arrow bitmap

	int     field_A0;               // Vista - New

    RECT    _rcExclude;             // Exclusion rectangle for when the menu comes up

    // More random stuff
    LONG    _lRef;                  // reference count


    int field_B4;
    DWORD dwordB8;
    int field_BC;

    TCHAR   _chMnem;                // Mnemonic
	WCHAR   _chMnemBack;            // Vista - New
    BOOL    _fMenuOpen;             // Is the menu open?

    IShellMenu *_psmPrograms;       // Cached ShellMenu for perf

    // Large things go at the end
    WCHAR  _szMessage[128];
	WCHAR _szMessageBack[128];
    WCHAR  _szTool[256];
	WCHAR  _szToolBack[256];
};
