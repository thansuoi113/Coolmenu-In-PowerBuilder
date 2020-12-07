// Coolmenu.h: interface for the CCoolmenu class.
//
//////////////////////////////////////////////////////////////////////
#include "MsgHook.h"
#include "CoolmenuGeneral.h"

#if !defined(AFX_COOLMENU_H__6A66ED00_F3C6_450D_B064_ED7230DBADF1__INCLUDED_)
#define AFX_COOLMENU_H__6A66ED00_F3C6_450D_B064_ED7230DBADF1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define IMGWIDTH 16
#define IMGHEIGHT 16
#define IMGPADDING 6
#define XPTEXTPADDING 8
#define TEXTPADDING 2
#define WM_UNINITMENUPOPUP 0x0125
#define HIMETRIC_INCH 2540

const CXGAP = 1;
const CXTEXTMARGIN = 2;
const CXBUTTONMARGIN = 2;
const CYBUTTONMARGIN = 2;

const DT_MYSTANDARD = DT_SINGLELINE|DT_LEFT|DT_VCENTER;

#define HIMETRIC_INCH 2540

#define countof(x)	(sizeof(x)/sizeof(x[0]))

#ifndef ULONG_PTR
typedef unsigned long ULONG_PTR, *PULONG_PTR;
#endif

#ifndef ODS_NOACCEL
#define ODS_NOACCEL         0x0100
#endif

// Additional flagdefinition for highlighting
#ifndef ODS_HOTLIGHT
#define ODS_HOTLIGHT        0x0040
#endif

#ifndef ODS_INACTIVE
#define ODS_INACTIVE        0x0080
#endif

#ifndef DT_HIDEPREFIX
#define DT_HIDEPREFIX  0x00100000
#endif

#ifndef DT_PREFIXONLY
#define DT_PREFIXONLY  0x00200000
#endif

#ifndef SPI_GETKEYBOARDCUES
#define SPI_GETKEYBOARDCUES 0x100A
#endif

#ifndef DFC_POPUPMENU
#define DFC_POPUPMENU	5
#endif

#ifndef DFCS_TRANSPARENT
#define DFCS_TRANSPARENT        0x0800
#endif

#ifndef WM_THEMECHANGED
#define WM_THEMECHANGED         0x031A
#endif

#define ODS_SELECTED_OPEN   0x0800
#define ODS_RIGHTJUSTIFY    0x1000
#define ODS_WRAP            0x2000
#define ODS_HIDDEN          0x4000
#define ODS_DRAW_VERTICAL   0x8000

class CCoolmenu : public CMsgHook
{
protected:
	DECLARE_DYNAMIC(CCoolmenu);
	virtual LRESULT WindowProc(UINT msg, WPARAM wp, LPARAM lp);
	void Initialize();
	BOOL ConvertMenu( HMENU pMenu, UINT nIndex, BOOL bSysMenu, BOOL bShowButtons, BOOL bMenubar );
	int GetButtonIndex( LPCTSTR sName );
	BOOL GetMenuItemTextBold( LPCTSTR sName );
	BOOL GetMenuItemTextItalic( LPCTSTR sName );
	BOOL GetMenuItemTextUnderline( LPCTSTR sName );
	COLORREF GetMenuItemTextColor( LPCTSTR sName );
	COLORREF GetBitmapTransColor( HBITMAP hBitmap, INT xPixel, INT yPixel );
	BOOL IsSystemMenu( BOOL bSysMenu, HMENU hMenu );

	//	DRAWING
	void FillRectCM( HDC dc, const RECT& rc, COLORREF color );
	void PLFillRect(HDC& pDC, const RECT& rc, COLORREF color);
	void PLFillRect3DLight(HDC& pDC, const RECT& rc);
	void DrawMenuText( HDC pDC, RECT& rc, COLORREF color, CString text );
	BOOL Draw3DCheckmark( HDC pDC, const RECT& rc, BOOL bSelected, BOOL bDisabled );
	BOOL DrawNormal( LPDRAWITEMSTRUCT lpds );
	BOOL DrawXP( LPDRAWITEMSTRUCT lpds );
	BOOL Draw2K3( LPDRAWITEMSTRUCT lpds );
	BOOL Draw2K7( LPDRAWITEMSTRUCT lpds );

	CBrush* GetMenuBarBrush();
	void UpdateMenuBarColor(HMENU hMenu);

	void SetLastMenuRect(HDC hDC, LPRECT pRect);
	COLORREF GetMenuBarColor2003();
	void GetMenuBarColor2003(COLORREF& color1, COLORREF& color2, BOOL bBackgroundColor /* = TRUE */);
	COLORREF GetMenuBarColor2007();
	void GetMenuBarColor2007(COLORREF& color1, COLORREF& color2, BOOL bBackgroundColor /* = TRUE */);
	COLORREF GetMenuBarColor(HMENU hMenu = NULL);
	static void DrawGradient(CDC* pDC,CRect& Rect,COLORREF StartColor,COLORREF EndColor, BOOL bHorizontal,BOOL bUseSolid=TRUE);

	void MenuDrawText(HDC hDC ,LPCTSTR lpString,int nCount,LPRECT lpRect,UINT uFormat);

	COLORREF GetBitmapBackground();
	void DrawSpecialCharStyle(CDC* pDC, LPCRECT pRect, TCHAR Sign, DWORD dwStyle);
	void DrawSpecialChar(CDC* pDC, LPCRECT pRect, TCHAR Sign, BOOL bBold);
	void DrawSpecial_WinXP(CDC* pDC, LPCRECT pRect, UINT nID, DWORD dwStyle);

	COLORREF GetThemeColor();
	void SetThemeColors();
	COLORREF GetMenuColor(HMENU hMenu=NULL);
	LPSTR FindMenuStringFromId( HMENU hMenu, UINT nID );
	BOOL OnMeasureItem( LPMEASUREITEMSTRUCT lpms );
	BOOL OnDrawItem( LPDRAWITEMSTRUCT lpms );
	void OnMenuSelect( UINT nItemID, UINT nFlags, HMENU hSysMenu );
	void OnInitMenuPopup( HMENU pPopupMenu, UINT nIndex );
	void OnInitMenu( HMENU hMenu );
	void OnUninitMenuPopup();
	void SetOpenMenu( HMENU hMenu, UINT nIndex, BOOL bInit, BOOL bAdd );
	LRESULT OnMenuChar( UINT nChar, UINT nFlags, CMenu* pMenu );

	HWND					m_hWnd;
	HMENU					m_hCurrentPopup, m_hMenu, m_hLastMenu;
	COLORREF				m_clrMenubar;
	COLORREF				m_bitmapBackground, menuColors[MAX_COLOR], menuDefaultColors[MAX_COLOR];
	DWORD					m_dwMsgPos, m_threadID;;
	BOOL					m_bShowButtons, m_bOnMenu, m_bSysMenu, m_bMouseSelect, m_bMenuOpened, m_bCoolMenubar, m_bIsToolbar, m_MenuBarUpdated, m_bNoWindowPosChanging;
	CImageList*			m_ilButtons;
	CImageList*			m_ilDisabledButtons;
	MENUITEMINFO		m_miInfo;
	CImageList*			m_checkmaps;
	CRect					m_LastActiveMenuRect;
	CMapStringToPtr*	m_imageIdList;
	CMapStringToPtr*	m_menuItemTextColorList;
	CMapStringToPtr*	m_menuItemTextBoldList;
	CMapStringToPtr*	m_menuItemTextItalicList;
	CMapStringToPtr*	m_menuItemTextUnderlineList;
	CPtrList				m_menuList;		// list of HMENU's initialized
	HFONT					m_fontMenu;
	int					m_selectcheck, m_unselectcheck, m_menuStyle;
	CCoolmenuGeneral*	m_CmGeneral;
	HMODULE				m_hLibrary, m_hThemeLibrary;
	BOOL					m_bKeyDown, m_bMenuSet, m_bShadowEnabled, m_bMenuRedraw, m_bReactOnThemeChange, m_bInitMenu, m_bPB105OrAbove, m_bNoPaint, m_bMenuClosed;
	int					m_LastMenubarIndex;

public:
	CCoolmenu();
	virtual ~CCoolmenu();

	HWND GetHwnd();
	DWORD GetThreadID();
	CRect GetLastMenuRect();
	COLORREF GetMenuColor( int colorIndex );
	bool GetHwnd( HWND hWnd );
	void SetHwnd( HWND hWnd, HMENU hMenu );
	void SetMenu( HMENU hMenu );
	HMENU GetHookedMenu();
	void RefreshMenubar(HMENU hMenu);
	void SetThreadID( DWORD threadID );
	void SetImagelist( CImageList* hImagelist, CImageList* hImagelistDisabled );
	void SetImageIdlist( CMapStringToPtr*	imageIdList );
	void SetMenuItemTextColorlist( CMapStringToPtr* menuItemTextColorList );
	void SetMenuItemTextBoldlist( CMapStringToPtr* menuItemTextBoldList );
	void SetMenuItemTextItaliclist( CMapStringToPtr* menuItemTextItalicList );
	void SetMenuItemTextUnderlinelist( CMapStringToPtr* menuItemTextUnderlineList );
	int GetMenuStyle();
	void SetMenuStyle(int menuStyle);
	void SetMenuColor(int colorIndex, COLORREF newColor);
	void SetCoolMenubar(BOOL bCoolMenubar);
	void SetKeyDown(BOOL bKeyDown);
	void SetPB105OrAbove(BOOL bPB105OrAbove);
	void SetInitMenu( BOOL bInitMenu );
	void SetMenuSet( BOOL bMenuSet );
	void SetReactOnThemeChange(BOOL bReactOnThemeChange);
	void SetShadowEnabled(BOOL bShadowEnabled);
	BOOL GetCoolMenubar();
	void SetMenuDefaults(COLORREF menuColors[MAX_COLOR]);
	void SetMenuRedraw( BOOL bRedraw );
};

struct CMenuItemInfo : public MENUITEMINFO {
	CMenuItemInfo()
	{ memset(this, 0, sizeof(MENUITEMINFO));
	  cbSize = sizeof(MENUITEMINFO);
	}
};

class CBoldDC
{
protected:
    CFont m_fontBold;
    HDC m_hDC;
    HFONT m_hDefFont;

public:
    CBoldDC (HDC hDC, bool bBold);
   ~CBoldDC ();
};

class CNewMemDC: public CDC
{
  CRect m_rect;
  CBitmap* m_pOldBitmap;
  CBitmap  m_memBitmap;
  BOOL m_bCancel;

  HDC m_hDcOriginal;

public:
  CNewMemDC(LPCRECT pRect, HDC hdc):m_rect(pRect),m_bCancel(FALSE),m_hDcOriginal(hdc)
  {
    CDC *pOrgDC = CDC::FromHandle(m_hDcOriginal);
    CreateCompatibleDC(pOrgDC);

    m_memBitmap.CreateCompatibleBitmap (pOrgDC,m_rect.Width (),m_rect.Height());
    m_pOldBitmap = SelectObject (&m_memBitmap);
    SetWindowOrg(m_rect.left, m_rect.top);
  }

  // Abborting to copy image from memory dc to client
  void DoCancel(){m_bCancel=TRUE;}

  ~CNewMemDC()
  {
    if(!m_bCancel)
    {
      CDC *pOrgDC = CDC::FromHandle(m_hDcOriginal);
      pOrgDC->BitBlt (m_rect.left,m_rect.top,m_rect.Width (),m_rect.Height (),this,m_rect.left,m_rect.top,SRCCOPY);
    }
    SelectObject (m_pOldBitmap);
  }
};


#endif // !defined(AFX_COOLMENU_H__6A66ED00_F3C6_450D_B064_ED7230DBADF1__INCLUDED_)
