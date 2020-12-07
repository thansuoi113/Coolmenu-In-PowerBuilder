// Coolmenu.cpp: implementation of the CCoolmenu class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Coolmenu.h"
#include <malloc.h>

struct CMyItemData {
	CMyItemData*	pnext;		// next in chain of all allocated
	CString			text;			// item text
	UINT				fType;		// original item type flags
	UINT				fState;		//	original item state
	UINT				nID;
	int				iButton;		// index of button image in image list
	int				iItem;
	BOOL				bInUse;		// item is being used in a menu
	int				itemHeight;
	BOOL				bMenuBar;
	BOOL				bHasSubmenu;
	DWORD				m_dwOpenMenu;
	BOOL				bBold;
	BOOL				bItalic;
	BOOL				bUnderline;
	COLORREF			clrText;
};

CMyItemData* GetNewItemData();
void DeleteItemData(CMyItemData* pmd);
BOOL IsMyItemData(CMyItemData* pmd);

CMyItemData*	m_pItemData;

CMapStringToPtr		m_SharedImageIdList;
Win32Type	g_Shell;

typedef BOOL (WINAPI* FktGradientFill)( IN HDC, IN PTRIVERTEX, IN ULONG, IN PVOID, IN ULONG, IN ULONG);
typedef BOOL (WINAPI* FktIsThemeActive)();
typedef HRESULT (WINAPI* FktSetWindowTheme)(HWND hwnd, LPCWSTR pszSubAppName, LPCWSTR pszSubIdList);

FktGradientFill pGradientFill = NULL;
FktIsThemeActive pIsThemeActive = NULL;
FktSetWindowTheme pSetWindowTheme = NULL;

BOOL IsMenuThemeActive()
{
  return pIsThemeActive?pIsThemeActive():FALSE;
}

IMPLEMENT_DYNAMIC(CCoolmenu, CMsgHook);

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CCoolmenu::CCoolmenu()
{
	Initialize();
}

CCoolmenu::~CCoolmenu()
{
	pGradientFill = NULL;
	FreeLibrary(m_hLibrary);
	pIsThemeActive = NULL;
	pSetWindowTheme = NULL;
	FreeLibrary(m_hThemeLibrary);
}

void CCoolmenu::Initialize()
{
	g_Shell = m_CmGeneral->IsShellType();
	m_bitmapBackground = CLR_DEFAULT;
	m_bShowButtons = TRUE;
	m_bCoolMenubar = TRUE;
	m_bReactOnThemeChange = FALSE;
	m_menuStyle = MenuStyleOffice2K3;
	m_bMenuSet = FALSE;
	m_LastActiveMenuRect.bottom = 0;
	m_LastActiveMenuRect.top = 0;
	m_LastActiveMenuRect.left = 0;
	m_LastActiveMenuRect.right = 0;
	m_selectcheck = 0;
	m_unselectcheck = 0;
	m_bInitMenu = FALSE;
	m_LastMenubarIndex = -1;
	// Try to load the library for gradient drawing
	m_hLibrary = LoadLibrary(_T("Msimg32.dll"));

	// Don't use the gradientfill under Win98 because it is buggy!!!
	if(g_Shell!=Win98 && m_hLibrary)
	{
		pGradientFill = (FktGradientFill)GetProcAddress(m_hLibrary,"GradientFill");
	}
	else
	{
		pGradientFill = NULL;
	}
	m_hThemeLibrary = LoadLibrary(_T("uxtheme.dll"));
	if(m_hThemeLibrary)
	{
		pIsThemeActive = (FktIsThemeActive)GetProcAddress(m_hThemeLibrary,"IsThemeActive");
		pSetWindowTheme = (FktSetWindowTheme)GetProcAddress(m_hThemeLibrary,"SetWindowTheme");
	}
	else
	{
		pIsThemeActive = NULL;
	}
}

int CCoolmenu::GetMenuStyle()
{
	return m_menuStyle;
}

void CCoolmenu::SetMenuStyle(int menuStyle)
{
	m_menuStyle = menuStyle;
	m_bShowButtons = (m_menuStyle!=MenuStyleNormal);
	SetMenu( m_hMenu );
}

void CCoolmenu::SetShadowEnabled(BOOL bShadowEnabled)
{
	m_bShadowEnabled = bShadowEnabled;
}

void CCoolmenu::SetCoolMenubar(BOOL bCoolMenubar)
{
	m_bCoolMenubar = bCoolMenubar;
	SetMenu( m_hMenu );
}

void CCoolmenu::SetKeyDown(BOOL bKeyDown)
{
	m_bKeyDown = bKeyDown;
	m_dwMsgPos = GetMessagePos();
}

void CCoolmenu::SetReactOnThemeChange(BOOL bReactOnThemeChange)
{
	m_bReactOnThemeChange = bReactOnThemeChange;
	SetThemeColors();
	SetMenu( m_hMenu );
}

BOOL CCoolmenu::GetCoolMenubar()
{
	return m_bCoolMenubar;
}

void CCoolmenu::SetPB105OrAbove(BOOL bPB105OrAbove)
{
	m_bPB105OrAbove = bPB105OrAbove;
}

void CCoolmenu::SetInitMenu( BOOL bInitMenu )
{
	m_bInitMenu = bInitMenu;
}

void CCoolmenu::SetMenuSet( BOOL bMenuSet )
{
	m_bNoPaint = FALSE;
	m_LastMenubarIndex = -1;
	m_bMenuSet = FALSE;
	m_bMouseSelect = FALSE;
	m_bKeyDown = FALSE;
//	RefreshMenubar(m_hMenu);
//	SetMenu(m_hMenu);
}

void CCoolmenu::SetMenuColor(int colorIndex, COLORREF newColor)
{
	if(colorIndex>=0 && colorIndex<=MAX_COLOR)
		menuColors[colorIndex] = newColor;
	if(colorIndex==TextColor)
		DrawMenuBar(m_pWndHooked);
	if(colorIndex==MenubarColor)
		UpdateMenuBarColor(m_hMenu);
}

COLORREF CCoolmenu::GetMenuColor( int colorIndex )
{
	if(colorIndex>=0 && colorIndex<=MAX_COLOR)
		return menuColors[colorIndex];
	return -1;
}

HWND CCoolmenu::GetHwnd()
{
	return m_pWndHooked;
}

bool CCoolmenu::GetHwnd( HWND hWnd )
{
	return ( m_pWndHooked && m_pWndHooked == hWnd );
}

void CCoolmenu::SetHwnd( HWND hWnd, HMENU hMenu )
{
	if(hMenu)
		m_hMenu = hMenu;
	else
		m_hMenu = GetMenu( m_pWndHooked );
	SetMenu(hMenu);
}

DWORD CCoolmenu::GetThreadID()
{
	return m_threadID;
}

void CCoolmenu::SetThreadID( DWORD threadID )
{
	m_threadID = threadID;
}

HMENU CCoolmenu::GetHookedMenu()
{
	return m_hMenu;
}

void CCoolmenu::SetMenu( HMENU hMenu )
{
	if(!IsMenu(hMenu)) return;
	//m_bMenuSet = TRUE;
	m_hMenu = hMenu;
	if( !m_bShowButtons || !m_bCoolMenubar )
	{
		ConvertMenu( hMenu, 0, FALSE, FALSE, FALSE );
	}
	else
		if ( ( m_menuStyle == MenuStyleOfficeXP || m_menuStyle == MenuStyleOffice2K3 || m_menuStyle == MenuStyleOffice2K7 ) && m_bCoolMenubar )
			ConvertMenu( hMenu, 0, FALSE, TRUE, TRUE );
		else
			ConvertMenu( hMenu, 0, FALSE, FALSE, FALSE );
	UpdateMenuBarColor( hMenu );
}

void CCoolmenu::SetMenuRedraw( BOOL bRedraw )
{
	m_bMenuRedraw = bRedraw;
}

void CCoolmenu::RefreshMenubar(HMENU hMenu)
{
	m_bCoolMenubar = FALSE;
	ConvertMenu( hMenu, 0, FALSE, FALSE, FALSE );
	m_bCoolMenubar = TRUE;
}

void CCoolmenu::SetImagelist( CImageList* hImagelist, CImageList* hImagelistDisabled )
{
	//	Use shared imagelist. If one already exists then destroy.
	m_ilButtons = hImagelist;
	m_ilDisabledButtons = hImagelistDisabled;
}

void CCoolmenu::SetImageIdlist( CMapStringToPtr* imageIdList )
{
	m_imageIdList = imageIdList;
}

void CCoolmenu::SetMenuItemTextColorlist( CMapStringToPtr* menuItemTextColorList )
{
	m_menuItemTextColorList = menuItemTextColorList;
}

void CCoolmenu::SetMenuItemTextBoldlist( CMapStringToPtr* menuItemTextBoldList )
{
	m_menuItemTextBoldList = menuItemTextBoldList;
}

void CCoolmenu::SetMenuItemTextItaliclist( CMapStringToPtr* menuItemTextItalicList )
{
	m_menuItemTextItalicList = menuItemTextItalicList;
}

void CCoolmenu::SetMenuItemTextUnderlinelist( CMapStringToPtr* menuItemTextUnderlineList )
{
	m_menuItemTextUnderlineList = menuItemTextUnderlineList;
}

LRESULT CCoolmenu::WindowProc(UINT msg, WPARAM wp, LPARAM lp)
{

	ASSERT(m_pOldWndProc);
	switch(msg) {
	case WM_WINDOWPOSCHANGING:
		{
			if (m_bMenuRedraw)
			{
				if(m_hMenu && m_bCoolMenubar )
				{
					RefreshMenubar(m_hMenu);
					SetMenu(m_hMenu);
				}
				m_bMenuRedraw = FALSE;
			}
		}
		break;
	case WM_MEASUREITEM:
		if (OnMeasureItem( (MEASUREITEMSTRUCT*)lp ))
			return TRUE;
		break;
	case WM_DRAWITEM:
		//	Draw a menuitem
		if (OnDrawItem( (DRAWITEMSTRUCT*)lp ) )
			return TRUE;
		break;
	case WM_UNINITMENUPOPUP:
		{
			OnUninitMenuPopup();
		}
		break;
	case WM_INITMENU:
		{
			OnInitMenu( (HMENU)wp );
		}
		break;
	case WM_INITMENUPOPUP:
		//	Systemmenu isn't converted. I like to handle them as well, but didn't succeed.
		if ( IsSystemMenu( (BOOL)HIWORD(lp), (HMENU)wp ) ) return 0;
		{
			m_hCurrentPopup = (HMENU)wp;
			if(IsMenu(m_hCurrentPopup) && ( m_bInitMenu | !m_bPB105OrAbove ) )
			{
				CallWindowProcA(m_pOldWndProc, m_pWndHooked, msg, wp, lp);
				ConvertMenu( m_hCurrentPopup, (UINT)LOWORD(lp), (BOOL)HIWORD(lp), m_bShowButtons, FALSE );
				OnInitMenuPopup( m_hCurrentPopup, (UINT)LOWORD(lp) );
			}
			m_bInitMenu = FALSE;
			return 0;
		}
		break;
	case WM_THEMECHANGED:
		{
			SetThemeColors();
			if(m_bReactOnThemeChange)
			{
				SetMenu(m_hMenu);
			}
		}
		break;
	case WM_SYSCOLORCHANGE:
		if(m_bReactOnThemeChange || ( m_menuStyle == MenuStyleOffice2K || m_menuStyle == MenuStyleNormal ))
		{
			SetMenu(m_hMenu);
			SetThemeColors();
			return 0;
		}
		return 0;
		break;
	case WM_PARENTNOTIFY:
		{
			m_bInitMenu = (LOWORD(wp)==WM_RBUTTONDOWN);
		}
		break;
	case WM_RBUTTONDOWN:
		{
			m_bInitMenu = m_bFixedBar && m_bPB105OrAbove;
		}
		break;
	case WM_NCPAINT:
 		if( m_bNoPaint )
 		{
 			m_bNoPaint = FALSE;
 			return 0;
 		}
		break;
	case WM_ERASEBKGND:
		{
 			if(!m_bMenuSet)
 			{
				RefreshMenubar(GetMenu(m_pWndHooked));
				SetMenu( GetMenu(m_pWndHooked) );
 			}
		}
		break;
	case WM_MENUSELECT:
		{
			m_bNoPaint = ( (HMENU)lp != m_hMenu && m_bPB105OrAbove );
			m_bMouseSelect = ( (UINT)HIWORD(wp) & MF_MOUSESELECT);
			OnMenuSelect((UINT)LOWORD(wp), (UINT)HIWORD(wp), (HMENU)lp);
			if ( GetSubMenu( (HMENU)lp, (UINT)LOWORD(wp) ) == m_hCurrentPopup )
			{
				LRESULT lR = CallWindowProcA(m_pOldWndProc, m_pWndHooked, msg, wp, lp);
				ConvertMenu( m_hCurrentPopup, (UINT)LOWORD(wp), FALSE, m_bShowButtons, FALSE );
				return lR;
			}
		}
		break;
	case WM_MENUCHAR:
		{
			//	When using ownerdraw menu's you have to handle WM_MENUCHAR, otherwise shortcut keys don't
			//	fire the event for that menuitem
			return OnMenuChar((TCHAR)LOWORD(wp), (UINT)HIWORD(wp), CMenu::FromHandle((HMENU)lp));
			//if (lr!=0)
			//	return lr;
		}
		break;
	}
	return m_pNext ? m_pNext->WindowProc(msg, wp, lp) :
		::CallWindowProc(m_pOldWndProc, m_pWndHooked, msg, wp, lp);
}

/*
	Function:		IsSystemMenu
	Returns:			BOOL
	Description:	Determines if menu handle is from a systemmenu
*/
BOOL CCoolmenu::IsSystemMenu( BOOL bSysMenu, HMENU hMenu )
{
	if ( bSysMenu ) return TRUE;

	HMENU sysMenu = GetSystemMenu( m_pWndHooked, FALSE );
	if ( hMenu == sysMenu ) return TRUE;

	MENUITEMINFO mii;
	mii.fMask = MIIM_ID | MIIM_SUBMENU;
	mii.cbSize = sizeof( MENUITEMINFO );
	GetMenuItemInfo( hMenu, 0, TRUE, &mii );
	if ( mii.hSubMenu > 0 ) return FALSE;
	if ( mii.wID >= 0xF000 )
		return TRUE;

	return FALSE;
}

COLORREF CCoolmenu::GetMenuColor(HMENU hMenu)
{
	if(hMenu!=NULL)
	{
		MENUINFO menuInfo={0};
		menuInfo.cbSize = sizeof(menuInfo);
		menuInfo.fMask = MIM_BACKGROUND;

		if(m_CmGeneral->GetMenuInfo(hMenu,&menuInfo) && menuInfo.hbrBack)
		{
			LOGBRUSH logBrush;
			if(GetObject(menuInfo.hbrBack,sizeof(LOGBRUSH),&logBrush))
			{
				return logBrush.lbColor;
			}
		}
	}

	if(g_Shell==WinXP)
	{
		BOOL bFlatMenu = FALSE;
		// theme ist not checket, that must be so
		if( (SystemParametersInfo(SPI_GETFLATMENU,0,&bFlatMenu,0) && bFlatMenu==TRUE) )
		{
			return GetSysColor(COLOR_MENUBAR);
		}
	}
	return GetSysColor(COLOR_MENU);
}


void CCoolmenu::DrawGradient(CDC* pDC,CRect& Rect,COLORREF StartColor,COLORREF EndColor, BOOL bHorizontal,BOOL bUseSolid)
{
	int Count = pDC->GetDeviceCaps(NUMCOLORS);
	if(Count==-1)
		bUseSolid = FALSE;

	// for running under win95 and WinNt 4.0 without loading Msimg32.dll
	if(!bUseSolid && pGradientFill )
	{
		TRIVERTEX vert[2] ;
		GRADIENT_RECT gRect;

		vert [0].y = Rect.top;
		vert [0].x = Rect.left;

//		vert [0].Red    = COLOR16(COLOR16(StartColor % 256)<<8);
//		StartColor /= 256;
//		vert [0].Green  = COLOR16(COLOR16(StartColor%256)<<8);
//		StartColor /= 256;
//		vert [0].Blue   = COLOR16(COLOR16(StartColor)<<8);
		vert [0].Red    = COLOR16( ( StartColor & 0xFF ) * 256 );
		vert [0].Green  = COLOR16( ( ( StartColor & 0xFF00 ) / 0x100 ) * 256 );
		vert [0].Blue   = COLOR16( ( ( StartColor & 0xFF0000 ) / 0x10000 ) * 256 );
		vert [0].Alpha  = 0x0000;

		vert [1].y = Rect.bottom;
		vert [1].x = Rect.right;

//		vert [1].Red    = COLOR16(COLOR16(GetRValue(EndColor))<<8);
//		vert [1].Green  = COLOR16(COLOR16(GetGValue(EndColor))<<8);
//		vert [1].Blue   = COLOR16(COLOR16(GetBValue(EndColor))<<8);
		vert [1].Red    = COLOR16( ( EndColor & 0xFF ) * 256 );
		vert [1].Green  = COLOR16( ( ( EndColor & 0xFF00 ) / 0x100 ) * 256 );
		vert [1].Blue   = COLOR16( ( ( EndColor & 0xFF0000 ) / 0x10000 ) * 256 );
		vert [1].Alpha  = 0x0000;

		gRect.UpperLeft  = 0;
		gRect.LowerRight = 1;

		if(bHorizontal)
		{
			pGradientFill(pDC->m_hDC,vert,2,&gRect,1,GRADIENT_FILL_RECT_H);
		}
		else
		{
			pGradientFill(pDC->m_hDC,vert,2,&gRect,1,GRADIENT_FILL_RECT_V);
		}
	}
	else
	{
		BYTE StartRed   = GetRValue(StartColor);
		BYTE StartGreen = GetGValue(StartColor);
		BYTE StartBlue  = GetBValue(StartColor);

		BYTE EndRed    = GetRValue(EndColor);
		BYTE EndGreen  = GetGValue(EndColor);
		BYTE EndBlue   = GetBValue(EndColor);

		int n = (bHorizontal)?Rect.Width():Rect.Height();

		// only need for the rest, can be optimized
		{
			if(bUseSolid)
			{
				// We need a solid brush (can not be doted)
				pDC->FillSolidRect(Rect,pDC->GetNearestColor(EndColor));
			}
			else
			{
				// We need a brush (can be doted)
				CBrush TempBrush(EndColor);
				pDC->FillRect(Rect,&TempBrush);
			}
		}
		int dy = 2;
		n-=dy;
		for(int dn=0;dn<=n;dn+=dy)
		{
			BYTE ActRed = (BYTE)(MulDiv(int(EndRed)-StartRed,dn,n)+StartRed);
			BYTE ActGreen = (BYTE)(MulDiv(int(EndGreen)-StartGreen,dn,n)+StartGreen);
			BYTE ActBlue = (BYTE)(MulDiv(int(EndBlue)-StartBlue,dn,n)+StartBlue);

			CRect TempRect;
			if(bHorizontal)
			{
				TempRect = CRect(CPoint(Rect.left+dn,Rect.top),CSize(dy,Rect.Height()));
			}
			else
			{
				TempRect = CRect(CPoint(Rect.left,Rect.top+dn),CSize(Rect.Width(),dy));
			}
			if(bUseSolid)
			{
				pDC->FillSolidRect(TempRect,pDC->GetNearestColor(RGB(ActRed,ActGreen,ActBlue)));
			}
			else
			{
				CBrush TempBrush(RGB(ActRed,ActGreen,ActBlue));
				pDC->FillRect(TempRect,&TempBrush);
			}
		}
	}
}

class CNewBrushList : public CObList
{
public:
	CNewBrushList(){}
	~CNewBrushList()
	{
		while(!IsEmpty())
		{
			delete RemoveTail();
		}
	}
};

class CNewBrush : public CBrush
{
public:
	UINT m_nMenuDrawMode;
	COLORREF m_BarColor;
	COLORREF m_BarColor2;

	CNewBrush(UINT menuDrawMode, COLORREF barColor, COLORREF barColor2):m_nMenuDrawMode(menuDrawMode),m_BarColor(barColor),m_BarColor2(barColor2)
	{
		if ( g_Shell!=WinNT4 && g_Shell!=Win95 )
		{
		// Get the desktop hDC...
			HDC hDcDsk = GetWindowDC(0) ;
			CDC* pDcDsk = CDC::FromHandle(hDcDsk);

			CDC clientDC;
			clientDC.CreateCompatibleDC(pDcDsk);

			CRect rect(0,0,(GetSystemMetrics(SM_CXFULLSCREEN)+16)&~7,20);

			CBitmap bitmap;
			bitmap.CreateCompatibleBitmap(pDcDsk,rect.Width(),rect.Height());
			CBitmap* pOldBitmap = clientDC.SelectObject(&bitmap);

			int nRight = rect.right;
			if(rect.right>700)
			{
				rect.right  = 700;
				CCoolmenu::DrawGradient( &clientDC, rect, barColor, barColor2, TRUE, TRUE );
				rect.left = rect.right;
				rect.right = nRight;
				clientDC.FillSolidRect( rect, barColor2 );
			}
			else
			{
				CCoolmenu::DrawGradient( &clientDC, rect, barColor, barColor2, TRUE, TRUE );
			}

			//      DrawGradient(&clientDC,rect,barColor,LightenColor(110,barColor),TRUE,TRUE);

			clientDC.SelectObject( pOldBitmap );

			// Release the desktopdc
			ReleaseDC( 0, hDcDsk );

			CreatePatternBrush( &bitmap );
		}
		else
		{
			CreateSolidBrush( barColor );
		}
	}
};

void CCoolmenu::SetLastMenuRect(HDC hDC, LPRECT pRect)
{
	HWND hWnd = WindowFromDC(hDC);
	if(hWnd && pRect)
	{
		CRect Temp;
		GetWindowRect(hWnd,Temp);
		m_LastActiveMenuRect = *pRect;
		m_LastActiveMenuRect.OffsetRect(Temp.TopLeft());
	}
}

CRect CCoolmenu::GetLastMenuRect()
{
	return m_LastActiveMenuRect;
}

COLORREF CCoolmenu::GetThemeColor()
{
	COLORREF color = CLR_NONE;
	if (IsMenuThemeActive() && m_bReactOnThemeChange )
	{
		// get the them from desktop window
		HANDLE hTheme = m_CmGeneral->localOpenThemeData(NULL,L"Window");
		if(hTheme)
		{
          // TM_PROP(3821, TMT, FILLCOLORHINT, COLOR)          // hint about fill color used (for custom controls)
			m_CmGeneral->localGetThemeColor(hTheme,1,1,3821,&color)==S_OK && (color!=1);
			m_CmGeneral->localCloseThemeData(hTheme);
		}
	}
	return color;
}

void CCoolmenu::SetThemeColors()
{
	COLORREF	themeColor = GetThemeColor();

	if(themeColor==CLR_NONE)
	{
		if(!IsMenuThemeActive() && m_bReactOnThemeChange)
		{
			// No theme active (Windows Classic Style selected)
			this->menuColors[SelectedColor] = m_CmGeneral->MidColor(GetSysColor(COLOR_WINDOW),GetSysColor(COLOR_HIGHLIGHT));//m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.7 );
			//COLORREF	colorSel2 = menuColors[SelectedColor];
			this->menuColors[PenColor] = GetSysColor(COLOR_HIGHLIGHT);
			this->menuColors[SelectedPenColor] = m_CmGeneral->DarkenColor( 50, GetSysColor( COLOR_HIGHLIGHT ) );;//(m_menuStyle==MenuStyleOfficeXP)?RGB(126,126,129):RGB(124,124,148);
			this->menuColors[CheckedColor] = m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.60 );
			this->menuColors[CheckedSelectedColor] = m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.60 );
			this->menuColors[MenubarGradientStartColor] = this->menuColors[SelectedColor];//m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.7 );
			this->menuColors[MenubarGradientEndColor] = this->menuColors[SelectedColor];//m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.7 );
			this->menuColors[MenubarGradientStartColorDown] = m_CmGeneral->LightenColor( GetMenuBarColor(), 0.1 );
			this->menuColors[MenubarGradientEndColorDown] = m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW));
			this->menuColors[SeparatorColor] = m_CmGeneral->LightenColor(GetSysColor(COLOR_HIGHLIGHT), 0.27);
			this->menuColors[MenuBckColor] = m_CmGeneral->MixedColor(m_CmGeneral->DarkenColor(1,GetSysColor(COLOR_WINDOW)),GetMenuColor());//(m_menuStyle==MenuStyleOfficeXP)?RGB(251,250,251):RGB(253,250,255);
			this->menuColors[GradientStartColor] = (m_menuStyle==MenuStyleOffice2K7)?RGB(255,255,255):this->menuColors[MenuBckColor];
			this->menuColors[GradientEndColor] = (m_menuStyle==MenuStyleOffice2K7)?RGB(255,255,255):m_CmGeneral->MixedColor(GetMenuBarColor(),m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW)));
			this->menuColors[BitmapBckColor] = (m_menuStyle==MenuStyleOffice2K7)?RGB(255,255,255):this->menuColors[MenubarGradientStartColorDown];//m_CmGeneral->MidColor(GetSysColor(COLOR_WINDOW),GetSysColor(COLOR_HIGHLIGHT));
			this->menuColors[CheckPenColor] = GetSysColor(COLOR_HIGHLIGHT);//(m_menuStyle==MenuStyleOffice2K3)?RGB(75,75,111):(m_menuStyle==MenuStyleOffice2K7)?RGB(255,171,63):RGB(178,180,191);
			this->menuColors[CheckSelectedPenColor] = GetSysColor(COLOR_HIGHLIGHT);//(m_menuStyle==MenuStyleOffice2K3)?RGB(75,75,111):(m_menuStyle==MenuStyleOffice2K7)?RGB(251,140,60):RGB(152,153,163);
			this->menuColors[CheckDisabledPenColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(147,160,112):RGB(141,141,141);
			return;
		}
		else
		{
			//	Reset to defaults
			for(int i = 0; i < MAX_COLOR; i++)
			{
				this->menuColors[i] = menuDefaultColors[i];
			}
			//this->menuColors[PenColor] = RGB(10,36,106);
			return;
		}
	}

	switch(themeColor)
	{
		case 0x00e55400: // blue
			{
				for(int i = 0; i < MAX_COLOR; i++)
				{
					this->menuColors[i] = menuDefaultColors[i];
				}
			}
			break;
		case 0x00bea3a4:  // silver
			this->menuColors[PenColor] = (m_menuStyle==MenuStyleOffice2K3)?RGB(75,75,111):(m_menuStyle==MenuStyleOffice2K7)?RGB(255,189,105):RGB(169,171,181);
			this->menuColors[SelectedPenColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(126,126,129):RGB(124,124,148);
			this->menuColors[SelectedColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(199,199,202):RGB(255,238,194);
			this->menuColors[CheckedSelectedColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(210,211,216):RGB(254,128,62);
			this->menuColors[CheckedColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(233,234,237):RGB(255,192,111);
			this->menuColors[MenubarGradientStartColor] = (m_menuStyle==MenuStyleOffice2K3)?RGB(255,214,154):RGB(255,223,132);
			this->menuColors[MenubarGradientEndColor] = RGB(255,245,204);
			this->menuColors[MenubarGradientStartColorDown] = (m_menuStyle==MenuStyleOfficeXP)?RGB(229,228,232):RGB(186,185,205);
			this->menuColors[MenubarGradientEndColorDown] = RGB(232,233,241);
			this->menuColors[SeparatorColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(186,186,189):RGB(110,109,143);
			this->menuColors[MenuBckColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(251,250,251):RGB(253,250,255);
			this->menuColors[GradientStartColor] = (m_menuStyle==MenuStyleOfficeXP)?menuDefaultColors[GradientStartColor]:RGB(249,249,255);
			this->menuColors[GradientEndColor] = (m_menuStyle==MenuStyleOfficeXP)?menuDefaultColors[GradientEndColor]:RGB(159,157,185);
			this->menuColors[BitmapBckColor] = (m_menuStyle==MenuStyleOffice2K7)?RGB(239,239,239):RGB(229,228,232);
			this->menuColors[CheckPenColor] = (m_menuStyle==MenuStyleOffice2K3)?RGB(75,75,111):(m_menuStyle==MenuStyleOffice2K7)?RGB(255,171,63):RGB(178,180,191);
			this->menuColors[CheckSelectedPenColor] = (m_menuStyle==MenuStyleOffice2K3)?RGB(75,75,111):(m_menuStyle==MenuStyleOffice2K7)?RGB(251,140,60):RGB(152,153,163);
			this->menuColors[CheckDisabledPenColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(147,160,112):RGB(141,141,141);
			break;
		case 0x0086b8aa:  //olive green
			this->menuColors[PenColor] = (m_menuStyle==MenuStyleOffice2K3)?RGB(63,93,56):(m_menuStyle==MenuStyleOffice2K7)?RGB(255,189,105):RGB(147,160,112);
			this->menuColors[SelectedPenColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(138,134,122):RGB(117,141,94);
			this->menuColors[SelectedColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(206,209,195):RGB(255,238,194);
			this->menuColors[CheckedSelectedColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(201,208,184):RGB(254,128,62);
			this->menuColors[CheckedColor] = (m_menuStyle==MenuStyleOfficeXP)?menuDefaultColors[CheckedColor]:RGB(255,192,111);
			this->menuColors[MenubarGradientStartColor] = (m_menuStyle==MenuStyleOffice2K3)?RGB(255,214,154):(m_menuStyle==MenuStyleOffice2K7)?RGB(255,223,132):RGB(239,237,222);
			this->menuColors[MenubarGradientEndColor] = (m_menuStyle==MenuStyleOfficeXP)?menuDefaultColors[MenubarGradientEndColor]:RGB(255,245,204);
			this->menuColors[MenubarGradientStartColorDown] = (m_menuStyle==MenuStyleOfficeXP)?RGB(239,237,222):RGB(194,206,159);
			this->menuColors[MenubarGradientEndColorDown] = (m_menuStyle==MenuStyleOfficeXP)?menuDefaultColors[MenubarGradientEndColorDown]:RGB(236,240,213);
			this->menuColors[SeparatorColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(197,194,184):RGB(96,128,88);
			this->menuColors[MenuBckColor] = (m_menuStyle==MenuStyleOfficeXP)?menuDefaultColors[MenuBckColor]:RGB(253,250,255);
			this->menuColors[GradientStartColor] = (m_menuStyle==MenuStyleOfficeXP)?menuDefaultColors[GradientStartColor]:RGB(255,255,237);
			this->menuColors[GradientEndColor] = (m_menuStyle==MenuStyleOfficeXP)?menuDefaultColors[GradientEndColor]:RGB(184,199,146);
			this->menuColors[CheckPenColor] = (m_menuStyle==MenuStyleOffice2K3)?RGB(63,93,56):(m_menuStyle==MenuStyleOffice2K7)?RGB(255,171,63):RGB(147,160,112);
			this->menuColors[CheckSelectedPenColor] = (m_menuStyle==MenuStyleOffice2K3)?RGB(63,93,56):(m_menuStyle==MenuStyleOffice2K7)?RGB(251,140,60):RGB(139,152,106);
			this->menuColors[CheckDisabledPenColor] = (m_menuStyle==MenuStyleOfficeXP)?RGB(180,177,163):RGB(141,141,141);
			break;
		default:
			{
				for(int i = 0; i < MAX_COLOR; i++)
				{
					this->menuColors[i] = menuDefaultColors[i];
				}
			}
			break;
	}
	return;
}

void CCoolmenu::GetMenuBarColor2003(COLORREF& color1, COLORREF& color2, BOOL bBackgroundColor /* = TRUE */)
{
// Win95 or WinNT do not support to change the menubarcolor
	if( g_Shell==Win95	|| g_Shell==WinNT4)
	{
		color1 = color2 = GetSysColor(COLOR_MENU);
		return;
	}

	if (IsMenuThemeActive() && m_bReactOnThemeChange )
	{
		COLORREF colorWindow = m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW));
		COLORREF colorCaption = GetSysColor(COLOR_ACTIVECAPTION);

		// get the them from desktop window
		HANDLE hTheme = m_CmGeneral->localOpenThemeData(NULL,L"Window");
		if(hTheme)
		{
			COLORREF color = CLR_NONE;
          // TM_PROP(3821, TMT, FILLCOLORHINT, COLOR)          // hint about fill color used (for custom controls)
			if(m_CmGeneral->localGetThemeColor(hTheme,1,1,3821,&color)==S_OK && (color!=1))
			{
				colorCaption = color;
			}
			//TM_PROP(3805, TMT, EDGEHIGHLIGHTCOLOR, COLOR)     // edge color
			if(m_CmGeneral->localGetThemeColor(hTheme,1,1,3805,&color)==S_OK && (color!=1))
			{
				colorWindow = color;
			}
			m_CmGeneral->localCloseThemeData(hTheme);

			// left side Menubar
			switch(colorCaption)
			{
				case 0x00e55400: // blue
					//color1 = RGB(163,194,245);
					color1 = RGB(158,190,245);
					break;

				case 0x00bea3a4:  // silver
					color1 = RGB(215,215,229);
					break;

				case 0x0086b8aa:  //olive green
					color1 = RGB(218,218,170);
					break;

				default:
				{
					CClientDC myDC(NULL);
					color1 = myDC.GetNearestColor(m_CmGeneral->MidColor(colorWindow,colorCaption));
				}
				break;
			}

			if (bBackgroundColor)
			{
				color2 = m_CmGeneral->LightenColor(110,color1);
			}
			else
			{
				color2 = m_CmGeneral->LightenColor(200,color1);
				color1 = m_CmGeneral->DarkenColor(20,color1);
			}
			return;
		}

		CClientDC myDC(NULL);
		COLORREF nearColor = myDC.GetNearestColor(m_CmGeneral->MidColor(colorWindow,colorCaption));

		// some colorscheme corrections (Andreas Schärer)
		if (nearColor == 15779244) //standartblau
		{ //entspricht (haar-)genau office 2003
//			color1 = RGB(163,194,245);
			color1 = RGB(158,190,245);
		}
		else if (nearColor == 15132390) //standartsilber
		{
			color1 = RGB(215,215,229);
		}
		else if (nearColor == 13425878) //olivgrün
		{
			color1 = RGB(218,218,170);
		}
		else
		{
			color1 = nearColor;
		}

		if (bBackgroundColor)
		{
			color2 = m_CmGeneral->LightenColor(100,color1);
		}
		else
		{
			color2 = m_CmGeneral->LightenColor(200,color1);
			color1 = m_CmGeneral->DarkenColor(20,color1);
		}
	}
	else
	{
		if(m_bReactOnThemeChange)
		{
			if(g_Shell==WinXP)
				color1 = GetSysColor( COLOR_MENUBAR );
			else
				color1 = GetSysColor( COLOR_MENU );
			color2 = ::GetSysColor(COLOR_WINDOW);
			color2 =  m_CmGeneral->GetAlphaBlendColor(color1,color2,220);
		}
		else
		{
			color1 = menuColors[MenubarColor];
			if (bBackgroundColor)
			{
				color2 = m_CmGeneral->LightenColor(110,color1);
			}
			else
			{
				color2 = m_CmGeneral->LightenColor(200,color1);
				color1 = m_CmGeneral->DarkenColor(20,color1);
			}
		}
	}
	return;
}

COLORREF CCoolmenu::GetMenuBarColor2003()
{
	COLORREF colorLeft, colorRight;
	GetMenuBarColor2003( colorLeft, colorRight, TRUE );
	return colorLeft;
}

void CCoolmenu::GetMenuBarColor2007(COLORREF& color1, COLORREF& color2, BOOL bBackgroundColor /* = TRUE */)
{
// Win95 or WinNT do not support to change the menubarcolor
	if( g_Shell==Win95	|| g_Shell==WinNT4)
	{
		color1 = color2 = GetSysColor(COLOR_MENU);
		return;
	}

	if (IsMenuThemeActive() && m_bReactOnThemeChange )
	{
		COLORREF colorWindow = m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW));
		COLORREF colorCaption = GetSysColor(COLOR_ACTIVECAPTION);

		// get the them from desktop window
		HANDLE hTheme = m_CmGeneral->localOpenThemeData(NULL,L"Window");
		if(hTheme)
		{
			COLORREF color = CLR_NONE;
          // TM_PROP(3821, TMT, FILLCOLORHINT, COLOR)          // hint about fill color used (for custom controls)
			if(m_CmGeneral->localGetThemeColor(hTheme,1,1,3821,&color)==S_OK && (color!=1))
			{
				colorCaption = color;
			}
			//TM_PROP(3805, TMT, EDGEHIGHLIGHTCOLOR, COLOR)     // edge color
			if(m_CmGeneral->localGetThemeColor(hTheme,1,1,3805,&color)==S_OK && (color!=1))
			{
				colorWindow = color;
			}
			m_CmGeneral->localCloseThemeData(hTheme);

			// left side Menubar
			switch(colorCaption)
			{
				case 0x00e55400: // blue
					//color1 = RGB(163,194,245);
					color1 = RGB(191,219,255);
					color2 = color1;
					break;

				case 0x00bea3a4:  // silver
					color1 = RGB(215,215,229);
					color2 = RGB(243,243,247);
					break;

				case 0x0086b8aa:  //olive green
					color1 = RGB(217,217,167);
					color2 = RGB(242,240,228);
					break;

				default:
				{
					CClientDC myDC(NULL);
					color1 = myDC.GetNearestColor(m_CmGeneral->MidColor(colorWindow,colorCaption));
					color2 = color1;
				}
				return;
			}

			/*if (bBackgroundColor)
			{
				color2 = m_CmGeneral->LightenColor(110,color1);
			}
			else
			{
				color2 = m_CmGeneral->LightenColor(200,color1);
				color1 = m_CmGeneral->DarkenColor(20,color1);
			}*/
			return;
		}

		CClientDC myDC(NULL);
		COLORREF nearColor = myDC.GetNearestColor(m_CmGeneral->MidColor(colorWindow,colorCaption));

		// some colorscheme corrections (Andreas Schärer)
		if (nearColor == 15779244) //standartblau
		{ //entspricht (haar-)genau office 2003
//			color1 = RGB(163,194,245);
			color1 = RGB(191,219,255);
		}
		else if (nearColor == 15132390) //standartsilber
		{
			color1 = RGB(215,215,229);
			color2 = RGB(243,243,247);
		}
		else if (nearColor == 13425878) //olivgrün
		{
			color1 = RGB(217,217,167);
			color2 = RGB(242,240,228);
		}
		else
		{
			color1 = nearColor;
			color2 = color1;
		}

		/*if (bBackgroundColor)
		{
			color2 = m_CmGeneral->LightenColor(100,color1);
		}
		else
		{
			color2 = m_CmGeneral->LightenColor(200,color1);
			color1 = m_CmGeneral->DarkenColor(20,color1);
		}*/
	}
	else
	{
		if(m_bReactOnThemeChange)
		{
			if(g_Shell==WinXP)
				color1 = GetSysColor( COLOR_MENUBAR );
			else
				color1 = GetSysColor( COLOR_MENU );
			color2 = color1;
		}
		else
		{
			color1 = menuColors[MenubarColor];
			if (bBackgroundColor)
			{
				color2 = color1;
			}
		}
	}
	return;
}

COLORREF CCoolmenu::GetMenuBarColor2007()
{
	COLORREF colorLeft, colorRight;
	GetMenuBarColor2007( colorLeft, colorRight, TRUE );
	return colorLeft;
}

COLORREF CCoolmenu::GetMenuBarColor(HMENU hMenu)
{
	if(hMenu!=NULL)
	{
		MENUINFO menuInfo = {0};
		menuInfo.cbSize = sizeof(menuInfo);
		menuInfo.fMask = MIM_BACKGROUND;

		if(m_CmGeneral->GetMenuInfo(hMenu,&menuInfo) && menuInfo.hbrBack)
		{
			LOGBRUSH logBrush;
			if(GetObject(menuInfo.hbrBack,sizeof(LOGBRUSH),&logBrush))
			{
				return logBrush.lbColor;
			}
		}
	}
	if(g_Shell==WinXP)
	{
		BOOL bFlatMenu = FALSE;
		if((SystemParametersInfo(SPI_GETFLATMENU,0,&bFlatMenu,0) && bFlatMenu==TRUE))
		{
			return GetSysColor(COLOR_MENUBAR);
		}
	}
	return GetSysColor(COLOR_MENU);
}

CBrush* CCoolmenu::GetMenuBarBrush()
{
	// The brushes will be destroyed at program-end => Not a memory-leak
	static CNewBrushList brushList;
	static CNewBrush* lastBrush = NULL;

	COLORREF menuBarColor;
	COLORREF menuBarColor2;		//JUS

	if ( m_menuStyle == MenuStyleOffice2K3 && m_bCoolMenubar )
	{
      GetMenuBarColor2003( menuBarColor, menuBarColor2, TRUE );
	}
	else
	{
		if ( m_menuStyle == MenuStyleOffice2K7 && m_bCoolMenubar )
		{
	      GetMenuBarColor2007( menuBarColor, menuBarColor2, TRUE );
			//menuBarColor = RGB(191,219,255);
			//menuBarColor2 = RGB(191,219,255);
		}
		else
		{
			if(g_Shell==WinXP)
			{
				BOOL bFlatMenu = FALSE;
				if((SystemParametersInfo(SPI_GETFLATMENU,0,&bFlatMenu,0) && bFlatMenu==TRUE))
					menuBarColor = GetSysColor(COLOR_MENUBAR);
				else
				{
					if(m_bReactOnThemeChange)
						menuBarColor = GetSysColor(COLOR_MENU);
					else
						menuBarColor = (m_menuStyle == MenuStyleOffice2K || m_menuStyle == MenuStyleNormal)?GetSysColor(COLOR_MENU):menuColors[MenubarColor];
				}
			}
			else
			{
				if(m_bReactOnThemeChange)
					menuBarColor = GetSysColor(COLOR_MENU);
				else
					menuBarColor = (m_menuStyle == MenuStyleOffice2K || m_menuStyle == MenuStyleNormal)?GetSysColor(COLOR_MENU):menuColors[MenubarColor];
			}
			menuBarColor2 = menuBarColor;
		}
	}
	// check if the last brush the one which we want
  // check if the last brush the one which we want
	if(lastBrush!=NULL &&
		lastBrush->m_BarColor==menuBarColor &&
		lastBrush->m_BarColor2==menuBarColor2)
	{
		return lastBrush;
	}

	// Check if the brush is allready created
	POSITION pos = brushList.GetHeadPosition();
	while (pos)
	{
		lastBrush = (CNewBrush*)brushList.GetNext(pos);
		if(lastBrush!=NULL &&
			lastBrush->m_BarColor==menuBarColor &&
			lastBrush->m_BarColor2==menuBarColor2)
		{
			return lastBrush;
		}
	}
	// create a new one and insert into the list
	brushList.AddHead(lastBrush = new CNewBrush( 0, menuBarColor, menuBarColor2 ) );
	return lastBrush;
}

void CCoolmenu::UpdateMenuBarColor(HMENU hMenu)
{
	CBrush* pBrush = GetMenuBarBrush();

	if(!hMenu) return;
	// for WindowsBlind activating it's better for not to change menubar background
	if(!pBrush /* || (pIsThemeActive && pIsThemeActive()) */ )
	{
		// menubackground hasn't been set
		return;
	}

	MENUINFO menuInfo = {0};
	menuInfo.cbSize = sizeof(menuInfo);
	menuInfo.hbrBack = *pBrush;
	menuInfo.fMask = MIM_BACKGROUND;

	// Change color only for CNewMenu and derived classes
	m_CmGeneral->SetMenuInfo(hMenu,&menuInfo);

	RedrawWindow(m_pWndHooked,NULL,0,RDW_FRAME|RDW_INVALIDATE|RDW_UPDATENOW|RDW_ERASENOW);
}

void CCoolmenu::MenuDrawText(HDC hDC ,LPCTSTR lpString,int nCount,LPRECT lpRect,UINT uFormat)
{
	if(nCount==-1)
	{
		nCount = lpString?(int)_tcslen(lpString):0;
	}
	LOGFONT logfont = {0};
	if(!GetObject(GetCurrentObject(hDC, OBJ_FONT),sizeof(logfont),&logfont))
	{
		logfont.lfOrientation = 0;
	}
	TCHAR* pBuffer = (TCHAR*)_alloca(nCount*sizeof(TCHAR)+2);
	_tcsncpy(pBuffer,lpString,nCount);
	pBuffer[nCount] = 0;

	UINT oldAlign =  GetTextAlign(hDC);

	const int nBorder=4;
	HGDIOBJ hOldFont = NULL;
	CFont TempFont;

	TEXTMETRIC textMetric;
	if (!GetTextMetrics(hDC, &textMetric))
	{
		textMetric.tmOverhang = 0;
		textMetric.tmAscent = 0;
	}
	else if ((textMetric.tmPitchAndFamily&(TMPF_VECTOR|TMPF_TRUETYPE))==0 )
	{
		// we have a bitmapfont it is not possible to rotate
		if(logfont.lfOrientation || logfont.lfEscapement)
		{
		hOldFont = GetCurrentObject(hDC,OBJ_FONT);
		_tcscpy(logfont.lfFaceName,_T("Arial"));
		// we need a truetype font for rotation
		logfont.lfOutPrecision = OUT_TT_ONLY_PRECIS;
		// the font will be destroyed at the end by destructor
		TempFont.CreateFontIndirect(&logfont);
		// we select at the end the old font
		SelectObject(hDC,TempFont);
		GetTextMetrics(hDC, &textMetric);
		}
	}

	SetTextAlign (hDC,TA_LEFT|TA_TOP|TA_UPDATECP);
	CPoint pos(lpRect->left,lpRect->top);

	if( (uFormat&DT_VCENTER) &&lpRect)
	{
		switch(logfont.lfOrientation%3600)
		{
			default:
			case 0:
				logfont.lfOrientation = 0;
				pos.y = (lpRect->top+lpRect->bottom-textMetric.tmHeight)/2;
				pos.x = lpRect->left + nBorder;
				break;

			case 1800:
			case -1800:
				logfont.lfOrientation = 1800;
				pos.y = (lpRect->top+textMetric.tmHeight +lpRect->bottom)/2;
				pos.x = lpRect->right - nBorder;
				break;

			case 900:
			case -2700:
				logfont.lfOrientation = 900;
				pos.x = (lpRect->left+lpRect->right-textMetric.tmHeight)/2;
				pos.y = lpRect->bottom - nBorder;
				break;

			case -900:
			case 2700:
				logfont.lfOrientation = 2700;
				pos.x = (lpRect->left+lpRect->right+textMetric.tmHeight)/2;
				pos.y = lpRect->top + nBorder;
				break;
		}
	}

	CPoint oldPos;
	MoveToEx(hDC,pos.x,pos.y,&oldPos);

	while(nCount)
	{
		TCHAR *pTemp =_tcsstr(pBuffer,_T("&"));
		if(pTemp)
		{
			// we found &
			if(*(pTemp+1)==_T('&'))
			{
				// the different is in character unicode and byte works
				int nTempCount = DWORD(pTemp-pBuffer)+1;
				ExtTextOut(hDC,pos.x,pos.y,ETO_CLIPPED,lpRect,pBuffer,nTempCount,NULL);
				nCount -= nTempCount+1;
				pBuffer = pTemp+2;
			}
			else
			{
				// draw underline the different is in character unicode and byte works
				int nTempCount = DWORD(pTemp-pBuffer);
				ExtTextOut(hDC,pos.x,pos.y,ETO_CLIPPED,lpRect,pBuffer,nTempCount,NULL);
				nCount -= nTempCount+1;
				pBuffer = pTemp+1;
				if(!(uFormat&DT_HIDEPREFIX) )
				{
					CSize size;
					GetTextExtentPoint(hDC, pTemp+1, 1, &size);
					GetCurrentPositionEx(hDC,&pos);
					COLORREF oldColor = SetBkColor(hDC, GetTextColor(hDC));
					LONG cx = size.cx - textMetric.tmOverhang / 2;
					LONG nTop;
					CRect rc;
					switch(logfont.lfOrientation)
					{
						case 0:
							// Get height of text so that underline is at bottom.
							nTop = pos.y + textMetric.tmAscent + 1;
							// Draw the underline using the foreground color.
							rc.SetRect(pos.x, nTop, pos.x+cx, nTop+1);
							ExtTextOut(hDC, pos.x, nTop, ETO_OPAQUE, &rc, _T(""), 0, NULL);
							break;

						case 1800:
							// Get height of text so that underline is at bottom.
							nTop = pos.y -(textMetric.tmAscent + 1);
							// Draw the underline using the foreground color.
							rc.SetRect(pos.x-cx, nTop-1, pos.x, nTop);
							ExtTextOut(hDC, pos.x, nTop, ETO_OPAQUE, &rc, _T(""), 0, NULL);
							break;

						case 900:
							// draw up
							// Get height of text so that underline is at bottom.
							nTop = pos.x + (textMetric.tmAscent + 1);
							// Draw the underline using the foreground color.
							rc.SetRect(nTop, pos.y-cx, nTop+1, pos.y);
							ExtTextOut(hDC, nTop, pos.y, ETO_OPAQUE, &rc, _T(""), 0, NULL);
							break;

						case 2700:
							// draw down
							// Get height of text so that underline is at bottom.
							nTop = pos.x -(textMetric.tmAscent + 1);
							// Draw the underline using the foreground color.
							rc.SetRect(nTop-1, pos.y, nTop, pos.y+cx);
							ExtTextOut(hDC, nTop, pos.y, ETO_OPAQUE, &rc, _T(""), 0, NULL);
							break;
					}
					SetBkColor(hDC, oldColor);
					// corect the actual drawingpoint
					MoveToEx(hDC,pos.x,pos.y,NULL);
				}
			}
		}
		else
		{
			// draw the rest of the string
			ExtTextOut(hDC,pos.x,pos.y,ETO_CLIPPED,lpRect,pBuffer,nCount,NULL);
			break;
		}
	}
	// restore old point
	MoveToEx(hDC,oldPos.x,oldPos.y,NULL);
	// restore old align
	SetTextAlign(hDC,oldAlign);

	if(hOldFont!=NULL)
	{
		SelectObject(hDC,hOldFont);
	}
}

void CCoolmenu::DrawSpecialCharStyle(CDC* pDC, LPCRECT pRect, TCHAR Sign, DWORD dwStyle)
{
	COLORREF oldColor;
	if( (dwStyle&ODS_GRAYED) || (dwStyle&ODS_INACTIVE))
	{
		oldColor = pDC->SetTextColor((m_menuStyle==MenuStyleOffice2K3)?RGB(107,154,232):(::GetSysColor(COLOR_GRAYTEXT)));
	}
	else
	{
		oldColor = pDC->SetTextColor(::GetSysColor(COLOR_MENUTEXT));
	}

	DrawSpecialChar(pDC,pRect,Sign,(dwStyle&ODS_DEFAULT) ? TRUE : FALSE);

	pDC->SetTextColor(oldColor);
}

void CCoolmenu::DrawSpecialChar(CDC* pDC, LPCRECT pRect, TCHAR Sign, BOOL bBold)
{
	//  48 Min
	//  49 Max
	//  50 Restore
	//	 52 Submenu
	//  98 Checkmark
	// 105 Bullet
	// 114 Close

	CFont MyFont;
	LOGFONT logfont;

	CRect rect(pRect) ;
	rect.DeflateRect(2,2);

	logfont.lfHeight = -rect.Height();
	logfont.lfWidth = 0;
	logfont.lfEscapement = 0;
	logfont.lfOrientation = 0;
	logfont.lfWeight = (bBold) ? FW_BOLD:FW_NORMAL;
	logfont.lfItalic = FALSE;
	logfont.lfUnderline = FALSE;
	logfont.lfStrikeOut = FALSE;
	logfont.lfCharSet = DEFAULT_CHARSET;
	logfont.lfOutPrecision = OUT_DEFAULT_PRECIS;
	logfont.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	logfont.lfQuality = DEFAULT_QUALITY;
	logfont.lfPitchAndFamily = DEFAULT_PITCH;

	_tcscpy(logfont.lfFaceName,_T("Marlett"));

	MyFont.CreateFontIndirect (&logfont);

	CFont* pOldFont = pDC->SelectObject (&MyFont);

	if(Sign==52)
	{
		int OldMode = pDC->SetBkMode(OPAQUE);
		pDC->DrawText (&Sign,1,rect,DT_RIGHT|DT_SINGLELINE);
		pDC->SetBkMode(OldMode);
	}
	else
	{
		int OldMode = pDC->SetBkMode(TRANSPARENT);
		pDC->DrawText (&Sign,1,rect,DT_CENTER|DT_SINGLELINE);
		pDC->SetBkMode(OldMode);
	}

	pDC->SelectObject(pOldFont);
}

/*
	Function:		GetButtonIndex
	Returns:			int
	Description:	Returns the iconindex within the ImageList for a menuname
*/
int CCoolmenu::GetButtonIndex( LPCTSTR sName )
{
	int val;
	if (!m_imageIdList->Lookup( sName, (void*&)val ) ) return -1;
	return val;
}

/*
	Function:		GetMenuItemTextBold
	Returns:			BOOL
	Description:	Returns if menuitem should be drawn bold
*/
BOOL CCoolmenu::GetMenuItemTextBold( LPCTSTR sName )
{
	BOOL val;
	if (!m_menuItemTextBoldList->Lookup( sName, (void*&)val ) ) return FALSE;
	return val;
}

/*
	Function:		GetMenuItemTextItalic
	Returns:			BOOL
	Description:	Returns if menuitem should be drawn italic
*/
BOOL CCoolmenu::GetMenuItemTextItalic( LPCTSTR sName )
{
	BOOL val = FALSE;
	if (!m_menuItemTextItalicList->Lookup( sName, (void*&)val ) ) return FALSE;
	return val;
}

/*
	Function:		GetMenuItemTextUnderline
	Returns:			BOOL
	Description:	Returns if menuitem should be drawn underlined
*/
BOOL CCoolmenu::GetMenuItemTextUnderline( LPCTSTR sName )
{
	BOOL val;
	if (!m_menuItemTextUnderlineList->Lookup( sName, (void*&)val ) ) return FALSE;
	return val;
}

/*
	Function:		GetMenuItemTextColor
	Returns:			COLORREF
	Description:	Returns the iconindex within the ImageList for a menuname
*/
COLORREF CCoolmenu::GetMenuItemTextColor( LPCTSTR sName )
{
	COLORREF val;
	if (!m_menuItemTextColorList->Lookup( sName, (void*&)val ) ) return -1;
	return val;
}

/*
	Function:		ConvertMenu
	Returns:			-
	Description:	Converts menu identified by pMenu to ownerdraw.
*/
BOOL CCoolmenu::ConvertMenu( HMENU pMenu, UINT nIndex, BOOL bSysMenu, BOOL bShowButtons, BOOL bMenubar )
{
	CString sItemName;
	if(pMenu==m_hMenu && !bShowButtons && ( m_menuStyle == MenuStyleOffice2K3 || m_menuStyle == MenuStyleOffice2K7 || m_menuStyle == MenuStyleOfficeXP ) && m_bCoolMenubar) return FALSE;
	int nItem = (int)GetMenuItemCount( pMenu );
	for (int i = 0; i < nItem; i++) {	// loop over each item in menu

		// get menu item info
		TCHAR itemname[256];
		CMenuItemInfo info;
		info.fMask = MIIM_DATA | MIIM_ID | MIIM_TYPE | MIIM_STATE | MIIM_SUBMENU;
		info.dwTypeData = itemname;
		info.cch = 255;
		::GetMenuItemInfo( pMenu, i, TRUE, &info);
		CMyItemData* pmd = (CMyItemData*)info.dwItemData;
		BOOL bPmdNotMine = FALSE;
		if (pmd && !IsMyItemData(pmd)) {
			continue;
		}

		if (bSysMenu && info.wID >= 0xF000) {
			continue; // don't do for system menu commands
		} else {
			if ((info.wID&0xffff)>=SC_SIZE && (info.wID&0xffff)<=SC_HOTKEY )
				continue;
		}

		// now that I have the info, I will modify it
		info.fMask = 0;	// assume nothing to change

		if (bShowButtons) {

			// I'm showing buttons: convert to owner-draw

			if (!(info.fType & MFT_OWNERDRAW)) {
				// If not already owner-draw, make it so. NOTE: If app calls
				// pCmdUI->SetText to change the text of a menu item, MFC will
				// turn the item to MFT_STRING. So I must set it back to
				// MFT_OWNERDRAW again. In this case, the menu item data (pmd)
				// will still be there.
				//
				info.fType |= MFT_OWNERDRAW;
				info.fMask |= MIIM_TYPE;
				if (!pmd)
				{									// if no item data:
					pmd = GetNewItemData();				// get new CMyItemData
					ASSERT(pmd);							//   (I hope)
					pmd->fType = info.fType;			//   handy when drawing
					// Build 102, +1 added for buffersize, as stated in MSDN
					LPSTR m_szBuffer = new CHAR [info.cch+1];
					CString menuText( info.dwTypeData, info.cch );
					strcpy( m_szBuffer, menuText );
					pmd->iButton = GetButtonIndex( (LPCTSTR)m_szBuffer );
					pmd->bBold = GetMenuItemTextBold( (LPCTSTR)m_szBuffer );
					pmd->bUnderline = GetMenuItemTextUnderline( (LPCTSTR)m_szBuffer );
					pmd->bItalic = GetMenuItemTextItalic( (LPCTSTR)m_szBuffer );
					pmd->clrText = GetMenuItemTextColor( (LPCTSTR)m_szBuffer );
					delete [] m_szBuffer;
					info.dwItemData = (DWORD)pmd;		//   set in menu item data
					info.fMask |= MIIM_DATA;			//   set item data
				}
				pmd->text = info.dwTypeData;			// copy menu item string
				pmd->bMenuBar = bMenubar;
				pmd->fState = info.fState;
				pmd->nID = info.wID;
				pmd->iItem = i;
				pmd->bHasSubmenu = (info.hSubMenu>0);
			}

			// now add the menu to list of "converted" menus
			ASSERT(pMenu);
			if (!m_menuList.Find(pMenu))
				m_menuList.AddHead(pMenu);

		} else {

			// no buttons -- I'm converting to strings
			if (info.fType & MFT_OWNERDRAW) {	// if ownerdraw:
				info.fType &= ~MFT_OWNERDRAW;		//   turn it off
				info.fMask |= MIIM_TYPE;			//   change item type
				ASSERT(pmd);							//   sanity check
				sItemName = pmd->text;				//   save name before deleting pmd
			} else if (info.fType & MFT_STRING)	// otherwise:
				sItemName = info.dwTypeData;		//   use name from MENUITEMINFO

			if (pmd) {
				// NOTE: pmd (item data) could still be left hanging around even
				// if MFT_OWNERDRAW is not set, in case mentioned above where app
				// calls pCmdUI->SetText to set text of item and MFC sets the type
				// to MFT_STRING.
				//
				info.dwItemData = NULL;				// item data is NULL
				info.fMask |= MIIM_DATA;			// change it
				DeleteItemData(pmd);
			}

			if (info.fMask & MIIM_TYPE) {
				// if setting name, copy name from CString to buffer and set cch
				_tcsncpy(itemname, sItemName, countof(itemname));
				info.dwTypeData = itemname;
				info.cch = sItemName.GetLength();
			}
		}

		// if after all the above, there is anything to change, change it
		if (info.fMask) {
			SetMenuItemInfo(pMenu, i, TRUE, &info);
		}
	}
	return TRUE;
}

void CCoolmenu::SetOpenMenu( HMENU hMenu, UINT nIndex, BOOL bInit, BOOL bAdd )
{
	CMenuItemInfo info;
	if( nIndex < 0 ) return;
	info.fMask = MIIM_DATA;
	::GetMenuItemInfo( hMenu, nIndex, TRUE, &info);
	CMyItemData* pmd = (CMyItemData*)info.dwItemData;
	if( !IsMyItemData(pmd) ) return;
	if(bInit)
	{
		pmd->m_dwOpenMenu = 0;
	}
	else
	{
		if(bAdd)
			pmd->m_dwOpenMenu += 1;
		else
			pmd->m_dwOpenMenu -= 1;
	}
	SetMenuItemInfo( hMenu, nIndex, TRUE, &info );
}

void CCoolmenu::OnInitMenu( HMENU hMenu )
{
	CMenuItemInfo info;
	int i = GetMenuItemCount( hMenu );
	for(int j=0;j<i;j++)
	{
		SetOpenMenu( hMenu, j, TRUE, FALSE );
	}
}

void CCoolmenu::OnInitMenuPopup( HMENU pPopupMenu, UINT nIndex )
{
	if( GetSubMenu( m_hMenu, nIndex ) == pPopupMenu )
	{
		CMenuItemInfo info;
		info.fMask = MIIM_DATA;
		::GetMenuItemInfo( m_hMenu, nIndex, TRUE, &info);
		CMyItemData* pmd = (CMyItemData*)info.dwItemData;
		if( !IsMyItemData(pmd) ) return;
		if(pmd && pmd->bMenuBar)
		{
			if(m_LastMenubarIndex != -1) SetOpenMenu( m_hMenu, m_LastMenubarIndex, TRUE, FALSE );
			m_LastMenubarIndex = nIndex;
			if(pmd->bHasSubmenu) pmd->m_dwOpenMenu += 1;
			SetMenuItemInfo( m_hMenu, nIndex, TRUE, &info );
			if (pmd->m_dwOpenMenu == 1)
			{
				// Redraw the menubar for the shade
				CRect rect = m_LastActiveMenuRect;
				if(!rect.IsRectEmpty())
				{
					rect.InflateRect(0,0,10,10);
					CPoint Point(0,0);
					ClientToScreen(m_pWndHooked,&Point);
					rect.OffsetRect(-Point);
					RedrawWindow(m_pWndHooked,rect,0,RDW_FRAME|RDW_INVALIDATE);
				}
			}
			return;
		}
	}
	if( m_LastMenubarIndex != -1 && m_hMenu )
	{
		SetOpenMenu( m_hMenu, m_LastMenubarIndex, FALSE, TRUE );
	}
}

void CCoolmenu::OnUninitMenuPopup()
{
	if(m_hMenu && m_LastMenubarIndex != -1)
	{
		CMenuItemInfo info;
		info.fMask = MIIM_DATA;
		::GetMenuItemInfo( m_hMenu, m_LastMenubarIndex, TRUE, &info);
		CMyItemData* pmd = (CMyItemData*)info.dwItemData;
		if( !IsMyItemData(pmd) ) return;
		pmd->m_dwOpenMenu -= 1;
		if(pmd->m_dwOpenMenu < 0) pmd->m_dwOpenMenu = 0;
		SetMenuItemInfo( m_hMenu, m_LastMenubarIndex, TRUE, &info );
		CRect rect = m_LastActiveMenuRect;
		if(!rect.IsRectEmpty() && pmd->m_dwOpenMenu == 0)
		{
			rect.InflateRect(0,0,10,10);
			CPoint Point(0,0);
			ClientToScreen(m_pWndHooked,&Point);
			rect.OffsetRect(-Point);
			RedrawWindow(m_pWndHooked,rect,0,RDW_FRAME|RDW_INVALIDATE);
		}
	}
}

/*
	Function:		OnMeasureItem
	Returns:			BOOL
	Description:	Sets width and height of menuitem in lpms
*/
BOOL CCoolmenu::OnMeasureItem( LPMEASUREITEMSTRUCT lpms )
{
	CMyItemData* pmd = (CMyItemData*)lpms->itemData;

	if ( !IsMyItemData(pmd) ) return FALSE;
	if ( lpms->CtlID != 0 ) return FALSE;
	if ( lpms->CtlType == ODT_MENU )
	{
		if ( pmd->fType & MFT_SEPARATOR )
		{
			lpms->itemHeight = 5;
			lpms->itemWidth = 3;
		}
		else
		{
			CFont fontMenu;
			LOGFONT logFontMenu;

			#ifdef _NEW_MENU_USER_FONT
				logFontMenu =  MENU_USER_FONT;
			#else
				NONCLIENTMETRICS nm = {0};
			nm.cbSize = sizeof (nm);
			VERIFY (SystemParametersInfo(SPI_GETNONCLIENTMETRICS,nm.cbSize,&nm,0));
			logFontMenu =  nm.lfMenuFont;
			#endif

			// Default selection?
			CMenuItemInfo info;
			info.fMask = MIIM_STATE;
			::GetMenuItemInfo( m_hMenu, pmd->nID, FALSE, &info);

			if( ( info.fState & MFS_DEFAULT ) == MFS_DEFAULT || pmd->bBold )
			{
			// Make the font bold
				logFontMenu.lfWeight = FW_BOLD;
			}

			fontMenu.CreateFontIndirect (&logFontMenu);

			// DC of the desktop
			CClientDC myDC(NULL);

			// Select menu font in...
			CFont* pOldFont = myDC.SelectObject (&fontMenu);
			//Get pointer to text SK
			CString itemText = pmd->text;

			if(pmd->bMenuBar)
			{
				LPSTR menuText = new CHAR[1024];
				GetMenuString(m_hMenu, pmd->iItem, menuText, 1024, MF_BYPOSITION);
				itemText = menuText;
				delete [] menuText;
			}


			SIZE size = {0,0};
			VERIFY(::GetTextExtentPoint32(myDC.m_hDC,itemText,itemText.GetLength(),&size));
			// Select old font in
			myDC.SelectObject(pOldFont);


			// Set width and height:
			if(pmd->bMenuBar)
			{
				if(itemText.Find(_T("&"))>=0)
				{
					lpms->itemWidth = size.cx - GetSystemMetrics(SM_CXMENUCHECK)/2;
				}
				else
				{
					lpms->itemWidth = size.cx;
				}
			}
			else
			{
				lpms->itemWidth = IMGWIDTH + size.cx + IMGWIDTH + 2;
			}

			int temp = GetSystemMetrics(SM_CYMENU);
			lpms->itemHeight = (temp>(IMGHEIGHT+4)) ? temp : (IMGHEIGHT+4);
			if(lpms->itemHeight<((UINT)size.cy) )
			{
				lpms->itemHeight=((UINT)size.cy);
 			}
			if (m_menuStyle==MenuStyleOffice2K3 || m_menuStyle == MenuStyleOffice2K7 || m_menuStyle == MenuStyleOfficeXP)
				lpms->itemHeight+=2;
		}
	}
	return TRUE;
}

/*
	Function:		OnMenuSelect
	Returns:			-
	Description:	If menu is closed it's converted back.
*/
void CCoolmenu::OnMenuSelect( UINT nItemID, UINT nFlags, HMENU hSysMenu )
{
	CPtrList	tempList;
	m_bInitMenu = FALSE;
	if (hSysMenu==NULL && nFlags==0xFFFF) {
		m_bNoPaint = FALSE;
		m_LastMenubarIndex = -1;
		m_bMenuSet = FALSE;
		m_bMouseSelect = FALSE;
		m_bKeyDown = FALSE;
		while (!m_menuList.IsEmpty()) {
			HMENU hMenu = (HMENU)m_menuList.RemoveHead();
 			if (!ConvertMenu( hMenu, 0, FALSE, FALSE, FALSE ))
				tempList.AddHead(hMenu);
		}
		while (!tempList.IsEmpty()) {
			m_menuList.AddHead((HMENU)tempList.RemoveHead());
		}
	}
	else
	{
		m_bMenuSet = TRUE;
		if( nFlags & MF_POPUP )
			m_bInitMenu = TRUE;
	}
}

/*
	Function:		OnMenuChar
	Returns:			LRESULT
	Description:	Find selected menuitem based on pressed key and execute it.
*/
LRESULT CCoolmenu::OnMenuChar( UINT nChar, UINT nFlags, CMenu* pMenu )
{
	ASSERT_VALID(pMenu);

	if ( !( nFlags & MF_POPUP ) )
		return 0; // not popup

	UINT iCurrentItem = (UINT)-1; // guaranteed higher than any command ID
	CUIntArray arItemsMatched;		// items that match the character typed

	UINT nItem = pMenu->GetMenuItemCount();
	for ( UINT i=0; i< nItem; i++ )
	{
		// get menu info
		CMenuItemInfo info;
		info.fMask = MIIM_DATA | MIIM_TYPE | MIIM_STATE;
		::GetMenuItemInfo( *pMenu, i, TRUE, &info );

		CMyItemData* pmd = (CMyItemData*)info.dwItemData;
		if ( ( info.fType & MFT_OWNERDRAW ) && ( pmd && IsMyItemData( pmd ) ) )
		{
			CString& text = pmd->text;
			int iAmpersand = text.Find('&');
			if (iAmpersand >=0 && toupper( nChar ) == toupper( text[iAmpersand+1] ) )
				arItemsMatched.Add(i);
		}
		if ( info.fState & MFS_HILITE )
			iCurrentItem = i; // note index of current item
	}

	// arItemsMatched now contains indexes of items that match the char typed.
	//
	//   * if none: beep
	//   * if one:  execute it
	//   * if more than one: hilite next
	//
	UINT nFound = arItemsMatched.GetSize();
	if (nFound == 0)
		return 0;

	else if (nFound==1)
		return MAKELONG( arItemsMatched[0], MNC_EXECUTE );

	// more than one found--return 1st one past current selected item;
	UINT iSelect = 0;
	for (i=0; i < nFound; i++) {
		if ( arItemsMatched[i] > iCurrentItem )
		{
			iSelect = i;
			break;
		}
	}
	return MAKELONG( arItemsMatched[iSelect], MNC_SELECT );
}

/*
	Function:		OnDrawItem
	Returns:			BOOL
	Description:	Calls Draw or XpDraw based on value of m_bXpStyle
*/
BOOL CCoolmenu::OnDrawItem( LPDRAWITEMSTRUCT lpds )
{
	if ( lpds->CtlType == ODT_MENU )
	{
		switch( m_menuStyle )
		{
		case MenuStyleOfficeXP:
			return DrawXP( lpds );
			break;
		case MenuStyleOffice2K3:
			return Draw2K3( lpds );
			break;
		case MenuStyleOffice2K7:
			return Draw2K7( lpds );
			break;
		default:
			return DrawNormal( lpds );
		}
	}
	return FALSE;
}

/*
	Function:		Draw
	Returns:			-
	Description:	Draws menuitem with old Office look
*/
BOOL CCoolmenu::DrawNormal( LPDRAWITEMSTRUCT lpds )
{
	CMyItemData* pmd = (CMyItemData*)lpds->itemData;

	if (lpds->CtlType != ODT_MENU || !IsMyItemData(pmd))
		return FALSE; // not handled by me

	CDC* pDC;

	pDC = CDC::FromHandle(lpds->hDC);

	RECT rcItem = lpds->rcItem;

	if ( pmd->fType & MFT_SEPARATOR )
	{
		// draw separator
		RECT rc = rcItem;
		rc.top += ( rc.bottom - rc.top )>>1;						// vertical center
		pDC->DrawEdge( &rc, EDGE_ETCHED, BF_TOP );		// draw separator line
	}
	else
	{													// not a separator
		BOOL bDisabled = lpds->itemState & ODS_GRAYED;
		BOOL bSelected = lpds->itemState & ODS_SELECTED;
		BOOL bChecked  = lpds->itemState & ODS_CHECKED;
		BOOL bHaveButn = ( pmd->iButton >= 0 );

		// Done with button, now paint text. First do background if needed.
		COLORREF colorBG;
		RECT rcBG = rcItem;							// whole rectangle

		colorBG = GetSysColor(bSelected ? COLOR_HIGHLIGHT : COLOR_MENU );

		// compute text rectangle and colors
		if (bHaveButn || bChecked )									// if there's a button:
			rcBG.left += ( rcBG.bottom - rcBG.top ) + CXGAP;			//  don't paint over it!
		if ( bSelected || lpds->itemAction==ODA_SELECT ) {
			// selected or selection state changed: paint text background
			FillRectCM(pDC->m_hDC, rcBG, colorBG);
		}

		pDC->SetBkMode(TRANSPARENT);			 // paint transparent text
		COLORREF colorText;
		colorText = GetSysColor(bDisabled && !bSelected ?  COLOR_GRAYTEXT :
			bSelected ? COLOR_HIGHLIGHTTEXT : COLOR_MENUTEXT);

		// Now paint menu item text.	No need to select font,
		// because windows sets it up before sending WM_DRAWITEM
		//
		CRect rectText;
		rectText.CopyRect(&lpds->rcItem);
		rectText.left += ( rcBG.bottom - rcBG.top ) + TEXTPADDING + 1;
		DrawMenuText( pDC->m_hDC, rectText, colorText, pmd->text ); // finally!

		// Paint button, or blank if none

		CRect rect;
		rect.CopyRect(&lpds->rcItem);
		CRect RectSel(lpds->rcItem);

      CRect IconRect( rect.left, rect.top, rect.left + ( rect.Height() ), rect.top + ( rect.Height() ) );

      CPoint ptImage = IconRect.TopLeft();

		int iButton = pmd->iButton;

		if (bHaveButn) {

			// this item has a button!

			// compute point to start drawing
			SIZE sz;
			sz.cx = (IconRect.Width()) - IMGWIDTH;
			sz.cy = (IconRect.Height()) - IMGHEIGHT;

			sz.cx >>= 1;
			sz.cy >>= 1;
			POINT p;
			p.x = IconRect.left + sz.cx;
			p.y = IconRect.top + sz.cy;

			// draw disabled or normal
			// normal: fill BG depending on state
			int sysColor;
			DWORD dFillColor;

			sysColor = (bChecked && !bSelected) ? COLOR_3DLIGHT : COLOR_MENU;
			dFillColor = GetSysColor(sysColor);
			if (sysColor == COLOR_3DLIGHT)
				PLFillRect3DLight( pDC->m_hDC, IconRect );
			else
				PLFillRect( pDC->m_hDC, IconRect, dFillColor );

			// draw pushed-in or popped-out edge
			if (bSelected || bChecked) {
				RECT rc2 = IconRect;
				pDC->DrawEdge( &rc2, bChecked ? BDR_SUNKENOUTER : BDR_RAISEDINNER, BF_RECT );
			}
			// draw the button!
			if(!bDisabled)
			{
				m_ilButtons->Draw( pDC, iButton, p, ILD_TRANSPARENT );
			}
			else
			{
				HICON hIcon = m_ilDisabledButtons->ExtractIcon( iButton );
				SIZE ls;
				ls.cx = 0;
				ls.cy = 0;

				int sysColor;
				DWORD dFillColor;

				sysColor = (bChecked && !bSelected) ? COLOR_3DLIGHT : COLOR_MENU;
				dFillColor = GetSysColor(sysColor);
				//pDC->DrawState((HBRUSH)NULL, (DRAWSTATEPROC)NULL, (LPARAM)hIcon, 0, p.x, p.y, ls.cx, ls.cy, DSS_NORMAL|DST_ICON );
				pDC->DrawState( p, ls, hIcon, DSS_NORMAL, (HBRUSH)NULL );
				DestroyIcon(hIcon);
			}

		} else {
			// no button: look for custom checked/unchecked bitmaps
			if ( bChecked ) Draw3DCheckmark(pDC->m_hDC, IconRect, bSelected, bDisabled );
		}

	}
	return TRUE; // handled
}

/*
	Function:		FillRect
	Returns:			-
	Description:	Fills a RECT with some color
*/
void CCoolmenu::FillRectCM( HDC dc, const RECT& rc, COLORREF color )
{
	HBRUSH hBrush = CreateSolidBrush( color );
	HBRUSH pOldBrush = (HBRUSH)SelectObject( dc, (HGDIOBJ)hBrush );
	PatBlt( dc, rc.left, rc.top, rc.right-rc.left, rc.bottom-rc.top, PATCOPY );
	SelectObject( dc, pOldBrush );
	DeleteObject( hBrush );
	return;
}

void CCoolmenu::DrawSpecial_WinXP(CDC* pDC, LPCRECT pRect, UINT nID, DWORD dwStyle)
{
	TCHAR cSign = 0;
	switch(nID&0xfff0)
	{
		case SC_MINIMIZE:
			cSign = 48; // Min
			break;
		case SC_MAXIMIZE:
			cSign = 49;// Max
			break;
		case SC_CLOSE:
			cSign = 114;// Close
			break;
		case SC_RESTORE:
			cSign = 50;// Restore
			break;
	}
	if(cSign)
	{
		COLORREF oldColor;
		BOOL bBold = (dwStyle&ODS_DEFAULT) ? TRUE : FALSE;
		CRect rect(pRect);
		rect.InflateRect(0,(IMGHEIGHT-rect.Height())>>1);

		if( (dwStyle&ODS_GRAYED) || (dwStyle&ODS_INACTIVE))
		{
			oldColor = pDC->SetTextColor(::GetSysColor(COLOR_GRAYTEXT));
		}
		else if(dwStyle&ODS_SELECTED)
		{
			oldColor = pDC->SetTextColor(menuColors[TextColor]);
			rect.OffsetRect(1,1);
			DrawSpecialChar(pDC,rect,cSign,bBold);
			pDC->SetTextColor(::GetSysColor(COLOR_MENUTEXT));
			rect.OffsetRect(-2,-2);
		}
		else
		{
			oldColor = pDC->SetTextColor(::GetSysColor(COLOR_MENUTEXT));
		}
		DrawSpecialChar(pDC,rect,cSign,bBold);

		pDC->SetTextColor(oldColor);
	}
}

/*
	Function:		DrawXP
	Returns:			-
	Description:	Draws menuitem with Office XP/2003/2007 look
*/
BOOL CCoolmenu::DrawXP( LPDRAWITEMSTRUCT lpds )
{
	ASSERT(lpds != NULL);

	CMyItemData* pmd = (CMyItemData*)lpds->itemData;
	if ( !IsMyItemData(pmd) )
		return FALSE; // not handled by me

	UINT state = lpds->itemState;

	CNewMemDC memDC(&lpds->rcItem,lpds->hDC);
	CDC* pDC;
	BOOL bIsMenuBar = pmd->bMenuBar;
	BOOL bHighContrast = FALSE;

	if( bIsMenuBar || ( pmd->fType & MFT_SEPARATOR ) )
	{ // For title and menubardrawing disable memory painting
		memDC.DoCancel();
		pDC = CDC::FromHandle(lpds->hDC);
	}
	else
	{
		pDC = &memDC;
	}

	BOOL bSilver = ( GetThemeColor()==0x00bea3a4);
// 	COLORREF colorWindow = m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW));
// 	COLORREF colorMenuBar = bIsMenuBar?GetMenuBarColor():GetMenuColor();
// 	COLORREF colorMenu = (m_bReactOnThemeChange)?(bSilver)?RGB( 251, 250, 251 ):(IsMenuThemeActive()?RGB( 252, 252, 249 ):m_CmGeneral->MixedColor(colorWindow,colorMenuBar)):m_CmGeneral->MixedColor(colorWindow,colorMenuBar);
// 	COLORREF colorBitmap = m_CmGeneral->MixedColor(GetMenuBarColor(),colorWindow);
// 	COLORREF colorSel = ((!IsMenuThemeActive()) && m_bReactOnThemeChange)?m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.7 ):menuColors[SelectedColor];
// 	COLORREF	colorSel2 = menuColors[SelectedColor];
// 	COLORREF colorBorder = (m_bReactOnThemeChange)?GetSysColor(COLOR_HIGHLIGHT):menuColors[PenColor];
// 	COLORREF colorCheck = ((!IsMenuThemeActive()) && m_bReactOnThemeChange)?m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.85 ):menuColors[CheckedColor];
// 	COLORREF colorCheckSel = ((!IsMenuThemeActive()) && m_bReactOnThemeChange)?m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.55 ):menuColors[CheckedSelectedColor];

/*	COLORREF colorWindow = m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW));
	COLORREF colorMenuBar = bIsMenuBar?GetMenuBarColor():GetMenuColor();
	COLORREF colorMenu = (m_bReactOnThemeChange)?(bSilver)?RGB( 251, 250, 251 ):(IsMenuThemeActive()?RGB( 252, 252, 249 ):m_CmGeneral->MixedColor(colorWindow,colorMenuBar)):m_CmGeneral->MixedColor(colorWindow,colorMenuBar);
	COLORREF colorBitmap = m_CmGeneral->MixedColor(GetMenuBarColor(),colorWindow);
	COLORREF colorSel = menuColors[SelectedColor];
	COLORREF	colorSel2 = menuColors[SelectedColor];
	COLORREF colorBorder = menuColors[PenColor];
	COLORREF colorCheck = menuColors[CheckedColor];
	COLORREF colorCheckSel = menuColors[CheckedSelectedColor];
*/

//	colorMenuBar = GetMenuBarColor();
//	colorBitmap = (g_Shell!=Win95 && g_Shell!=WinNT4)?colorMenuBar:RGB(163,194,245);

	// Better contrast when you have less than 256 colors
/*	if(pDC->GetNearestColor(colorMenu)==pDC->GetNearestColor(colorBitmap))
	{
		colorMenu = colorWindow;
		colorBitmap = colorMenuBar;
	}
*/
	CPen Pen(PS_SOLID,0,(state&ODS_GRAYED)?RGB(146,143,136):menuColors[PenColor]);

	if(bIsMenuBar)
	{
		if( !IsMenu((HMENU)(UINT_PTR)lpds->itemID) && lpds->itemState&ODS_SELECTED_OPEN)
		{
			lpds->itemState = (lpds->itemState & ~(ODS_SELECTED|ODS_SELECTED_OPEN))|ODS_HOTLIGHT;
		}
		else if( !(lpds->itemState&ODS_SELECTED_OPEN) && !pmd->m_dwOpenMenu && lpds->itemState&ODS_SELECTED)
		{
			lpds->itemState = (lpds->itemState&~ODS_SELECTED)|ODS_HOTLIGHT;
		}
// 		if(!(lpds->itemState&ODS_HOTLIGHT))
// 		{
// 			colorSel = (m_bReactOnThemeChange)?(bSilver)?RGB( 229, 228, 232 ):(IsMenuThemeActive())?RGB( 239, 237, 232 ):m_CmGeneral->LightenColor( 110, GetSysColor( COLOR_MENU ) ):menuColors[BitmapBckColor];
// 			colorSel2 = colorMenuBar;
// 		}
//		colorMenu = colorMenuBar;
	}

	CRect RectL(lpds->rcItem);
	CRect RectR(lpds->rcItem);
	CRect RectSel(lpds->rcItem);

	LOGFONT logFontMenu;
	CFont fontMenu;

#ifdef _NEW_MENU_USER_FONT
	logFontMenu = MENU_USER_FONT;
#else
	NONCLIENTMETRICS nm = {0};
	nm.cbSize = sizeof (nm);
	VERIFY (SystemParametersInfo(SPI_GETNONCLIENTMETRICS,nm.cbSize,&nm,0));
	logFontMenu = nm.lfMenuFont;
#endif

	if(bIsMenuBar)
	{
		RectR.InflateRect (0,0,0,0);
		if(lpds->itemState&ODS_DRAW_VERTICAL)
			RectSel.InflateRect (0,0,0, -4);
		else
			RectSel.InflateRect (0,(g_Shell==WinXP)?-1:0,-2 -2,0);

		if(lpds->itemState&ODS_SELECTED)
		{
			if(lpds->itemState&ODS_DRAW_VERTICAL)
				RectR.bottom -=4;
			else
				RectR.right -=4;
			if(m_CmGeneral->NumScreenColors() <= 256)
			{
				pDC->FillSolidRect(RectR,menuColors[MenuBckColor]);
			}
			else
			{
//				DrawGradient(pDC,RectR,colorMenu,colorBitmap,FALSE,TRUE);
				DrawGradient(pDC,RectR,menuColors[MenuBckColor],menuColors[BitmapBckColor],FALSE,TRUE);
			}
			if(lpds->itemState&ODS_DRAW_VERTICAL)
				RectR.bottom +=4;
			else
				RectR.right +=4;
		}
		else
		{
			MENUINFO menuInfo = {0};
			menuInfo.cbSize = sizeof(menuInfo);
			menuInfo.fMask = MIM_BACKGROUND;

			if(m_CmGeneral->GetMenuInfo(m_hMenu,&menuInfo) && menuInfo.hbrBack)
			{
				CBrush *pBrush = CBrush::FromHandle(menuInfo.hbrBack);
				CPoint brushOrg(0,0);

				VERIFY(pBrush->UnrealizeObject());
				CPoint oldOrg = pDC->SetBrushOrg(brushOrg);
				pDC->FillRect(RectR,pBrush);
				pDC->SetBrushOrg(oldOrg);
			}
			else
			{
				pDC->FillSolidRect(RectR,menuColors[MenuBckColor]);
			}
		}
	}
	else
	{
		RectL.right = RectL.left - (logFontMenu.lfHeight) + 12;
		RectR.left  = RectL.right;
		// Draw for Bitmapbackground
//		pDC->FillSolidRect(RectL,(m_bReactOnThemeChange)?(bSilver)?RGB( 229, 228, 232 ):(IsMenuThemeActive())?RGB( 239, 237, 232 ):m_CmGeneral->LightenColor( 110, GetSysColor( COLOR_MENU ) ):menuColors[BitmapBckColor]);
		pDC->FillSolidRect(RectL,menuColors[BitmapBckColor]);
		// Draw for Textbackground
		pDC->FillSolidRect(RectR,menuColors[MenuBckColor]);
	}

	// Spacing for submenu only in popups
	if(!bIsMenuBar)
	{
		RectR.left += 4;
		RectR.right -= 15;
	}

	//  Flag for highlighted item
	if(lpds->itemState & (ODS_HOTLIGHT|ODS_INACTIVE) )
	{
		lpds->itemState |= ODS_SELECTED;
	}

	if(bIsMenuBar && (lpds->itemState&ODS_SELECTED) )
	{
		if(!(lpds->itemState&ODS_INACTIVE) && pmd->bHasSubmenu)
		{
			SetLastMenuRect(lpds->hDC,RectSel);
			if(!(lpds->itemState&ODS_HOTLIGHT))
			{
        // Create a new pen for the special color
				Pen.DeleteObject();
				Pen.CreatePen(PS_SOLID,0,menuColors[PenColor]);

				if(m_bShadowEnabled )
				{
					int X,Y;
					CRect rect = RectR;
					int winH = rect.Height();

					// Simulate a shadow on right edge...
					if (m_CmGeneral->NumScreenColors() <= 256)
					{
						DWORD clr = GetSysColor(COLOR_BTNSHADOW);
						if(lpds->itemState&ODS_DRAW_VERTICAL)
						{
							int winW = rect.Width();
							for (Y=0; Y<=1 ;Y++)
							{
								for (X=4; X<=(winW-1) ;X++)
								{
									pDC->SetPixel(rect.left+X,rect.bottom-Y,clr);
								}
							}
						}
						else
						{
							for (X=3; X<=4 ;X++)
							{
								for (Y=4; Y<=(winH-1) ;Y++)
								{
									pDC->SetPixel(rect.right - X, Y + rect.top, clr );
								}
							}
						}
					}
					else
					{
						if(lpds->itemState&ODS_DRAW_VERTICAL)
						{
							int winW = rect.Width();
							COLORREF barColor = pDC->GetPixel(rect.left+4,rect.bottom-4);
							for (Y=1; Y<=4 ;Y++)
							{
								for (X=0; X<4 ;X++)
								{
									if(barColor==CLR_INVALID)
									{
										barColor = pDC->GetPixel(rect.left+X,rect.bottom-Y);
									}
									pDC->SetPixel(rect.left+X,rect.bottom-Y,barColor);
								}
								for (X=4; X<8 ;X++)
								{
									pDC->SetPixel(rect.left+X,rect.bottom-Y,m_CmGeneral->DarkenColor(2* 3 * Y * (X - 3), barColor)) ;
								}
								for (X=8; X<=(winW-1) ;X++)
								{
									pDC->SetPixel(rect.left+X,rect.bottom-Y, m_CmGeneral->DarkenColor(2*15 * Y, barColor) );
								}
							}
						}
						else
						{
							COLORREF barColor = pDC->GetPixel(rect.right-1,rect.top);
							for (X=1; X<=4 ;X++)
							{
								for (Y=0; Y<4 ;Y++)
								{
									if(barColor==CLR_INVALID)
									{
										barColor = pDC->GetPixel(rect.right-X,Y+rect.top);
									}
									pDC->SetPixel(rect.right-X,Y+rect.top, barColor );
								}
								for (Y=4; Y<8 ;Y++)
								{
									pDC->SetPixel(rect.right-X,Y+rect.top,m_CmGeneral->DarkenColor(2* 3 * X * (Y - 3), barColor)) ;
								}
								for (Y=8; Y<=(winH-1) ;Y++)
								{
									pDC->SetPixel(rect.right - X, Y+rect.top, m_CmGeneral->DarkenColor(2*15 * X, barColor) );
								}
							}
						}
					}
				}
			}
		}
	}
	// For keyboard navigation only
	BOOL bDrawSmallSelection = FALSE;
	// remove the selected bit if it's grayed out
	if( (lpds->itemState&ODS_GRAYED) )	//&& !m_bSelectDisable)
	{
		if( lpds->itemState & ODS_SELECTED )
		{
			lpds->itemState = lpds->itemState & (~ODS_SELECTED);
			DWORD MsgPos = ::GetMessagePos();
			if( MsgPos==m_dwMsgPos )
			{
				bDrawSmallSelection = TRUE;
			}
			else
			{
				m_dwMsgPos = MsgPos;
			}
		}
	}

	if( !(lpds->itemState&ODS_HOTLIGHT) && (lpds->itemState & ODS_SELECTED) && bIsMenuBar )
	{
		Pen.DeleteObject();
		Pen.CreatePen(PS_SOLID,0,menuColors[SelectedPenColor]);
	}

	// Draw the seperator
	if( pmd->fType & MFT_SEPARATOR )	//state & MFT_SEPARATOR )
	{
		// Draw only the seperator
		CRect rect;
		rect.top = RectR.CenterPoint().y;
		rect.bottom = rect.top+1;
		rect.right = lpds->rcItem.right;
		rect.left = RectR.left;
		pDC->FillSolidRect(rect,menuColors[SeparatorColor]);
	}
	else
	{
		if( (lpds->itemState & ODS_SELECTED) && !(lpds->itemState&ODS_INACTIVE) )
		{
			if(bIsMenuBar)
			{
				/*if(m_CmGeneral->NumScreenColors() <= 256)
					pDC->FillSolidRect(RectSel,colorWindow);
				else
				{*/
				if(lpds->itemState&ODS_HOTLIGHT)
					pDC->FillSolidRect(RectSel,menuColors[SelectedColor]);//:colorSel);
				else
					pDC->FillSolidRect(RectSel,menuColors[MenubarGradientStartColorDown]);//:colorSel);
				//}
			}
			else
			{
				pDC->FillSolidRect(RectSel,menuColors[SelectedColor]);
			}
			// Draw the selection
			CPen* pOldPen = pDC->SelectObject(&Pen);
			CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);
			pDC->Rectangle(RectSel);
			pDC->SelectObject(pOldBrush);
			pDC->SelectObject(pOldPen);
		}
		else if (bDrawSmallSelection)
		{
			pDC->FillSolidRect(RectSel,menuColors[MenuBckColor]);
			// Draw the selection for keyboardnavigation
			CPen* pOldPen = pDC->SelectObject(&Pen);
			CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);
			pDC->Rectangle(RectSel);
			pDC->SelectObject(pOldBrush);
			pDC->SelectObject(pOldPen);
		}

		UINT state = lpds->itemState;

		BOOL standardflag=FALSE;
		BOOL selectedflag=FALSE;
		BOOL disableflag=FALSE;
		BOOL checkflag=FALSE;

		CString strText = pmd->text;

		if(pmd->bMenuBar)
		{
			CMenuItemInfo infoMb;
			infoMb.fMask = MIIM_DATA;
			::GetMenuItemInfo( m_hMenu, pmd->iItem, TRUE, &infoMb);
			LPSTR menuText = new CHAR[1024];
			GetMenuString(m_hMenu, pmd->iItem, menuText, 1024, MF_BYPOSITION);
			strText = menuText;
			delete [] menuText;
		}


		if( (state&ODS_CHECKED) && (pmd->iButton<0) )
		{
			if(state&ODS_SELECTED && m_selectcheck>0)
			{
				checkflag=TRUE;
			}
			else if(m_unselectcheck>0)
			{
				checkflag=TRUE;
			}
		}
		else if(pmd->iButton >= 0)
		{
			standardflag = TRUE;
			if(state&ODS_SELECTED)
			{
				selectedflag=TRUE;
			}
			else if(state&ODS_GRAYED)
			{
				disableflag=TRUE;
			}
		}

		// draw the menutext
		if(!strText.IsEmpty())
		{
			// Default selection?
			if(state&ODS_DEFAULT || pmd->bBold)
			{
				// Make the font bold
				logFontMenu.lfWeight = FW_BOLD;
			}
			if(pmd->bItalic)
				logFontMenu.lfItalic = TRUE;
			if(pmd->bUnderline)
				logFontMenu.lfUnderline= TRUE;
			if(state&ODS_DRAW_VERTICAL)
			{
				// rotate font 90°
				logFontMenu.lfOrientation = -900;
				logFontMenu.lfEscapement = -900;
			}

			fontMenu.CreateFontIndirect(&logFontMenu);

			CString leftStr;
			CString rightStr;
			leftStr.Empty();
			rightStr.Empty();

			int tablocr = strText.ReverseFind(_T('\t'));
			if(tablocr!=-1)
			{
				rightStr = strText.Mid(tablocr+1);
				leftStr = strText.Left(strText.Find(_T('\t')));
			}
			else
			{
				leftStr=strText;
			}

			// Draw the text in the correct color:
			UINT nFormat  = DT_LEFT| DT_SINGLELINE|DT_VCENTER;
			UINT nFormatr = DT_RIGHT|DT_SINGLELINE|DT_VCENTER;

			int iOldMode = pDC->SetBkMode( TRANSPARENT);
			CFont* pOldFont = pDC->SelectObject (&fontMenu);

			COLORREF OldTextColor;
			if( (lpds->itemState&ODS_GRAYED) ||
				(bIsMenuBar && lpds->itemState&ODS_INACTIVE) )
			{
				// Draw the text disabled?
/*				if(bIsMenuBar && (m_CmGeneral->NumScreenColors() <= 256) )
				{
					OldTextColor = pDC->SetTextColor((m_bReactOnThemeChange)?colorWindow:menuColors[DisabledTextColor]);
				}
				else
				{*/
//				OldTextColor = pDC->SetTextColor((m_bReactOnThemeChange)?GetSysColor(COLOR_GRAYTEXT):menuColors[DisabledTextColor]);
				OldTextColor = pDC->SetTextColor(menuColors[DisabledTextColor]);
				//}
			}
			else
			{
				// Draw the text normal
				if( bHighContrast && !bIsMenuBar && !(state&ODS_SELECTED) )
				{
					OldTextColor = pDC->SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
				}
				else
				{
					OldTextColor = pDC->SetTextColor((pmd->clrText!=-1)?pmd->clrText:menuColors[TextColor]);
				}
			}
			BOOL bEnabled = FALSE;
			SystemParametersInfo(SPI_GETKEYBOARDCUES,0,&bEnabled,0);
			UINT dt_Hide = (lpds->itemState&ODS_NOACCEL && !bEnabled)?DT_HIDEPREFIX:0;
			if(bIsMenuBar)
			{
				MenuDrawText(pDC->m_hDC,leftStr,-1,RectSel, DT_SINGLELINE|DT_VCENTER|DT_CENTER|dt_Hide);
			}
			else
			{
				pDC->DrawText(leftStr,RectR, nFormat);
				if(tablocr!=-1)
				{
					pDC->DrawText (rightStr,RectR,nFormatr|dt_Hide);
				}
			}
			pDC->SetTextColor(OldTextColor);
			pDC->SelectObject(pOldFont);
			pDC->SetBkMode(iOldMode);
		}

		// Draw the bitmap or checkmarks
		if(!bIsMenuBar)
		{
			CRect rect2 = RectR;

			if(checkflag||standardflag||selectedflag||disableflag)
			{
				if(checkflag && m_checkmaps)
				{
					CPoint ptImage(RectL.left+3,RectL.top+4);

					if(state&ODS_SELECTED)
					{
						m_checkmaps->Draw(pDC,1,ptImage,ILD_TRANSPARENT);
					}
					else
					{
						m_checkmaps->Draw(pDC,0,ptImage,ILD_TRANSPARENT);
					}
				}
				else
				{
					CSize size;
					size.cx=16;
					size.cy=16;
					HICON hDrawIcon = (state & ODS_DISABLED)?m_ilDisabledButtons->ExtractIcon( pmd->iButton ):m_ilButtons->ExtractIcon( pmd->iButton );
					CPoint ptImage( RectL.left + ((RectL.Width()-size.cx)>>1), RectL.top + ((RectL.Height()-size.cy)>>1) );

					// Need to draw the checked state
					if (state&ODS_CHECKED)
					{
						CRect rect = RectL;
						rect.InflateRect (-1,-1,-2,-1);
						if(!(state&ODS_DISABLED))
						{
							if(selectedflag)
							{
								pDC->FillSolidRect(rect,menuColors[CheckedSelectedColor]);
							}
							else
							{
								pDC->FillSolidRect(rect,menuColors[CheckedColor]);
							}
						}
						CPen PenBorder(PS_SOLID,0,(state&ODS_GRAYED)?menuColors[CheckDisabledPenColor]:(state&ODS_SELECTED)?menuColors[CheckSelectedPenColor]:menuColors[CheckPenColor] );
						//CPen PenBorder(PS_SOLID,0,colorBorder );
						CPen* pOldPen = pDC->SelectObject(&PenBorder);
						CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);

						pDC->Rectangle(rect);

						pDC->SelectObject(pOldBrush);
						pDC->SelectObject(pOldPen);
					}

					// Correcting of a smaler icon
					if(size.cx<IMGWIDTH)
					{
						ptImage.x += (IMGWIDTH-size.cx)>>1;
					}

					if(state & ODS_DISABLED)
					{
						pDC->DrawState( ptImage, size, hDrawIcon, DSS_NORMAL, (HBRUSH)NULL );
					}
					else
					{
						if(selectedflag)
						{
							if(!(state & ODS_CHECKED))
							{
								ptImage.x += 2; ptImage.y += 2;
								CBrush Brush;
								// Color of the shade
								Brush.CreateSolidBrush(pDC->GetNearestColor( RGB( MulDiv(GetRValue(menuColors[SelectedColor]),7,10),
																								 MulDiv(GetGValue(menuColors[SelectedColor]),7,10),
																								 MulDiv(GetBValue(menuColors[SelectedColor])+55,7,10))));				//					Brush.CreateSolidBrush(pDC->GetNearestColor( MulDiv(GetRValue(GetSysColor(COLOR_WINDOW)),7,10)));

								pDC->DrawState(ptImage, size, hDrawIcon, DSS_NORMAL | DSS_MONO, &Brush);
								ptImage.x-=2; ptImage.y-=2;
							}
							pDC->DrawState(ptImage, size, hDrawIcon, DSS_NORMAL,(HBRUSH)NULL);
						}
						else
						{
							// draws the icon with normal color
							pDC->DrawState(ptImage, size, hDrawIcon, DSS_NORMAL,(HBRUSH)NULL);
						}
					}
					DestroyIcon(hDrawIcon);
				}
			}

			if(pmd->iButton<0 /*&& state&ODS_CHECKED */ && !checkflag)
			{
				MENUITEMINFO info = {0};
				info.cbSize = sizeof(info);
				info.fMask = MIIM_CHECKMARKS;

				::GetMenuItemInfo((HMENU)(lpds->hwndItem),lpds->itemID,MF_BYCOMMAND, &info);

				if(state&ODS_CHECKED || info.hbmpUnchecked)
				{
					CRect rect = RectL;
					rect.InflateRect (-1,-1,-2,-1);
					// draw the color behind checkmarks
					if(!(state&ODS_GRAYED))
					{
						if(state&ODS_SELECTED)
						{
							pDC->FillSolidRect(rect,menuColors[CheckedSelectedColor]);//!! menuColors[CheckedSelectedColor]);
						}
						else
						{
							pDC->FillSolidRect(rect,menuColors[CheckedColor]);	//!!menuColors[CheckedColor]);
						}
					}
					//CPen PenBorder(PS_SOLID,0,menuColors[PenColor] );
					CPen PenBorder(PS_SOLID,0,(state&ODS_GRAYED)?menuColors[CheckDisabledPenColor]:(state&ODS_SELECTED)?menuColors[CheckSelectedPenColor]:menuColors[CheckPenColor] );
					CPen* pOldPen = pDC->SelectObject(&PenBorder);
					CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);

					pDC->Rectangle(rect);

					pDC->SelectObject(pOldBrush);
					pDC->SelectObject(pOldPen);
					if (state&ODS_CHECKED)
					{
						CRect rect(RectL);
						rect.left++;
						rect.top += 2;

						if (!info.hbmpChecked)
						{ // Checkmark
							DrawSpecialCharStyle(pDC,rect,98,state);
						}
						else if(!info.hbmpUnchecked)
						{ // Bullet
							DrawSpecialCharStyle(pDC,rect,105,state);
						}
						else
						{ // Draw Bitmap
							BITMAP myInfo = {0};
							GetObject((HGDIOBJ)info.hbmpChecked,sizeof(myInfo),&myInfo);
							CPoint Offset = RectL.TopLeft() + CPoint((RectL.Width()-myInfo.bmWidth)/2,(RectL.Height()-myInfo.bmHeight)/2);
							pDC->DrawState(Offset,CSize(0,0),info.hbmpChecked,DST_BITMAP|DSS_MONO);
						}
					}
					else
					{
						// Draw Bitmap
						BITMAP myInfo = {0};
						GetObject((HGDIOBJ)info.hbmpUnchecked,sizeof(myInfo),&myInfo);
						CPoint Offset = RectL.TopLeft() + CPoint((RectL.Width()-myInfo.bmWidth)/2,(RectL.Height()-myInfo.bmHeight)/2);
						if(state & ODS_DISABLED)
						{
							pDC->DrawState(Offset,CSize(0,0),info.hbmpUnchecked,DST_BITMAP|DSS_MONO|DSS_DISABLED);
						}
						else
						{
							pDC->DrawState(Offset,CSize(0,0),info.hbmpUnchecked,DST_BITMAP|DSS_MONO);
						}
					}
				}
			}
		}
	}
	return TRUE;
}

/*
	Function:		Draw2K3
	Returns:			-
	Description:	Draws menuitem with Office 2003 look
*/
BOOL CCoolmenu::Draw2K3( LPDRAWITEMSTRUCT lpds )
{
	ASSERT(lpds != NULL);

	CMyItemData* pmd = (CMyItemData*)lpds->itemData;
	if ( !IsMyItemData(pmd) )
		return FALSE; // not handled by me

	UINT state = lpds->itemState;

	CNewMemDC memDC(&lpds->rcItem,lpds->hDC);
	CDC* pDC;
	BOOL bIsMenuBar = pmd->bMenuBar;
	BOOL bHighContrast = FALSE;

	if( bIsMenuBar || ( pmd->fType & MFT_SEPARATOR ) )
	{ // For title and menubardrawing disable memory painting
		memDC.DoCancel();
		pDC = CDC::FromHandle(lpds->hDC);
	}
	else
	{
		pDC = &memDC;
	}

	BOOL bSilver = ( GetThemeColor()==0x00bea3a4);
	/*COLORREF colorWindow = m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW));
	COLORREF colorMenuBar = bIsMenuBar?GetMenuBarColor():GetMenuColor();
	COLORREF colorMenu = m_CmGeneral->MixedColor(colorWindow,colorMenuBar);
	COLORREF colorBitmap = m_CmGeneral->MixedColor(GetMenuBarColor(),colorWindow);
	COLORREF colorSel = ((!IsMenuThemeActive()) && m_bReactOnThemeChange)?m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.7 ):menuColors[SelectedColor];
	COLORREF	colorSel2 = menuColors[SelectedColor];
	COLORREF colorBorder = (m_bReactOnThemeChange)?GetSysColor(COLOR_HIGHLIGHT):menuColors[PenColor];
	COLORREF colorCheck = ((!IsMenuThemeActive()) && m_bReactOnThemeChange)?m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.55 ):menuColors[CheckedColor];
	COLORREF colorCheckSel = ((!IsMenuThemeActive()) && m_bReactOnThemeChange)?m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.55 ):menuColors[CheckedSelectedColor];
	*/
	/*COLORREF colorWindow = m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW));
	COLORREF colorMenuBar = bIsMenuBar?GetMenuBarColor():GetMenuColor();
	COLORREF colorMenu = m_CmGeneral->MixedColor(colorWindow,colorMenuBar);
	COLORREF colorBitmap = m_CmGeneral->MixedColor(GetMenuBarColor(),colorWindow);

	colorMenuBar = GetMenuBarColor2003();
	colorBitmap = (g_Shell!=Win95 && g_Shell!=WinNT4)?colorMenuBar:RGB(163,194,245);
*/
	// Better contrast when you have less than 256 colors
	/*if(pDC->GetNearestColor(colorMenu)==pDC->GetNearestColor(colorBitmap))
	{
		colorMenu = colorWindow;
		colorBitmap = colorMenuBar;
	}*/

	CPen Pen(PS_SOLID,0,menuColors[PenColor]);

	if(bIsMenuBar)
	{
		if( !IsMenu((HMENU)(UINT_PTR)lpds->itemID) && lpds->itemState&ODS_SELECTED_OPEN)
		{
			lpds->itemState = (lpds->itemState & ~(ODS_SELECTED|ODS_SELECTED_OPEN))|ODS_HOTLIGHT;
		}
		else if( !(lpds->itemState&ODS_SELECTED_OPEN) && !pmd->m_dwOpenMenu && lpds->itemState&ODS_SELECTED)
		{
			lpds->itemState = (lpds->itemState&~ODS_SELECTED)|ODS_HOTLIGHT;
		}
		if(!(lpds->itemState&ODS_HOTLIGHT))
		{
			//colorSel = colorBitmap;
			//colorSel2 = colorMenuBar;
		}
		//colorMenu = colorMenuBar;
	}

	CRect RectL(lpds->rcItem);
	CRect RectR(lpds->rcItem);
	CRect RectSel(lpds->rcItem);

	LOGFONT logFontMenu;
	CFont fontMenu;

#ifdef _NEW_MENU_USER_FONT
	logFontMenu = MENU_USER_FONT;
#else
	NONCLIENTMETRICS nm = {0};
	nm.cbSize = sizeof (nm);
	VERIFY (SystemParametersInfo(SPI_GETNONCLIENTMETRICS,nm.cbSize,&nm,0));
	logFontMenu = nm.lfMenuFont;
#endif

	if(bIsMenuBar)
	{
		RectR.InflateRect (0,0,0,0);
		if(lpds->itemState&ODS_DRAW_VERTICAL)
			RectSel.InflateRect (0,0,0, -4);
		else
			RectSel.InflateRect (0,(g_Shell==WinXP)?-1:0,-2 -2,0);

		if(lpds->itemState&ODS_SELECTED)
		{
			if(lpds->itemState&ODS_DRAW_VERTICAL)
				RectR.bottom -=4;
			else
				RectR.right -=4;
			if(m_CmGeneral->NumScreenColors() <= 256)
			{
				//pDC->FillSolidRect(RectR,colorMenu);
			}
			else
			{
				//DrawGradient(pDC,RectR,colorMenu,colorBitmap,FALSE,TRUE);
			}
			if(lpds->itemState&ODS_DRAW_VERTICAL)
				RectR.bottom +=4;
			else
				RectR.right +=4;
		}
		else
		{
			MENUINFO menuInfo = {0};
			menuInfo.cbSize = sizeof(menuInfo);
			menuInfo.fMask = MIM_BACKGROUND;

			if(m_CmGeneral->GetMenuInfo(m_hMenu,&menuInfo) && menuInfo.hbrBack)
			{
				CBrush *pBrush = CBrush::FromHandle(menuInfo.hbrBack);
				CPoint brushOrg(0,0);

				VERIFY(pBrush->UnrealizeObject());
				CPoint oldOrg = pDC->SetBrushOrg(brushOrg);
				pDC->FillRect(RectR,pBrush);
				pDC->SetBrushOrg(oldOrg);
			}
			else
			{
				//pDC->FillSolidRect(RectR,colorMenu);
			}
		}
	}
	else
	{
		RectL.right = RectL.left - (logFontMenu.lfHeight) + 12;
		RectR.left  = RectL.right;
		// Draw for Bitmapbackground
		//!!DrawGradient(pDC,RectL,(m_bReactOnThemeChange)?colorMenu:menuColors[GradientStartColor],(m_bReactOnThemeChange)?colorBitmap:menuColors[GradientEndColor],TRUE);
		DrawGradient(pDC,RectL,menuColors[GradientStartColor], menuColors[GradientEndColor],TRUE);
		// Draw for Textbackground
		pDC->FillSolidRect(RectR,menuColors[MenuBckColor]);
	}

	// Spacing for submenu only in popups
	if(!bIsMenuBar)
	{
		RectR.left += 4;
		RectR.right -= 15;
	}

	//  Flag for highlighted item
	if(lpds->itemState & (ODS_HOTLIGHT|ODS_INACTIVE) )
	{
		lpds->itemState |= ODS_SELECTED;
	}

	if(bIsMenuBar && (lpds->itemState&ODS_SELECTED) )
	{
		if(!(lpds->itemState&ODS_INACTIVE) && pmd->bHasSubmenu)
		{
			SetLastMenuRect(lpds->hDC,RectSel);
			if(!(lpds->itemState&ODS_HOTLIGHT))
			{
        // Create a new pen for the special color
				Pen.DeleteObject();
				Pen.CreatePen(PS_SOLID,0,menuColors[PenColor]);

				if(m_bShadowEnabled )
				{
					int X,Y;
					CRect rect = RectR;
					int winH = rect.Height();

					// Simulate a shadow on right edge...
					if (m_CmGeneral->NumScreenColors() <= 256)
					{
						DWORD clr = GetSysColor(COLOR_BTNSHADOW);
						if(lpds->itemState&ODS_DRAW_VERTICAL)
						{
							int winW = rect.Width();
							for (Y=0; Y<=1 ;Y++)
							{
								for (X=4; X<=(winW-1) ;X++)
								{
									pDC->SetPixel(rect.left+X,rect.bottom-Y,clr);
								}
							}
						}
						else
						{
							for (X=3; X<=4 ;X++)
							{
								for (Y=4; Y<=(winH-1) ;Y++)
								{
									pDC->SetPixel(rect.right - X, Y + rect.top, clr );
								}
							}
						}
					}
					else
					{
						if(lpds->itemState&ODS_DRAW_VERTICAL)
						{
							int winW = rect.Width();
							COLORREF barColor = pDC->GetPixel(rect.left+4,rect.bottom-4);
							for (Y=1; Y<=4 ;Y++)
							{
								for (X=0; X<4 ;X++)
								{
									if(barColor==CLR_INVALID)
									{
										barColor = pDC->GetPixel(rect.left+X,rect.bottom-Y);
									}
									pDC->SetPixel(rect.left+X,rect.bottom-Y,barColor);
								}
								for (X=4; X<8 ;X++)
								{
									pDC->SetPixel(rect.left+X,rect.bottom-Y,m_CmGeneral->DarkenColor(2* 3 * Y * (X - 3), barColor)) ;
								}
								for (X=8; X<=(winW-1) ;X++)
								{
									pDC->SetPixel(rect.left+X,rect.bottom-Y, m_CmGeneral->DarkenColor(2*15 * Y, barColor) );
								}
							}
						}
						else
						{
							COLORREF barColor = pDC->GetPixel(rect.right-1,rect.top);
							for (X=1; X<=4 ;X++)
							{
								for (Y=0; Y<4 ;Y++)
								{
									if(barColor==CLR_INVALID)
									{
										barColor = pDC->GetPixel(rect.right-X,Y+rect.top);
									}
									pDC->SetPixel(rect.right-X,Y+rect.top, barColor );
								}
								for (Y=4; Y<8 ;Y++)
								{
									pDC->SetPixel(rect.right-X,Y+rect.top,m_CmGeneral->DarkenColor(2* 3 * X * (Y - 3), barColor)) ;
								}
								for (Y=8; Y<=(winH-1) ;Y++)
								{
									pDC->SetPixel(rect.right - X, Y+rect.top, m_CmGeneral->DarkenColor(2*15 * X, barColor) );
								}
							}
						}
					}
				}
			}
		}
	}
	// For keyboard navigation only
	BOOL bDrawSmallSelection = FALSE;
	// remove the selected bit if it's grayed out
	if( (lpds->itemState&ODS_GRAYED) )	//&& !m_bSelectDisable)
	{
		if( lpds->itemState & ODS_SELECTED )
		{
			lpds->itemState = lpds->itemState & (~ODS_SELECTED);
			if(m_bKeyDown)
			{
				bDrawSmallSelection = TRUE;
				m_bMouseSelect = FALSE;
			}
			DWORD MsgPos = ::GetMessagePos();
			if( MsgPos==m_dwMsgPos )
			{
				bDrawSmallSelection = TRUE;
			}
			else
			{
				m_dwMsgPos = MsgPos;
			}
		}
	}

	if( !(lpds->itemState&ODS_HOTLIGHT) && (lpds->itemState & ODS_SELECTED) && bIsMenuBar )
	{
		Pen.DeleteObject();
		Pen.CreatePen(PS_SOLID,0,menuColors[SelectedPenColor]);
	}
	// Draw the seperator
	if( pmd->fType & MFT_SEPARATOR )	//state & MFT_SEPARATOR )
	{
		// Draw only the seperator
		CRect rect;
		rect.top = RectR.CenterPoint().y;
		rect.bottom = rect.top+1;
		rect.right = lpds->rcItem.right;
		rect.left = RectR.left;
		pDC->FillSolidRect(rect,menuColors[SeparatorColor]);
	}
	else
	{
		if( (lpds->itemState & ODS_SELECTED) && !(lpds->itemState&ODS_INACTIVE) )
		{
			if(bIsMenuBar)
			{
				/*if(m_CmGeneral->NumScreenColors() <= 256)
				{
					pDC->FillSolidRect(RectSel,colorWindow);
				}
				else
				{*/
					if(lpds->itemState&ODS_HOTLIGHT)
						DrawGradient(pDC,RectSel,menuColors[MenubarGradientEndColor],menuColors[MenubarGradientStartColor],FALSE,TRUE);
					else
						DrawGradient(pDC,RectSel,menuColors[MenubarGradientEndColorDown],menuColors[MenubarGradientStartColorDown],FALSE,TRUE);
/*					if(pmd->bHasSubmenu)
						DrawGradient(pDC,RectSel,(IsMenuThemeActive())?colorWindow:colorSel,colorSel,FALSE,TRUE);
					else
						DrawGradient(pDC,RectSel,colorWindow,menuColors[SelectedColor],FALSE,TRUE);*/
				//}
			}
			else
			{
				pDC->FillSolidRect(RectSel,menuColors[SelectedColor]);
				/*!!pDC->FillSolidRect(RectSel,colorSel);*/
			}
			// Draw the selection
			CPen* pOldPen = pDC->SelectObject(&Pen);
			CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);
			pDC->Rectangle(RectSel);
			pDC->SelectObject(pOldBrush);
			pDC->SelectObject(pOldPen);
		}
		else if (bDrawSmallSelection && !bIsMenuBar && !m_bMouseSelect)
		{
			pDC->FillSolidRect(RectSel,menuColors[MenuBckColor]);
			// Draw the selection for keyboardnavigation
			CPen* pOldPen = pDC->SelectObject(&Pen);
			CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);
			pDC->Rectangle(RectSel);
			pDC->SelectObject(pOldBrush);
			pDC->SelectObject(pOldPen);
			m_bKeyDown = FALSE;
		}

		UINT state = lpds->itemState;

		BOOL standardflag=FALSE;
		BOOL selectedflag=FALSE;
		BOOL disableflag=FALSE;
		BOOL checkflag=FALSE;

		CString strText = pmd->text;

		if(pmd->bMenuBar)
		{
			CMenuItemInfo infoMb;
			infoMb.fMask = MIIM_DATA;
			::GetMenuItemInfo( m_hMenu, pmd->iItem, TRUE, &infoMb);
			LPSTR menuText = new CHAR[1024];
			GetMenuString(m_hMenu, pmd->iItem, menuText, 1024, MF_BYPOSITION);
			strText = menuText;
			delete [] menuText;
		}


		if( (state&ODS_CHECKED) && (pmd->iButton<0) )
		{
			if(state&ODS_SELECTED && m_selectcheck>0)
			{
				checkflag=TRUE;
			}
			else if(m_unselectcheck>0)
			{
				checkflag=TRUE;
			}
		}
		else if(pmd->iButton >= 0)
		{
			standardflag = TRUE;
			if(state&ODS_SELECTED)
			{
				selectedflag=TRUE;
			}
			else if(state&ODS_GRAYED)
			{
				disableflag=TRUE;
			}
		}

		// draw the menutext
		if(!strText.IsEmpty())
		{
			// Default selection?
			if(state&ODS_DEFAULT || pmd->bBold)
			{
				// Make the font bold
				logFontMenu.lfWeight = FW_BOLD;
			}
			if(pmd->bItalic)
				logFontMenu.lfItalic = TRUE;
			if(pmd->bUnderline)
				logFontMenu.lfUnderline= TRUE;
			if(state&ODS_DRAW_VERTICAL)
			{
				// rotate font 90°
				logFontMenu.lfOrientation = -900;
				logFontMenu.lfEscapement = -900;
			}

			fontMenu.CreateFontIndirect(&logFontMenu);

			CString leftStr;
			CString rightStr;
			leftStr.Empty();
			rightStr.Empty();

			int tablocr = strText.ReverseFind(_T('\t'));
			if(tablocr!=-1)
			{
				rightStr = strText.Mid(tablocr+1);
				leftStr = strText.Left(strText.Find(_T('\t')));
			}
			else
			{
				leftStr=strText;
			}

			// Draw the text in the correct color:
			UINT nFormat  = DT_LEFT| DT_SINGLELINE|DT_VCENTER;
			UINT nFormatr = DT_RIGHT|DT_SINGLELINE|DT_VCENTER;

			int iOldMode = pDC->SetBkMode( TRANSPARENT);
			CFont* pOldFont = pDC->SelectObject (&fontMenu);

			COLORREF OldTextColor;
			if( (lpds->itemState&ODS_GRAYED) ||
				(bIsMenuBar && lpds->itemState&ODS_INACTIVE) )
			{
				// Draw the text disabled?
				if(bIsMenuBar && (m_CmGeneral->NumScreenColors() <= 256) )
				{
					//OldTextColor = pDC->SetTextColor((m_bReactOnThemeChange)?colorWindow:menuColors[DisabledTextColor]);
					OldTextColor = pDC->SetTextColor(menuColors[DisabledTextColor]);
				}
				else
				{
					//OldTextColor = pDC->SetTextColor((m_bReactOnThemeChange)?GetSysColor(COLOR_GRAYTEXT):menuColors[DisabledTextColor]);
					OldTextColor = pDC->SetTextColor(menuColors[DisabledTextColor]);
				}
			}
			else
			{
				// Draw the text normal
				if( bHighContrast && !bIsMenuBar && !(state&ODS_SELECTED) )
				{
					OldTextColor = pDC->SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
				}
				else
				{
					OldTextColor = pDC->SetTextColor((pmd->clrText!=-1)?pmd->clrText:menuColors[TextColor]);
				}
			}
			BOOL bEnabled = FALSE;
			SystemParametersInfo(SPI_GETKEYBOARDCUES,0,&bEnabled,0);
			UINT dt_Hide = (lpds->itemState&ODS_NOACCEL && !bEnabled)?DT_HIDEPREFIX:0;
			if(bIsMenuBar)
			{
				MenuDrawText(pDC->m_hDC,leftStr,-1,RectSel, DT_SINGLELINE|DT_VCENTER|DT_CENTER|dt_Hide);
			}
			else
			{
				pDC->DrawText(leftStr,RectR, nFormat);
				if(tablocr!=-1)
				{
					pDC->DrawText (rightStr,RectR,nFormatr|dt_Hide);
				}
			}
			pDC->SetTextColor(OldTextColor);
			pDC->SelectObject(pOldFont);
			pDC->SetBkMode(iOldMode);
		}

		// Draw the bitmap or checkmarks
		if(!bIsMenuBar)
		{
			CRect rect2 = RectR;

			if(checkflag||standardflag||selectedflag||disableflag)
			{
				if(checkflag && m_checkmaps)
				{
					CPoint ptImage(RectL.left+3,RectL.top+4);

					if(state&ODS_SELECTED)
					{
						m_checkmaps->Draw(pDC,1,ptImage,ILD_TRANSPARENT);
					}
					else
					{
						m_checkmaps->Draw(pDC,0,ptImage,ILD_TRANSPARENT);
					}
				}
				else
				{
					CSize size;
					size.cx=16;
					size.cy=16;
					HICON hDrawIcon = (state & ODS_DISABLED)?m_ilDisabledButtons->ExtractIcon( pmd->iButton ):m_ilButtons->ExtractIcon( pmd->iButton );
					CPoint ptImage( RectL.left + ((RectL.Width()-size.cx)>>1), RectL.top + ((RectL.Height()-size.cy)>>1) );

					// Need to draw the checked state
					if (state&ODS_CHECKED)
					{
						CRect rect = RectL;
						rect.InflateRect (-1,-1,-2,-1);
						if(!(state&ODS_DISABLED))
						{
							if(selectedflag)
							{
								pDC->FillSolidRect(rect,menuColors[CheckedSelectedColor]);
							}
							else
							{
								pDC->FillSolidRect(rect,menuColors[CheckedColor]);
							}
						}
						//CPen PenBorder(PS_SOLID,0,menuColors[PenColor] );
						CPen PenBorder(PS_SOLID,0,(state&ODS_GRAYED)?menuColors[CheckDisabledPenColor]:(state&ODS_SELECTED)?menuColors[CheckSelectedPenColor]:menuColors[CheckPenColor] );
						CPen* pOldPen = pDC->SelectObject(&PenBorder);
						CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);

						pDC->Rectangle(rect);

						pDC->SelectObject(pOldBrush);
						pDC->SelectObject(pOldPen);
					}

					// Correcting of a smaler icon
					if(size.cx<IMGWIDTH)
					{
						ptImage.x += (IMGWIDTH-size.cx)>>1;
					}

					if(state & ODS_DISABLED)
					{
						pDC->DrawState( ptImage, size, hDrawIcon, DSS_NORMAL, (HBRUSH)NULL );
					}
					else
					{
						if(selectedflag)
						{
							pDC->DrawState(ptImage, size, hDrawIcon, DSS_NORMAL,(HBRUSH)NULL);
						}
						else
						{
							// draws the icon with normal color
							pDC->DrawState(ptImage, size, hDrawIcon, DSS_NORMAL,(HBRUSH)NULL);
						}
					}
					DestroyIcon(hDrawIcon);
				}
			}

			if(pmd->iButton<0 /*&& state&ODS_CHECKED */ && !checkflag)
			{
				MENUITEMINFO info = {0};
				info.cbSize = sizeof(info);
				info.fMask = MIIM_CHECKMARKS;

				::GetMenuItemInfo((HMENU)(lpds->hwndItem),lpds->itemID,MF_BYCOMMAND, &info);

				if(state&ODS_CHECKED || info.hbmpUnchecked)
				{
					CRect rect = RectL;
					rect.InflateRect (-1,-1,-2,-1);
					// draw the color behind checkmarks
					if(!(state&ODS_GRAYED))
					{
						if(state&ODS_SELECTED)
						{
							pDC->FillSolidRect(rect,menuColors[CheckedSelectedColor]);//!! menuColors[CheckedSelectedColor]);
						}
						else
						{
							pDC->FillSolidRect(rect,menuColors[CheckedColor]);	//!!menuColors[CheckedColor]);
						}
					}
					//CPen PenBorder(PS_SOLID,0,menuColors[PenColor] );
					CPen PenBorder(PS_SOLID,0,(state&ODS_GRAYED)?menuColors[CheckDisabledPenColor]:(state&ODS_SELECTED)?menuColors[CheckSelectedPenColor]:menuColors[CheckPenColor] );
					CPen* pOldPen = pDC->SelectObject(&PenBorder);
					CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);

					pDC->Rectangle(rect);

					pDC->SelectObject(pOldBrush);
					pDC->SelectObject(pOldPen);
					if (state&ODS_CHECKED)
					{
						CRect rect(RectL);
						rect.left++;
						rect.top += 2;

						if (!info.hbmpChecked)
						{ // Checkmark
							DrawSpecialCharStyle(pDC,rect,98,state);
						}
						else if(!info.hbmpUnchecked)
						{ // Bullet
							DrawSpecialCharStyle(pDC,rect,105,state);
						}
						else
						{ // Draw Bitmap
							BITMAP myInfo = {0};
							GetObject((HGDIOBJ)info.hbmpChecked,sizeof(myInfo),&myInfo);
							CPoint Offset = RectL.TopLeft() + CPoint((RectL.Width()-myInfo.bmWidth)/2,(RectL.Height()-myInfo.bmHeight)/2);
							pDC->DrawState(Offset,CSize(0,0),info.hbmpChecked,DST_BITMAP|DSS_MONO);
						}
					}
					else
					{
						// Draw Bitmap
						BITMAP myInfo = {0};
						GetObject((HGDIOBJ)info.hbmpUnchecked,sizeof(myInfo),&myInfo);
						CPoint Offset = RectL.TopLeft() + CPoint((RectL.Width()-myInfo.bmWidth)/2,(RectL.Height()-myInfo.bmHeight)/2);
						if(state & ODS_DISABLED)
						{
							pDC->DrawState(Offset,CSize(0,0),info.hbmpUnchecked,DST_BITMAP|DSS_MONO|DSS_DISABLED);
						}
						else
						{
							pDC->DrawState(Offset,CSize(0,0),info.hbmpUnchecked,DST_BITMAP|DSS_MONO);
						}
					}
				}
			}
		}
	}
	return TRUE;
}

/*
	Function:		Draw2K7
	Returns:			-
	Description:	Draws menuitem with Office XP/2003/2007 look
*/
BOOL CCoolmenu::Draw2K7( LPDRAWITEMSTRUCT lpds )
{
	ASSERT(lpds != NULL);

	CMyItemData* pmd = (CMyItemData*)lpds->itemData;
	if ( !IsMyItemData(pmd) )
		return FALSE; // not handled by me

	UINT state = lpds->itemState;

	CNewMemDC memDC(&lpds->rcItem,lpds->hDC);
	CDC* pDC;
	BOOL bIsMenuBar = pmd->bMenuBar;
	BOOL bHighContrast = FALSE;

	if( bIsMenuBar || ( pmd->fType & MFT_SEPARATOR ) )
	{ // For title and menubardrawing disable memory painting
		memDC.DoCancel();
		pDC = CDC::FromHandle(lpds->hDC);
	}
	else
	{
		pDC = &memDC;
	}

//	BOOL bSilver = ( GetThemeColor()==0x00bea3a4);
	BOOL bGrun = ( GetThemeColor()==0x0086b8aa);
	/*COLORREF colorWindow = m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW));
	COLORREF colorMenuBar = bIsMenuBar?GetMenuBarColor():GetMenuColor();
	COLORREF colorMenu = m_CmGeneral->MixedColor(colorWindow,colorMenuBar);
	COLORREF colorBitmap = m_CmGeneral->MixedColor(GetMenuBarColor(),colorWindow);
	COLORREF colorSel = ((!IsMenuThemeActive()) && m_bReactOnThemeChange)?m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.7 ):menuColors[SelectedColor];
	COLORREF	colorSel2 = menuColors[SelectedColor];
	COLORREF colorBorder = (m_bReactOnThemeChange && !IsMenuThemeActive())?GetSysColor(COLOR_HIGHLIGHT):menuColors[PenColor];
	colorBorder = (m_bReactOnThemeChange)?GetSysColor(COLOR_HIGHLIGHT):menuColors[PenColor];
	COLORREF colorCheck    = ((!IsMenuThemeActive()) && m_bReactOnThemeChange)?m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.55 ):menuColors[CheckedColor];
	COLORREF colorCheckSel = ((!IsMenuThemeActive()) && m_bReactOnThemeChange)?m_CmGeneral->LightenColor( GetSysColor( COLOR_HIGHLIGHT ), 0.55 ):menuColors[CheckedSelectedColor];*/
/*	COLORREF colorWindow = m_CmGeneral->DarkenColor(10,GetSysColor(COLOR_WINDOW));
	COLORREF colorMenuBar = bIsMenuBar?GetMenuBarColor():GetMenuColor();
	COLORREF colorMenu = m_CmGeneral->MixedColor(colorWindow,colorMenuBar);
	COLORREF colorBitmap = m_CmGeneral->MixedColor(GetMenuBarColor(),colorWindow);

	colorMenuBar = GetMenuBarColor2007();//RGB(191,219,255);
	colorBitmap = (g_Shell!=Win95 && g_Shell!=WinNT4)?colorMenuBar:RGB(163,194,245);
*/
	// Better contrast when you have less than 256 colors
/*	if(pDC->GetNearestColor(colorMenu)==pDC->GetNearestColor(colorBitmap))
	{
		colorMenu = colorWindow;
		colorBitmap = colorMenuBar;
	}
*/
//!!	CPen Pen(PS_SOLID,0,(state&ODS_GRAYED)?RGB(141,141,141):menuColors[PenColor]);
	CPen Pen(PS_SOLID,0,menuColors[PenColor]);

	if(bIsMenuBar)
	{
		if( !IsMenu((HMENU)(UINT_PTR)lpds->itemID) && lpds->itemState&ODS_SELECTED_OPEN)
		{
			lpds->itemState = (lpds->itemState & ~(ODS_SELECTED|ODS_SELECTED_OPEN))|ODS_HOTLIGHT;
		}
		else if( !(lpds->itemState&ODS_SELECTED_OPEN) && !pmd->m_dwOpenMenu && lpds->itemState&ODS_SELECTED)
		{
			lpds->itemState = (lpds->itemState&~ODS_SELECTED)|ODS_HOTLIGHT;
		}
		//colorMenu = colorMenuBar;
	}

	CRect RectL(lpds->rcItem);
	CRect RectR(lpds->rcItem);
	CRect RectSel(lpds->rcItem);

	LOGFONT logFontMenu;
	CFont fontMenu;

#ifdef _NEW_MENU_USER_FONT
	logFontMenu = MENU_USER_FONT;
#else
	NONCLIENTMETRICS nm = {0};
	nm.cbSize = sizeof (nm);
	VERIFY (SystemParametersInfo(SPI_GETNONCLIENTMETRICS,nm.cbSize,&nm,0));
	logFontMenu = nm.lfMenuFont;
#endif

	if(bIsMenuBar)
	{
		RectR.InflateRect (0,0,0,0);
		if(lpds->itemState&ODS_DRAW_VERTICAL)
			RectSel.InflateRect (0,0,0, -4);
		else
			RectSel.InflateRect (0,(g_Shell==WinXP)?-1:0,-2 -2,0);

		if(lpds->itemState&ODS_SELECTED)
		{
			if(lpds->itemState&ODS_DRAW_VERTICAL)
				RectR.bottom -=4;
			else
				RectR.right -=4;
			/*if(m_CmGeneral->NumScreenColors() <= 256)
			{
				pDC->FillSolidRect(RectR,colorMenu);
			}
			else
			{
				DrawGradient(pDC,RectR,colorMenu,colorBitmap,FALSE,TRUE);
			}*/
			if(lpds->itemState&ODS_DRAW_VERTICAL)
				RectR.bottom +=4;
			else
				RectR.right +=4;
		}
		else
		{
			MENUINFO menuInfo = {0};
			menuInfo.cbSize = sizeof(menuInfo);
			menuInfo.fMask = MIM_BACKGROUND;

			if(m_CmGeneral->GetMenuInfo(m_hMenu,&menuInfo) && menuInfo.hbrBack)
			{
				CBrush *pBrush = CBrush::FromHandle(menuInfo.hbrBack);
				CPoint brushOrg(0,0);

				VERIFY(pBrush->UnrealizeObject());
				CPoint oldOrg = pDC->SetBrushOrg(brushOrg);
				pDC->FillRect(RectR,pBrush);
				pDC->SetBrushOrg(oldOrg);
			}
			else
			{
				//pDC->FillSolidRect(RectR,colorMenu);
			}
		}
	}
	else
	{
		RectL.right = RectL.left - (logFontMenu.lfHeight) + 12;
		RectR.left  = RectL.right;
		// Draw for Bitmapbackground
		if(m_bReactOnThemeChange)
		{
			if(!bGrun)
				pDC->FillSolidRect(RectL,menuColors[BitmapBckColor]);
			else
				DrawGradient(pDC,RectL,menuColors[GradientStartColor],menuColors[GradientEndColor],TRUE);
		}
		else
		{
			pDC->FillSolidRect(RectL,menuColors[BitmapBckColor]);
		}
		// Draw for Textbackground
		pDC->FillSolidRect(RectR,menuColors[MenuBckColor]);
	}

	// Spacing for submenu only in popups
	if(!bIsMenuBar)
	{
		RectR.left += 4;
		RectR.right -= 15;
	}

	//  Flag for highlighted item
	if(lpds->itemState & (ODS_HOTLIGHT|ODS_INACTIVE) )
	{
		lpds->itemState |= ODS_SELECTED;
	}

	if(bIsMenuBar && (lpds->itemState&ODS_SELECTED) )
	{
		if(!(lpds->itemState&ODS_INACTIVE) && pmd->bHasSubmenu)
		{
			SetLastMenuRect(lpds->hDC,RectSel);
			if(!(lpds->itemState&ODS_HOTLIGHT))
			{
        // Create a new pen for the special color

				if(m_bShadowEnabled )
				{
					int X,Y;
					CRect rect = RectR;
					int winH = rect.Height();

					// Simulate a shadow on right edge...
					if (m_CmGeneral->NumScreenColors() <= 256)
					{
						DWORD clr = GetSysColor(COLOR_BTNSHADOW);
						if(lpds->itemState&ODS_DRAW_VERTICAL)
						{
							int winW = rect.Width();
							for (Y=0; Y<=1 ;Y++)
							{
								for (X=4; X<=(winW-1) ;X++)
								{
									pDC->SetPixel(rect.left+X,rect.bottom-Y,clr);
								}
							}
						}
						else
						{
							for (X=3; X<=4 ;X++)
							{
								for (Y=4; Y<=(winH-1) ;Y++)
								{
									pDC->SetPixel(rect.right - X, Y + rect.top, clr );
								}
							}
						}
					}
					else
					{
						if(lpds->itemState&ODS_DRAW_VERTICAL)
						{
							int winW = rect.Width();
							COLORREF barColor = pDC->GetPixel(rect.left+4,rect.bottom-4);
							for (Y=1; Y<=4 ;Y++)
							{
								for (X=0; X<4 ;X++)
								{
									if(barColor==CLR_INVALID)
									{
										barColor = pDC->GetPixel(rect.left+X,rect.bottom-Y);
									}
									pDC->SetPixel(rect.left+X,rect.bottom-Y,barColor);
								}
								for (X=4; X<8 ;X++)
								{
									pDC->SetPixel(rect.left+X,rect.bottom-Y,m_CmGeneral->DarkenColor(2* 3 * Y * (X - 3), barColor)) ;
								}
								for (X=8; X<=(winW-1) ;X++)
								{
									pDC->SetPixel(rect.left+X,rect.bottom-Y, m_CmGeneral->DarkenColor(2*15 * Y, barColor) );
								}
							}
						}
						else
						{
							COLORREF barColor = pDC->GetPixel(rect.right-1,rect.top);
							for (X=1; X<=4 ;X++)
							{
								for (Y=0; Y<4 ;Y++)
								{
									if(barColor==CLR_INVALID)
									{
										barColor = pDC->GetPixel(rect.right-X,Y+rect.top);
									}
									pDC->SetPixel(rect.right-X,Y+rect.top, barColor );
								}
								for (Y=4; Y<8 ;Y++)
								{
									pDC->SetPixel(rect.right-X,Y+rect.top,m_CmGeneral->DarkenColor(2* 3 * X * (Y - 3), barColor)) ;
								}
								for (Y=8; Y<=(winH-1) ;Y++)
								{
									pDC->SetPixel(rect.right - X, Y+rect.top, m_CmGeneral->DarkenColor(2*15 * X, barColor) );
								}
							}
						}
					}
				}
			}
		}
	}
	// For keyboard navigation only
	BOOL bDrawSmallSelection = FALSE;
	// remove the selected bit if it's grayed out
	if( (lpds->itemState&ODS_GRAYED))	//&& !m_bSelectDisable)
	{
		if( lpds->itemState & ODS_SELECTED )
		{
			lpds->itemState = lpds->itemState & (~ODS_SELECTED);
			DWORD MsgPos = ::GetMessagePos();
			if(m_bKeyDown)
			{
				bDrawSmallSelection = TRUE;
				m_bMouseSelect = FALSE;
			}
			if( MsgPos==m_dwMsgPos )
			{
				bDrawSmallSelection = TRUE;
			}
			else
			{
				m_dwMsgPos = MsgPos;
			}
		}
	}

	if( !(lpds->itemState&ODS_HOTLIGHT) && (lpds->itemState & ODS_SELECTED) && bIsMenuBar )
	{
		Pen.DeleteObject();
		Pen.CreatePen(PS_SOLID,0,menuColors[SelectedPenColor]);
	}
	// Draw the seperator
	if( pmd->fType & MFT_SEPARATOR )	//state & MFT_SEPARATOR )
	{
		// Draw only the seperator
		CRect rect;
		rect.top = RectR.CenterPoint().y;
		rect.bottom = rect.top+1;
		rect.right = lpds->rcItem.right;
		rect.left = RectR.left;
		pDC->FillSolidRect(rect,menuColors[SeparatorColor]);
	}
	else
	{
		if( (lpds->itemState & ODS_SELECTED) && !(lpds->itemState&ODS_INACTIVE) )
		{
			if(bIsMenuBar)
			{
				/*if(m_CmGeneral->NumScreenColors() <= 256)
				{
					pDC->FillSolidRect(RectSel,colorWindow);
				}
				else
				{*/
					//DrawGradient(pDC,RectSel,colorSel,colorSel2,FALSE,TRUE);
					if(lpds->itemState&ODS_HOTLIGHT)
						DrawGradient(pDC,RectSel,menuColors[MenubarGradientEndColor],menuColors[MenubarGradientStartColor],FALSE,TRUE);
					else
						DrawGradient(pDC,RectSel,menuColors[MenubarGradientEndColorDown],menuColors[MenubarGradientStartColorDown],FALSE,TRUE);
				//}
			}
			else
			{
				pDC->FillSolidRect(RectSel,menuColors[SelectedColor]);
			}
			// Draw the selection
			CPen* pOldPen = pDC->SelectObject(&Pen);
			CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);
			pDC->Rectangle(RectSel);
			pDC->SelectObject(pOldBrush);
			pDC->SelectObject(pOldPen);
		}
		else if (bDrawSmallSelection && !bIsMenuBar && !m_bMouseSelect)
		{
			pDC->FillSolidRect(RectSel,menuColors[MenuBckColor]);
			// Draw the selection for keyboardnavigation
			CPen* pOldPen = pDC->SelectObject(&Pen);
			CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);
			pDC->Rectangle(RectSel);
			pDC->SelectObject(pOldBrush);
			pDC->SelectObject(pOldPen);
			m_bKeyDown = FALSE;
		}

		UINT state = lpds->itemState;

		BOOL standardflag=FALSE;
		BOOL selectedflag=FALSE;
		BOOL disableflag=FALSE;
		BOOL checkflag=FALSE;

		CString strText = pmd->text;

		if(pmd->bMenuBar)
		{
			CMenuItemInfo infoMb;
			infoMb.fMask = MIIM_DATA;
			::GetMenuItemInfo( m_hMenu, pmd->iItem, TRUE, &infoMb);
			LPSTR menuText = new CHAR[1024];
			GetMenuString(m_hMenu, pmd->iItem, menuText, 1024, MF_BYPOSITION);
			strText = menuText;
			delete [] menuText;
		}


		if( (state&ODS_CHECKED) && (pmd->iButton<0) )
		{
			if(state&ODS_SELECTED && m_selectcheck>0)
			{
				checkflag=TRUE;
			}
			else if(m_unselectcheck>0)
			{
				checkflag=TRUE;
			}
		}
		else if(pmd->iButton >= 0)
		{
			standardflag = TRUE;
			if(state&ODS_SELECTED)
			{
				selectedflag=TRUE;
			}
			else if(state&ODS_GRAYED)
			{
				disableflag=TRUE;
			}
		}

		// draw the menutext
		if(!strText.IsEmpty())
		{
			// Default selection?
			if(state&ODS_DEFAULT || pmd->bBold)
			{
				// Make the font bold
				logFontMenu.lfWeight = FW_BOLD;
			}
			if(pmd->bItalic)
				logFontMenu.lfItalic = TRUE;
			if(pmd->bUnderline)
				logFontMenu.lfUnderline= TRUE;
			if(state&ODS_DRAW_VERTICAL)
			{
				// rotate font 90°
				logFontMenu.lfOrientation = -900;
				logFontMenu.lfEscapement = -900;
			}

			fontMenu.CreateFontIndirect(&logFontMenu);

			CString leftStr;
			CString rightStr;
			leftStr.Empty();
			rightStr.Empty();

			int tablocr = strText.ReverseFind(_T('\t'));
			if(tablocr!=-1)
			{
				rightStr = strText.Mid(tablocr+1);
				leftStr = strText.Left(strText.Find(_T('\t')));
			}
			else
			{
				leftStr=strText;
			}

			// Draw the text in the correct color:
			UINT nFormat  = DT_LEFT| DT_SINGLELINE|DT_VCENTER;
			UINT nFormatr = DT_RIGHT|DT_SINGLELINE|DT_VCENTER;

			int iOldMode = pDC->SetBkMode( TRANSPARENT);
			CFont* pOldFont = pDC->SelectObject (&fontMenu);

			COLORREF OldTextColor;
			if( (lpds->itemState&ODS_GRAYED) ||
				(bIsMenuBar && lpds->itemState&ODS_INACTIVE) )
			{
				// Draw the text disabled?
				if(bIsMenuBar && (m_CmGeneral->NumScreenColors() <= 256) )
				{
					OldTextColor = pDC->SetTextColor(menuColors[DisabledTextColor]);
				}
				else
				{
					OldTextColor = pDC->SetTextColor(menuColors[DisabledTextColor]);
				}
			}
			else
			{
				// Draw the text normal
				if( bHighContrast && !bIsMenuBar && !(state&ODS_SELECTED) )
				{
					OldTextColor = pDC->SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
				}
				else
				{
					OldTextColor = pDC->SetTextColor((pmd->clrText!=-1)?pmd->clrText:menuColors[TextColor]);
				}
			}
			BOOL bEnabled = FALSE;
			SystemParametersInfo(SPI_GETKEYBOARDCUES,0,&bEnabled,0);
			UINT dt_Hide = (lpds->itemState&ODS_NOACCEL && !bEnabled)?DT_HIDEPREFIX:0;
			if(bIsMenuBar)
			{
				MenuDrawText(pDC->m_hDC,leftStr,-1,RectSel, DT_SINGLELINE|DT_VCENTER|DT_CENTER|dt_Hide);
			}
			else
			{
				pDC->DrawText(leftStr,RectR, nFormat);
				if(tablocr!=-1)
				{
					pDC->DrawText (rightStr,RectR,nFormatr|dt_Hide);
				}
			}
			pDC->SetTextColor(OldTextColor);
			pDC->SelectObject(pOldFont);
			pDC->SetBkMode(iOldMode);
		}

		// Draw the bitmap or checkmarks
		if(!bIsMenuBar)
		{
			CRect rect2 = RectR;

			if(checkflag||standardflag||selectedflag||disableflag)
			{
				if(checkflag && m_checkmaps)
				{
					CPoint ptImage(RectL.left+3,RectL.top+4);

					if(state&ODS_SELECTED)
					{
						m_checkmaps->Draw(pDC,1,ptImage,ILD_TRANSPARENT);
					}
					else
					{
						m_checkmaps->Draw(pDC,0,ptImage,ILD_TRANSPARENT);
					}
				}
				else
				{
					CSize size;
					size.cx=16;
					size.cy=16;
					HICON hDrawIcon = (state & ODS_DISABLED)?m_ilDisabledButtons->ExtractIcon( pmd->iButton ):m_ilButtons->ExtractIcon( pmd->iButton );
					CPoint ptImage( RectL.left + ((RectL.Width()-size.cx)>>1), RectL.top + ((RectL.Height()-size.cy)>>1) );

					// Need to draw the checked state
					if (state&ODS_CHECKED)
					{
						CRect rect = RectL;
						rect.InflateRect (-1,-1,-2,-1);
						if(!(state&ODS_DISABLED))
						{
							if(selectedflag)
							{
								pDC->FillSolidRect(rect,menuColors[CheckedSelectedColor]);
							}
							else
							{
								pDC->FillSolidRect(rect,menuColors[CheckedColor]);
							}
						}
						CPen PenBorder(PS_SOLID,0,(state&ODS_GRAYED)?menuColors[CheckDisabledPenColor]:(state&ODS_SELECTED)?menuColors[CheckSelectedPenColor]:menuColors[CheckPenColor] );
						CPen* pOldPen = pDC->SelectObject(&PenBorder);
						CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);

						pDC->Rectangle(rect);

						pDC->SelectObject(pOldBrush);
						pDC->SelectObject(pOldPen);
					}

					// Correcting of a smaler icon
					if(size.cx<IMGWIDTH)
					{
						ptImage.x += (IMGWIDTH-size.cx)>>1;
					}

					if(state & ODS_DISABLED)
					{
						pDC->DrawState( ptImage, size, hDrawIcon, DSS_NORMAL, (HBRUSH)NULL );
					}
					else
					{
						if(selectedflag)
						{
							pDC->DrawState(ptImage, size, hDrawIcon, DSS_NORMAL,(HBRUSH)NULL);
						}
						else
						{
							// draws the icon with normal color
							pDC->DrawState(ptImage, size, hDrawIcon, DSS_NORMAL,(HBRUSH)NULL);
						}
					}
					DestroyIcon(hDrawIcon);
				}
			}

			if(pmd->iButton<0 /*&& state&ODS_CHECKED */ && !checkflag)
			{
				MENUITEMINFO info = {0};
				info.cbSize = sizeof(info);
				info.fMask = MIIM_CHECKMARKS;

				::GetMenuItemInfo((HMENU)(lpds->hwndItem),lpds->itemID,MF_BYCOMMAND, &info);

				if(state&ODS_CHECKED || info.hbmpUnchecked)
				{
					CRect rect = RectL;
					rect.InflateRect (-1,-1,-2,-1);
					// draw the color behind checkmarks
					if(!(state&ODS_GRAYED))
					{
						if(state&ODS_SELECTED)
						{
							pDC->FillSolidRect(rect,menuColors[CheckedSelectedColor]);//!! menuColors[CheckedSelectedColor]);
						}
						else
						{
							pDC->FillSolidRect(rect,menuColors[CheckedColor]);	//!!menuColors[CheckedColor]);
						}
					}
					CPen PenBorder(PS_SOLID,0,(state&ODS_GRAYED)?menuColors[CheckDisabledPenColor]:(state&ODS_SELECTED)?menuColors[CheckSelectedPenColor]:menuColors[CheckPenColor] );
					//CPen PenBorder(PS_SOLID,0,(m_bReactOnThemeChange && !IsMenuThemeActive())?menuColors[PenColor]:(state&ODS_SELECTED)?RGB( 251, 140, 60 ):RGB( 255, 171, 63 ) );
					CPen* pOldPen = pDC->SelectObject(&PenBorder);
					CBrush* pOldBrush = (CBrush*)pDC->SelectStockObject(HOLLOW_BRUSH);

					pDC->Rectangle(rect);

					pDC->SelectObject(pOldBrush);
					pDC->SelectObject(pOldPen);
					if (state&ODS_CHECKED)
					{
						CRect rect(RectL);
						rect.left++;
						rect.top += 2;

						if (!info.hbmpChecked)
						{ // Checkmark
							DrawSpecialCharStyle(pDC,rect,98,state);
						}
						else if(!info.hbmpUnchecked)
						{ // Bullet
							DrawSpecialCharStyle(pDC,rect,105,state);
						}
						else
						{ // Draw Bitmap
							BITMAP myInfo = {0};
							GetObject((HGDIOBJ)info.hbmpChecked,sizeof(myInfo),&myInfo);
							CPoint Offset = RectL.TopLeft() + CPoint((RectL.Width()-myInfo.bmWidth)/2,(RectL.Height()-myInfo.bmHeight)/2);
							pDC->DrawState(Offset,CSize(0,0),info.hbmpChecked,DST_BITMAP|DSS_MONO);
						}
					}
					else
					{
						// Draw Bitmap
						BITMAP myInfo = {0};
						GetObject((HGDIOBJ)info.hbmpUnchecked,sizeof(myInfo),&myInfo);
						CPoint Offset = RectL.TopLeft() + CPoint((RectL.Width()-myInfo.bmWidth)/2,(RectL.Height()-myInfo.bmHeight)/2);
						if(state & ODS_DISABLED)
						{
							pDC->DrawState(Offset,CSize(0,0),info.hbmpUnchecked,DST_BITMAP|DSS_MONO|DSS_DISABLED);
						}
						else
						{
							pDC->DrawState(Offset,CSize(0,0),info.hbmpUnchecked,DST_BITMAP|DSS_MONO);
						}
					}
				}
			}
		}
	}
	return TRUE;
}

/*
	Function:		PLFillRect
	Returns:			-
	Description:	Fills a RECT with a color
*/
void CCoolmenu::PLFillRect( HDC& pDC, const RECT& rc, COLORREF color)
{
	HBRUSH brush = CreateSolidBrush( color );
	HBRUSH pOldBrush = (HBRUSH)SelectObject( pDC, (HGDIOBJ)brush );
	PatBlt( pDC, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, PATCOPY );
	SelectObject( pDC, (HGDIOBJ)pOldBrush );
	DeleteObject( brush );
}

/*
	Function:		PLFillRect3DLight
	Returns:			-
	Description:	Fills a RECT with a color
*/
void CCoolmenu::PLFillRect3DLight( HDC& pDC, const RECT& rc)
{
	// checked but not hot or pressed:, fill with light bg color
	// since some color schemes don't have COLOR_3DLIGHT, need to
	// create 8x8 mono bitmap alternating black/white every pixel, then
	// use it as a pattern brush with 3dface/3dhilite colors as bg/fg.
	// Each row is a WORD instead of byte because windows pads each row
	// to neares WORD boundary
	//
	static HBITMAP bmLight;
	if (!(HBITMAP)bmLight) {
		// 0xaa = 10101010
		// 0x55 = 01010101
		WORD rows[8] = { 0xaa, 0x55, 0xaa, 0x55, 0xaa, 0x55, 0xaa, 0x55 };
		bmLight = CreateBitmap( 8, 8, 1, 1, rows );
	}
	HBRUSH br = CreatePatternBrush( bmLight ); // pattern brush
	HBRUSH pOldBrush = (HBRUSH)SelectObject( pDC, (HGDIOBJ)br );
	COLORREF cBG = SetBkColor( pDC, GetSysColor( COLOR_3DFACE ) );
	COLORREF cFG = SetTextColor( pDC, GetSysColor( COLOR_3DHIGHLIGHT ) );
	PatBlt( pDC, rc.left, rc.top, rc.right - rc.left, rc.bottom - rc.top, PATCOPY );
	SelectObject( pDC, (HGDIOBJ)pOldBrush );
	SetBkColor( pDC, cBG );
	DeleteObject( (HBRUSH)br );
	DeleteObject( bmLight );
	SetTextColor( pDC, cFG );
}

/*
	Function:		DrawMenuText
	Returns:			-
	Description:	Draw the menuitemtext
*/
void  CCoolmenu::DrawMenuText( HDC pDC, RECT& rc, COLORREF color, CString text )
{
	CString sShortCut;
	CString sCaption;

	sCaption = text;
	int iTabPos = text.Find('\t');

	if ( iTabPos >= 0 )
	{
		sCaption = text.Left(iTabPos);
		sShortCut = text.Right( text.GetLength() - iTabPos - 1 );
	}
	SetTextColor( pDC, color );
	CRect rcText = CRect( rc );
	DrawText( pDC, sCaption, sCaption.GetLength(), &rcText, DT_MYSTANDARD );
	rcText.right -= TEXTPADDING + 4;
	if (sShortCut.GetLength() > 0)
		DrawText( pDC, sShortCut, sShortCut.GetLength(), &rcText, DT_MYSTANDARD | DT_RIGHT );
}

/*
	Function:		Draw3DCheckmark
	Returns:			-
	Description:	Draws a checkmark
*/
BOOL CCoolmenu::Draw3DCheckmark( HDC pDC, const RECT& rc, BOOL bSelected, BOOL bDisabled )
{
	RECT rect = rc;

	COLORREF crHighLight = GetSysColor( COLOR_HIGHLIGHT );

	COLORREF crBackImg = crBackImg = ( bDisabled || bSelected ? GetSysColor( COLOR_MENU ) : GetSysColor( COLOR_3DFACE ) );

	int sysColor = !bSelected ? COLOR_3DLIGHT : COLOR_MENU;
	COLORREF dFillColor = GetSysColor(sysColor);
	if ( sysColor == COLOR_3DLIGHT )
		PLFillRect3DLight( pDC, rect );
	else
		PLFillRect( pDC, rect, dFillColor );

	DrawEdge( pDC, &rect, BDR_SUNKENOUTER, BF_RECT);

	rect.left += ( ( ( rect.right - rect.left ) - IMGWIDTH )>>1 ) + 2;
	rect.right = rect.left + IMGWIDTH + IMGPADDING;
	SetBkColor( pDC, crBackImg );
	HBITMAP hBmp = LoadBitmap( NULL, MAKEINTRESOURCE( 32760 ) );
	if ( bDisabled )
	{
		HBRUSH hBrush = CreateSolidBrush( GetSysColor( COLOR_GRAYTEXT ) );
		DrawState( pDC, hBrush, (DRAWSTATEPROC)NULL, (LPARAM)hBmp, 0, rect.left, rect.top + 3, rect.right - rect.left, rect.bottom - rect.top, DSS_MONO | DST_BITMAP );
		DeleteObject( hBrush );
	}
	else
		DrawState( pDC, (HBRUSH)NULL, (DRAWSTATEPROC)NULL, (LPARAM)hBmp, 0, rect.left, rect.top + 1 + ( ( ( rect.bottom - rect.top ) - IMGHEIGHT )>>1), rect.right - rect.left, rect.bottom - rect.top, DSS_MONO | DST_BITMAP );
	DeleteObject( hBmp );

	return TRUE;
}

/////////////////
// Get a new CMyItemData from free chain or heap
//
CMyItemData* GetNewItemData()
{
	for (CMyItemData* pmd = m_pItemData; pmd; pmd=pmd->pnext) {
		if (!pmd->bInUse)
			break; // found a free one
	}
	if (pmd==NULL) {
		// no free item data: allocate a new one
		pmd = new CMyItemData;
		pmd->pnext = m_pItemData;
		m_pItemData = pmd;
	}
	pmd->bInUse = TRUE;
	return pmd;
}

/////////////////
// Delete CMyItemData: just free the string and mark as free
//
void DeleteItemData(CMyItemData* pmd)
{
	pmd->text.Empty();	// free string
	pmd->bInUse = FALSE; // mark as free
}

//////////////////
// See if a given item data pointer is one of mine.
// Ideally this should be a hash lookup, not a list--but how many
// menu items can there be?
//
BOOL IsMyItemData(CMyItemData* pItemData)
{
	for (CMyItemData* pmd = m_pItemData; pmd; pmd=pmd->pnext) {
		if (pmd == pItemData)
			return pmd->bInUse;
	}
	return FALSE;
}

CBoldDC::CBoldDC (HDC hDC, bool bBold) : m_hDC (hDC), m_hDefFont (NULL)
{
	LOGFONT lf;

	CFont::FromHandle((HFONT)::GetCurrentObject(m_hDC, OBJ_FONT))->GetLogFont (&lf);
	if ( ( bBold && lf.lfWeight != FW_BOLD) ||
		(!bBold && lf.lfWeight == FW_BOLD) )
	{
		lf.lfWeight = bBold ? FW_BOLD : FW_NORMAL;

		m_fontBold.CreateFontIndirect (&lf);
		m_hDefFont = (HFONT)::SelectObject (m_hDC, m_fontBold);
    }
}

/////////////////////////////////////////////////////////////////////////////
CBoldDC::~CBoldDC ()
{
    if ( m_hDefFont != NULL )
    {
        ::SelectObject (m_hDC, m_hDefFont);
    }
}

void CCoolmenu::SetMenuDefaults(COLORREF menuColors[MAX_COLOR])
{
	for(int i = 0; i < MAX_COLOR; i++)
	{
		this->menuColors[i] = menuColors[i];
		this->menuDefaultColors[i] = menuColors[i];
	}
	SetThemeColors();
	UpdateMenuBarColor(m_hMenu);
}
