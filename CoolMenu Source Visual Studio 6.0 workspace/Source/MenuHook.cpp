// MenuHook.cpp: implementation of the CMenuHook class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MenuHook.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMenuHook::CMenuHook()
{
	Init();
}

CMenuHook::~CMenuHook()
{

}

void CMenuHook::Init()
{
	mScreen.DeleteObject();
	bDrawBorder = TRUE;
	bSubMenu = FALSE;
	bMenuAbove = FALSE;
	mPoint = CPoint(0,0);
	xPosParent = -1;
	m_bReactOnThemeChange = FALSE;
}

HWND CMenuHook::GetHwnd()
{
	return m_pWndHooked;
}

void CMenuHook::SetSubMenu( BOOL bIsSubMenu )
{
	bSubMenu = bIsSubMenu;
}

int CMenuHook::GetXPos()
{
	return xPos;
}

void CMenuHook::SetXPosParent(int xPos)
{
	xPosParent = xPos;
}

LRESULT CMenuHook::WindowProc(UINT msg, WPARAM wp, LPARAM lp)
{
#ifdef _USRDLL
	// If this is a DLL, need to set up MFC state
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
#endif

	BOOL bCallDefault = TRUE;
	LRESULT result = NULL;
	WNDPROC wndProc = m_pOldWndProc;
	switch(msg)
	{
	case WM_NCCALCSIZE:
		{
			NCCALCSIZE_PARAMS* lpncsp = (NCCALCSIZE_PARAMS*)lp;
			int cx=0,cy=0;
			if(!m_CmGeneral->IsShadowEnabled())
			{
				cx = 5;
				cy = 6;
			}
			else
			{
				cy = cx = 2;
			}
			lpncsp->rgrc[0].left += 2;
			lpncsp->rgrc[0].top += 2;
			lpncsp->rgrc[0].right -= cx;
			lpncsp->rgrc[0].bottom -= cy;
			result = 0;
			return 0;
		}
		break;
	case WM_WINDOWPOSCHANGED:
	case WM_WINDOWPOSCHANGING:
		{
			LPWINDOWPOS pPos = (LPWINDOWPOS)lp;
			if(msg==WM_WINDOWPOSCHANGING)
			{
				OnWindowPosChanging( pPos );
			}
			if(!(pPos->flags&SWP_NOMOVE) )
			{
				if(mPoint==CPoint(0,0))
				{
					mPoint = CPoint(pPos->x,pPos->y);
				}
				else if(mPoint!=CPoint(pPos->x,pPos->y))
				{
					if(!(pPos->flags&SWP_NOSIZE))
					{
					}
					else
					{
						mPoint=CPoint(pPos->x,pPos->y);
						mScreen.DeleteObject();
					}
				}
			}
		}
		result = 0;
		return 1;
		break;
	case WM_CREATE:
		OnCreate();
		break;
	case WM_NCPAINT:
		{
			if(bDrawBorder)
			{
				HDC hDC;
				hDC = GetWindowDC (GetHwnd());
				if (OnDrawBorder( hDC, false ) )
				{
					OnDrawBorder(hDC,true);
				}
				ReleaseDC(GetHwnd(),hDC);
			}
		}
		return 0;
		break;
	case WM_ERASEBKGND:
		OnEraseBckground( (HDC)wp );
		result = 1;
		return 1;
		break;
	}
	if(bCallDefault)
		result = CallWindowProc( wndProc, GetHwnd(), msg, wp, lp );
	return result;
}

void CMenuHook::SetCurrentCoolmenu( CCoolmenu* cm )
{
	m_currentCoolmenu = cm;
}

void CMenuHook::SetMenuStyle( int menuStyle )
{
	m_menuStyle = menuStyle;
}

void CMenuHook::SetReactOnThemeChange( BOOL bReactOnThemeChange )
{
	m_bReactOnThemeChange = bReactOnThemeChange;
}

void CMenuHook::OnCreate()
{
	HWND hwnd = GetHwnd();
	LONG style = GetWindowLong( hwnd, GWL_STYLE );
	lStyle = style;
	style &= ~WS_BORDER;
	SetWindowLong( hwnd, GWL_STYLE , style);
	LONG exStyle = GetWindowLong( hwnd, GWL_EXSTYLE );
	lExStyle = exStyle;
	exStyle &= ~WS_EX_WINDOWEDGE;
	exStyle &= ~WS_EX_DLGMODALFRAME;
	SetWindowLong( hwnd, GWL_EXSTYLE , exStyle);
}

void CMenuHook::OnNcPaint( WPARAM wp )
{
	OnDrawBorder( NULL, TRUE );
}

BOOL CMenuHook::OnCalcFrameRect(LPRECT pRect)
{
	if(GetWindowRect(GetHwnd(),pRect))
	{
		pRect->top += 2;
		pRect->left += 2;
		if(!m_CmGeneral->IsShadowEnabled() && m_currentCoolmenu && m_currentCoolmenu->m_bShadowEnabled )
		{
			pRect->bottom -= 7;
			pRect->right -= 7;
		}
		else
		{
			pRect->bottom -= 3;
			pRect->right -= 3;
		}
		return TRUE;
	}
	return FALSE;
}

void CMenuHook::DrawShade( HDC hDC, CPoint screen)
{
	if(m_CmGeneral->IsShadowEnabled())
		return;

	// Get the size of the menu... 
	CRect Rect;
	GetWindowRect(GetHwnd(), Rect );

	int xStart = 0;
	if(m_currentCoolmenu && bMenuAbove && !bSubMenu)
	{
		CRect rectMenu = m_currentCoolmenu->GetLastMenuRect();
		xStart = rectMenu.Width() - 3;
	}

	long winW = Rect.Width(); 
	long winH = Rect.Height(); 
	long xOrg = screen.x;
	long yOrg = screen.y;

	CDC* pDC = CDC::FromHandle(hDC);
	CDC memDC;
	memDC.CreateCompatibleDC(pDC);
	CBitmap* pOldBitmap = memDC.SelectObject(&mScreen);

	HDC hDcDsk = memDC.m_hDC;
	xOrg = 0;
	yOrg = 0;

	int X,Y;
	// Simulate a shadow on right edge... 
	if (m_CmGeneral->NumScreenColors() <= 256)
	{
		DWORD rgb = ::GetSysColor(COLOR_BTNSHADOW);
		BitBlt(hDC,winW-2,0,2,winH,hDcDsk,xOrg+winW-2,0,SRCCOPY);
		BitBlt(hDC,0,winH-2,winW,2,hDcDsk,0,yOrg+winH-2,SRCCOPY);
		for (X=3; X<=4 ;X++)
		{
			for (Y=0; Y<4 ;Y++)
			{
				SetPixel(hDC,winW-X,Y,GetPixel(hDcDsk,xOrg+winW-X,yOrg+Y));
			}
			for (Y=4; Y<8 ;Y++)
			{
				SetPixel(hDC,winW-X,Y,rgb) ;
			}
			for (Y=8; Y<=(winH-5) ;Y++)
			{
				SetPixel( hDC, winW - X, Y, rgb);
			}
			for (Y=(winH-4); Y<=(winH-3) ;Y++)
			{
				SetPixel( hDC, winW - X, Y, rgb);
			}
		}
		// Simulate a shadow on the bottom edge... 
		for(Y=3; Y<=4 ;Y++)
		{
			for(X=0; X<=3 ;X++)
			{
				SetPixel(hDC,X,winH-Y, GetPixel(hDcDsk,xOrg+X,yOrg+winH-Y)) ;
			}
			for(X=4; X<=7 ;X++)
			{
				SetPixel( hDC, X, winH - Y, rgb) ;
			}
			for(X=8; X<=(winW-5) ;X++)
			{
				SetPixel( hDC, X, winH - Y, rgb) ;
			}
		}
	}
	else 
	{
		for (X=1; X<=4 ;X++)
		{
			for (Y=0; Y<4 ;Y++)
			{
				SetPixel(hDC,winW-X,Y, GetPixel(hDcDsk,xOrg+winW-X,yOrg+Y) );
			}
			for (Y=4; Y<8 ;Y++)
			{
				COLORREF c = GetPixel(hDcDsk, xOrg + winW - X, yOrg + Y) ;
				SetPixel(hDC,winW-X,Y,m_CmGeneral->DarkenColor(2* 3 * X * (Y - 3), c)) ;
			}
			for (Y=8; Y<=(winH-5) ;Y++)
			{
				COLORREF c = GetPixel(hDcDsk, xOrg + winW - X, yOrg + Y) ;
				SetPixel( hDC, winW - X, Y, m_CmGeneral->DarkenColor(2* 15 * X, c) );
			}
			for (Y=(winH-4); Y<=(winH-1) ;Y++)
			{
				COLORREF c = GetPixel(hDcDsk, xOrg + winW - X, yOrg + Y) ;
				SetPixel( hDC, winW - X, Y, m_CmGeneral->DarkenColor(2* 3 * X * -(Y - winH), c)) ;
			}
		}

		// Simulate a shadow on the bottom edge... 
		for(Y=1; Y<=4 ;Y++)
		{
			for(X=0; X<=3 ;X++)
			{
				if(X+xStart>winW-5) break;
				SetPixel(hDC,X + xStart,winH-Y, GetPixel(hDcDsk,xOrg+X+xStart,yOrg+winH-Y)) ;
			}
			for(X=4; X<=7 ;X++)
			{
				COLORREF c = GetPixel(hDcDsk, xOrg + X+xStart, yOrg + winH - Y) ;
				if(X+xStart>winW-5) break;
				SetPixel( hDC, X + xStart, winH - Y, m_CmGeneral->DarkenColor(2*3 * (X - 3) * Y, c)) ;
			}
			for(X=8; X<=(winW-5) ;X++)
			{
				COLORREF  c = GetPixel(hDcDsk, xOrg + X+xStart, yOrg + winH - Y); 
				if(X+xStart>winW-5) break;
				SetPixel( hDC, X + xStart, winH - Y, m_CmGeneral->DarkenColor(2* 15 * Y, c)) ;
			}
		}
	}
	memDC.SelectObject(pOldBitmap);
}

void CMenuHook::DrawFrame(CDC* pDC, CRect rectOuter, CRect rectInner, COLORREF crBorder)
{
	CRect Temp;
	rectInner.right -= 1;
	// Border top
	Temp.SetRect(rectOuter.TopLeft(),CPoint(rectOuter.right,rectInner.top));
	pDC->FillSolidRect(Temp,crBorder);
	// Border bottom
	Temp.SetRect(CPoint(rectOuter.left,rectInner.bottom),rectOuter.BottomRight());
	pDC->FillSolidRect(Temp,crBorder);

	// Border left
	Temp.SetRect(rectOuter.TopLeft(),CPoint(rectInner.left,rectOuter.bottom));
	pDC->FillSolidRect(Temp,crBorder);
	// Border right
	Temp.SetRect(CPoint(rectInner.right,rectOuter.top),rectOuter.BottomRight());
	pDC->FillSolidRect(Temp,crBorder);
}

void CMenuHook::SetScreenBitmap()
{
	if(mScreen.m_hObject==NULL)
	{
		// Get the desktop hDC... 
		HDC hDcDsk = GetWindowDC(0) ;
		CDC* pDcDsk = CDC::FromHandle(hDcDsk);

		CDC dc;
		dc.CreateCompatibleDC(pDcDsk);

		CRect rect;
		GetWindowRect(GetHwnd(),rect);
		mScreen.CreateCompatibleBitmap(pDcDsk,rect.Width()+10,rect.Height()+10);
		CBitmap* pOldBitmap = dc.SelectObject(&mScreen);
		dc.BitBlt(0,0,rect.Width()+10,rect.Height()+10,pDcDsk,mPoint.x,mPoint.y,SRCCOPY);

		dc.SelectObject(pOldBitmap);
		// Release the desktop hDC...
		ReleaseDC(0,hDcDsk);
	}
	return;
}

void CMenuHook::DrawSmalBorder( HDC hDC )
{
	
	CRect Rect, RectLastMenu;
	// Get the size of the menu... 
	GetWindowRect( GetHwnd(), Rect );
	Rect.OffsetRect(mPoint - Rect.TopLeft());
	RectLastMenu = m_currentCoolmenu->GetLastMenuRect();
	Rect &= RectLastMenu;
	if(!Rect.IsRectEmpty())
	{
		if(Rect.Width()>Rect.Height())
		{
			Rect.InflateRect(-1,0);
		}
		else
		{
			Rect.InflateRect(0,-1);
		}
		if(bMenuAbove && m_CmGeneral->IsShellType() != WinXP)
			Rect.bottom-=4;
		Rect.OffsetRect(-mPoint);
		CDC* pDC = CDC::FromHandle(hDC);

		COLORREF colorSmalBorder;
		if (m_CmGeneral->NumScreenColors() > 256) 
		{
			colorSmalBorder = m_currentCoolmenu->GetMenuColor( MenuBckColor );//m_CmGeneral->MixedColor(RGB(0,45,150),GetSysColor(COLOR_WINDOW));
		}
		else
		{
			colorSmalBorder = GetSysColor(COLOR_BTNFACE);
		}
		pDC->FillSolidRect(Rect,colorSmalBorder);
	}
}


BOOL CMenuHook::OnDrawBorder( HDC hDC, BOOL bOnlyBorder )
{
	HWND hwnd = GetHwnd();
	SetScreenBitmap();
	CRect rect,RectLastMenu;
	CRect client;
	CDC* pDC = CDC::FromHandle (hDC);
	BOOL bMenuAbove = FALSE;

	// Get the size of the menu... 
	GetWindowRect( hwnd, rect );
	GetClientRect( hwnd, client);
	CPoint offset( 0,0 );
	ClientToScreen( hwnd, &offset );
	client.OffsetRect(offset-rect.TopLeft());

	long winW = rect.Width(); 
	long winH = rect.Height(); 

	COLORREF crMenuBar = (m_currentCoolmenu)?m_currentCoolmenu->GetMenuColor( MenuBckColor ):RGB(163,194,245);//RGB(163,194,245);//m_CmGeneral->GetMenuColor(GetMenu(hwnd));
	COLORREF crWindow = GetSysColor(COLOR_WINDOW);
	COLORREF crThinBorder = m_CmGeneral->MixedColor(crWindow,crMenuBar);
	COLORREF clrBorder = (m_currentCoolmenu)?m_currentCoolmenu->GetMenuColor( SelectedPenColor ):m_CmGeneral->DarkenColor(128,crMenuBar);
//	COLORREF clrBorder = (m_currentCoolmenu)?m_currentCoolmenu->GetMenuColor( PenColor ):m_CmGeneral->DarkenColor(128,crMenuBar);
	COLORREF colorBitmap;

	if (m_CmGeneral->NumScreenColors() > 256) 
	{
		colorBitmap = m_CmGeneral->MixedColor(crMenuBar,crWindow);
	}
	else
	{
		colorBitmap = GetSysColor(COLOR_BTNFACE);
	}

	// Better contrast when you have less than 256 colors
	if(pDC->GetNearestColor(crThinBorder)==pDC->GetNearestColor(colorBitmap))
	{
		crThinBorder = crWindow;
		colorBitmap = crMenuBar;
	} 

	 BOOL bHighContrast = FALSE;

	if(!m_CmGeneral->IsShadowEnabled())
	{
		if(!bOnlyBorder)
		{
			DrawFrame(pDC,CRect(CPoint(1,1),CSize(winW-6,winH-6)),client,crThinBorder );
		}
		if(bHighContrast)
		{
			pDC->Draw3dRect(CRect(CPoint(0,0),CSize(winW-4,winH-4)),GetSysColor(COLOR_BTNTEXT ),GetSysColor(COLOR_BTNTEXT ));
		}
		else
		{
			pDC->Draw3dRect(CRect(CPoint(1,1),CSize(winW-6,winH-6)),crThinBorder,crThinBorder);
//			pDC->Draw3dRect(CRect(CPoint(0,0),CSize(winW-4,winH-4)),(m_bReactOnThemeChange?clrBorder:(m_menuStyle==MenuStyleOffice2K3?RGB(0,45,150):RGB(138,134,122))),m_menuStyle==MenuStyleOffice2K3?RGB(0,45,150):RGB(138,134,122));
			pDC->Draw3dRect(CRect(CPoint(0,0),CSize(winW-4,winH-4)),(m_bReactOnThemeChange?clrBorder:(m_currentCoolmenu?m_currentCoolmenu->GetMenuColor( SelectedPenColor ):RGB(138,134,122))),m_currentCoolmenu?m_currentCoolmenu->GetMenuColor( SelectedPenColor ):RGB(138,134,122));
			if(m_currentCoolmenu && m_currentCoolmenu->m_bShadowEnabled ) DrawShade(hDC,mPoint);
		}
	}
	else
	{
		if(!bOnlyBorder)
		{
			DrawFrame(pDC,CRect(CPoint(1,1),CSize(winW-2,winH-2)),client,GetSysColor( COLOR_MENU ));
		}

		if(bHighContrast)
		{
			pDC->Draw3dRect(CRect(CPoint(0,0),CSize(winW-0,winH-0)),GetSysColor(COLOR_BTNTEXT ),GetSysColor(COLOR_BTNTEXT ));
		}
		else
		{
			pDC->Draw3dRect(CRect(CPoint(1,1),CSize(winW-2,winH-2)),crThinBorder,crThinBorder);
			if(m_bReactOnThemeChange)
				pDC->Draw3dRect(CRect(CPoint(0,0),CSize(winW-0,winH-0)),clrBorder,clrBorder);
			else
			{
//				pDC->Draw3dRect(CRect(CPoint(0,0),CSize(winW-0,winH-0)),m_menuStyle==MenuStyleOffice2K3?RGB(0,45,150):RGB(138,134,122),m_menuStyle==MenuStyleOffice2K3?RGB(0,45,150):RGB(138,134,122));
				pDC->Draw3dRect(CRect(CPoint(0,0),CSize(winW-0,winH-0)),(m_bReactOnThemeChange?clrBorder:(m_currentCoolmenu?m_currentCoolmenu->GetMenuColor( SelectedPenColor ):RGB(138,134,122))),m_currentCoolmenu?m_currentCoolmenu->GetMenuColor( SelectedPenColor ):RGB(138,134,122));
			}
		}
	}
	if(m_currentCoolmenu && m_currentCoolmenu->GetCoolMenubar()) DrawSmalBorder(  hDC );
	return TRUE;
}

void CMenuHook::OnEraseBckground( HDC hDC )
{
	CDC* pDC = CDC::FromHandle (hDC);
	CRect Rect;
	GetClientRect( GetHwnd(), Rect );
	Rect.InflateRect(+2,0,-1,0);
	pDC->FillSolidRect (Rect,GetSysColor(COLOR_MENU));
}

BOOL CMenuHook::OnWindowPosChanging( LPWINDOWPOS wp )
{
	bMenuAbove = FALSE;
	POINT lp;
	lp.x = wp->x;
	lp.y = wp->y;
	ClientToScreen(GetHwnd(),&lp);
	xPos = lp.x;

	if(m_currentCoolmenu)
	{
		CRect rect = m_currentCoolmenu->GetLastMenuRect();
		if( rect.top > lp.y )
			bMenuAbove = TRUE & !bSubMenu;
	}

	if(bSubMenu)
	{
		if( xPos > xPosParent )
			wp->x += 3;
		wp->y += 3;
	}
	if(m_CmGeneral->IsShadowEnabled())
	{
		wp->cx -= 2;
		wp->cy -= 2;
	}
	else
	{
		wp->cx += 2;
		wp->cy += 2;
	}
	(bMenuAbove)?wp->y+=3:wp->y--;
	return TRUE;
}
