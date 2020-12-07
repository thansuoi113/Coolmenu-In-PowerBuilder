// MenuHook.h: interface for the CMenuHook class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_MENUHOOK_H__16356D3F_DEE5_4E7B_A55A_BDB972DA7473__INCLUDED_)
#define AFX_MENUHOOK_H__16356D3F_DEE5_4E7B_A55A_BDB972DA7473__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Coolmenu.h"
#include "CoolmenuGeneral.h"

class CMenuHook : public CMsgHook  
{
public:
	CMenuHook();
	virtual ~CMenuHook();

	void Init();
	HWND GetHwnd();
	void SetSubMenu( BOOL bIsSubMenu );
	void SetCurrentCoolmenu( CCoolmenu* cm );
	void SetMenuStyle( int menuStyle );
	void SetReactOnThemeChange( BOOL bReactOnThemeChange );
	int GetXPos();
	void SetXPosParent(int xPos);

protected:
	virtual LRESULT WindowProc(UINT msg, WPARAM wp, LPARAM lp);
	void OnCreate();
	void OnNcPaint(WPARAM wp );
	void OnEraseBckground(HDC hDC );
	BOOL OnWindowPosChanging( LPWINDOWPOS wp );
	BOOL OnDrawBorder(HDC hDC, BOOL bOnlyBorder );
	void DrawSmalBorder(HDC hDC);
	void DrawFrame(CDC* pDC, CRect rectOuter, CRect rectInner, COLORREF crBorder);
	void DrawShade(HDC hDC, CPoint screen);
	BOOL OnCalcFrameRect(LPRECT pRect);
	void SetScreenBitmap();

	CPoint				mPoint;
	CBitmap				mScreen;
	BOOL					bDrawBorder, bSubMenu, bMenuAbove, m_bReactOnThemeChange;
	LONG					lStyle;
	LONG					lExStyle;
	CCoolmenuGeneral*	m_CmGeneral;
	int					m_menuStyle, xPos, xPosParent;
	CCoolmenu*			m_currentCoolmenu;
};

#endif // !defined(AFX_MENUHOOK_H__16356D3F_DEE5_4E7B_A55A_BDB972DA7473__INCLUDED_)
