/*	This file is the entrypoint for the coolmenu dll.

	To make coolmenu work with more than one thread and with more than
	one main window, I've made some changes.

	For every thread a hook is set by calling SetWindowsHookEx. This is used
	to examine what windows are opened within the thread. Messages are sent
	through hook function before they are sent through. By looking at the
	class of the window to be created a hook is set for the window. This is
	done by creating an instance of the CCoolmenu class. All messages for
	this window are then handled by that instance.

	When all windows for a certain thread are closed, the thread is unhooked.

	A third hook mechanism is used for menuwindows. By looking for window class
	"#32768" I know a menuwindow is about to be created. When the bDrawFlatBorder
	boolean is set to TRUE, that is, give menus a flat border as it is within
	Office 2003, the menuwindow is hooked and the borderstyle is changed (see
	MenuHook.cpp).

	All settings that could be set through functions on n_coolmenu are applied
	to all windows within a thread. You can't change settings for one window
	within a thread without changing it for any other windows within that thread.

	Also there is one imagelist for every thread. The handle from this imagelist
	is set within the CCoolMenu class instance.
*/
#include "stdafx.h"
#include "CoolmenuMain.h"

///////////////////////////////////////////////////////////////////////////////
//	CHookData
///////////////////////////////////////////////////////////////////////////////
CHookData* GetNewHookData( DWORD threadID )
{
	for (CHookData* hd = m_pHookData; hd; hd=hd->pnext) {
		if (!hd->bInUse)
			break; // found a free one
		else if( hd->threadID == threadID )
			return hd;
	}
	if (hd==NULL) {
		// no free item data: allocate a new one
		hd = new CHookData;
		hd->pnext = m_pHookData;
		m_pHookData = hd;
	}
	//	Do some initializations
	//	Build 120: ILC_COLOR24 replaced by ILC_COLOR32.
//	hd->hImagelist = ImageList_Create( 16, 16, ILC_MASK | ILC_COLOR32, 30, 5 );
	hd->hImagelist.Create( 16, 16, ILC_MASK | ILC_COLORDDB, 0, 10 );
	hd->hImlDisabled.Create( 16, 16, ILC_MASK | ILC_COLORDDB, 0, 10 );
	hd->bInUse = TRUE;
	hd->hwndList = new CDWordArray();
	hd->menuStyle = MenuStyleOffice2K3;
	hd->bDrawFlatBorder = TRUE;
	hd->bCoolMenubar = TRUE;
	//	Build 125: bReactOnThemeChange added.
	hd->bReactOnThemeChange = FALSE;
	hd->bShadowEnabled = TRUE;
	hd->bOrigCoolMenubar = TRUE;
	hd->bOrigDrawFlatBorder = TRUE;
	hd->hResources = NULL;
	for(int i=0;i<MAX_COLOR;i++)
		hd->colors[i] = m_default2K3Colors[i];
	return hd;
}

//	Get the CHookData instance for the current thread.
CHookData* GetHookData(DWORD threadID)
{
	for (CHookData* hd = m_pHookData; hd; hd=hd->pnext) {
		if (hd->threadID == threadID)
			return hd;
	}
	return NULL;
}

void DeleteHookData(CHookData* hd)
{
	hd->bInUse = FALSE; // mark as free
	hd->threadID = -1;
	hd->hHook = NULL;
	hd->hHookMsg = NULL;
	if(hd->hResources) FreeLibrary(hd->hResources);
	//ImageList_Destroy( hd->hImagelist );
	hd->hImagelist.DeleteImageList();
	hd->hImlDisabled.DeleteImageList();
	delete hd->hwndList;
}

BOOL IsHooked(DWORD threadID)
{
	for (CHookData* hd = m_pHookData; hd; hd=hd->pnext) {
		if (hd->bInUse && hd->threadID == threadID && hd->hHook > 0 )
		{
			return TRUE;
		}
	}
	return FALSE;
}

HHOOK GetHook(DWORD threadID)
{
	for (CHookData* hd = m_pHookData; hd; hd=hd->pnext) {
		if (hd->threadID == threadID)
			return hd->hHook;
	}
	return NULL;
}

void AddWindow(CHookData* hd, HWND hwnd)
{
	BOOL bHwndFound = FALSE;
	for(int i =0;i<hd->hwndList->GetSize();i++)
	{
		if(hd->hwndList->GetAt(i) == (DWORD)hwnd )
		{
			bHwndFound = TRUE;
			break;
		}
	}
	if(!bHwndFound)
		hd->hwndList->Add( (DWORD)hwnd );
}

CHookData* Hook(DWORD threadID)
{
	CHookData*	hd = GetNewHookData(threadID);
	hd->threadID = threadID;
	HINSTANCE hInst = AfxGetInstanceHandle();
	hd->hHook = SetWindowsHookEx( WH_CALLWNDPROC, (HOOKPROC)CallWindowProc, hInst, threadID );
	hd->hHookMsg = SetWindowsHookEx( WH_GETMESSAGE, (HOOKPROC)GetMsgProc, hInst, threadID );
	return hd;
}

void Unhook(DWORD threadID)
{
	CHookData* hd = GetHookData( threadID );
	UnhookWindowsHookEx( hd->hHook );
	UnhookWindowsHookEx( hd->hHookMsg );
	DeleteHookData(hd);
}

///////////////////////////////////////////////////////////////////////////////
//	CMenuHookList
///////////////////////////////////////////////////////////////////////////////
CMenuHookList::CMenuHookList()
{
}

CMenuHookList::~CMenuHookList()
{
	CMenuHook* mh;
	while (!IsEmpty())
	{
		mh = (CMenuHook*)RemoveTail();
		delete mh;
	}
}

CMenuHook* CMenuHookList::GetNewMenuHook( HWND hWnd )
{
	POSITION pos;
	CMenuHook* mh;
	BOOL mhFound = FALSE;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		mh = (CMenuHook*)GetNext( pos );
		if ( mh->GetHwnd() == hWnd )
		{
			mhFound = TRUE;
			break;
		}
	}
	if (!mhFound)
	{
		//	See if there's a menuhook object for a window that's no longer valid.
		for( pos = GetHeadPosition(); pos != NULL; )
		{

			mh = (CMenuHook*)GetNext( pos );
			if ( !IsWindow( mh->GetHwnd() ) )
			{
				mh->Init();
				return mh;
			}
		}
		mh = new CMenuHook();
		AddTail( mh );
		return mh;
	}
	return NULL;
}

CMenuHook* CMenuHookList::HookWindow( HWND hWnd )
{
	CMenuHook* mh = GetNewMenuHook( hWnd );
	if(!mh) return NULL;
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData(threadID);
	int xPos;
	mh->SetReactOnThemeChange(hd->bReactOnThemeChange);
	mh->SetSubMenu( IsSubMenu(xPos) );
	mh->SetXPosParent(xPos);
	mh->HookWindow( hWnd );
	mh->SetCurrentCoolmenu(m_currentCoolmenu);
	mh->SetMenuStyle(hd->menuStyle);
	return mh;
}

CMenuHook* CMenuHookList::GetMenuHook( HWND hWnd )
{
	POSITION pos;
	CMenuHook* mh;
	BOOL mhFound = FALSE;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		mh = (CMenuHook*)GetNext( pos );
		if ( mh->GetHwnd() == hWnd )
		{
			mhFound = TRUE;
			break;
		}
	}
	if (mhFound) return mh;
	return NULL;
}

BOOL CMenuHookList::IsSubMenu(int& xPos)
{
	POSITION pos;
	CMenuHook* mh;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		mh = (CMenuHook*)GetNext( pos );
		if ( mh->m_pOldWndProc != NULL )
		{
			xPos = mh->GetXPos();
			return TRUE;
		}
	}
	return FALSE;
}
///////////////////////////////////////////////////////////////////////////////
//	CCoolmenuList
///////////////////////////////////////////////////////////////////////////////
CCoolmenuList::CCoolmenuList()
{
}

CCoolmenuList::~CCoolmenuList()
{
	CCoolmenu* cm;
	while (!IsEmpty())
	{
		cm = (CCoolmenu*)RemoveTail();
		delete cm;
	}
}

void CCoolmenuList::SetImagelist( HIMAGELIST hSharedImagelist )
{
	POSITION pos;
	CCoolmenu* cm;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if (cm)
		{
			//cm->SetImagelist( hSharedImagelist );
		}
	}
}

void CCoolmenuList::SetImageIdlist( CMapStringToPtr* imageIdlist )
{
	POSITION pos;
	CCoolmenu* cm;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if (cm)
		{
			cm->SetImageIdlist( imageIdlist );
		}
	}
}

CCoolmenu* CCoolmenuList::GetNewCoolmenu( HWND hWnd )
{
	POSITION pos;
	CCoolmenu* cm;
	BOOL cmFound = FALSE;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if ( cm->GetHwnd() == hWnd )
		{
			cmFound = TRUE;
			break;
		}
	}
	//	Current window handle not found, so installcoolmenu function not called before for this window
	if (!cmFound)
	{
		//	See if there's a coolmenu object for a window that's no longer valid.
		for( pos = GetHeadPosition(); pos != NULL; )
		{

			cm = (CCoolmenu*)GetNext( pos );
			if ( !IsWindow( cm->GetHwnd() ) )
			{
				return cm;
			}
		}
		cm = new CCoolmenu();
		AddTail( cm );
		return cm;
	}
	return NULL;
}

BOOL CCoolmenuList::IsHooked( HWND hWnd )
{
	POSITION pos;
	CCoolmenu* cm;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if ( cm->GetHwnd() == hWnd )
		{
			return TRUE;
		}
	}
	return FALSE;
}

CCoolmenu* CCoolmenuList::HookWindow( HWND hWnd, HMENU hMenu )
{
	if(IsHooked(hWnd)) return NULL;
	CCoolmenu* cm = GetNewCoolmenu( hWnd );
	if(!cm) return NULL;
	cm->Initialize();
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData(threadID);
	AddWindow(hd,hWnd);
	cm->HookWindow( hWnd );
	cm->SetReactOnThemeChange(hd->bReactOnThemeChange);
	cm->SetCoolMenubar(hd->bCoolMenubar);
	cm->SetPB105OrAbove( m_bPB105OrAbove );
	cm->SetShadowEnabled(hd->bShadowEnabled);
	cm->SetMenuStyle(hd->menuStyle);
	cm->SetMenuDefaults(hd->colors);
	cm->SetImagelist( &hd->hImagelist, &hd->hImlDisabled );
	cm->SetMenuItemTextBoldlist( &hd->m_menuItemTextBoldList );
	cm->SetMenuItemTextItaliclist( &hd->m_menuItemTextItalicList );
	cm->SetMenuItemTextUnderlinelist( &hd->m_menuItemTextUnderlineList);
	cm->SetMenuItemTextColorlist( &hd->m_menuItemTextColorList );
	cm->SetImageIdlist( &hd->m_imageIdList );
	cm->SetThreadID( threadID );
	return cm;
}

CCoolmenu* CCoolmenuList::GetCoolmenu( HWND hWnd )
{
	POSITION pos;
	CCoolmenu* cm;
	BOOL cmFound = FALSE;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if ( cm->GetHwnd() == hWnd )
		{
			cmFound = TRUE;
			break;
		}
	}
	if (cmFound) return cm;
	if(GetParent(hWnd))
		return GetCoolmenu(GetParent(hWnd));
	return NULL;
}

void CCoolmenuList::SetMenuStyle( int menuStyle )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData*	hd = GetHookData(threadID);

	if(menuStyle==hd->menuStyle) return;

	if(menuStyle==MenuStyleNormal || menuStyle == MenuStyleOffice2K)
	{
		if(hd->menuStyle==MenuStyleOfficeXP || hd->menuStyle == MenuStyleOffice2K3 || hd->menuStyle == MenuStyleOffice2K7 )
		{
			hd->bOrigCoolMenubar = hd->bCoolMenubar;
			hd->bOrigDrawFlatBorder = hd->bDrawFlatBorder;
			hd->bCoolMenubar = FALSE;
			hd->bDrawFlatBorder = FALSE;
		}
	}
	else
	{
		hd->bCoolMenubar = hd->bOrigCoolMenubar;
		hd->bDrawFlatBorder = hd->bOrigDrawFlatBorder;
	}
	hd->menuStyle = menuStyle;
	POSITION pos;
	CCoolmenu* cm;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if(cm->GetThreadID()==threadID)
		{
			cm->SetCoolMenubar(hd->bCoolMenubar);
			cm->SetMenuStyle(menuStyle);
			cm->SetMenuDefaults(menuStyle==MenuStyleOffice2K3?m_default2K3Colors:(menuStyle==MenuStyleOffice2K7)?m_default2K7Colors:m_defaultXpColors);
			cm->SetThemeColors();
		}
	}
	for(int i=0;i<MAX_COLOR;i++)
		hd->colors[i] = (hd->menuStyle==MenuStyleOffice2K3?m_default2K3Colors[i]:(menuStyle==MenuStyleOffice2K7)?m_default2K7Colors[i]:m_defaultXpColors[i]);
}

int CCoolmenuList::GetMenuStyle()
{
	DWORD threadID = GetCurrentThreadId();
	CHookData*	hd = GetHookData(threadID);

	if(hd) return hd->menuStyle;
	return -1;
}

void CCoolmenuList::SetMenuColor( int colorIndex, COLORREF newColor )
{
	if(colorIndex<0 || colorIndex > MAX_COLOR) return;
	DWORD threadID = GetCurrentThreadId();
	POSITION pos;
	CCoolmenu* cm;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if(cm->GetThreadID()==threadID)
		{
			cm->SetMenuColor(colorIndex,newColor);
		}
	}
	CHookData*	hd = GetHookData(threadID);
	hd->colors[colorIndex] = newColor;
}

void CCoolmenuList::SetShadowEnabled( BOOL bShadowEnabled )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	if(hd->bShadowEnabled == bShadowEnabled) return;
	hd->bShadowEnabled = bShadowEnabled;
	POSITION pos;
	CCoolmenu* cm;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if(cm->GetThreadID()==threadID)
		{
			cm->SetShadowEnabled(bShadowEnabled);
		}
	}
	if(m_currentCoolmenu) m_currentCoolmenu->SetShadowEnabled(bShadowEnabled);
}

void CCoolmenuList::SetCoolMenubar( BOOL bCoolMenubar )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	if(hd->bCoolMenubar == bCoolMenubar) return;
	POSITION pos;
	CCoolmenu* cm;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if(cm->GetThreadID()==threadID)
		{
			cm->SetCoolMenubar(bCoolMenubar);
		}
	}
	if(hd->menuStyle==MenuStyleNormal || hd->menuStyle==MenuStyleOffice2K)
	{
		hd->bOrigCoolMenubar = bCoolMenubar;
	}
	else
	{
		hd->bOrigCoolMenubar = bCoolMenubar;
		hd->bCoolMenubar = bCoolMenubar;
	}
}

//	Build 125: added to enable coolmenu to react on theme or color changes,
void CCoolmenuList::SetReactOnThemeChange( BOOL bReactOnThemeChange )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	if(hd->bReactOnThemeChange == bReactOnThemeChange) return;
	POSITION pos;
	CCoolmenu* cm;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if(cm->GetThreadID()==threadID)
		{
			cm->SetReactOnThemeChange(bReactOnThemeChange);
		}
	}
	hd->bReactOnThemeChange = bReactOnThemeChange;
}

BOOL CCoolmenuList::GetReactOnThemeChange()
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	return hd->bReactOnThemeChange;
}

int CCoolmenuList::AddStockImage( int iImage, BOOL bBitmap, HMODULE hModule, COLORREF bckColor, INT xPixel, INT yPixel )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	if(hd) return m_CmGeneral->AddStockImage( &hd->hImagelist, &hd->hImlDisabled, iImage, bBitmap, hModule, bckColor, xPixel, yPixel );
	return -1;
}

void CCoolmenuList::SetImageNameIdW( WCHAR* iName, int iImage )
{
	DWORD threadID = GetCurrentThreadId();
	CString cName( iName, lstrlenW( iName) );
	LPSTR m_szBuffer = new CHAR [lstrlenW( iName)];
	strcpy( m_szBuffer, cName );
	CCoolmenuList::SetImageNameId( (LPCTSTR)m_szBuffer, iImage );
	delete [] m_szBuffer;
}

void CCoolmenuList::SetImageNameId(  LPCTSTR iName, int iImage )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	hd->m_imageIdList.SetAt( iName, (void*&)iImage );
}

int CCoolmenuList::AddMenuImageW( WCHAR* sImage, HICON hIcon )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	int iIcon = m_CmGeneral->AddMenuImageW( &hd->hImagelist, &hd->hImlDisabled, sImage, hIcon );
	// Build 199: add menuitemtext and iconindex to list which is used to determine
	// what image to show for which menuitem.
	if(iIcon>=0) SetImageNameIdW( sImage, iIcon );
	return iIcon;
}

int CCoolmenuList::AddMenuImage( LPCTSTR sImage, HICON hIcon )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	int iIcon = m_CmGeneral->AddMenuImage( &hd->hImagelist, &hd->hImlDisabled, sImage, hIcon );
	// Build 199: add menuitemtext and iconindex to list which is used to determine
	// what image to show for which menuitem.
	if(iIcon>=0) SetImageNameId( sImage, iIcon );
	return iIcon;
}

int CCoolmenuList::AddImageW( WCHAR* sImage, COLORREF bckColor, INT xPixel, INT yPixel )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	return m_CmGeneral->AddImageW( &hd->hImagelist, &hd->hImlDisabled, hd->hResources, sImage, bckColor, xPixel, yPixel );
}

int CCoolmenuList::AddImage( LPCTSTR sImage, COLORREF bckColor, INT xPixel, INT yPixel )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	return m_CmGeneral->AddImage( &hd->hImagelist, &hd->hImlDisabled, hd->hResources, sImage, bckColor, xPixel, yPixel );
}

int CCoolmenuList::AddImageID( WORD wImage, COLORREF bckColor, INT xPixel, INT yPixel )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	return m_CmGeneral->AddImageID( &hd->hImagelist, &hd->hImlDisabled, hd->hResources, wImage, bckColor, xPixel, yPixel );
}

void CCoolmenuList::SetDefaults()
{
	DWORD threadID = GetCurrentThreadId();
	CHookData*	hd = GetHookData(threadID);
	for(int i=0;i<MAX_COLOR;i++)
		hd->colors[i] = (hd->menuStyle==MenuStyleOffice2K3?m_default2K3Colors[i]:(MenuStyleOffice2K7)?m_default2K7Colors[i]:m_defaultXpColors[i]);
	POSITION pos;
	CCoolmenu* cm;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cm = (CCoolmenu*)GetNext( pos );
		if(cm->GetThreadID()==threadID)
			cm->SetMenuDefaults(hd->colors);
	}
}

void CCoolmenuList::SetMenuItemProps( LPCTSTR sName, BOOL bBold, BOOL bItalic, BOOL bUnderline, COLORREF clrText )
{
	DWORD threadID = GetCurrentThreadId();
	CHookData* hd = GetHookData( threadID );
	hd->m_menuItemTextColorList.SetAt( sName, (void*&)clrText );
	hd->m_menuItemTextBoldList.SetAt( sName, (void*&)bBold );
	hd->m_menuItemTextItalicList.SetAt( sName, (void*&)bItalic );
	hd->m_menuItemTextUnderlineList.SetAt( sName, (void*&)bUnderline );
}

///////////////////////////////////////////////////////////////////////////////
//	CCoolbarImageList
///////////////////////////////////////////////////////////////////////////////
CCoolbarImageList::CCoolbarImageList( int ctlID )
{
	m_ctlID = ctlID;
}

CCoolbarImageList::~CCoolbarImageList()
{

}

void CCoolbarImageList::CreateImageLists( int cx, int cy, int cxLarge, int cyLarge )
{
	m_cx = cx;
	m_cy = cy;
	m_cxLarge = cxLarge;
	m_cyLarge = cyLarge;
	CCoolbarImagelist.Create( cx, cy, ILC_MASK | ILC_COLOR32, 0, 10 );
	CCoolbarHotImagelist.Create( cx, cy, ILC_MASK | ILC_COLOR32, 0, 10 );
	CCoolbarDisabledImagelist.Create( cx, cy, ILC_MASK | ILC_COLOR32, 0, 10 );
	CCoolbarLargeImagelist.Create( cxLarge, cyLarge, ILC_MASK | ILC_COLORDDB, 0, 10 );
	CCoolbarLargeHotImagelist.Create( cxLarge, cyLarge, ILC_MASK | ILC_COLOR32, 0, 10 );
	CCoolbarLargeDisabledImagelist.Create( cxLarge, cyLarge, ILC_MASK | ILC_COLOR32, 0, 10 );
}

void CCoolbarImageList::DestroyImageLists()
{
	CCoolbarImagelist.DeleteImageList();
	CCoolbarHotImagelist.DeleteImageList();
	CCoolbarDisabledImagelist.DeleteImageList();
	CCoolbarLargeImagelist.DeleteImageList();
	CCoolbarLargeHotImagelist.DeleteImageList();
	CCoolbarLargeDisabledImagelist.DeleteImageList();
}

void CCoolbarImageList::SetIconSize( BOOL bSmall, int imagelistType, int cx, int cy )
{
	if( bSmall)
	{
		if(imagelistType == il_Normal)
			ImageList_SetIconSize( CCoolbarImagelist.m_hImageList, cx, cy );
		else
			ImageList_SetIconSize( CCoolbarDisabledImagelist.m_hImageList, cx, cy );
	}
	else
	{
		if(imagelistType == il_Normal)
			ImageList_SetIconSize( CCoolbarLargeImagelist.m_hImageList, cx, cy );
		else
			ImageList_SetIconSize( CCoolbarLargeDisabledImagelist.m_hImageList, cx, cy );
	}
}

int CCoolbarImageList::GetCtlID()
{
	return m_ctlID;
}

///////////////////////////////////////////////////////////////////////////////
//	CCoolbarImageListList
///////////////////////////////////////////////////////////////////////////////
CCoolbarImageListList::CCoolbarImageListList()
{
	ctlID = 0;
}

CCoolbarImageListList::~CCoolbarImageListList()
{
}

CCoolbarImageList* CCoolbarImageListList::GetNewCoolbarImagelist()
{
	//POSITION pos;
	CCoolbarImageList* cbIl;
	//	See if there's a coolmenu object for a window that's no longer valid.
/*	for( pos = GetHeadPosition(); pos != NULL; )
	{

		cbIl = (CCoolbarImageList*)GetNext( pos );
		if ( cbIl->GetCtlID() == 0 )
		{
			return cbIl;
		}
	}*/
	ctlID++;
	cbIl = new CCoolbarImageList( ctlID );
	AddTail( cbIl );
	return cbIl;
}

CCoolbarImageList* CCoolbarImageListList::GetCoolbarImagelist( int CtlID )
{
	POSITION pos;
	CCoolbarImageList* cbIl;
	BOOL cbIlFound = FALSE;
	for( pos = GetHeadPosition(); pos != NULL; )
	{
		cbIl = (CCoolbarImageList*)GetNext( pos );
		if ( cbIl->GetCtlID() == CtlID )
		{
			cbIlFound = TRUE;
			break;
		}
	}
	if (cbIlFound) return cbIl;
	return NULL;
}


BOOL IsSystemMenu( HWND hWnd, BOOL bSysMenu, HMENU hMenu )
{
	if ( bSysMenu ) return TRUE;

	HMENU sysMenu = GetSystemMenu( hWnd, FALSE );
	if ( hMenu == sysMenu ) return TRUE;

	MENUITEMINFO mii;
	mii.fMask = MIIM_ID | MIIM_SUBMENU;
	mii.cbSize = sizeof( MENUITEMINFO );
	GetMenuItemInfo( hMenu, 0, TRUE, &mii );
	if ( mii.hSubMenu > 0 ) return FALSE;

	if ( mii.wID >= 0xF000 ) return TRUE;

	return FALSE;
}

//	Messages posted to the current thread are sent through this function.
LRESULT CALLBACK GetMsgProc( int nCode, WPARAM wParam, LPARAM lParam )
{
	CHookData* hd = GetHookData( GetCurrentThreadId() );
	HHOOK hHook = hd->hHookMsg;

	if (nCode < 0 )
		return CallNextHookEx( hHook, nCode, wParam, lParam );

	if (nCode == HC_ACTION)
	{
		MSG* msg = (MSG*)lParam;
		switch (msg->message)
		{
  		case WM_MDISETMENU:
 			{
 				CCoolmenu* cm = m_coolmenuList.GetCoolmenu(msg->hwnd);
 				if(cm)
 				{
					if( (HMENU)msg->wParam || (HMENU)msg->lParam )
						cm->SetMenuSet( TRUE );
 					/*if((HMENU)msg->wParam)
 					{
 						cm->SetMenu((HMENU)msg->wParam);
 					}
 					else if((HMENU)msg->lParam)
					{
 						cm->SetMenu((HMENU)msg->lParam);
					}*/
 				}
 			}
			break;
		case WM_LBUTTONDOWN:
			{
				if(m_isContextMenu) m_isContextMenu = FALSE;
				CCoolmenu* cm = m_coolmenuList.GetCoolmenu(msg->hwnd);
				if(cm)
					cm->SetInitMenu( TRUE );
			}
			break;
// 		case WM_KEYDOWN:
//			{
//				CCoolmenu* cm = m_coolmenuList.GetCoolmenu(msg->hwnd);
//				if(cm)
//				{
//					cm->SetKeyDown( TRUE );
//				}
//			}
//			break;
		}
	}
	return CallNextHookEx( hHook, nCode, wParam, lParam );
}

//	Messages sent to the current thread are sent through this function.
LRESULT CALLBACK CallWindowProc( int nCode, WPARAM wParam, LPARAM lParam )
{
	CHookData* hd = GetHookData( GetCurrentThreadId() );
	HHOOK hHook = hd->hHook;
	CWPSTRUCT* cwps = (CWPSTRUCT*)lParam;

	if (nCode < 0 )
		return CallNextHookEx( hHook, nCode, wParam, lParam );
	if (nCode == HC_ACTION)
	{
		switch (cwps->message)
		{
		case WM_CREATE:
			{
				TCHAR Name[20];
				int Count = GetClassName (cwps->hwnd,Name,ARRAY_SIZE(Name));
				// Check for the menu-class
				CREATESTRUCT* cs = (CREATESTRUCT*)lParam;
				if(_tcscmp(Name,_T("#32768"))==0 && !m_isSystemMenu && !m_isContextMenu && hd->bDrawFlatBorder)
				{
					m_menuHookList.HookWindow(cwps->hwnd);
				}
				else if (_tcscmp(Name,_T(m_sWindowClass))==0)
				{
					long style = GetWindowLong( cwps->hwnd, GWL_STYLE );
					if (style&WS_CHILDWINDOW)
					{
					}
					else
					{
						m_coolmenuList.HookWindow( cwps->hwnd, cs->hMenu );
					}
				}
				else if (_tcscmp(Name,_T(m_sPopupWindowClass))==0)
				{
					long exStyle = GetWindowLong( cwps->hwnd, GWL_EXSTYLE );
					CREATESTRUCT* cs = (CREATESTRUCT*)lParam;
					if (cs->hwndParent)
					{
						m_coolmenuList.HookWindow( cwps->hwnd, cs->hMenu );
					}
				}
				else if(_tcscmp(Name,_T(m_sToolbarClass))==0)
				{
					m_coolmenuList.HookWindow( cwps->hwnd, cs->hMenu );
				}
				else if(_tcscmp(Name,_T(m_sFloatToolbarClass))==0)
				{
					m_coolmenuList.HookWindow( cwps->hwnd, cs->hMenu );
				}
			}
			break;
		case WM_INITMENUPOPUP:
			{
				CCoolmenu* cmTemp = m_coolmenuList.GetCoolmenu( cwps->hwnd );
				m_isSystemMenu = IsSystemMenu(cwps->hwnd, (BOOL)HIWORD(cwps->lParam), (HMENU)cwps->wParam );
				if(cmTemp)
				{
					m_currentCoolmenu = cmTemp;
					if(GetSubMenu( GetMenu(m_currentCoolmenu->GetHwnd()), (UINT)LOWORD(cwps->lParam))==(HMENU)cwps->wParam)
						m_isContextMenu = FALSE;
				}
			}
			break;
		case WM_UNINITMENUPOPUP:
			{
				//MessageBox( NULL, "", "WM_UNINITMENUPOPUP", 0 );
				m_isSystemMenu = FALSE;
				m_isContextMenu = FALSE;
			}
			break;
		case WM_PARENTNOTIFY:
			if(LOWORD(cwps->wParam)==WM_RBUTTONDOWN)
				m_isContextMenu = FALSE;
			break;
		case WM_CONTEXTMENU:
			{
				m_isContextMenu = TRUE;
			}
			break;
		case WM_WINDOWPOSCHANGING:
			{
				LONG exStyle = GetWindowLong(cwps->hwnd,GWL_EXSTYLE);
				if(exStyle & WS_EX_MDICHILD)
				{
					CCoolmenu* cmTemp = m_coolmenuList.GetCoolmenu( cwps->hwnd );
					if(cmTemp)
					{
						cmTemp->SetMenuRedraw(TRUE);
					}
				}
			}
			break;
 		case WM_MDISETMENU:
			{
				CCoolmenu* cm = m_coolmenuList.GetCoolmenu(cwps->hwnd);
				if(cm)
				{
					if((HMENU)cwps->wParam)
					{
						cm->SetMenu((HMENU)cwps->wParam);
					}
					else if((HMENU)cwps->lParam)
					{
						cm->SetMenu((HMENU)cwps->lParam);
					}
					//cm->SetMenuSet( FALSE );
				}
			}
			break;
 		case WM_KEYDOWN:
			{
				CCoolmenu* cm = m_coolmenuList.GetCoolmenu(cwps->hwnd);
				if(cm)
				{
					cm->SetKeyDown( TRUE );
				}
			}
			break;
 		case WM_KEYUP:
			{
				CCoolmenu* cm = m_coolmenuList.GetCoolmenu(cwps->hwnd);
				if(cm)
				{
					cm->SetKeyDown( FALSE );
				}
			}
			break;
		case WM_DESTROY:
			{
				LONG exStyle = GetWindowLong(cwps->hwnd,GWL_EXSTYLE);
				if(exStyle & WS_EX_MDICHILD)
				{
					CCoolmenu* cmTemp = m_coolmenuList.GetCoolmenu( cwps->hwnd );
					if(cmTemp)
					{
						cmTemp->SetMenuRedraw(TRUE);
					}
				}
			}
			break;
		}
	}
	return CallNextHookEx( hHook, nCode, wParam, lParam );
}

void CoolMenuUninitialize( HWND hwnd )
{
	DWORD threadID = GetCurrentThreadId();

	CHookData*	hd = GetHookData(threadID);
	if(hd)
	{
		for(int i=0;i<hd->hwndList->GetSize();i++)
		{
			if(hd->hwndList->GetAt(i) == (DWORD)hwnd )
			{
				hd->hwndList->RemoveAt(i);
			}
		}
	}
	if(hd->hwndList->GetSize()==0)
	{
		Unhook( threadID );
	}
}

//	InstallCoolMenu initializes coolmenu for a thread. You can call
//	it from within a window by passing the window handle or from, for
//	example, the application object (of_Initialize) with NULL for
//	the window handle.
COOLMENUMAIN_API void __stdcall InstallCoolMenu( HWND hWnd )
{
	DWORD	threadID = GetCurrentThreadId();
	m_CmGeneral = new CCoolmenuGeneral();
	if(!IsHooked(threadID))
		Hook( threadID );

	if(hWnd) m_coolmenuList.HookWindow(hWnd,GetMenu(hWnd));

	return;
}

COOLMENUMAIN_API void __stdcall Uninstall()
{
	DWORD	threadID = GetCurrentThreadId();
	if(!IsHooked(threadID)) return;
	Unhook( threadID );
	return;
}

COOLMENUMAIN_API void __stdcall SetShadowEnabled( BOOL bShadowEnabled )
{
	m_coolmenuList.SetShadowEnabled( bShadowEnabled );
}

COOLMENUMAIN_API void __stdcall SetCoolMenubar( BOOL bCoolMenubar )
{
	m_coolmenuList.SetCoolMenubar( bCoolMenubar );
}

COOLMENUMAIN_API void __stdcall SetReactOnThemeChange( BOOL bReactOnThemeChange )
{
	m_coolmenuList.SetReactOnThemeChange( bReactOnThemeChange );
}

COOLMENUMAIN_API BOOL __stdcall GetReactOnThemeChange()
{
	return m_coolmenuList.GetReactOnThemeChange();
}

COOLMENUMAIN_API BOOL __stdcall GetCoolMenubar()
{
	CHookData* hd = GetHookData(GetCurrentThreadId());
	if(hd) return hd->bCoolMenubar;
	return FALSE;
}

COOLMENUMAIN_API void __stdcall SetFlatMenu( BOOL bFlatMenu )
{
	CHookData* hd = GetHookData(GetCurrentThreadId());
	if(hd->bDrawFlatBorder == bFlatMenu) return;
	if(hd->menuStyle==MenuStyleNormal || hd->menuStyle==MenuStyleOffice2K)
	{
		hd->bOrigDrawFlatBorder = bFlatMenu;
	}
	else
		hd->bDrawFlatBorder = bFlatMenu;
}

COOLMENUMAIN_API BOOL __stdcall SetResourceFile( LPCTSTR sResourceFile )
{
	CHookData* hd = GetHookData(GetCurrentThreadId());
	if(hd)
	{
		hd->hResources = LoadLibraryA(sResourceFile);
		if(hd->hResources) return TRUE;
	}
	return FALSE;
}

COOLMENUMAIN_API void __stdcall SetResourceFileW( WCHAR* sResourceFile )
{
	//m_sResourceFile = sResourceFile;
}

//	Names of window classes depend upon the PowerBuilder version. This function
//	is called form the constructor event of n_coolmenu.
COOLMENUMAIN_API void __stdcall SetPbVersion( LPTSTR lpszVersion )
{
	m_sWindowClass = CString( "FNWND3" );
	m_sWindowClass += lpszVersion;
	m_sPopupWindowClass = CString( "FNWNS3" );
	m_sPopupWindowClass += lpszVersion;
	m_sToolbarClass = CString( "FNFIXEDBAR" );
	m_sToolbarClass += lpszVersion;
	m_sFloatToolbarClass = CString( "FNFLOATBAR" );
	m_sFloatToolbarClass += lpszVersion;
	m_bPB105OrAbove = ( strlen( lpszVersion ) > 2 );
}

COOLMENUMAIN_API int __stdcall AddStockImage( int iImage, BOOL bBitmap, HMODULE hModule, COLORREF bckColor, INT xPixel, INT yPixel )
{
	return m_coolmenuList.AddStockImage( iImage, bBitmap, hModule, bckColor, xPixel, yPixel );
}

/*
	Function:		SetImageNameId
	Returns:			-
	Description:	After calling AddImage to add bitmaps and stockitems to the imagelist
						this function is called to associate the menuitemtext with the index
						of the bitmap within the imagelist
*/
COOLMENUMAIN_API void __stdcall SetImageNameIdW( WCHAR* iName, int iImage )
{
	m_coolmenuList.SetImageNameIdW( iName, iImage );
}

COOLMENUMAIN_API void __stdcall SetImageNameId(  LPCTSTR iName, int iImage )
{
	m_coolmenuList.SetImageNameId(  iName, iImage );
}

COOLMENUMAIN_API int __stdcall AddMenuImageW( WCHAR* sImage, HICON hIcon )
{
	return m_coolmenuList.AddMenuImageW( sImage, hIcon );
}

/*
	Function:		AddMenuImage
	Returns:			int
	Description:	Add an image to the imagelist by passing a handle to a loaded bitmap
						or icon. Use this to load an image within your PB application and add
						it to the imagelist by passing the handle. This can be used to add
						images from a dll you created with resources.
*/
COOLMENUMAIN_API int __stdcall AddMenuImage( LPCTSTR sImage, HICON hIcon )
{
	return m_coolmenuList.AddMenuImage( sImage, hIcon );
}

/*
	Function:		AddImage
	Returns:			int
	Description:	Add an image to the imagelist by passing the name of a bitmap, icon or
						stockicon. bckColor can be used to set the backcolor for your image to
						be transparent.
	Build 145:		Support added for jpg and gif.
*/
COOLMENUMAIN_API int __stdcall AddImageW( WCHAR* sImage, COLORREF bckColor, INT xPixel, INT yPixel )
{
	return m_coolmenuList.AddImageW( sImage, bckColor, xPixel, yPixel );

}

COOLMENUMAIN_API int __stdcall AddImage( LPCTSTR sImage, COLORREF bckColor, INT xPixel, INT yPixel )
{
	return m_coolmenuList.AddImage( sImage, bckColor, xPixel, yPixel );
}

COOLMENUMAIN_API int __stdcall AddImageID( WORD wImage, COLORREF bckColor, INT xPixel, INT yPixel )
{
	return m_coolmenuList.AddImageID( wImage, bckColor, xPixel, yPixel );
}

/*
	Function:		SetMenuStyle
	Returns:			void
	Description:	Display or don't display images in the menu.
*/
COOLMENUMAIN_API void __stdcall SetMenuStyle(int menuStyle)
{
	m_coolmenuList.SetMenuStyle( menuStyle );
}

/*
	Function:		GetMenuStyle
	Returns:			int
	Description:	Return current menustyle.
*/
COOLMENUMAIN_API int __stdcall GetMenuStyle()
{
	return m_coolmenuList.GetMenuStyle();
}

COOLMENUMAIN_API void __stdcall SetMenuColor( int colorIndex, COLORREF clr )
{
	m_coolmenuList.SetMenuColor( colorIndex, clr );
}

COOLMENUMAIN_API void __stdcall SetDefaults()
{
	m_coolmenuList.SetDefaults();
}

// Build 130. Set text characteristics for specific menuitems.
COOLMENUMAIN_API void __stdcall SetMenuItemProps( LPCTSTR sName, BOOL bBold, BOOL bItalic, BOOL bUnderline, COLORREF clrText )
{
	m_coolmenuList.SetMenuItemProps( sName, bBold, bItalic, bUnderline, clrText );
}

COOLMENUMAIN_API HFONT __stdcall GetFont()
{
	LOGFONT logFontMenu;
	NONCLIENTMETRICS nm = {0};
	nm.cbSize = sizeof (nm);
	VERIFY (SystemParametersInfo(SPI_GETNONCLIENTMETRICS,nm.cbSize,&nm,0));
	logFontMenu =  nm.lfMenuFont;
	HFONT hFont = CreateFontIndirect(&logFontMenu);
	return hFont;
}

COOLMENUMAIN_API int __stdcall CoolbarCreateImageLists( int cx, int cy, int cxLarge, int cyLarge )
{
	CCoolbarImageList* cbIL = m_coolbarList.GetNewCoolbarImagelist();
	if(cbIL)
	{
		cbIL->CreateImageLists( cx, cy, cxLarge, cyLarge );
		return cbIL->GetCtlID();
	}
	return -1;
}

COOLMENUMAIN_API void __stdcall CoolbarDestroyImageLists( int ctlID )
{
	CCoolbarImageList* cbIL = m_coolbarList.GetCoolbarImagelist( ctlID );
	if(cbIL)
		cbIL->DestroyImageLists();
}

COOLMENUMAIN_API void __stdcall CoolbarSetIconSize( int ctlID, BOOL bSmall, int imagelistType, int cx, int cy )
{
	CCoolbarImageList* cbIL = m_coolbarList.GetCoolbarImagelist( ctlID );
	if(cbIL)
		cbIL->SetIconSize( bSmall, imagelistType, cx, cy );
}

COOLMENUMAIN_API int __stdcall CoolbarAddStockImage( int ctlID, int iImage, BOOL bBitmap, HMODULE hModule, COLORREF bckColor, INT xPixel, INT yPixel, BOOL bSmall, int imagelistType )
{
	CCoolbarImageList* cbIL = m_coolbarList.GetCoolbarImagelist( ctlID );
	if(cbIL)
	{
		if(bSmall)
		{
			return m_CmGeneral->AddStockImage( (imagelistType==il_Normal?&cbIL->CCoolbarImagelist:&cbIL->CCoolbarHotImagelist), (imagelistType==il_Normal?&cbIL->CCoolbarDisabledImagelist:NULL), iImage, bBitmap, hModule, bckColor, xPixel, yPixel );
		}
		else
		{
			if(xPixel == cbIL->m_cx - 1 ) xPixel = cbIL->m_cxLarge - 1;
			if(yPixel == cbIL->m_cy - 1 ) yPixel = cbIL->m_cyLarge - 1;
			return m_CmGeneral->AddStockImage( (imagelistType==il_Normal?&cbIL->CCoolbarLargeImagelist:&cbIL->CCoolbarLargeHotImagelist), (imagelistType==il_Normal?&cbIL->CCoolbarLargeDisabledImagelist:NULL), iImage, bBitmap, hModule, bckColor, xPixel, yPixel );
		}
	}
	return -1;
}

COOLMENUMAIN_API int __stdcall CoolbarAddImage( int ctlID, LPCTSTR sImage, COLORREF bckColor, INT xPixel, INT yPixel, BOOL bSmall, int imagelistType )
{
	CCoolbarImageList* cbIL = m_coolbarList.GetCoolbarImagelist( ctlID );
	if(cbIL)
	{
		if(bSmall)
			return m_CmGeneral->AddImage( (imagelistType==il_Normal?&cbIL->CCoolbarImagelist:&cbIL->CCoolbarHotImagelist), (imagelistType==il_Normal?&cbIL->CCoolbarDisabledImagelist:NULL), cbIL->hCoolbarResources, sImage, bckColor, xPixel, yPixel );
		else
		{
			if(xPixel == cbIL->m_cx - 1 ) xPixel = cbIL->m_cxLarge - 1;
			if(yPixel == cbIL->m_cy - 1 ) yPixel = cbIL->m_cyLarge - 1;
			return m_CmGeneral->AddImage( (imagelistType==il_Normal?&cbIL->CCoolbarLargeImagelist:&cbIL->CCoolbarLargeHotImagelist), (imagelistType==il_Normal?&cbIL->CCoolbarLargeDisabledImagelist:NULL), cbIL->hCoolbarResources, sImage, bckColor, xPixel, yPixel );
		}
	}
	return -1;
}

COOLMENUMAIN_API int __stdcall CoolbarAddImageID( int ctlID, WORD wImage, COLORREF bckColor, INT xPixel, INT yPixel, BOOL bSmall, int imagelistType )
{
	CCoolbarImageList* cbIL = m_coolbarList.GetCoolbarImagelist( ctlID );
	if(cbIL)
	{
		if(bSmall)
			return m_CmGeneral->AddImageID( (imagelistType==il_Normal?&cbIL->CCoolbarImagelist:&cbIL->CCoolbarHotImagelist), (imagelistType==il_Normal?&cbIL->CCoolbarDisabledImagelist:NULL), cbIL->hCoolbarResources, wImage, bckColor, xPixel, yPixel );
		else
		{
			if(xPixel == cbIL->m_cx - 1 ) xPixel = cbIL->m_cxLarge - 1;
			if(yPixel == cbIL->m_cy - 1 ) yPixel = cbIL->m_cyLarge - 1;
			return m_CmGeneral->AddImageID( (imagelistType==il_Normal?&cbIL->CCoolbarLargeImagelist:&cbIL->CCoolbarLargeHotImagelist), (imagelistType==il_Normal?&cbIL->CCoolbarLargeDisabledImagelist:NULL), cbIL->hCoolbarResources, wImage, bckColor, xPixel, yPixel );
		}
	}
	return -1;
}

COOLMENUMAIN_API BOOL __stdcall CoolbarSetResourceFile( int ctlID, LPCTSTR sResourceFile )
{
	CCoolbarImageList* cbIL = m_coolbarList.GetCoolbarImagelist( ctlID );
	if(cbIL)
	{
		cbIL->hCoolbarResources = LoadLibraryA(sResourceFile);
		if(cbIL->hCoolbarResources) return TRUE;
	}
	return FALSE;
}

COOLMENUMAIN_API HIMAGELIST __stdcall CoolbarGetImageListHandle( int ctlID, BOOL bSmall, int imagelistType )
{
	CCoolbarImageList* cbIL = m_coolbarList.GetCoolbarImagelist( ctlID );
	if(cbIL)
	{
		if( bSmall)
		{
			if( imagelistType == il_Normal )
				return cbIL->CCoolbarImagelist.m_hImageList;
			else
			{
				if( imagelistType == il_Disabled )
					return cbIL->CCoolbarDisabledImagelist.m_hImageList;
				else
					return cbIL->CCoolbarHotImagelist.m_hImageList;
			}
		}
		else
		{
			if( imagelistType == il_Normal )
				return cbIL->CCoolbarLargeImagelist.m_hImageList;
			else
			{
				if( imagelistType == il_Disabled )
					return cbIL->CCoolbarLargeDisabledImagelist.m_hImageList;
				else
					return cbIL->CCoolbarLargeHotImagelist.m_hImageList;
			}
		}
	}
	return (HIMAGELIST)-1;
}

COOLMENUMAIN_API HBITMAP __stdcall GetFocusBitmap( BOOL bFloating )
{
	HBITMAP hbm;
	static WORD _dotPatternBmp1[] =
	{
		0x00aa, 0x0055, 0x00aa, 0x0055, 0x00aa, 0x0055, 0x00aa, 0x0055
	};
	static WORD _dotPatternBmp2[] =
	{
		0xffff, 0xffff, 0xffff, 0xffff, 0xffff, 0xffff, 0xffff, 0xffff
	};

	hbm = CreateBitmap(8, 8, 1, 1, (bFloating)?_dotPatternBmp1:_dotPatternBmp2);

	return hbm;
}