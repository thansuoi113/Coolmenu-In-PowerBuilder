// CoolmenuGeneral.cpp: implementation of the CCoolmenuGeneral class.
//
//	This class implements some functions of general use.
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "CoolmenuGeneral.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

class CNewLoadLib
{
public:

  HMODULE m_hModule;
  void* m_pProg;

  CNewLoadLib(LPCTSTR pName,LPCSTR pProgName)
  {
    m_hModule = LoadLibrary(pName);
    m_pProg = m_hModule ? (void*)GetProcAddress (m_hModule,pProgName) : NULL;
  }

  ~CNewLoadLib()
  {
    if(m_hModule)
    {
      FreeLibrary(m_hModule);
      m_hModule = NULL;
    }
  }
};

#if(WINVER < 0x0500)

BOOL CCoolmenuGeneral::GetMenuInfo( HMENU hMenu, LPMENUINFO pInfo)
{
	static CNewLoadLib menuInfo(_T("user32.dll"),"GetMenuInfo");
	if(menuInfo.m_pProg)
	{
		typedef BOOL (WINAPI* FktGetMenuInfo)(HMENU, LPMENUINFO);
		return ((FktGetMenuInfo)menuInfo.m_pProg)(hMenu,pInfo);
	}
	return FALSE;
}

BOOL CCoolmenuGeneral::SetMenuInfo( HMENU hMenu, LPCMENUINFO pInfo)
{
	static CNewLoadLib menuInfo(_T("user32.dll"),"SetMenuInfo");
	if(menuInfo.m_pProg)
	{
		typedef BOOL (WINAPI* FktSetMenuInfo)(HMENU, LPCMENUINFO);
		return ((FktSetMenuInfo)menuInfo.m_pProg)(hMenu,pInfo);
	}
	return FALSE;
}
#endif

CCoolmenuGeneral::CCoolmenuGeneral()
{
	ilDummy = new CImageList();
	ilDummy->Create( 16, 16, ILC_MASK | ILC_COLOR32, 0, 10 );
}

CCoolmenuGeneral::~CCoolmenuGeneral()
{
	ilDummy->DeleteImageList();
}

Win32Type CCoolmenuGeneral::IsShellType()
{
	OSVERSIONINFO osvi = {0};
	osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);

	DWORD winVer=GetVersion();
	if(winVer<0x80000000)
	{/*NT */
		if(!GetVersionEx(&osvi))
		{
		//ShowLastError();
		}
		if(osvi.dwMajorVersion==4L)
		{
			return WinNT4;
		}
		else if(osvi.dwMajorVersion==5L && osvi.dwMinorVersion==0L)
		{
			return Win2000;
		}
		// thanks to John Zegers
		else if(osvi.dwMajorVersion>=5L)// && osvi.dwMinorVersion==1L)
		{
			return WinXP;
		}
		return WinNT3;
	}
	else if (LOBYTE(LOWORD(winVer))<4)
	{
		return Win32s;
	}

	if(!GetVersionEx(&osvi))
	{
		//ShowLastError();
	}
	if(osvi.dwMajorVersion==4L && osvi.dwMinorVersion==10L)
	{
		return Win98;
	}
	else if(osvi.dwMajorVersion==4L && osvi.dwMinorVersion==90L)
	{
		return WinME;
	}
	return Win95;
}

BOOL CCoolmenuGeneral::IsShadowEnabled()
{
	BOOL bEnabled = FALSE;
	if(SystemParametersInfo(SPI_GETDROPSHADOW,0,&bEnabled,0))
	{
		return bEnabled;
	}
	return FALSE;
}

int CCoolmenuGeneral::NumScreenColors()
{
	static int nColors = 0;
	if (!nColors)
	{
		// DC of the desktop
		CClientDC myDC(NULL);
		nColors = myDC.GetDeviceCaps(NUMCOLORS);
		if (nColors == -1)
		{
			nColors = 64000;
		}
	}
	return nColors;
}

COLORREF CCoolmenuGeneral::MidColor(COLORREF colorA,COLORREF colorB)
{
	// (7a + 3b)/10
	int red   = MulDiv(7,GetRValue(colorA),10) + MulDiv(3,GetRValue(colorB),10);
	int green = MulDiv(7,GetGValue(colorA),10) + MulDiv(3,GetGValue(colorB),10);
	int blue  = MulDiv(7,GetBValue(colorA),10) + MulDiv(3,GetBValue(colorB),10);

	return RGB( red, green, blue);
}

COLORREF CCoolmenuGeneral::MixedColor(COLORREF colorA,COLORREF colorB)
{
	// ( 86a + 14b ) / 100
	int red   = MulDiv(86,GetRValue(colorA),100) + MulDiv(14,GetRValue(colorB),100);
	int green = MulDiv(86,GetGValue(colorA),100) + MulDiv(14,GetGValue(colorB),100);
	int blue  = MulDiv(86,GetBValue(colorA),100) + MulDiv(14,GetBValue(colorB),100);

	return RGB( red,green,blue);
}

COLORREF CCoolmenuGeneral::DarkenColor( long lScale, COLORREF lColor)
{
	long red   = MulDiv(GetRValue(lColor),(255-lScale),255);
	long green = MulDiv(GetGValue(lColor),(255-lScale),255);
	long blue  = MulDiv(GetBValue(lColor),(255-lScale),255);

	return RGB(red, green, blue);
}

COLORREF CCoolmenuGeneral::LightenColor( long lScale, COLORREF lColor)
{
	long R = MulDiv(255-GetRValue(lColor),lScale,255)+GetRValue(lColor);
	long G = MulDiv(255-GetGValue(lColor),lScale,255)+GetGValue(lColor);
	long B = MulDiv(255-GetBValue(lColor),lScale,255)+GetBValue(lColor);

	return RGB(R, G, B);
}

COLORREF CCoolmenuGeneral::LightenColor( COLORREF col, double factor )
{
	if(factor>0.0&&factor<=1.0){
		BYTE red,green,blue,lightred,lightgreen,lightblue;
		red = GetRValue(col);
		green = GetGValue(col);
		blue = GetBValue(col);
		lightred = (BYTE)((factor*(255-red)) + red);
		lightgreen = (BYTE)((factor*(255-green)) + green);
		lightblue = (BYTE)((factor*(255-blue)) + blue);
		col = RGB(lightred,lightgreen,lightblue);
	}
	return(col);
}

COLORREF CCoolmenuGeneral::GrayColor(COLORREF crColor)
{
	int Gray  = (((int)GetRValue(crColor)) + GetGValue(crColor) + GetBValue(crColor))/3;
	return RGB( Gray,Gray,Gray);
}


COLORREF CCoolmenuGeneral::GetAlphaBlendColor(COLORREF blendColor, COLORREF pixelColor,int weighting)
{
  if(pixelColor==CLR_NONE)
  {
    return CLR_NONE;
  }
  // Algorithme for alphablending
  //dest' = ((weighting * source) + ((255-weighting) * dest)) / 256
  DWORD refR = ((weighting * GetRValue(pixelColor)) + ((255-weighting) * GetRValue(blendColor))) / 256;
  DWORD refG = ((weighting * GetGValue(pixelColor)) + ((255-weighting) * GetGValue(blendColor))) / 256;;
  DWORD refB = ((weighting * GetBValue(pixelColor)) + ((255-weighting) * GetBValue(blendColor))) / 256;;

  return RGB(refR,refG,refB);
}

COLORREF CCoolmenuGeneral::GetMenuColor(HMENU hMenu)
{
	if(hMenu!=NULL)
	{
		MENUINFO menuInfo={0};
		menuInfo.cbSize = sizeof(menuInfo);
		menuInfo.fMask = MIM_BACKGROUND;

		if(GetMenuInfo(hMenu,&menuInfo) && menuInfo.hbrBack)
		{
			LOGBRUSH logBrush;
			if(GetObject(menuInfo.hbrBack,sizeof(LOGBRUSH),&logBrush))
			{
				return logBrush.lbColor;
			}
		}
	}

	//  if(m_CmGeneral->IsShellType()==WinXP)
	if(IsShellType()==WinXP)
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

COLORREF CCoolmenuGeneral::GetBitmapTransColor( HBITMAP hBitmap, INT xPixel, INT yPixel )
{
	COLORREF	cr;
	HDC hdcSrc;

	hdcSrc = CreateCompatibleDC( NULL );

	// Load the bitmap into memory DC
	HBITMAP hbmSrcT = (HBITMAP)SelectObject( hdcSrc, hBitmap );

	// Dynamically get the transparent color
	cr = GetPixel( hdcSrc, xPixel, yPixel );

	SelectObject( hdcSrc, hbmSrcT );
	DeleteDC( hdcSrc );

	return cr;
}

COLORREF MakeGrayAlphablend(CBitmap* pBitmap, int weighting, COLORREF blendcolor)
{
	CDC myDC;
	// Create a compatible bitmap to the screen
	myDC.CreateCompatibleDC(0);
	// Select the bitmap into the DC
	CBitmap* pOldBitmap = myDC.SelectObject(pBitmap);

	BITMAP myInfo = {0};
	GetObject((HGDIOBJ)pBitmap->m_hObject,sizeof(myInfo),&myInfo);

	for (int nHIndex = 0; nHIndex < myInfo.bmHeight; nHIndex++)
	{
		for (int nWIndex = 0; nWIndex < myInfo.bmWidth; nWIndex++)
		{
			COLORREF ref = myDC.GetPixel(nWIndex,nHIndex);

			// make it gray
			DWORD nAvg =  (GetRValue(ref) + GetGValue(ref) + GetBValue(ref))/3;

			// Algorithme for alphablending
			//dest' = ((weighting * source) + ((255-weighting) * dest)) / 256
			DWORD refR = ((weighting * nAvg) + ((255-weighting) * GetRValue(blendcolor))) / 256;
			DWORD refG = ((weighting * nAvg) + ((255-weighting) * GetGValue(blendcolor))) / 256;;
			DWORD refB = ((weighting * nAvg) + ((255-weighting) * GetBValue(blendcolor))) / 256;;

			myDC.SetPixel(nWIndex,nHIndex,RGB(refR,refG,refB));
		}
	}
	COLORREF topLeftColor = myDC.GetPixel(0,0);
	myDC.SelectObject(pOldBitmap);
	return topLeftColor;
}

int CCoolmenuGeneral::AddGrayIcon( CImageList* hIml, HICON hIcon, int nIndex )
{
	int cx, cy;
	ICONINFO iconInfo = {0};
	if(!GetIconInfo(hIcon,&iconInfo))
	{
		return -1;
	}
	ImageList_GetIconSize( hIml->m_hImageList, &cx, &cy );
	CSize size = CSize(cx,cy);
	CDC myDC;
	myDC.CreateCompatibleDC(0);

	CBitmap bmColor;
	bmColor.Attach(iconInfo.hbmColor);
	CBitmap bmMask;
	bmMask.Attach(iconInfo.hbmMask);

	CBitmap* pOldBitmap = myDC.SelectObject(&bmColor);
	COLORREF crMenu = GetSysColor(COLOR_MENUBAR);
	COLORREF crPixel;
	for(int i=0;i<size.cx;++i)
	{
		for(int j=0;j<size.cy;++j)
		{
			crPixel = myDC.GetPixel(i,j);
			myDC.SetPixel(i,j,MixedColor(LightenColor(100,GrayColor(crPixel)),crMenu));
		}
	}
	myDC.SelectObject(pOldBitmap);
	if(nIndex==-1)
	{
		return hIml->Add( &bmColor, &bmMask );
	}
	return (hIml->Replace(nIndex,&bmColor,&bmMask)) ? nIndex: -1;
}

int CCoolmenuGeneral::AddGloomIcon( CImageList* hIml, HICON hIcon, int nIndex, int cx, int cy )
{
	ICONINFO iconInfo = {0};
	if(!GetIconInfo(hIcon,&iconInfo))
	{
		return -1;
	}

	CSize size = CSize( cx, cy );
	CDC myDC;
	myDC.CreateCompatibleDC(0);

	CBitmap bmColor;
	bmColor.Attach(iconInfo.hbmColor);
	CBitmap bmMask;
	bmMask.Attach(iconInfo.hbmMask);

	CBitmap* pOldBitmap = myDC.SelectObject(&bmColor);
	COLORREF crPixel;
	for(int i=0;i<size.cx;++i)
	{
		for(int j=0;j<size.cy;++j)
		{
			crPixel = myDC.GetPixel(i,j);
			myDC.SetPixel(i,j,DarkenColor(50,crPixel));
		}
	}
	myDC.SelectObject(pOldBitmap);
	if(nIndex==-1)
	{
		return hIml->Add(&bmColor,&bmMask);
	}

	return (hIml->Replace(nIndex,&bmColor,&bmMask)) ? nIndex: -1;
}

int CCoolmenuGeneral::AddGrayIcon( CImageList* hIml, HBITMAP hBitmap, COLORREF transColor, int nIndex )
{
	CBitmap bmColor;
	bmColor.Attach(hBitmap);

	COLORREF blendcolor;
	if(IsShellType()==Win95 || IsShellType()==WinNT4)
	{
		blendcolor = LightenColor( 115, GetSysColor(COLOR_MENU) );
	}
	else
		blendcolor = LightenColor( 115, GetSysColor(COLOR_3DFACE) );

	MakeGrayAlphablend(&bmColor,110, blendcolor);

	//if(transColor!=CLR_DEFAULT) transColor = GetSysColor( COLOR_3DFACE );
	if(nIndex==-1)
	{
		return hIml->Add( &bmColor, transColor );//RGB(0,0,0) );
	}

	return -1;//(ImageList_ReplaceIcon(hIml,nIndex,&bmColor,&bmMask)) ? nIndex: -1;
}

int CCoolmenuGeneral::AddStockImage( CImageList* hIml, CImageList* hImlDisabled, int iImage, BOOL bBitmap, HMODULE hModule, COLORREF bckColor, INT xPixel, INT yPixel )
{
	int iIcon = -1, cx, cy;
	CBitmap bmColor;

	ImageList_GetIconSize( hIml->m_hImageList, &cx, &cy );

	if ( bBitmap )
	{
		HBITMAP hBitmap = (HBITMAP)LoadImageA( hModule, MAKEINTRESOURCE( iImage ), IMAGE_BITMAP, cx, cy, LR_LOADMAP3DCOLORS );
		if( hBitmap != NULL ) {
			bmColor.Attach(hBitmap);
			if (bckColor == NULL)
				bckColor = GetBitmapTransColor( hBitmap, xPixel, yPixel );
			if (bckColor == -1)
			{
				iIcon = hIml->Add( &bmColor, bckColor );
				HICON hIcon = hIml->ExtractIcon( iIcon );
				//AddGloomIcon( hIml, hIcon, iIcon, cx, cy );
				if(hImlDisabled) AddGrayIcon( hImlDisabled, hIcon, -1 );
				DestroyIcon( hIcon );
			}
			else
			{
				iIcon = hIml->Add( &bmColor, bckColor );
				HICON hIcon = hIml->ExtractIcon( iIcon );
				//AddGloomIcon( hIml, hIcon, iIcon, cx, cy );
				if(hImlDisabled) AddGrayIcon( hImlDisabled, hIcon, -1 );
				DestroyIcon( hIcon );
			}
			//AddGrayIcon( hIml, hBitmap, bckColor, -1);
			DeleteObject( hBitmap );
		} else {
			return -1;
		}
	}
	else {
		HICON hIcon = (HICON)LoadImageA( hModule, MAKEINTRESOURCE( iImage ), IMAGE_ICON, cx, cy, LR_LOADMAP3DCOLORS );
		if (hIcon != NULL )
		{
			iIcon = hIml->Replace( -1, hIcon );
			if(hImlDisabled) AddGrayIcon ( hImlDisabled, hIcon, -1 );
			DestroyIcon( hIcon );
		} else {
			return -1;
		}
	}
	return iIcon;
}

int CCoolmenuGeneral::AddMenuImageW( CImageList* hIml, CImageList* hImlDisabled, WCHAR* sImage, HICON hIcon )
{
	LPSTR m_szBuffer = new CHAR [lstrlenW( sImage)];
	CString cName( sImage, lstrlenW( sImage) );
	strcpy( m_szBuffer, cName );
	int iIcon = AddMenuImage( hIml, hImlDisabled, (LPCTSTR)m_szBuffer, hIcon );
	delete [] m_szBuffer;
	return iIcon;
}

/*
	Function:		AddMenuImage
	Returns:			int
	Description:	Add an image to the imagelist by passing a handle to a loaded bitmap
						or icon. Use this to load an image within your PB application and add
						it to the imagelist by passing the handle. This can be used to add
						images from a dll you created with resources.
*/
int CCoolmenuGeneral::AddMenuImage( CImageList* hIml, CImageList* hImlDisabled, LPCTSTR sImage, HICON hIcon )
{
	int iIcon = -1;
	if (hIcon != NULL ) {
		iIcon = hIml->Replace( -1, hIcon );
		if(hImlDisabled) AddGrayIcon( hImlDisabled, hIcon, -1 );
	}
	return iIcon;
}

/*
	Function:		AddImage
	Returns:			int
	Description:	Add an image to the imagelist by passing the name of a bitmap, icon or
						stockicon. bckColor can be used to set the backcolor for your image to
						be transparent.
	Build 145:		Support added for jpg and gif.
*/
int CCoolmenuGeneral::AddImageW( CImageList* hIml, CImageList* hImlDisabled, HMODULE hResourceFile, WCHAR* sImage, COLORREF bckColor, INT xPixel, INT yPixel )
{
	int iIcon = -1;

	CString cName( sImage, lstrlenW( sImage) );
	LPSTR m_szBuffer = new CHAR [lstrlenW( sImage)];
	strcpy( m_szBuffer, cName );
	iIcon = AddImage( hIml, hImlDisabled, hResourceFile, (LPCTSTR)m_szBuffer, bckColor, xPixel, yPixel );
	delete [] m_szBuffer;
	return iIcon;
}

int CCoolmenuGeneral::AddImage( CImageList* hIml, CImageList* hImlDisabled, HMODULE hResourceFile, LPCTSTR sImage, COLORREF bckColor, INT xPixel, INT yPixel )
{
	HBITMAP hBitmap;
	CBitmap bmColor;
	HICON hIcon;
	CString* lstr = new CString( sImage );
	lstr->MakeLower();
	CString lstrExtension;
	int iIcon = -1;
	int	cx, cy, countBefore, countAfter;
	int	iStart = lstr->Find( '.', 0) + 1;
	if ( iStart == 0 )
	{
		lstrExtension = lstr->Right( 1 );
	}
	else
	{
		lstrExtension = lstr->Right( 3 );
	}
	ImageList_GetIconSize( hIml->m_hImageList, &cx, &cy );
	if ( lstrExtension == "bmp" )
	{
		hBitmap = (HBITMAP)LoadImageA( NULL, sImage, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE + LR_LOADMAP3DCOLORS );
		if (hBitmap != NULL )
		{
			if(ilDummy->GetSafeHandle())
			{
				ImageList_Add( ilDummy->m_hImageList, hBitmap, NULL );
				if(ilDummy->GetImageCount() == 1)
				{
					DeleteObject( hBitmap );
					hBitmap = (HBITMAP)LoadImageA( NULL, sImage, IMAGE_BITMAP, cx, cy, LR_LOADFROMFILE );
				}
				for (int i=0;i < ilDummy->GetImageCount();i++)
				{
					ilDummy->Remove(i);
				}
			}

			bmColor.Attach(hBitmap);
			countBefore = hIml->GetImageCount();
			if (bckColor == NULL)
				bckColor = GetBitmapTransColor( hBitmap, xPixel, yPixel );
			if (bckColor == -1)
			{
				iIcon = hIml->Add( &bmColor, bckColor );
			}
			else
			{
				iIcon = hIml->Add( &bmColor, bckColor );
			}
			countAfter = hIml->GetImageCount();
			if(hImlDisabled)
			{
				for(int i = iIcon; i < iIcon + ( countAfter - countBefore ); i++)
				{
					HICON hIcon = hIml->ExtractIcon( i );
					AddGrayIcon( hImlDisabled, hIcon, -1 );
					DestroyIcon( hIcon );
				}
			}
			DeleteObject( hBitmap );
		} else {
			return -1;
		}
	} else {
		if (lstrExtension == "ico" )
		{
			hIcon = (HICON)LoadImageA( NULL, sImage, IMAGE_ICON, cx, cy, LR_LOADFROMFILE + LR_LOADMAP3DCOLORS );
			if (hIcon != NULL )
			{
				//countBefore = hIml->GetImageCount();
				iIcon = hIml->Replace( -1, hIcon );
				//countAfter = hIml->GetImageCount();
				/*if(hImlDisabled)
				{
					for(int i = iIcon; i < iIcon + ( countAfter - countBefore ); i++)
					{
						HICON hIconNew = hIml->ExtractIcon( i );
						AddGrayIcon( hImlDisabled, hIconNew, -1 );
						DestroyIcon( hIconNew );
					}
				}*/
				if(hImlDisabled) AddGrayIcon( hImlDisabled, hIcon, -1 );
				DestroyIcon( hIcon );
			} else {
				return -1;
			}
		} else {
			if ( lstrExtension == "jpg" | lstrExtension == "gif" ) {
				//	Use CPicture class to load jpg and gif.
				CPicture* jpgPic = new CPicture();
				hBitmap = jpgPic->Load( sImage );
				if ( hBitmap != NULL && hBitmap != 0 )
				{
					bmColor.Attach(hBitmap);
					bckColor = -1;//GetBitmapTransColor( hBitmap, 15, 14 );
					iIcon = hIml->Add( &bmColor, bckColor );
					//iIcon = ImageList_Add( hIml, hBitmap, NULL );
//					AddGrayIcon( hImlDisabled, hBitmap, CLR_DEFAULT, -1 );
					hIcon = hIml->ExtractIcon( iIcon );
					//hIml->Replace( hIml, iIcon, hIcon );
					if(hImlDisabled) AddGrayIcon( hImlDisabled, hIcon, -1 );
					DestroyIcon( hIcon );
					DeleteObject( hBitmap );
				}
				else {
					return -1;
				}
			}
		}
	}
	return iIcon;
}

int CCoolmenuGeneral::AddImageID( CImageList* hIml, CImageList* hImlDisabled, HMODULE hResourceFile, WORD wImage, COLORREF bckColor, INT xPixel, INT yPixel )
{
	int iIcon = -1, cx, cy;
	CBitmap bmColor;
	ImageList_GetIconSize( hIml->m_hImageList, &cx, &cy );
	HBITMAP hBitmap = (HBITMAP)LoadImageA( hResourceFile, MAKEINTRESOURCE(wImage), IMAGE_BITMAP, cx, cy, LR_LOADMAP3DCOLORS );
	if (hBitmap != NULL ) {
		bmColor.Attach(hBitmap);
		if (bckColor == NULL)
			bckColor = GetBitmapTransColor( hBitmap, xPixel, yPixel );
		if (bckColor == -1)
		{
			iIcon = hIml->Add( &bmColor, (COLORREF)NULL );
			if(hImlDisabled) AddGrayIcon( hImlDisabled, hBitmap, CLR_DEFAULT, -1 );
		}
		else
		{
			iIcon = hIml->Add( &bmColor, bckColor );
			if(hImlDisabled) AddGrayIcon( hImlDisabled, hBitmap, bckColor, -1 );
		}
		DeleteObject( hBitmap );
	} else {
		HICON hIcon = (HICON)LoadImageA( hResourceFile, MAKEINTRESOURCE(wImage), IMAGE_ICON, cx, cy, LR_LOADMAP3DCOLORS );
		if (hIcon != NULL ) {
			iIcon = hIml->Replace( -1, hIcon );
			if(hImlDisabled) AddGrayIcon( hImlDisabled, hIcon, -1 );
			DestroyIcon( hIcon );
		} else {
			return -1;
		}
	}
	return iIcon;
}

HANDLE CCoolmenuGeneral::localOpenThemeData( HWND hWnd, LPCWSTR pszClassList)
{
  static CNewLoadLib menuInfo(_T("UxTheme.dll"),"OpenThemeData");
  if(menuInfo.m_pProg)
  {
    typedef HANDLE (WINAPI* FktOpenThemeData)(HWND hWnd, LPCWSTR pszClassList);
    return ((FktOpenThemeData)menuInfo.m_pProg)(hWnd,pszClassList);
  }
  return NULL;
}

HRESULT CCoolmenuGeneral::localCloseThemeData(HANDLE hTheme)
{
  static CNewLoadLib menuInfo(_T("UxTheme.dll"),"CloseThemeData");
  if(menuInfo.m_pProg)
  {
    typedef HRESULT (WINAPI* FktCloseThemeData)(HANDLE hTheme);
    return ((FktCloseThemeData)menuInfo.m_pProg)(hTheme);
  }
  return NULL;
}

HRESULT CCoolmenuGeneral::localGetThemeColor(HANDLE hTheme,
    int iPartId,
    int iStateId,
    int iColorId,
    COLORREF *pColor)
{
  static CNewLoadLib menuInfo(_T("UxTheme.dll"),"GetThemeColor");
  if(menuInfo.m_pProg)
  {
    typedef HRESULT (WINAPI* FktGetThemeColor)(HANDLE hTheme,int iPartId,int iStateId,int iColorId, COLORREF *pColor);
    return ((FktGetThemeColor)menuInfo.m_pProg)(hTheme, iPartId, iStateId, iColorId, pColor);
  }
  return S_FALSE;
}

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CPicture::CPicture()
//=============================================================================
{
	m_IPicture = NULL;
	m_Height = 0;
	m_Weight = 0;
	m_Width = 0;
}

CPicture::~CPicture()
//=============================================================================
{
	if(m_IPicture != NULL) FreePictureData(); // Important - Avoid Leaks...
}

void CPicture::FreePictureData()
//=============================================================================
{
	if(m_IPicture != NULL)
	{
		m_IPicture->Release();
		m_IPicture = NULL;
		m_Height = 0;
		m_Weight = 0;
		m_Width = 0;
	}
}

HBITMAP CPicture::Load( CString sFilePathName )
//=============================================================================
{
	HBITMAP hB = NULL;
	BOOL bResult;
	CFile PictureFile;
	CFileException e;
	int	nSize = 0;

	if ( m_IPicture != NULL) FreePictureData(); // Important - Avoid Leaks...

	if ( PictureFile.Open( sFilePathName, CFile::modeRead | CFile::typeBinary, &e ) )
	{
		nSize = PictureFile.GetLength();
		BYTE* pBuffer = new BYTE[nSize];

		if(PictureFile.Read(pBuffer, nSize) > 0)
		{
			if(LoadPictureData(pBuffer, nSize))	bResult = TRUE;
		}

		PictureFile.Close();
		delete [] pBuffer;
	}
	else // Open Failed...
	{
		bResult = FALSE;
	}

	m_Weight = nSize; // Update Picture Size Info...

	if(m_IPicture != NULL) // Do Not Try To Read From Memory That Is Not Exist...
	{
		m_IPicture->get_Height(&m_Height);
		m_IPicture->get_Width(&m_Width);
		// Calculate Its Size On a "Standard" (96 DPI) Device Context
		m_Height = MulDiv(m_Height, 96, HIMETRIC_INCH);
		m_Width  = MulDiv(m_Width,  96, HIMETRIC_INCH);
		m_IPicture->get_Handle((unsigned int*)&hB);
	}
	else // Picture Data Is Not a Known Picture Type
	{
		m_Height = 0;
		m_Width = 0;
		bResult = FALSE;
	}
	return hB;
}

BOOL CPicture::LoadPictureData(BYTE *pBuffer, int nSize)
//=============================================================================
{
	BOOL bResult = FALSE;

	HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, nSize);

	if(hGlobal == NULL)
	{
		return(FALSE);
	}

	void* pData = GlobalLock(hGlobal);
	memcpy(pData, pBuffer, nSize);
	GlobalUnlock(hGlobal);

	IStream* pStream = NULL;

	if(CreateStreamOnHGlobal(hGlobal, TRUE, &pStream) == S_OK)
	{
		HRESULT hr;
		if((hr = OleLoadPicture(pStream, nSize, TRUE, IID_IPicture, (LPVOID *)&m_IPicture)) == E_NOINTERFACE)
		{
			return(FALSE);
		}
		else // S_OK
		{
			pStream->Release();
			pStream = NULL;
			bResult = TRUE;
		}
	}

	FreeResource(hGlobal); // 16Bit Windows Needs This (32Bit - Automatic Release)

	return(bResult);
}

