// CoolmenuGeneral.h: interface for the CCoolmenuGeneral class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_COOLMENUGENERAL_H__77097D6E_6413_4148_B6F8_0E6083645D3D__INCLUDED_)
#define AFX_COOLMENUGENERAL_H__77097D6E_6413_4148_B6F8_0E6083645D3D__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define HIMETRIC_INCH 2540

#ifndef COLOR_MENUBAR
#define COLOR_MENUBAR       30
#endif

#ifndef SPI_GETDROPSHADOW
#define SPI_GETDROPSHADOW   0x1024
#endif

#ifndef SPI_GETFLATMENU
#define SPI_GETFLATMENU     0x1022
#endif

#ifndef MIIM_STRING
#define MIIM_STRING			0x00000040
#endif

#ifndef MIIM_FTYPE
#define MIIM_FTYPE			0x00000100
#endif

#define WM_REGISTERNEWWINDOW			0xC252

#ifndef ARRAY_SIZE
#define ARRAY_SIZE(a) (sizeof(a)/sizeof(a[0]))
#endif

#ifndef ULONG_PTR
typedef unsigned long ULONG_PTR, *PULONG_PTR;
#endif

#if(WINVER < 0x0500)

#define MNS_NOCHECK         0x80000000
#define MNS_MODELESS        0x40000000
#define MNS_DRAGDROP        0x20000000
#define MNS_AUTODISMISS     0x10000000
#define MNS_NOTIFYBYPOS     0x08000000
#define MNS_CHECKORBMP      0x04000000

#define MIM_MAXHEIGHT               0x00000001
#define MIM_BACKGROUND              0x00000002
#define MIM_HELPID                  0x00000004
#define MIM_MENUDATA                0x00000008
#define MIM_STYLE                   0x00000010
#define MIM_APPLYTOSUBMENUS         0x80000000

typedef struct tagMENUINFO
{
  DWORD   cbSize;
  DWORD   fMask;
  DWORD   dwStyle;
  UINT    cyMax;
  HBRUSH  hbrBack;
  DWORD   dwContextHelpID;
  ULONG_PTR dwMenuData;
}   MENUINFO, FAR *LPMENUINFO;
typedef MENUINFO CONST FAR *LPCMENUINFO;

//BOOL GUILIBDLLEXPORT GetMenuInfo( HMENU hMenu, LPMENUINFO pInfo);
//BOOL GUILIBDLLEXPORT SetMenuInfo( HMENU hMenu, LPCMENUINFO pInfo);

#endif

// Constans for detecting OS-Type
enum Win32Type
{
  Win32s,
  WinNT3,
  Win95,
  Win98,
  WinME,
  WinNT4,
  Win2000,
  WinXP
};

enum MenuStyle
{
	MenuStyleNormal,
	MenuStyleOffice2K,
	MenuStyleOfficeXP,
	MenuStyleOffice2K3,
	MenuStyleOffice2K7
};

enum MenuColor
{
	SelectedColor,
	MenuBckColor,
	BitmapBckColor,
	CheckedSelectedColor,
	CheckedColor,
	PenColor,
	TextColor,
	HighlightTextColor,
	DisabledTextColor,
	GradientStartColor,
	GradientEndColor,
	MenubarColor,
	MenubarGradientStartColor,
	MenubarGradientEndColor,
	MenubarGradientStartColorDown,
	MenubarGradientEndColorDown,
	SelectedPenColor,
	SeparatorColor,
	CheckPenColor,
	CheckSelectedPenColor,
	CheckDisabledPenColor
};

#define MAX_COLOR 21

class CCoolmenuGeneral
{
public:
	CCoolmenuGeneral();
	virtual ~CCoolmenuGeneral();

	Win32Type IsShellType();
	BOOL IsShadowEnabled();
	int NumScreenColors();
	BOOL GetMenuInfo( HMENU hMenu, LPMENUINFO pInfo);
	BOOL SetMenuInfo( HMENU hMenu, LPCMENUINFO pInfo);
	COLORREF LightenColor( long lScale, COLORREF lColor);
	COLORREF LightenColor( COLORREF col, double factor );
	COLORREF MidColor(COLORREF colorA,COLORREF colorB);
	COLORREF MixedColor(COLORREF colorA,COLORREF colorB);
	COLORREF DarkenColor( long lScale, COLORREF lColor);
	COLORREF GrayColor(COLORREF crColor);
	COLORREF GetAlphaBlendColor(COLORREF blendColor, COLORREF pixelColor,int weighting);
	COLORREF GetMenuColor(HMENU hMenu);
	COLORREF GetBitmapTransColor( HBITMAP hBitmap, INT xPixel, INT yPixel );
	int AddGrayIcon( CImageList* hIml, HICON hIcon, int nIndex );
	int AddGrayIcon( CImageList* hIml, HBITMAP hIcon, COLORREF transColor, int nIndex );
	int AddGloomIcon( CImageList* hIml, HICON hIcon, int nIndex, int cx, int cy );
	int AddStockImage( CImageList* hIml, CImageList* hImlDisabled, int iImage, BOOL bBitmap, HMODULE hModule, COLORREF bckColor, INT xPixel, INT yPixel );
	int AddMenuImageW( CImageList* hIml, CImageList* hImlDisabled, WCHAR* sImage, HICON hIcon );
	int AddMenuImage( CImageList* hIml, CImageList* hImlDisabled, LPCTSTR sImage, HICON hIcon );
	int AddImageW( CImageList* hIml, CImageList* hImlDisabled, HMODULE hResourceFile, WCHAR* sImage, COLORREF bckColor, INT xPixel, INT yPixel );
	int AddImage( CImageList* hIml, CImageList* hImlDisabled, HMODULE hResourceFile, LPCTSTR sImage, COLORREF bckColor, INT xPixel, INT yPixel );
	int AddImageID( CImageList* hIml, CImageList* hImlDisabled, HMODULE hResourceFile, WORD wImage, COLORREF bckColor, INT xPixel, INT yPixel );
	HANDLE localOpenThemeData( HWND hWnd, LPCWSTR pszClassList);
	HRESULT localCloseThemeData(HANDLE hTheme);
	HRESULT localGetThemeColor(HANDLE hTheme, int iPartId, int iStateId, int iColorId, COLORREF *pColor);

//protected:
	CImageList*	ilDummy;
};

class CPicture
{
public:
	void FreePictureData();
	HBITMAP Load(CString sFilePathName);
	BOOL LoadPictureData(BYTE* pBuffer, int nSize);
	BOOL Show(CDC* pDC, CPoint LeftTop, CPoint WidthHeight, int MagnifyX, int MagnifyY);
	BOOL Show(CDC* pDC, CRect DrawRect);

	CPicture();
	virtual ~CPicture();

	IPicture* m_IPicture; // Same As LPPICTURE (typedef IPicture __RPC_FAR *LPPICTURE)

	LONG      m_Height; // Height (In Pixels Ignor What Current Device Context Uses)
	LONG      m_Weight; // Size Of The Image Object In Bytes (File OR Resource)
	LONG      m_Width;  // Width (In Pixels Ignor What Current Device Context Uses)
};

#endif // !defined(AFX_COOLMENUGENERAL_H__77097D6E_6413_4148_B6F8_0E6083645D3D__INCLUDED_)
