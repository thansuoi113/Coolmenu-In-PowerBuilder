#include "coolmenu.h"
#include "menuhook.h"

// The following ifdef block is the standard way of creating macros which make exporting
// from a DLL simpler. All files within this DLL are compiled with the CM_MT_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see
// COOLMENUMAIN_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef COOLMENUMAIN_EXPORTS
#define COOLMENUMAIN_API __declspec(dllexport)
#else
#define COOLMENUMAIN_API __declspec(dllimport)
#endif

struct CHookData {
	CHookData*			pnext;		// next in chain of all allocated
	BOOL					bInUse;		// item is being used in a menu
	DWORD					threadID;
	HHOOK					hHook, hHookMsg;
	int					menuStyle;
	COLORREF				colors[MAX_COLOR];
	BOOL					bCoolMenubar;
	BOOL					bShadowEnabled;
	BOOL					bDrawFlatBorder;
	BOOL					bOrigCoolMenubar;
	BOOL					bOrigDrawFlatBorder;
	CImageList			hImagelist;
	CImageList			hImlDisabled;
	CMapStringToPtr	m_imageIdList;
	CMapStringToPtr	m_menuItemTextColorList;
	CMapStringToPtr	m_menuItemTextBoldList;
	CMapStringToPtr	m_menuItemTextItalicList;
	CMapStringToPtr	m_menuItemTextUnderlineList;
	CDWordArray*		hwndList;
	HMODULE				hResources;
	BOOL					bReactOnThemeChange;
};

CHookData* GetNewHookData( DWORD threadID );
CHookData* GetHookData( DWORD threadID );
void DeleteHookData(CHookData* hd);
BOOL IsHooked(DWORD threadID);
HHOOK GetHook(DWORD threadID);
void Unhook(DWORD threadID);
CHookData* Hook(DWORD threadID);
void AddWindow(CHookData* hd, HWND hwnd);

class CMenuHookList : CPtrList
{
public:
	CMenuHookList();
	~CMenuHookList();

	CMenuHook* GetNewMenuHook( HWND hWnd );
	CMenuHook* HookWindow( HWND hWnd );
	CMenuHook* GetMenuHook( HWND hWnd );
	BOOL IsSubMenu(int& xPos);
};

class CCoolmenuList : CPtrList
{
public:
	CCoolmenuList();
	~CCoolmenuList();

	CCoolmenu* GetNewCoolmenu( HWND hWnd );
	BOOL IsHooked( HWND hWnd );
	CCoolmenu* HookWindow( HWND hWnd, HMENU hMenu );
	CCoolmenu* GetCoolmenu( HWND hWnd );
	void SetImagelist( HIMAGELIST hSharedImagelist );
	void SetImageIdlist( CMapStringToPtr* imageIdlist );
	void SetMenuStyle( int menuStyle );
	int GetMenuStyle();
	void SetCoolMenubar( BOOL bCoolMenubar );
	void SetReactOnThemeChange( BOOL bReactOnThemeChange );
	BOOL GetReactOnThemeChange();
	void SetShadowEnabled( BOOL bShadowEnabled );
	void SetMenuColor( int colorIndex, COLORREF newColor );
	int AddStockImage( int iImage, BOOL bBitmap, HMODULE hModule, COLORREF bckColor, INT xPixel, INT yPixel );
	void SetImageNameIdW( WCHAR* iName, int iImage );
	void SetImageNameId(  LPCTSTR iName, int iImage );
	int AddMenuImageW( WCHAR* sImage, HICON hIcon );
	int AddMenuImage( LPCTSTR sImage, HICON hIcon );
	int AddImageW( WCHAR* sImage, COLORREF bckColor, INT xPixel, INT yPixel );
	int AddImage( LPCTSTR sImage, COLORREF bckColor, INT xPixel, INT yPixel );
	int AddImageID( WORD wImage, COLORREF bckColor, INT xPixel, INT yPixel );

	void SetDefaults();
	void SetMenuItemProps( LPCTSTR sName, BOOL bBold, BOOL bItalic, BOOL bUnderline, COLORREF clrText );
};

CHookData*			m_pHookData;
CCoolmenuList		m_coolmenuList;
CMenuHookList		m_menuHookList;
BOOL					m_isSystemMenu = FALSE, m_isContextMenu = FALSE, m_bInitMenu = FALSE, m_bPB105OrAbove;
CCoolmenuGeneral*	m_CmGeneral;
CString				m_sWindowClass, m_sPopupWindowClass, m_sToolbarClass, m_sFloatToolbarClass;
CCoolmenu*			m_currentCoolmenu;
COLORREF				m_default2K3Colors[MAX_COLOR] = { RGB(255,238,194), RGB(249,248,247), RGB(135,173,228), RGB(254,128,62),
																	 RGB(255,192,111), RGB(0,0,128), GetSysColor(COLOR_MENUTEXT), GetSysColor(COLOR_MENUTEXT),
																	 GetSysColor( COLOR_GRAYTEXT ), RGB(246,246,246), RGB( 150, 180, 240 ), RGB(163,194,245),
																	 RGB(255,214,154), RGB(255,244,204), RGB(147,181,231), RGB(227,239,255), RGB(0,45,150), RGB(106,140,203),
																	 RGB(0,0,128), RGB(0,0,128), RGB(141,141,141) };
COLORREF				m_default2K7Colors[MAX_COLOR] = { RGB(255,231,162), RGB(246,246,246), RGB(233,238,238), RGB(255,171,63),
																	 RGB(255,192,111), RGB(255,189,105), GetSysColor(COLOR_MENUTEXT), GetSysColor(COLOR_MENUTEXT),
																	 GetSysColor( COLOR_GRAYTEXT ), RGB(246,246,246), RGB( 150, 180, 240 ), RGB(191,219,255),
																	 RGB(255,223,132), RGB(255,245,204), RGB(153,191,240), RGB(227,239,254), RGB(101,147,207), RGB(154,198,255),
																	 RGB(251,140,60), RGB(255,171,63), RGB(141,141,141) };
COLORREF				m_defaultXpColors[MAX_COLOR] =  { RGB(193,210,238), RGB(252,252,249), RGB(239,237,222), RGB(152,181,226),
																	 RGB(225,230,232), RGB(49,106,197), GetSysColor(COLOR_MENUTEXT), GetSysColor(COLOR_MENUTEXT),
																	 RGB(166,166,166), NULL, NULL, RGB(212,208,200), RGB(193,210,238), RGB(193,210,238), RGB(239,237,222), RGB(193,210,238), RGB(138,134,122), RGB(197,194,184),
																	 RGB(49,106,197), RGB(49,106,197), RGB(180,177,163) };
enum ImageListType
{
  il_Normal,
  il_Disabled,
  il_Hot
};

//	Coolbar specific stuff
class CCoolbarImageList : public CObject
{
public:
	CCoolbarImageList( int ctlID );
	~CCoolbarImageList();

	void CreateImageLists( int cx, int cy, int cxLarge, int cyLarge );
	void DestroyImageLists();
	void SetIconSize( BOOL bSmall, int imagelistType, int cx, int cy );
	int GetCtlID();

	CImageList		CCoolbarImagelist, CCoolbarHotImagelist, CCoolbarDisabledImagelist, CCoolbarLargeImagelist, CCoolbarLargeHotImagelist, CCoolbarLargeDisabledImagelist;
	HMODULE			hCoolbarResources;
	int				m_ctlID, m_cx, m_cy, m_cxLarge, m_cyLarge;
};

class CCoolbarImageListList : CPtrList
{
public:
	CCoolbarImageListList();
	~CCoolbarImageListList();

	CCoolbarImageList* GetNewCoolbarImagelist();
	CCoolbarImageList* GetCoolbarImagelist( int ctlID );

	int	ctlID;
};

CCoolbarImageListList		m_coolbarList;

LRESULT CALLBACK GetMsgProc( int nCode, WPARAM wParam, LPARAM lParam );
LRESULT CALLBACK CallWindowProc( int nCode, WPARAM wParam, LPARAM lParam );
CHookData* CoolMenuInitialize( HWND hwnd );
void CoolMenuUninitialize( HWND hwnd );
BOOL IsSystemMenu( HWND hWnd, BOOL bSysMenu, HMENU hMenu );

extern "C"
{
	COOLMENUMAIN_API void __stdcall InstallCoolMenu( HWND hWnd );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall Uninstall();}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetShadowEnabled( BOOL bShadowEnabled );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetCoolMenubar( BOOL bCoolMenubar );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetReactOnThemeChange( BOOL bReactOnThemeChange );}

extern "C"
{
	COOLMENUMAIN_API BOOL __stdcall GetReactOnThemeChange();}

extern "C"
{
	COOLMENUMAIN_API BOOL __stdcall GetCoolMenubar();}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetFlatMenu( BOOL bFlatMenu );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetPbVersion( LPCTSTR lpszVersion );}

extern "C"
{
	COOLMENUMAIN_API BOOL __stdcall SetResourceFile( LPCTSTR sResourceFile );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetResourceFileW( WCHAR* sResourceFile );}

extern "C"
{
	COOLMENUMAIN_API int __stdcall AddStockImage( int iImage, BOOL bBitmap, HMODULE hModule, COLORREF bckColor, INT xPixel, INT yPixel );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetImageNameIdW( WCHAR* iName, int iImage );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetImageNameId( LPCTSTR iName, int iImage );}

extern "C"
{
	COOLMENUMAIN_API int __stdcall AddImageW( WCHAR* sImage, COLORREF bckColor, INT xPixel = 0, INT yPixel = 0 );}

extern "C"
{
	COOLMENUMAIN_API int __stdcall AddImage( LPCTSTR sImage, COLORREF bckColor, INT xPixel = 0, INT yPixel = 0 );}

extern "C"
{
	COOLMENUMAIN_API int __stdcall AddImageID( WORD wImage, COLORREF bckColor, INT xPixel = 0, INT yPixel = 0 );}

extern "C"
{
	COOLMENUMAIN_API int __stdcall AddMenuImageW( WCHAR* sImage, HICON hIcon );}

extern "C"
{
	COOLMENUMAIN_API int __stdcall AddMenuImage( LPCTSTR sImage, HICON hIcon );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetMenuStyle(int menuStyle);}

extern "C"
{
	COOLMENUMAIN_API int __stdcall GetMenuStyle();}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetMenuColor( int colorIndex, COLORREF clr );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetDefaults();}

extern "C"
{
	COOLMENUMAIN_API void __stdcall SetMenuItemProps( LPCTSTR sName, BOOL bBold, BOOL bItalic, BOOL bUnderline, COLORREF clrText );}

extern "C"
{
	COOLMENUMAIN_API HFONT __stdcall GetFont();}

extern "C"
{
	COOLMENUMAIN_API int __stdcall CoolbarCreateImageLists( int cx, int cy, int cxLarge, int cyLarge );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall CoolbarDestroyImageLists( int cltID );}

extern "C"
{
	COOLMENUMAIN_API void __stdcall CoolbarSetIconSize( int ctlID, BOOL bSmall, int imagelistType, int cltID, int cx, int cy );}

extern "C"
{
	COOLMENUMAIN_API int __stdcall CoolbarAddStockImage( int ctlID, int iImage, BOOL bBitmap, HMODULE hModule, COLORREF bckColor, INT xPixel, INT yPixel, BOOL bSmall = TRUE, int imagelistType = il_Normal );}

extern "C"
{
	COOLMENUMAIN_API int __stdcall CoolbarAddImage( int ctlID, LPCTSTR sImage, COLORREF bckColor, INT xPixel = 0, INT yPixel = 0, BOOL bSmall = TRUE, int imagelistType = il_Normal );}

extern "C"
{
	COOLMENUMAIN_API int __stdcall CoolbarAddImageID( int ctlID, WORD wImage, COLORREF bckColor, INT xPixel = 0, INT yPixel = 0, BOOL bSmall = TRUE, int imagelistType = il_Normal );}

extern "C"
{
	COOLMENUMAIN_API BOOL __stdcall CoolbarSetResourceFile( int ctlID, LPCTSTR sResourceFile );}

extern "C"
{
	COOLMENUMAIN_API HIMAGELIST __stdcall CoolbarGetImageListHandle( int ctlID, BOOL bSmall, int imagelistType );}

extern "C"
{
	COOLMENUMAIN_API HBITMAP __stdcall GetFocusBitmap( BOOL bFloating );}