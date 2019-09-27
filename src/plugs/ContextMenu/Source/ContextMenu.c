#define WIN32_LEAN_AND_MEAN
#define _WIN32_IE 0x0500
#include <windows.h>
#include <richedit.h>
#include <shellapi.h>
#include <objbase.h>
#include <shlobj.h>
#include "AkelEdit.h"
#include "AkelDLL.h"
#include "StackFunc.h"
#include "StrFunc.h"
#include "WideFunc.h"
#include "IconMenu.h"
#include "Resources\Resource.h"


/*
//Include stack functions
#define StackGetElement
#define StackInsertAfter
#define StackInsertBefore
#define StackInsertIndex
#define StackMoveAfter
#define StackMoveBefore
#define StackMoveIndex
#define StackDelete
#define StackClear
#define StackSize
#include "StackFunc.h"

//Include string functions
#define WideCharLower
#define xmemcpy
#define xmemset
#define xstrcmpW
#define xstrcmpiW
#define xstrcmpinW
#define xstrcmpnW
#define xstrlenW
#define xstrcpyW
#define xstrcpynW
#define xstrstrW
#define xstrrepW
#define xatoiA
#define xatoiW
#define xitoaW
#define xuitoaW
#define dec2hexW
#define hex2decW
#define xprintfW
#include "StrFunc.h"

//Include wide functions
#define AppendMenuWide
#define ComboBox_AddStringWide
#define CreateDialogWide
#define CreateProcessWide
#define CreateWindowExWide
#define DialogBoxParamWide
#define DialogBoxWide
#define DispatchMessageWide
#define ExpandEnvironmentStringsWide
#define ExtractIconExWide
#define FileExistsWide
#define GetFileAttributesWide
#define GetMenuStringWide
#define GetMessageWide
#define GetWindowLongPtrWide
#define GetWindowTextWide
#define InsertMenuWide
#define IsDialogMessageWide
#define ListBox_GetTextWide
#define ListBox_InsertStringWide
#define ModifyMenuWide
#define SearchPathWide
#define SetDlgItemTextWide
#define SetWindowLongPtrWide
#define SetWindowTextWide
#define ShellExecuteWide
#define TranslateAcceleratorWide
#include "WideFunc.h"
//*/

//Include icon menu functions
#define ICONMENU_INCLUDE
#include "IconMenu.h"

//Defines
#define DLLA_CONTEXTMENU_INDEX        1
#define DLLA_CONTEXTMENU_STARTSTOP    10
#define DLLA_CONTEXTMENU_SHOWSUBMENU  1
#define DLLA_CONTEXTMENU_SHOWMAINMENU 2

#define STRID_MANUAL                          1
#define STRID_MENUMAIN                        2
#define STRID_MENUEDIT                        3
#define STRID_MENUTAB                         4
#define STRID_MENUURL                         5
#define STRID_MENURECENTFILES                 6
#define STRID_ENABLE                          7
#define STRID_HIDE                            8
#define STRID_SHOW                            9
#define STRID_AUTOLOAD                        10
#define STRID_NOSUBMENUINDEX                  11
#define STRID_NOSUBMENUNAME                   12
#define STRID_PARSEMSG_NONPARENT              13
#define STRID_PARSEMSG_UNKNOWNMETHOD          14
#define STRID_PARSEMSG_METHODALREADYDEFINED   15
#define STRID_PARSEMSG_NOMETHOD               16
#define STRID_PARSEMSG_UNNAMEDSUBMENU         17
#define STRID_PARSEMSG_NOOPENBRACKET          18
#define STRID_MENU_OPEN                       19
#define STRID_MENU_MOVEUP                     20
#define STRID_MENU_MOVEDOWN                   21
#define STRID_MENU_SORT                       22
#define STRID_MENU_DELETE                     23
#define STRID_MENU_DELETEOLD                  24
#define STRID_MENU_EDIT                       25
#define STRID_FAVOURITES                      26
#define STRID_SHOWFILE                        27
#define STRID_FAVADDING                       28
#define STRID_FAVEDITING                      29
#define STRID_FAVNAME                         30
#define STRID_FAVFILE                         31
#define STRID_PLUGIN                          32
#define STRID_OK                              33
#define STRID_CANCEL                          34
#define STRID_CLOSE                           35
#define STRID_DEFAULTMANUAL                   36
#define STRID_DEFAULTMAIN                     37
#define STRID_DEFAULTEDIT                     38
#define STRID_DEFAULTTAB                      39
#define STRID_DEFAULTURL                      40
#define STRID_DEFAULTRECENTFILES              41

#define AKDLL_MENUINDEX   (WM_USER + 100)

#define OF_MENUSTEXT      0x1
#define OF_MAINRECT       0x2
#define OF_FAVTEXT        0x4
#define OF_FAVRECT        0x8

#define BUFFER_SIZE      1024

#define EXTACT_COMMAND    1
#define EXTACT_CALL       2
#define EXTACT_EXEC       3
#define EXTACT_OPENFILE   4
#define EXTACT_SAVEFILE   5
#define EXTACT_FONT       6
#define EXTACT_RECODE     7
#define EXTACT_INSERT     8
#define EXTACT_LINK       9
#define EXTACT_FAVOURITE  10
#define EXTACT_MENU       11

#define EXTPARAM_CHAR     1
#define EXTPARAM_INT      2

#define URL_OPEN          1
#define URL_COPY          2
#define URL_SELECT        3
#define URL_CUT           4
#define URL_PASTE         5
#define URL_DELETE        6

#define FAV_ADDDIALOG     1
#define FAV_ADDSILENT     2
#define FAV_MANAGE        3
#define FAV_DELSILENT     4

#define TYPE_UNKNOWN     -1
#define TYPE_MANUAL       0
#define TYPE_MAIN         1
#define TYPE_EDIT         2
#define TYPE_TAB          3
#define TYPE_URL          4
#define TYPE_RECENTFILES  5
#define TYPE_MAX          6

//CreateContextMenu SET() flags
#define CCMS_NOSDI       0x01
#define CCMS_NOMDI       0x02
#define CCMS_NOPMDI      0x04
#define CCMS_NOSHORTCUT  0x08
#define CCMS_NOICONS     0x10
#define CCMS_NOFILEEXIST 0x20

#define IDM_MIN_EXPLORER    20001
#define IDM_MAX_EXPLORER    20999
#define IDM_SEPARATOR       21000
#define IDM_SEPARATOR1      21001
#define IDM_MIN_FAVOURITES  21002
#define IDM_MAX_FAVOURITES  22000
#define IDM_MIN_MANUAL      22001
#define IDM_MAX_MANUAL      25000
#define IDM_MIN_MENU        25001
#define IDM_MAX_MENU        30000

#define FS_NONE            0
#define FS_FONTNORMAL      1
#define FS_FONTBOLD        2
#define FS_FONTITALIC      3
#define FS_FONTBOLDITALIC  4

#define INIT_EXPLORER      1
#define INIT_FAVOURITES    2
#define INIT_RECENTFILES   3
#define INIT_LANGUAGES     4
#define INIT_OPENCODEPAGES 5
#define INIT_SAVECODEPAGES 6

#define IMENU_EDIT     0x00000001
#define IMENU_CHECKS   0x00000004

//GetEditPos type
#define GEP_LEFTTOP     -1
#define GEP_RIGHTTOP    -2
#define GEP_RIGHTBOTTOM -3
#define GEP_LEFTBOTTOM  -4
#define GEP_CARET       -5
#define GEP_CURSOR      -6

#ifndef WM_MENURBUTTONUP
  #define WM_MENURBUTTONUP 0x0122
#endif
#ifndef TPM_RECURSE
  #define TPM_RECURSE 0x0001L
#endif

typedef struct _FAVITEM {
  struct _FAVITEM *next;
  struct _FAVITEM *prev;
  wchar_t wszName[MAX_PATH];
  wchar_t wszFile[MAX_PATH];
} FAVITEM;

typedef struct {
  FAVITEM *first;
  FAVITEM *last;
  int nElements;
} STACKFAV;

typedef struct _MAININDEXITEM {
  struct _MAININDEXITEM *next;
  struct _MAININDEXITEM *prev;
  int nMainMenuIndex;
  BOOL bMenuBreak;
  HMENU hSubMenu;
} MAININDEXITEM;

typedef struct {
  MAININDEXITEM *first;
  MAININDEXITEM *last;
} STACKMAININDEX;

typedef struct _EXTPARAM {
  struct _EXTPARAM *next;
  struct _EXTPARAM *prev;
  DWORD dwType;
  INT_PTR nNumber;
  char *pString;
  wchar_t *wpString;
  char *pExpanded;
  int nExpandedAnsiLen;
  wchar_t *wpExpanded;
  int nExpandedUnicodeLen;
} EXTPARAM;

typedef struct {
  EXTPARAM *first;
  EXTPARAM *last;
  int nElements;
} STACKEXTPARAM;

typedef struct _MENUITEM {
  struct _MENUITEM *next;
  struct _MENUITEM *prev;
  BOOL bUpdateItem;
  BOOL bAutoLoad;
  DWORD dwAction;
  int nTextOffset;
  int nItem;
  int nMenuType;
  STACKEXTPARAM hParamStack;

  //SubMenu
  HMENU hSubMenu;
  int nSubMenuIndex;

  //Icon
  wchar_t wszIconFile[MAX_PATH];
  int nFileIconIndex;
  int nImageListIconIndex;
} MENUITEM;

typedef struct {
  MENUITEM *first;
  MENUITEM *last;
} STACKMENUITEM;

typedef struct _SUBMENUITEM {
  struct _SUBMENUITEM *next;
  struct _SUBMENUITEM *prev;
  HMENU hSubMenu;
  int nSubMenuIndex;
  int nSubMenuCodeItem;
} SUBMENUITEM;

typedef struct {
  SUBMENUITEM *first;
  SUBMENUITEM *last;
} STACKSUBMENU;

typedef struct _SHOWSUBMENUITEM {
  struct _SHOWSUBMENUITEM *next;
  struct _SHOWSUBMENUITEM *prev;
  MENUITEM *lpMenuItem;
  wchar_t *wpShowMenuName;
  wchar_t wszMenuItem[MAX_PATH];
} SHOWSUBMENUITEM;

typedef struct {
  SHOWSUBMENUITEM *first;
  SHOWSUBMENUITEM *last;
} STACKSHOWSUBMENU;

typedef struct _POPUPMENU {
  HICONMENU hIconMenu;
  HIMAGELIST hImageList;
  int nImageListIconCount;
  HMENU hPopupMenu;
  STACKMENUITEM hMenuItemStack;
  STACKMAININDEX hMainMenuIndexStack;
  STACKSHOWSUBMENU hShowSubmenuStack;
  HMENU hMainMenu;
  BOOL bClearMainMenu;
  BOOL bLinkToManualMenu;
  HMENU hExplorerSubMenu;
  int nExplorerFirstIndex;
  int nExplorerLastIndex;
  HMENU hFavouritesSubMenu;
  int nFavouritesFirstIndex;
  int nFavouritesLastIndex;
  HMENU hRecentFilesSubMenu;
  int nRecentFilesFirstIndex;
  int nRecentFilesLastIndex;
  HMENU hLanguagesSubMenu;
  int nLanguagesFirstIndex;
  int nLanguagesLastIndex;
  HMENU hOpenCodepagesSubMenu;
  int nOpenCodepagesFirstIndex;
  int nOpenCodepagesLastIndex;
  HMENU hSaveCodepagesSubMenu;
  int nSaveCodepagesFirstIndex;
  int nSaveCodepagesLastIndex;
  HMENU hMdiDocumentsSubMenu;
  LPCONTEXTMENU pExplorerMenu;
  LPCONTEXTMENU2 pExplorerSubMenu2;
  LPCONTEXTMENU3 pExplorerSubMenu3;
  wchar_t wszExplorerFile[MAX_PATH];
} POPUPMENU;

typedef struct _PLUGINCALL {
  unsigned char *lpStruct;
  BOOL bAutoLoad;
} PLUGINCALL;

typedef struct {
  wchar_t *wpText;
  BOOL bEnable;
} CONTROLTYPE;

//Coder external call
typedef struct {
  UINT_PTR dwStructSize;
  INT_PTR nAction;
  wchar_t *wpAlias;
  BOOL *lpbActive;
} DLLEXTCODERCHECKALIAS;

typedef struct {
  UINT_PTR dwStructSize;
  INT_PTR nAction;
  wchar_t *wpVarTheme;
  BOOL *lpbActive;
} DLLEXTCODERCHECKVARTHEME;

typedef struct {
  UINT_PTR dwStructSize;
  INT_PTR nAction;
  INT_PTR nMarkID;
  wchar_t *wpColorText;
  wchar_t *wpColorBk;
  BOOL *lpbActive;
} DLLEXTCODERCHECKMARK;

#define DLLA_CODER_SETEXTENSION            1
#define DLLA_CODER_SETVARTHEME             5
#define DLLA_CODER_CHECKALIAS              12
#define DLLA_CODER_CHECKVARTHEME           14
#define DLLA_HIGHLIGHT_MARK                2
#define DLLA_HIGHLIGHT_CHECKMARK           11

//SpecialChar external call
typedef struct {
  UINT_PTR dwStructSize;
  INT_PTR nAction;
  unsigned char *pSpecialChar;
  COLORREF *lpcrColor;
  COLORREF *lpcrSelColor;
  BOOL *lpbColorEnable;
  BOOL *lpbSelColorEnable;
} DLLEXTSPECIALCHAR;

#define DLLA_SPECIALCHAR_OLDSET 1
#define DLLA_SPECIALCHAR_OLDGET 2

//LineBoard external call
typedef struct {
  UINT_PTR dwStructSize;
  INT_PTR nAction;
  int *lpnRulerHeight;
  BOOL *lpbRulerEnable;
} DLLEXTLINEBOARD;

#define DLLA_LINEBOARD_GETRULERHEIGHT   2
#define DLLA_LINEBOARD_SETRULERHEIGHT   3

//Functions prototypes
DWORD WINAPI ThreadProc(LPVOID lpParameter);
LRESULT CALLBACK NewMainProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK NewFrameProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK EditParentMessages(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK MainDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK NewEditDlgProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

LRESULT CALLBACK FavListDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK NewListBoxProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
BOOL CALLBACK FavEditDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
void CreateFavStack(STACKFAV *hStack, const wchar_t *wpText);
FAVITEM* StackInsertFavourite(STACKFAV *hStack);
FAVITEM* StackGetFavouriteByIndex(STACKFAV *hStack, int nIndex);
FAVITEM* StackGetFavouriteByFile(STACKFAV *hStack, const wchar_t *wpFile, int *nIndex);
int StackSortFavourites(STACKFAV *hStack, int nUpDown, BOOL bShowFile);
void StackFreeFavourites(STACKFAV *hStack);
void FillFavouritesListBox(STACKFAV *hStack, HWND hWnd, BOOL bShowFile);
int MoveListBoxItem(STACKFAV *hStack, HWND hWnd, int nOldIndex, int nNewIndex);
BOOL ShiftListBoxSelItems(STACKFAV *hStack, HWND hWnd, BOOL bMoveDown);
int DeleteListBoxSelItems(STACKFAV *hStack, HWND hWnd);
int DeleteListBoxOldItems(STACKFAV *hStack, HWND hWnd);
int GetListBoxSelItems(HWND hWnd, int **lpSelItems);
void FreeListBoxSelItems(int **lpSelItems);

BOOL CreateContextMenu(POPUPMENU *hMenuStack, const wchar_t *wpText, int nType);
DWORD IsFlagOn(DWORD dwSetFlags, DWORD dwCheckFlags);
void UpdateContextMenu(POPUPMENU *hMenuStack, int nType, HMENU hSubMenu);
int ShowContextMenu(POPUPMENU *hMenuStack, HWND hWnd, int nType, int x, int y, HMENU hMenu);
void ShowUrlMenu(HWND hWnd, CHARRANGE64 *crUrl, int x, int y);
void InitMenuPopup(POPUPMENU *hMenuStack, int nMenuType);
HMENU InsertMainMenu(POPUPMENU *hMenuStack);
void SetMainMenu(POPUPMENU *hMenuStack, HMENU hMenu, BOOL bDrawMenuBar);
void UnsetMainMenu(POPUPMENU *hMenuStack);
void ShowMainMenu(BOOL bShow);
MENUITEM* GetContextMenuItem(POPUPMENU *hMenuStack, int nItem);
MENUITEM* GetCommandItem(POPUPMENU *hMenuStack, int nCommand);
void ViewItemCode(MENUITEM *lpElement);
void CallContextMenu(POPUPMENU *hMenuStack, int nItem);
void FreeContextMenu(POPUPMENU *hMenuStack);

BOOL InsertMenuCommon(HICONMENU hIconMenu, HIMAGELIST hImageList, INT_PTR nIconIndex, int nIconWidth, int nIconHeight, HMENU hMenu, int uPosition, UINT uFlags, UINT_PTR uIDNewItem, const wchar_t *lpNewItem);
BOOL ModifyMenuCommon(HICONMENU hIconMenu, HIMAGELIST hImageList, INT_PTR nIconIndex, int nIconWidth, int nIconHeight, HMENU hMenu, int uPosition, UINT uFlags, UINT_PTR uIDNewItem, const wchar_t *lpNewItem);
BOOL DeleteMenuCommon(HICONMENU hIconMenu, HMENU hMenu, int uPosition, UINT uFlags);
BOOL GetExplorerMenu(LPCONTEXTMENU *pContextMenu, LPCONTEXTMENU2 *pContextSubMenu2, LPCONTEXTMENU3 *pContextSubMenu3, HWND hWnd, wchar_t *wpFile);
LPITEMIDLIST NextPIDL(LPCITEMIDLIST pidl);
int GetSubMenuIndex(HMENU hMenu, HMENU hSubMenu);
int GetSubMenuCount(HMENU hMenu);
HMENU FindRootSubMenuByName(POPUPMENU *hMenuStack, const wchar_t *wpSubMenuName);
void RemoveSeparator1(POPUPMENU *hMenuStack, HMENU hSubMenu);
void GetEditPos(HWND hWnd, POINT *pt, int nType);
void ShowStandardEditMenu(HWND hWnd, HMENU hMenu, BOOL bMouse);

void ParseMethodParameters(STACKEXTPARAM *hParamStack, const wchar_t *wpText, const wchar_t **wppText);
void ExpandMethodParameters(STACKEXTPARAM *hParamStack, const wchar_t *wpFile, const wchar_t *wpExeDir, HMENU hPopupMenu, int nMenuItem, const wchar_t *wpUrlLink);
int StructMethodParameters(STACKEXTPARAM *hParamStack, unsigned char *lpStruct);
EXTPARAM* GetMethodParameter(STACKEXTPARAM *hParamStack, int nIndex);
void GetIconParameters(const wchar_t *wpText, wchar_t *wszIconFile, int nMaxIconFile, int *nIconIndex, const wchar_t **wppText);
void FreeMethodParameters(STACKEXTPARAM *hParamStack);
int GetMethodName(const wchar_t *wpText, wchar_t *wszStr, int nStrLen, const wchar_t **wppText);
int NextString(const wchar_t *wpText, wchar_t *wszStr, int nStrLen, const wchar_t **wppText, int *nMinus);
BOOL SkipComment(const wchar_t **wpText);
int GetFileDir(const wchar_t *wpFile, int nFileLen, wchar_t *wszFileDir, DWORD dwFileDirLen);
INT_PTR TranslateEscapeString(HWND hWndEdit, const wchar_t *wpInput, wchar_t *wszOutput, DWORD *lpdwCaret);
int TranslateFileString(const wchar_t *wpString, wchar_t *wszBuffer, int nBufferSize);

int GetCurFile(wchar_t *wszFile, int nMaxFile);
const wchar_t* GetFileName(const wchar_t *wpFile, int nFileLen);
INT_PTR GetEditText(HWND hWnd, wchar_t **wpText);
BOOL GetWindowPos(HWND hWnd, HWND hWndOwner, RECT *rc);
DWORD ScrollCaret(HWND hWnd);

INT_PTR WideOption(HANDLE hOptions, const wchar_t *pOptionName, DWORD dwType, BYTE *lpData, DWORD dwData);
void ReadOptions(DWORD dwFlags);
void SaveOptions(DWORD dwFlags);
wchar_t* GetDefaultMenu(int nStringID);
const char* GetLangStringA(LANGID wLangID, int nStringID);
const wchar_t* GetLangStringW(LANGID wLangID, int nStringID);
BOOL IsExtCallParamValid(LPARAM lParam, int nIndex);
INT_PTR GetExtCallParam(LPARAM lParam, int nIndex);
void InitCommon(PLUGINDATA *pd);
void InitMain();
void UninitMain();

//Global variables
wchar_t wszBuffer[BUFFER_SIZE];
wchar_t wszPluginName[MAX_PATH];
wchar_t wszPluginTitle[MAX_PATH];
HANDLE hHeap;
HINSTANCE hInstanceDLL;
HWND hMainWnd;
HWND hMdiClient;
HMENU hMainMenu;
HMENU hNewMainMenu;
HMENU hMenuWindow;
HMENU hMenuRecentFiles;
HMENU hMenuLanguage;
HMENU hPopupEdit;
HMENU hPopupCodepage;
HMENU hPopupOpenCodepage;
HMENU hPopupSaveCodepage;
HICON hMainIcon;
HACCEL hGlobalAccel;
BOOL bOldWindows;
BOOL bAkelEdit;
int nMDI;
LANGID wLangModule;
BOOL bInitCommon=FALSE;
BOOL bInitMain=FALSE;
DWORD dwSaveFlags=0;
HANDLE hThread=NULL;
DWORD dwThreadId;
HWND hWndMainDlg=NULL;
RECT rcMainMinMaxDialog={272, 186, 0, 0};
RECT rcMainCurrentDialog={0};
BOOL bMainDialogRectSave=FALSE;
BOOL bMenuMainOnStart=FALSE;
BOOL bMenuMainEnable=TRUE;
BOOL bMenuEditEnable=TRUE;
BOOL bMenuTabEnable=TRUE;
BOOL bMenuUrlEnable=TRUE;
BOOL bMenuRecentFilesEnable=TRUE;
POPUPMENU hMenuManualStack={0};
POPUPMENU hMenuMainStack={0};
POPUPMENU hMenuEditStack={0};
POPUPMENU hMenuTabStack={0};
POPUPMENU hMenuUrlStack={0};
POPUPMENU hMenuRecentFilesStack={0};
POPUPMENU hMenuDialogStack={0};
POPUPMENU *lpCurrentMenuStack=NULL;
int nCurrentMenuType=TYPE_UNKNOWN;
wchar_t wszCurrentFile[MAX_PATH];
wchar_t wszExeDir[MAX_PATH];
wchar_t wszUrlLink[MAX_PATH];
wchar_t *wszManualText=NULL;
wchar_t *wszMainText=NULL;
wchar_t *wszEditText=NULL;
wchar_t *wszTabText=NULL;
wchar_t *wszUrlText=NULL;
wchar_t *wszRecentFilesText=NULL;
wchar_t wszExtMenuSearch[MAX_PATH]=L"";
CHARRANGE64 crExtSetSel={0};
CHARRANGE64 crUrlMenuShow={0};
int nExtMenuIndex=0;
BOOL bExtFocusEdit=FALSE;
BOOL bMenuUrlShow=FALSE;
BOOL bUseIcons=TRUE;
BOOL bBigIcons=FALSE;
SIZE sizeIcon;
HMENU hMenuMainHide=NULL;
BOOL bMenuMainHide=FALSE;
WNDPROC OldEditDlgProc=NULL;
WNDPROCDATA *NewMainProcData=NULL;
WNDPROCDATA *NewFrameProcData=NULL;

HWND hWndFavDlg=NULL;
HICON hIconFav=NULL;
HICON hIconHollow=NULL;
RECT rcFavMinMaxDialog={305, 183, 0, 0};
RECT rcFavCurrentDialog={0};
BOOL bFavDialogRectSave=FALSE;
STACKFAV hFavStack={0};
wchar_t *wszFavText=NULL;
HWND hFavToolbar=NULL;
HIMAGELIST hFavImageList=NULL;
BOOL bFavShowFile=FALSE;
WNDPROC OldListBoxProc;


//Identification
void __declspec(dllexport) DllAkelPadID(PLUGINVERSION *pv)
{
  pv->dwAkelDllVersion=AKELDLL;
  pv->dwExeMinVersion3x=MAKE_IDENTIFIER(-1, -1, -1, -1);
  pv->dwExeMinVersion4x=MAKE_IDENTIFIER(4, 8, 8, 0);
  pv->pPluginName="ContextMenu";
}

//Plugin extern function
void __declspec(dllexport) Main(PLUGINDATA *pd)
{
  pd->dwSupport|=PDS_SUPPORTALL;
  if (pd->dwSupport & PDS_GETSUPPORT)
    return;

  if (!bInitCommon) InitCommon(pd);

  if (pd->lParam)
  {
    INT_PTR nAction=GetExtCallParam(pd->lParam, 1);

    if (nAction == DLLA_CONTEXTMENU_INDEX)
    {
      unsigned char *pSearchString=NULL;

      if (IsExtCallParamValid(pd->lParam, 2))
        nExtMenuIndex=(int)GetExtCallParam(pd->lParam, 2);
      if (IsExtCallParamValid(pd->lParam, 3))
        pSearchString=(unsigned char *)GetExtCallParam(pd->lParam, 3);
      if (IsExtCallParamValid(pd->lParam, 4))
        crExtSetSel.cpMin=GetExtCallParam(pd->lParam, 4);
      if (IsExtCallParamValid(pd->lParam, 5))
        crExtSetSel.cpMax=GetExtCallParam(pd->lParam, 5);

      if (pSearchString)
      {
        if (pd->dwSupport & PDS_STRANSI)
          MultiByteToWideChar(CP_ACP, 0, (char *)pSearchString, -1, wszExtMenuSearch, MAX_PATH);
        else
          xstrcpynW(wszExtMenuSearch, (wchar_t *)pSearchString, MAX_PATH);
      }
      bExtFocusEdit=TRUE;
    }
    else if (nAction == DLLA_CONTEXTMENU_STARTSTOP)
    {
      if (bInitMain)
      {
        if (hWndMainDlg) SendMessage(hWndMainDlg, WM_COMMAND, IDCANCEL, 0);
        UnsetMainMenu(&hMenuMainStack);
        if (bMenuMainHide) ShowMainMenu(TRUE);
        UninitMain();

        //Unload plugin
        pd->nUnload=UD_UNLOAD;
      }
      else
      {
        InitMain();

        if (bMenuMainEnable)
        {
          CreateContextMenu(&hMenuMainStack, wszMainText, TYPE_MAIN);
          hNewMainMenu=InsertMainMenu(&hMenuMainStack);
          SetMainMenu(&hMenuMainStack, hNewMainMenu, !bMenuMainHide);
        }
        if (bMenuMainHide) ShowMainMenu(FALSE);

        //Stay in memory, and show as active
        pd->nUnload=UD_NONUNLOAD_ACTIVE;
      }
      return;
    }
  }

  //Initialize
  if (!bInitMain) InitMain();

  if (!pd->bOnStart)
  {
    if (!hThread)
      hThread=CreateThread(NULL, 0, ThreadProc, NULL, 0, &dwThreadId);
    else
      PostMessage(hWndMainDlg, AKDLL_MENUINDEX, nExtMenuIndex, 0);
  }

  //Stay in memory, and show as active
  pd->nUnload=UD_NONUNLOAD_ACTIVE;
}

void __declspec(dllexport) Show(PLUGINDATA *pd)
{
  wchar_t wszSubMenuName[MAX_PATH];
  POINT ptPos;
  int nSubMenuIndex;

  //Function doesn't support autoload
  pd->dwSupport|=PDS_NOAUTOLOAD;
  if (pd->dwSupport & PDS_GETSUPPORT)
    return;

  if (!bInitCommon) InitCommon(pd);

  //Initialize
  if (!bInitMain) InitMain();

  //Default values
  ptPos.x=GEP_LEFTTOP;
  ptPos.y=GEP_LEFTTOP;
  nSubMenuIndex=-1;
  wszSubMenuName[0]=L'\0';

  if (pd->lParam)
  {
    INT_PTR nAction=GetExtCallParam(pd->lParam, 1);

    if (nAction == DLLA_CONTEXTMENU_SHOWSUBMENU)
    {
      unsigned char *pPosX=NULL;
      unsigned char *pPosY=NULL;
      unsigned char *pSubMenuName=NULL;

      if (IsExtCallParamValid(pd->lParam, 2))
        pPosX=(unsigned char *)GetExtCallParam(pd->lParam, 2);
      if (IsExtCallParamValid(pd->lParam, 3))
        pPosY=(unsigned char *)GetExtCallParam(pd->lParam, 3);
      if (IsExtCallParamValid(pd->lParam, 4))
        nSubMenuIndex=(int)GetExtCallParam(pd->lParam, 4);
      if (IsExtCallParamValid(pd->lParam, 5))
        pSubMenuName=(unsigned char *)GetExtCallParam(pd->lParam, 5);

      if (pPosX && pPosY)
      {
        if (pd->dwSupport & PDS_STRANSI)
        {
          ptPos.x=(int)xatoiA((char *)pPosX, NULL);
          ptPos.y=(int)xatoiA((char *)pPosY, NULL);
        }
        else
        {
          ptPos.x=(int)xatoiW((wchar_t *)pPosX, NULL);
          ptPos.y=(int)xatoiW((wchar_t *)pPosY, NULL);
        }
        if (ptPos.x == ptPos.y && ptPos.x < 0)
        {
          if (pd->hWndEdit)
            GetEditPos(pd->hWndEdit, &ptPos, ptPos.x);
          else
            GetEditPos(pd->hMdiClient, &ptPos, ptPos.x == GEP_CARET?GEP_LEFTTOP:ptPos.x);
        }
      }
      if (pSubMenuName)
      {
        if (pd->dwSupport & PDS_STRANSI)
          MultiByteToWideChar(CP_ACP, 0, (char *)pSubMenuName, -1, wszSubMenuName, MAX_PATH);
        else
          xstrcpynW(wszSubMenuName, (wchar_t *)pSubMenuName, MAX_PATH);
      }
    }
    else if (nAction == DLLA_CONTEXTMENU_SHOWMAINMENU)
    {
      unsigned char *pPosX=NULL;
      unsigned char *pPosY=NULL;

      if (IsExtCallParamValid(pd->lParam, 2))
        pPosX=(unsigned char *)GetExtCallParam(pd->lParam, 2);
      if (IsExtCallParamValid(pd->lParam, 3))
        pPosY=(unsigned char *)GetExtCallParam(pd->lParam, 3);

      if (pPosX && pPosY)
      {
        if (pd->dwSupport & PDS_STRANSI)
        {
          ptPos.x=(int)xatoiA((char *)pPosX, NULL);
          ptPos.y=(int)xatoiA((char *)pPosY, NULL);
        }
        else
        {
          ptPos.x=(int)xatoiW((wchar_t *)pPosX, NULL);
          ptPos.y=(int)xatoiW((wchar_t *)pPosY, NULL);
        }
        if (ptPos.x == ptPos.y && ptPos.x < 0)
        {
          if (pd->hWndEdit)
            GetEditPos(pd->hWndEdit, &ptPos, ptPos.x);
          else
            GetEditPos(pd->hMdiClient, &ptPos, ptPos.x == GEP_CARET?GEP_LEFTTOP:ptPos.x);
        }

        if (bMenuMainEnable)
        {
          GetCurFile(wszCurrentFile, MAX_PATH);
          lpCurrentMenuStack=&hMenuMainStack;
          nCurrentMenuType=TYPE_MAIN;
          UpdateContextMenu(lpCurrentMenuStack, TYPE_MAIN, NULL);
        }
        if (!bMenuMainEnable || !hMenuMainStack.bClearMainMenu)
        {
          wchar_t wszMenuItem[MAX_PATH];
          HMENU hPopupMenu;
          HMENU hSubMenu;
          int nSubMenuIndex;

          if (hPopupMenu=CreatePopupMenu())
          {
            for (nSubMenuIndex=0; hSubMenu=GetSubMenu(hMainMenu, nSubMenuIndex); ++nSubMenuIndex)
            {
              if (GetMenuStringWide(hMainMenu, nSubMenuIndex, wszMenuItem, MAX_PATH, MF_BYPOSITION))
              {
                InsertMenuWide(hPopupMenu, nSubMenuIndex, MF_BYPOSITION|MF_POPUP, (UINT_PTR)hSubMenu, wszMenuItem);
              }
            }
            TrackPopupMenu(hPopupMenu, TPM_LEFTBUTTON, ptPos.x, ptPos.y, 0, hMainWnd, NULL);

            //Detach submenus
            while (RemoveMenu(hPopupMenu, 0, MF_BYPOSITION));

            //Destroy popup menu
            DestroyMenu(hPopupMenu);
          }
        }
        else TrackPopupMenu(hMenuMainStack.hPopupMenu, TPM_LEFTBUTTON, ptPos.x, ptPos.y, 0, hMainWnd, NULL);

        if (bMenuMainEnable)
        {
          lpCurrentMenuStack=NULL;
          nCurrentMenuType=TYPE_UNKNOWN;
        }
      }

      //If plugin already loaded, stay in memory, but show as non-active
      if (pd->bInMemory) pd->nUnload=UD_NONUNLOAD_NONACTIVE;
      return;
    }
  }

  if (!pd->bOnStart)
  {
    HMENU hMenu=NULL;
    int nCmd;
    BOOL bResult=TRUE;

    GetCurFile(wszCurrentFile, MAX_PATH);
    if (!hMenuManualStack.hPopupMenu)
      bResult=CreateContextMenu(&hMenuManualStack, wszManualText, TYPE_MANUAL);

    if (bResult)
    {
      if (ptPos.x == ptPos.y && ptPos.x < 0)
        GetEditPos(pd->hWndEdit, &ptPos, ptPos.x);

      if (nSubMenuIndex >= 0)
      {
        hMenu=GetSubMenu(hMenuManualStack.hPopupMenu, nSubMenuIndex);

        if (!hMenu)
        {
          xprintfW(wszBuffer, GetLangStringW(wLangModule, STRID_NOSUBMENUINDEX), nSubMenuIndex);
          MessageBoxW(hMainWnd, wszBuffer, wszPluginTitle, MB_OK|MB_ICONERROR);
        }
      }
      else if (wszSubMenuName[0])
      {
        hMenu=FindRootSubMenuByName(&hMenuManualStack, wszSubMenuName);
      }
      else hMenu=hMenuManualStack.hPopupMenu;

      if (hMenu)
      {
        if (nCmd=ShowContextMenu(&hMenuManualStack, hMainWnd, TYPE_MANUAL, ptPos.x, ptPos.y, hMenu))
          CallContextMenu(&hMenuManualStack, nCmd);
      }
    }
  }

  //Mark "ContextMenu::Main" as active
  {
    PLUGINFUNCTION *pf;
    PLUGINADDW pa;

    if (!(pf=(PLUGINFUNCTION *)SendMessage(hMainWnd, AKD_DLLFINDW, (WPARAM)L"ContextMenu::Main", 0)))
    {
      xmemset(&pa, 0, sizeof(PLUGINADDW));
      pa.pFunction=L"ContextMenu::Main";
      pf=(PLUGINFUNCTION *)SendMessage(hMainWnd, AKD_DLLADDW, 0, (LPARAM)&pa);
    }
    if (pf && !pf->bRunning) pf->bRunning=TRUE;
  }

  //Stay in memory, but show as non-active
  pd->nUnload=UD_NONUNLOAD_NONACTIVE;
}

DWORD WINAPI ThreadProc(LPVOID lpParameter)
{
  MSG msg;

  OleInitialize(NULL);

  hWndMainDlg=CreateDialogWide(hInstanceDLL, MAKEINTRESOURCEW(IDD_CONTEXT), hMainWnd, (DLGPROC)MainDlgProc);

  if (hWndMainDlg)
  {
    while (GetMessageWide(&msg, NULL, 0, 0) > 0)
    {
      if (SendMessage(hMainWnd, AKD_TRANSLATEMESSAGE, TMSG_GLOBAL|TMSG_PLUGIN|TMSG_HOTKEY, (LPARAM)&msg))
        continue;

      if (hWndMainDlg && !IsDialogMessageWide(hWndMainDlg, &msg))
      {
        TranslateMessage(&msg);
        DispatchMessageWide(&msg);
      }
    }
  }
  if (hThread)
  {
    CloseHandle(hThread);
    hThread=NULL;
  }
  return 0;
}

LRESULT CALLBACK NewMainProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  LRESULT lResult;

  if (lResult=IconMenu_Messages(hWnd, uMsg, wParam, lParam))
    return lResult;

  if (uMsg == AKDN_MAIN_ONSTART_PRESHOW)
  {
    if (bMenuMainEnable)
    {
      BOOL bResult=TRUE;

      if (!hMenuMainStack.hPopupMenu)
        bResult=CreateContextMenu(&hMenuMainStack, wszMainText, TYPE_MAIN);

      if (bResult)
      {
        hNewMainMenu=InsertMainMenu(&hMenuMainStack);
        bMenuMainOnStart=TRUE;
      }
    }
    if (bMenuMainHide)
    {
      bMenuMainOnStart=TRUE;
    }
  }
  else if (uMsg == AKDN_SIZE_ONSTART)
  {
    if (bMenuMainOnStart)
    {
      RECT rcBefore;
      RECT rcAfter;
      NSIZE *ns=(NSIZE *)lParam;

      bMenuMainOnStart=FALSE;
      GetClientRect(hMainWnd, &rcBefore);
      if (bMenuMainEnable)
        SetMainMenu(&hMenuMainStack, hNewMainMenu, !bMenuMainHide);
      if (bMenuMainHide)
        ShowMainMenu(FALSE);
      GetClientRect(hMainWnd, &rcAfter);
      ns->rcCurrent.bottom+=rcAfter.bottom - rcBefore.bottom;
    }
  }
  else if (uMsg == AKDN_CONTEXTMENU)
  {
    NCONTEXTMENU *ncm=(NCONTEXTMENU *)lParam;

    if (ncm->bProcess)
    {
      if (ncm->uType == NCM_EDIT)
      {
        if (bMenuUrlEnable)
        {
          AECHARINDEX ciCaret;
          AECHARRANGE aecrUrl;
          CHARRANGE64 recrUrl;

          if (SendMessage(ncm->hWnd, AEM_EXGETOPTIONS, 0, 0) & AECOE_DETECTURL)
          {
            if (bMenuUrlShow)
            {
              ShowUrlMenu(ncm->hWnd, &crUrlMenuShow, ncm->pt.x, ncm->pt.y);
              ncm->bProcess=FALSE;
            }
            else if (SendMessage(ncm->hWnd, AEM_GETINDEX, AEGI_CARETCHAR, (LPARAM)&ciCaret) &&
                     SendMessage(ncm->hWnd, AEM_INDEXINURL, (WPARAM)&ciCaret, (LPARAM)&aecrUrl))
            {
              recrUrl.cpMin=SendMessage(ncm->hWnd, AEM_INDEXTORICHOFFSET, 0, (LPARAM)&aecrUrl.ciMin);
              recrUrl.cpMax=SendMessage(ncm->hWnd, AEM_INDEXTORICHOFFSET, 0, (LPARAM)&aecrUrl.ciMax);
              ShowUrlMenu(ncm->hWnd, &recrUrl, ncm->pt.x, ncm->pt.y);
              ncm->bProcess=FALSE;
            }
          }
        }
        if (bMenuEditEnable && ncm->bProcess)
        {
          int nCmd;
          BOOL bResult=TRUE;

          if (GetCurFile(wszCurrentFile, MAX_PATH))
          {
            if (!hMenuEditStack.hPopupMenu)
              bResult=CreateContextMenu(&hMenuEditStack, wszEditText, TYPE_EDIT);

            if (bResult)
            {
              if (nCmd=ShowContextMenu(&hMenuEditStack, hMainWnd, TYPE_EDIT, ncm->pt.x, ncm->pt.y, NULL))
                CallContextMenu(&hMenuEditStack, nCmd);
            }
            ncm->bProcess=FALSE;
          }
        }
      }
      else if (ncm->uType == NCM_TAB)
      {
        if (bMenuTabEnable)
        {
          int nCmd;
          BOOL bResult=TRUE;

          if (GetCurFile(wszCurrentFile, MAX_PATH))
          {
            if (!hMenuTabStack.hPopupMenu)
              bResult=CreateContextMenu(&hMenuTabStack, wszTabText, TYPE_TAB);

            if (bResult)
            {
              if (nCmd=ShowContextMenu(&hMenuTabStack, hMainWnd, TYPE_TAB, ncm->pt.x, ncm->pt.y, NULL))
                CallContextMenu(&hMenuTabStack, nCmd);
            }
            ncm->bProcess=FALSE;
          }
        }
      }
    }
    else ncm->bProcess=FALSE;

    bMenuUrlShow=FALSE;
  }
  else if (uMsg == WM_MENURBUTTONUP)
  {
    if (bMenuRecentFilesEnable)
    {
      int nMenuItemID=GetMenuItemID((HMENU)lParam, (int)wParam);
      int nMenuItemIndex=nMenuItemID - IDM_RECENT_FILES - 1;

      if (nMenuItemID > IDM_RECENT_FILES && nMenuItemID < IDM_LANGUAGE)
      {
        POINT pt;
        int nCmd;
        RECENTFILE *rf;
        BOOL bResult=TRUE;

        GetCursorPos(&pt);

        if (MenuItemFromPoint(hMainWnd, (HMENU)lParam, pt) != -1)
        {
          if (rf=(RECENTFILE *)SendMessage(hMainWnd, AKD_RECENTFILES, RF_FINDITEMBYINDEX, nMenuItemIndex))
          {
            xstrcpynW(wszCurrentFile, rf->wszFile, MAX_PATH);

            if (!hMenuRecentFilesStack.hPopupMenu)
              bResult=CreateContextMenu(&hMenuRecentFilesStack, wszRecentFilesText, TYPE_RECENTFILES);

            if (bResult)
            {
              if (nCmd=ShowContextMenu(&hMenuRecentFilesStack, hMainWnd, TYPE_RECENTFILES, pt.x, pt.y, NULL))
              {
                CallContextMenu(&hMenuRecentFilesStack, nCmd);
                SendMessage(hMainWnd, WM_CANCELMODE, 0, 0);
              }
            }
          }
        }
      }
    }
  }
  else if (uMsg == WM_ENTERMENULOOP)
  {
    if (bMenuMainEnable)
    {
      if (wParam == FALSE)
      {
        GetCurFile(wszCurrentFile, MAX_PATH);
        lpCurrentMenuStack=&hMenuMainStack;
        nCurrentMenuType=TYPE_MAIN;
        UpdateContextMenu(lpCurrentMenuStack, TYPE_MAIN, NULL);
      }
    }
    if (bMenuMainHide && wParam == FALSE) ShowMainMenu(TRUE);
  }
  else if (uMsg == WM_EXITMENULOOP)
  {
    if (bMenuMainEnable)
    {
      if (wParam == FALSE)
      {
        lpCurrentMenuStack=NULL;
        nCurrentMenuType=TYPE_UNKNOWN;
      }
    }
    if (bMenuMainHide && wParam == FALSE) ShowMainMenu(FALSE);
  }
  else if (uMsg == WM_INITMENUPOPUP)
  {
    if (lpCurrentMenuStack)
    {
      UpdateContextMenu(lpCurrentMenuStack, nCurrentMenuType, (HMENU)wParam);

      if ((HMENU)wParam == lpCurrentMenuStack->hExplorerSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_EXPLORER);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hFavouritesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_FAVOURITES);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hRecentFilesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_RECENTFILES);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hLanguagesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_LANGUAGES);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hOpenCodepagesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_OPENCODEPAGES);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hSaveCodepagesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_SAVECODEPAGES);
      }
    }
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) >= IDM_MIN_EXPLORER && LOWORD(wParam) <= IDM_MAX_MENU)
    {
      CallContextMenu(&hMenuMainStack, LOWORD(wParam));
    }
    else
    {
      if (hMenuMainStack.hMainMenu)
      {
        //TranslateAccelerator
        if (HIWORD(wParam) == 1)
        {
          DWORD dwMenuState;

          if (GetMenuState(hMainMenu, LOWORD(wParam), MF_BYCOMMAND) != (DWORD)-1)
          {
            SendMessage(hMainWnd, WM_INITMENU, (WPARAM)hMainMenu, IMENU_EDIT|IMENU_CHECKS);

            if ((dwMenuState=GetMenuState(hMainMenu, LOWORD(wParam), MF_BYCOMMAND)) != (DWORD)-1)
            {
              if (dwMenuState & (MF_DISABLED | MF_GRAYED))
                return 0;
            }
          }
        }
      }
    }
  }
  else if (uMsg == AKDN_MAIN_ONFINISH)
  {
    NewMainProcData->NextProc(hWnd, uMsg, wParam, lParam);
    if (hWndMainDlg) SendMessage(hWndMainDlg, WM_COMMAND, IDCANCEL, 0);
    UninitMain();
    return FALSE;
  }

  //Edit parent
  if (lResult=EditParentMessages(hWnd, uMsg, wParam, lParam))
    return lResult;

  //Explorer menu
  if (lpCurrentMenuStack)
  {
    if (lpCurrentMenuStack->pExplorerSubMenu3)
    {
      LRESULT lResult=0;

      if (SUCCEEDED(lpCurrentMenuStack->pExplorerSubMenu3->lpVtbl->HandleMenuMsg2(lpCurrentMenuStack->pExplorerSubMenu3, uMsg, wParam, lParam, &lResult)))
        return lResult;
    }
    else if (lpCurrentMenuStack->pExplorerSubMenu2)
    {
      if (SUCCEEDED(lpCurrentMenuStack->pExplorerSubMenu2->lpVtbl->HandleMenuMsg(lpCurrentMenuStack->pExplorerSubMenu2, uMsg, wParam, lParam)))
        return 0;
    }
  }

  //Call next procedure
  return NewMainProcData->NextProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK NewFrameProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  LRESULT lResult;

  if (lResult=IconMenu_Messages(hWnd, uMsg, wParam, lParam))
    return lResult;

  if (lResult=EditParentMessages(hWnd, uMsg, wParam, lParam))
    return lResult;

  //Call next procedure
  return NewFrameProcData->NextProc(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK EditParentMessages(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  if (uMsg == WM_NOTIFY)
  {
    if (wParam == ID_EDIT)
    {
      if (((NMHDR *)lParam)->code == EN_LINK)
      {
        ENLINK64 *enl=(ENLINK64 *)lParam;

        if (enl->msg == WM_RBUTTONUP)
        {
          if (bMenuUrlEnable)
          {
            crUrlMenuShow.cpMin=enl->chrg64.cpMin;
            crUrlMenuShow.cpMax=enl->chrg64.cpMax;
            bMenuUrlShow=TRUE;
          }
        }
      }
    }
  }
  return FALSE;
}

LRESULT CALLBACK MainDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndType;
  static HWND hWndEnable;
  static HWND hWndHide;
  static HWND hWndText;
  static HWND hWndShow;
  static HWND hWndOK;
  static HWND hWndCancel;
  static int nMenuType;
  static CONTROLTYPE ct[TYPE_MAX]={0};
  static DIALOGRESIZE drs[]={{&hWndType,   DRS_SIZE|DRS_X, 0},
                             {&hWndEnable, DRS_MOVE|DRS_X, 0},
                             {&hWndHide,   DRS_MOVE|DRS_X, 0},
                             {&hWndText,   DRS_SIZE|DRS_X, 0},
                             {&hWndText,   DRS_SIZE|DRS_Y, 0},
                             {&hWndShow,   DRS_MOVE|DRS_Y, 0},
                             {&hWndOK,     DRS_MOVE|DRS_X, 0},
                             {&hWndOK,     DRS_MOVE|DRS_Y, 0},
                             {&hWndCancel, DRS_MOVE|DRS_X, 0},
                             {&hWndCancel, DRS_MOVE|DRS_Y, 0},
                             {0, 0, 0}};
  LRESULT lResult;

  if (lResult=IconMenu_Messages(hDlg, uMsg, wParam, lParam))
  {
    SetWindowLongPtrWide(hDlg, DWLP_MSGRESULT, lResult);
    return TRUE;
  }

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndType=GetDlgItem(hDlg, IDC_MENUTYPE);
    hWndEnable=GetDlgItem(hDlg, IDC_MENUENABLE);
    hWndHide=GetDlgItem(hDlg, IDC_MENUHIDE);
    hWndText=GetDlgItem(hDlg, IDC_MENUTEXT);
    hWndShow=GetDlgItem(hDlg, IDC_MENUSHOW);
    hWndOK=GetDlgItem(hDlg, IDOK);
    hWndCancel=GetDlgItem(hDlg, IDCANCEL);

    SetWindowTextWide(hDlg, wszPluginTitle);
    SetDlgItemTextWide(hDlg, IDC_MENUENABLE, GetLangStringW(wLangModule, STRID_ENABLE));
    SetDlgItemTextWide(hDlg, IDC_MENUHIDE, GetLangStringW(wLangModule, STRID_HIDE));
    SetDlgItemTextWide(hDlg, IDC_MENUSHOW, GetLangStringW(wLangModule, STRID_SHOW));
    SetDlgItemTextWide(hDlg, IDOK, GetLangStringW(wLangModule, STRID_OK));
    SetDlgItemTextWide(hDlg, IDCANCEL, GetLangStringW(wLangModule, STRID_CANCEL));

    ComboBox_AddStringWide(hWndType, GetLangStringW(wLangModule, STRID_MANUAL));
    ComboBox_AddStringWide(hWndType, GetLangStringW(wLangModule, STRID_MENUMAIN));
    ComboBox_AddStringWide(hWndType, GetLangStringW(wLangModule, STRID_MENUEDIT));
    ComboBox_AddStringWide(hWndType, GetLangStringW(wLangModule, STRID_MENUTAB));
    ComboBox_AddStringWide(hWndType, GetLangStringW(wLangModule, STRID_MENUURL));
    ComboBox_AddStringWide(hWndType, GetLangStringW(wLangModule, STRID_MENURECENTFILES));

    if (ct[TYPE_MANUAL].wpText=(wchar_t *)HeapAlloc(hHeap, 0, (xstrlenW(wszManualText) + 1) * sizeof(wchar_t)))
    {
      xstrcpyW(ct[TYPE_MANUAL].wpText, wszManualText);
      ct[TYPE_MANUAL].bEnable=FALSE;
    }
    if (ct[TYPE_MAIN].wpText=(wchar_t *)HeapAlloc(hHeap, 0, (xstrlenW(wszMainText) + 1) * sizeof(wchar_t)))
    {
      xstrcpyW(ct[TYPE_MAIN].wpText, wszMainText);
      ct[TYPE_MAIN].bEnable=bMenuMainEnable;
    }
    if (ct[TYPE_EDIT].wpText=(wchar_t *)HeapAlloc(hHeap, 0, (xstrlenW(wszEditText) + 1) * sizeof(wchar_t)))
    {
      xstrcpyW(ct[TYPE_EDIT].wpText, wszEditText);
      ct[TYPE_EDIT].bEnable=bMenuEditEnable;
    }
    if (ct[TYPE_TAB].wpText=(wchar_t *)HeapAlloc(hHeap, 0, (xstrlenW(wszTabText) + 1) * sizeof(wchar_t)))
    {
      xstrcpyW(ct[TYPE_TAB].wpText, wszTabText);
      ct[TYPE_TAB].bEnable=bMenuTabEnable;
    }
    if (ct[TYPE_URL].wpText=(wchar_t *)HeapAlloc(hHeap, 0, (xstrlenW(wszUrlText) + 1) * sizeof(wchar_t)))
    {
      xstrcpyW(ct[TYPE_URL].wpText, wszUrlText);
      ct[TYPE_URL].bEnable=bMenuUrlEnable;
    }
    if (ct[TYPE_RECENTFILES].wpText=(wchar_t *)HeapAlloc(hHeap, 0, (xstrlenW(wszRecentFilesText) + 1) * sizeof(wchar_t)))
    {
      xstrcpyW(ct[TYPE_RECENTFILES].wpText, wszRecentFilesText);
      ct[TYPE_RECENTFILES].bEnable=bMenuRecentFilesEnable;
    }
    if (bMenuMainHide)
      SendMessage(hWndHide, BM_SETCHECK, BST_CHECKED, 0);

    OldEditDlgProc=(WNDPROC)GetWindowLongPtrWide(hWndText, GWLP_WNDPROC);
    SetWindowLongPtrWide(hWndText, GWLP_WNDPROC, (UINT_PTR)NewEditDlgProc);

    nMenuType=TYPE_UNKNOWN;
    SendMessage(hWndType, CB_SETCURSEL, nExtMenuIndex, 0);
    PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MENUTYPE, CBN_SELCHANGE), (LPARAM)hWndType);
  }
  else if (uMsg == WM_INITMENUPOPUP)
  {
    if (lpCurrentMenuStack)
    {
      UpdateContextMenu(lpCurrentMenuStack, nCurrentMenuType, (HMENU)wParam);

      if ((HMENU)wParam == lpCurrentMenuStack->hExplorerSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_EXPLORER);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hFavouritesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_FAVOURITES);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hRecentFilesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_RECENTFILES);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hLanguagesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_LANGUAGES);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hOpenCodepagesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_OPENCODEPAGES);
      }
      else if ((HMENU)wParam == lpCurrentMenuStack->hSaveCodepagesSubMenu)
      {
        InitMenuPopup(lpCurrentMenuStack, INIT_SAVECODEPAGES);
      }
    }
  }
  else if (uMsg == AKDLL_MENUINDEX)
  {
    int nCurSel=(int)SendMessage(hWndType, CB_GETCURSEL, 0, 0);

    if (nCurSel != (int)wParam)
      SendMessage(hWndType, CB_SETCURSEL, wParam, 0);
    PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MENUTYPE, CBN_SELCHANGE), (LPARAM)hWndType);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_MENUTYPE && HIWORD(wParam) == CBN_SELCHANGE)
    {
      int nItem=(int)SendMessage(hWndType, CB_GETCURSEL, 0, 0);

      if (nItem != nMenuType)
      {
        if (SendMessage(hWndText, EM_GETMODIFY, 0, 0))
        {
          if (ct[nMenuType].wpText)
          {
            HeapFree(hHeap, 0, ct[nMenuType].wpText);
            ct[nMenuType].wpText=NULL;
          }
          GetEditText(hWndText, &ct[nMenuType].wpText);
        }
        nMenuType=nItem;

        if (nMenuType == TYPE_MAIN)
          ShowWindow(hWndHide, TRUE);
        else
          ShowWindow(hWndHide, FALSE);

        if (nMenuType == TYPE_MANUAL)
        {
          ShowWindow(hWndEnable, FALSE);
          EnableWindow(hWndText, TRUE);
        }
        else
        {
          ShowWindow(hWndEnable, TRUE);
          SendMessage(hWndEnable, BM_SETCHECK, ct[nMenuType].bEnable?BST_CHECKED:BST_UNCHECKED, 0);
          EnableWindow(hWndText, ct[nMenuType].bEnable);
        }
        SetWindowTextWide(hWndText, ct[nMenuType].wpText);
      }

      if (*wszExtMenuSearch)
      {
        FINDTEXTEX64W fte;

        fte.chrg.cpMin=0;
        fte.chrg.cpMax=-1;
        fte.lpstrText=wszExtMenuSearch;

        if (SendMessage(hWndText, EM_FINDTEXTEX64W, FR_DOWN|FR_MATCHCASE, (LPARAM)&fte) != -1)
        {
          SendMessage(hWndText, AEM_LOCKSCROLL, SB_BOTH, TRUE);
          SendMessage(hWndText, EM_SETSEL, fte.chrgText.cpMax, fte.chrgText.cpMin);
          SendMessage(hWndText, AEM_LOCKSCROLL, SB_BOTH, FALSE);
          ScrollCaret(hWndText);
        }
        wszExtMenuSearch[0]=L'\0';
      }
      if (crExtSetSel.cpMin || crExtSetSel.cpMax)
      {
        if (crExtSetSel.cpMax == -1)
        {
          int nLine;

          nLine=(int)SendMessage(hWndText, EM_EXLINEFROMCHAR, 0, crExtSetSel.cpMin);
          crExtSetSel.cpMax=SendMessage(hWndText, EM_LINEINDEX, nLine, 0) + SendMessage(hWndText, EM_LINELENGTH, crExtSetSel.cpMin, 0);
        }

        SendMessage(hWndText, AEM_LOCKSCROLL, SB_BOTH, TRUE);
        SendMessage(hWndText, EM_SETSEL, crExtSetSel.cpMax, crExtSetSel.cpMin);
        SendMessage(hWndText, AEM_LOCKSCROLL, SB_BOTH, FALSE);
        ScrollCaret(hWndText);

        crExtSetSel.cpMin=0;
        crExtSetSel.cpMax=0;
      }
      if (bExtFocusEdit)
      {
        bExtFocusEdit=FALSE;
        SetFocus(hWndText);
      }
      nExtMenuIndex=nItem;
    }
    else if (LOWORD(wParam) == IDC_MENUENABLE)
    {
      if (SendMessage(hWndEnable, BM_GETCHECK, 0, 0) == BST_CHECKED)
        ct[nMenuType].bEnable=TRUE;
      else
        ct[nMenuType].bEnable=FALSE;

      EnableWindow(hWndText, ct[nMenuType].bEnable);
    }
    else if (LOWORD(wParam) == IDC_MENUSHOW)
    {
      wchar_t *wpText;

      if (GetEditText(hWndText, &wpText))
      {
        POINT pt;
        int nCmd;

        GetCurFile(wszCurrentFile, MAX_PATH);
        if (CreateContextMenu(&hMenuDialogStack, wpText, nMenuType))
        {
          GetCursorPos(&pt);
          if (nCmd=ShowContextMenu(&hMenuDialogStack, hDlg, TYPE_MANUAL, pt.x, pt.y, NULL))
            CallContextMenu(&hMenuDialogStack, nCmd);
          FreeContextMenu(&hMenuDialogStack);
        }
        HeapFree(hHeap, 0, wpText);
      }
    }
    else if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
    {
      if (LOWORD(wParam) == IDOK)
      {
        int i;

        if (SendMessage(hWndText, EM_GETMODIFY, 0, 0))
        {
          if (ct[nMenuType].wpText)
          {
            HeapFree(hHeap, 0, ct[nMenuType].wpText);
            ct[nMenuType].wpText=NULL;
          }
          GetEditText(hWndText, &ct[nMenuType].wpText);
        }

        //Test for errors
        for (i=0; i < TYPE_MAX; ++i)
        {
          if (!CreateContextMenu(&hMenuDialogStack, ct[i].wpText, i))
            return 0;
          FreeContextMenu(&hMenuDialogStack);
        }

        if (wszManualText) HeapFree(hHeap, 0, wszManualText);
        wszManualText=ct[TYPE_MANUAL].wpText;
        if (!*wszManualText)
        {
          HeapFree(hHeap, 0, wszManualText);
          wszManualText=GetDefaultMenu(STRID_DEFAULTMANUAL);
        }

        bMenuMainEnable=ct[TYPE_MAIN].bEnable;
        if (wszMainText) HeapFree(hHeap, 0, wszMainText);
        wszMainText=ct[TYPE_MAIN].wpText;
        if (!*wszMainText)
        {
          HeapFree(hHeap, 0, wszMainText);
          wszMainText=GetDefaultMenu(STRID_DEFAULTMAIN);
        }

        bMenuEditEnable=ct[TYPE_EDIT].bEnable;
        if (wszEditText) HeapFree(hHeap, 0, wszEditText);
        wszEditText=ct[TYPE_EDIT].wpText;
        if (!*wszEditText)
        {
          HeapFree(hHeap, 0, wszEditText);
          wszEditText=GetDefaultMenu(STRID_DEFAULTEDIT);
        }

        bMenuTabEnable=ct[TYPE_TAB].bEnable;
        if (wszTabText) HeapFree(hHeap, 0, wszTabText);
        wszTabText=ct[TYPE_TAB].wpText;
        if (!*wszTabText)
        {
          HeapFree(hHeap, 0, wszTabText);
          wszTabText=GetDefaultMenu(STRID_DEFAULTTAB);
        }

        bMenuUrlEnable=ct[TYPE_URL].bEnable;
        if (wszUrlText) HeapFree(hHeap, 0, wszUrlText);
        wszUrlText=ct[TYPE_URL].wpText;
        if (!*wszUrlText)
        {
          HeapFree(hHeap, 0, wszUrlText);
          wszUrlText=GetDefaultMenu(STRID_DEFAULTURL);
        }

        bMenuRecentFilesEnable=ct[TYPE_RECENTFILES].bEnable;
        if (wszRecentFilesText) HeapFree(hHeap, 0, wszRecentFilesText);
        wszRecentFilesText=ct[TYPE_RECENTFILES].wpText;
        if (!*wszRecentFilesText)
        {
          HeapFree(hHeap, 0, wszRecentFilesText);
          wszRecentFilesText=GetDefaultMenu(STRID_DEFAULTRECENTFILES);
        }

        UnsetMainMenu(&hMenuMainStack);
        FreeContextMenu(&hMenuMainStack);
        FreeContextMenu(&hMenuEditStack);
        FreeContextMenu(&hMenuTabStack);
        FreeContextMenu(&hMenuUrlStack);
        FreeContextMenu(&hMenuRecentFilesStack);
        FreeContextMenu(&hMenuManualStack);

        if (bMenuMainEnable)
        {
          CreateContextMenu(&hMenuMainStack, wszMainText, TYPE_MAIN);
          hNewMainMenu=InsertMainMenu(&hMenuMainStack);
          SetMainMenu(&hMenuMainStack, hNewMainMenu, TRUE);
        }

        //Main menu visibility
        i=(BOOL)SendMessage(hWndHide, BM_GETCHECK, 0, 0);
        if (bMenuMainHide != i)
        {
          bMenuMainHide=i;
          ShowMainMenu(!bMenuMainHide);
        }

        dwSaveFlags|=OF_MENUSTEXT;
      }
      else if (LOWORD(wParam) == IDCANCEL)
      {
        int i;

        for (i=0; i < TYPE_MAX; ++i)
        {
          if (ct[i].wpText) HeapFree(hHeap, 0, ct[i].wpText);
        }
      }

      if (dwSaveFlags)
      {
        SaveOptions(dwSaveFlags);
        dwSaveFlags=0;
      }

      DestroyWindow(hWndMainDlg);
      hWndMainDlg=NULL;
      CloseHandle(hThread);
      hThread=NULL;
      ExitThread(0);
    }
  }
  else if (uMsg == WM_CLOSE)
  {
    PostMessage(hDlg, WM_COMMAND, IDCANCEL, 0);
    return TRUE;
  }

  //Explorer menu
  if (lpCurrentMenuStack)
  {
    if (lpCurrentMenuStack->pExplorerSubMenu3)
    {
      LRESULT lResult=0;

      if (SUCCEEDED(lpCurrentMenuStack->pExplorerSubMenu3->lpVtbl->HandleMenuMsg2(lpCurrentMenuStack->pExplorerSubMenu3, uMsg, wParam, lParam, &lResult)))
        return lResult;
    }
    else if (lpCurrentMenuStack->pExplorerSubMenu2)
    {
      if (SUCCEEDED(lpCurrentMenuStack->pExplorerSubMenu2->lpVtbl->HandleMenuMsg(lpCurrentMenuStack->pExplorerSubMenu2, uMsg, wParam, lParam)))
        return 0;
    }
  }

  //Dialog resize messages
  {
    DIALOGRESIZEMSG drsm={&drs[0], &rcMainMinMaxDialog, &rcMainCurrentDialog, DRM_PAINTSIZEGRIP, hDlg, uMsg, wParam, lParam};

    if (SendMessage(hMainWnd, AKD_DIALOGRESIZE, 0, (LPARAM)&drsm))
      dwSaveFlags|=OF_MAINRECT;
  }

  return 0;
}

LRESULT CALLBACK NewEditDlgProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  LRESULT lResult;

  if (uMsg == WM_GETDLGCODE)
  {
    if (bOldWindows)
      lResult=CallWindowProcA(OldEditDlgProc, hWnd, uMsg, wParam, lParam);
    else
      lResult=CallWindowProcW(OldEditDlgProc, hWnd, uMsg, wParam, lParam);

    if (lResult & DLGC_HASSETSEL)
    {
      lResult&=~DLGC_HASSETSEL;
    }
    return lResult;
  }
  else if (uMsg == WM_CONTEXTMENU)
  {
    ShowStandardEditMenu((HWND)wParam, hPopupEdit, lParam != -1);
  }

  if (bOldWindows)
    return CallWindowProcA(OldEditDlgProc, hWnd, uMsg, wParam, lParam);
  else
    return CallWindowProcW(OldEditDlgProc, hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK FavListDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndItemsList;
  static HWND hWndStats;
  static HWND hWndSearch;
  static HWND hWndShowFile;
  static HWND hWndClose;
  static HICONMENU hIconMenuList;
  static HMENU hMenuList;
  static BOOL bListChanged;
  static BOOL bFavShowFileDlg;
  static DIALOGRESIZE drs[]={{&hWndItemsList,  DRS_SIZE|DRS_X, 0},
                             {&hWndItemsList,  DRS_SIZE|DRS_Y, 0},
                             {&hFavToolbar,    DRS_SIZE|DRS_X, 0},
                             {&hFavToolbar,    DRS_MOVE|DRS_Y, 0},
                             {&hWndStats,      DRS_MOVE|DRS_X, 0},
                             {&hWndStats,      DRS_MOVE|DRS_Y, 0},
                             {&hWndSearch,     DRS_SIZE|DRS_X, 0},
                             {&hWndSearch,     DRS_MOVE|DRS_Y, 0},
                             {&hWndShowFile,   DRS_MOVE|DRS_Y, 0},
                             {&hWndClose,      DRS_MOVE|DRS_X, 0},
                             {&hWndClose,      DRS_MOVE|DRS_Y, 0},
                             {0, 0, 0}};
  int nItem;
  LRESULT lResult;

  if (lResult=IconMenu_Messages(hDlg, uMsg, wParam, lParam))
  {
    SetWindowLongPtrWide(hDlg, DWLP_MSGRESULT, lResult);
    return TRUE;
  }

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hIconFav);
    hWndFavDlg=hDlg;
    hWndItemsList=GetDlgItem(hDlg, IDC_ITEM_LIST);
    hWndStats=GetDlgItem(hDlg, IDC_ITEM_STATS);
    hWndSearch=GetDlgItem(hDlg, IDC_SEARCH);
    hWndShowFile=GetDlgItem(hDlg, IDC_SHOWFILE);
    hWndClose=GetDlgItem(hDlg, IDCANCEL);

    SetWindowTextWide(hDlg, GetLangStringW(wLangModule, STRID_FAVOURITES));
    SetDlgItemTextWide(hDlg, IDC_SHOWFILE, GetLangStringW(wLangModule, STRID_SHOWFILE));
    SetDlgItemTextWide(hDlg, IDCANCEL, GetLangStringW(wLangModule, STRID_CLOSE));
    SendMessage(hWndSearch, EM_LIMITTEXT, MAX_PATH, 0);
    if (bFavShowFile) SendMessage(hWndShowFile, BM_SETCHECK, BST_CHECKED, 0);
    bFavShowFileDlg=bFavShowFile;

    //Create toolbar
    {
      TBBUTTON tbb;
      RECT rcList;
      RECT rcStats;
      HICON hIcon;
      int nIcon;
      int nCommand;

      hFavImageList=ImageList_Create(16, 16, ILC_COLOR32|ILC_MASK, 0, 0);
      ImageList_SetBkColor(hFavImageList, GetSysColor(COLOR_MENU));

      for (nIcon=IDI_ICON_OPEN; nIcon <= IDI_ICON_EDIT; ++nIcon)
      {
        hIcon=(HICON)LoadImageA(hInstanceDLL, MAKEINTRESOURCEA(nIcon), IMAGE_ICON, 16, 16, 0);
        ImageList_AddIcon(hFavImageList, hIcon);
        DestroyIcon(hIcon);
      }

      GetWindowPos(hWndItemsList, hDlg, &rcList);
      GetWindowPos(hWndStats, hDlg, &rcStats);
      hFavToolbar=CreateWindowExWide(0,
                            L"ToolbarWindow32",
                            NULL,
                            WS_CHILD|WS_VISIBLE|TBSTYLE_TOOLTIPS|TBSTYLE_FLAT|CCS_NORESIZE|CCS_NOPARENTALIGN|CCS_NODIVIDER,
                            rcList.left, rcList.top + rcList.bottom + 1, rcStats.left - rcList.left, 28,
                            hDlg,
                            (HMENU)(UINT_PTR)IDC_TOOLBAR,
                            hInstanceDLL,
                            NULL);

      SendMessage(hFavToolbar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
      SendMessage(hFavToolbar, TB_SETBITMAPSIZE, 0, MAKELONG(16, 16));
      SendMessage(hFavToolbar, TB_SETBUTTONSIZE, 0, MAKELONG(24, 22));
      SendMessage(hFavToolbar, TB_SETIMAGELIST, 0, (LPARAM)hFavImageList);

      for (nCommand=IDC_ITEM_OPEN; nCommand <= IDC_ITEM_EDIT; ++nCommand)
      {
        tbb.iBitmap=nCommand - IDC_ITEM_OPEN;
        tbb.idCommand=nCommand;
        tbb.fsState=TBSTATE_ENABLED;
        tbb.fsStyle=TBSTYLE_BUTTON;
        tbb.dwData=0;
        tbb.iString=0;
        SendMessage(hFavToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);

        if (nCommand == IDC_ITEM_OPEN ||
            nCommand == IDC_ITEM_SORT ||
            nCommand == IDC_ITEM_DELETEOLD)
        {
          tbb.iBitmap=-1;
          tbb.idCommand=0;
          tbb.fsState=0;
          tbb.fsStyle=TBSTYLE_SEP;
          tbb.dwData=0;
          tbb.iString=0;
          SendMessage(hFavToolbar, TB_ADDBUTTONS, 1, (LPARAM)&tbb);
        }
      }
    }

    //Popup menu
    if (hMenuList=CreatePopupMenu())
    {
      if (hIconMenuList=IconMenu_Alloc(hDlg))
      {
        IconMenu_AddItemW(hIconMenuList, hFavImageList, 0, 16, 16, hMenuList, (UINT)-1, MF_BYPOSITION|MF_STRING, IDC_ITEM_OPEN, GetLangStringW(wLangModule, STRID_MENU_OPEN));
        IconMenu_AddItemW(hIconMenuList, hFavImageList, -1, 0, 0, hMenuList, (UINT)-1, MF_BYPOSITION|MF_SEPARATOR, (UINT)-1, NULL);
        IconMenu_AddItemW(hIconMenuList, hFavImageList, 1, 16, 16, hMenuList, (UINT)-1, MF_BYPOSITION|MF_STRING, IDC_ITEM_MOVEUP, GetLangStringW(wLangModule, STRID_MENU_MOVEUP));
        IconMenu_AddItemW(hIconMenuList, hFavImageList, 2, 16, 16, hMenuList, (UINT)-1, MF_BYPOSITION|MF_STRING, IDC_ITEM_MOVEDOWN, GetLangStringW(wLangModule, STRID_MENU_MOVEDOWN));
        IconMenu_AddItemW(hIconMenuList, hFavImageList, 3, 16, 16, hMenuList, (UINT)-1, MF_BYPOSITION|MF_STRING, IDC_ITEM_SORT, GetLangStringW(wLangModule, STRID_MENU_SORT));
        IconMenu_AddItemW(hIconMenuList, hFavImageList, -1, 0, 0, hMenuList, (UINT)-1, MF_BYPOSITION|MF_SEPARATOR, (UINT)-1, NULL);
        IconMenu_AddItemW(hIconMenuList, hFavImageList, 4, 16, 16, hMenuList, (UINT)-1, MF_BYPOSITION|MF_STRING, IDC_ITEM_DELETE, GetLangStringW(wLangModule, STRID_MENU_DELETE));
        IconMenu_AddItemW(hIconMenuList, hFavImageList, 5, 16, 16, hMenuList, (UINT)-1, MF_BYPOSITION|MF_STRING, IDC_ITEM_DELETEOLD, GetLangStringW(wLangModule, STRID_MENU_DELETEOLD));
        IconMenu_AddItemW(hIconMenuList, hFavImageList, -1, 0, 0, hMenuList, (UINT)-1, MF_BYPOSITION|MF_SEPARATOR, (UINT)-1, NULL);
        IconMenu_AddItemW(hIconMenuList, hFavImageList, 6, 16, 16, hMenuList, (UINT)-1, MF_BYPOSITION|MF_STRING, IDC_ITEM_EDIT, GetLangStringW(wLangModule, STRID_MENU_EDIT));
      }
    }
    bListChanged=FALSE;
    FillFavouritesListBox(&hFavStack, hWndItemsList, bFavShowFileDlg);
    if (StackGetFavouriteByFile(&hFavStack, wszCurrentFile, &nItem))
      SendMessage(hWndItemsList, LB_SETSEL, TRUE, nItem);
    PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_ITEM_LIST, LBN_SELCHANGE), 0);

    //SubClass listbox
    OldListBoxProc=(WNDPROC)GetWindowLongPtrWide(hWndItemsList, GWLP_WNDPROC);
    SetWindowLongPtrWide(hWndItemsList, GWLP_WNDPROC, (UINT_PTR)NewListBoxProc);
  }
  else if (uMsg == WM_CONTEXTMENU)
  {
    if ((HWND)wParam == hWndItemsList)
    {
      POINT pt;
      int nCount;
      int nSelCount;

      if (lParam == -1)
      {
        pt.x=0;
        pt.y=0;
        ClientToScreen(hWndItemsList, &pt);
      }
      else
      {
        GetCursorPos(&pt);
      }

      nCount=(int)SendMessage(hWndItemsList, LB_GETCOUNT, 0, 0);
      nSelCount=(int)SendMessage(hWndItemsList, LB_GETSELCOUNT, 0, 0);
      EnableMenuItem(hMenuList, IDC_ITEM_OPEN, nSelCount?MF_ENABLED:MF_GRAYED);
      EnableMenuItem(hMenuList, IDC_ITEM_MOVEUP, nSelCount?MF_ENABLED:MF_GRAYED);
      EnableMenuItem(hMenuList, IDC_ITEM_MOVEDOWN, nSelCount?MF_ENABLED:MF_GRAYED);
      EnableMenuItem(hMenuList, IDC_ITEM_SORT, nCount?MF_ENABLED:MF_GRAYED);
      EnableMenuItem(hMenuList, IDC_ITEM_DELETE, nSelCount?MF_ENABLED:MF_GRAYED);
      EnableMenuItem(hMenuList, IDC_ITEM_EDIT, (nSelCount == 1)?MF_ENABLED:MF_GRAYED);
      TrackPopupMenu(hMenuList, TPM_LEFTBUTTON|TPM_RIGHTBUTTON, pt.x, pt.y, 0, hDlg, NULL);
    }
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (((NMHDR *)lParam)->code == TTN_GETDISPINFOA)
    {
      int nStrID=((int)((NMHDR *)lParam)->idFrom - IDC_ITEM_OPEN) + STRID_MENU_OPEN;

      WideCharToMultiByte(CP_ACP, 0, GetLangStringW(wLangModule, nStrID), -1, ((NMTTDISPINFOA *)lParam)->szText, 80, NULL, NULL);
    }
    else if (((NMHDR *)lParam)->code == TTN_GETDISPINFOW)
    {
      int nStrID=((int)((NMHDR *)lParam)->idFrom - IDC_ITEM_OPEN) + STRID_MENU_OPEN;

      xstrcpynW(((NMTTDISPINFOW *)lParam)->szText, GetLangStringW(wLangModule, nStrID), 80);
    }
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_SHOWFILE)
    {
      int *lpSelItems=NULL;
      int nSelCount;
      int i;

      bFavShowFileDlg=(BOOL)SendMessage(hWndShowFile, BM_GETCHECK, 0, 0);

      //Fill listbox
      nSelCount=GetListBoxSelItems(hWndItemsList, &lpSelItems);
      FillFavouritesListBox(&hFavStack, hWndItemsList, bFavShowFileDlg);
      for (i=0; i < nSelCount; ++i)
        SendMessage(hWndItemsList, LB_SETSEL, TRUE, lpSelItems[i]);
      if (lpSelItems) FreeListBoxSelItems(&lpSelItems);
    }
    else if (LOWORD(wParam) == IDC_SEARCH)
    {
      if (HIWORD(wParam) == EN_CHANGE)
      {
        wchar_t wszSearch[MAX_PATH];

        SendMessage(hWndItemsList, LB_SETSEL, FALSE, -1);

        if (GetWindowTextWide(hWndSearch, wszSearch, MAX_PATH))
        {
          for (nItem=0; ; ++nItem)
          {
            if (ListBox_GetTextWide(hWndItemsList, nItem, wszBuffer) == LB_ERR)
              break;
            if (xstrstrW(wszBuffer, -1, wszSearch, -1, FALSE, NULL, NULL))
              SendMessage(hWndItemsList, LB_SETSEL, TRUE, nItem);
          }
        }
        PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_ITEM_LIST, LBN_SELCHANGE), 0);
      }
    }
    else if (LOWORD(wParam) == IDC_ITEM_MOVEUP)
    {
      if (ShiftListBoxSelItems(&hFavStack, hWndItemsList, FALSE))
        bListChanged=TRUE;
    }
    else if (LOWORD(wParam) == IDC_ITEM_MOVEDOWN)
    {
      if (ShiftListBoxSelItems(&hFavStack, hWndItemsList, TRUE))
        bListChanged=TRUE;
    }
    else if (LOWORD(wParam) == IDC_ITEM_SORT)
    {
      StackSortFavourites(&hFavStack, 1, bFavShowFileDlg);
      FillFavouritesListBox(&hFavStack, hWndItemsList, bFavShowFileDlg);
      bListChanged=TRUE;
      PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_ITEM_LIST, LBN_SELCHANGE), 0);
    }
    else if (LOWORD(wParam) == IDC_ITEM_DELETE)
    {
      if (DeleteListBoxSelItems(&hFavStack, hWndItemsList))
        bListChanged=TRUE;
      PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_ITEM_LIST, LBN_SELCHANGE), 0);
    }
    else if (LOWORD(wParam) == IDC_ITEM_DELETEOLD)
    {
      if (DeleteListBoxOldItems(&hFavStack, hWndItemsList))
        bListChanged=TRUE;
      PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_ITEM_LIST, LBN_SELCHANGE), 0);
    }
    else if (LOWORD(wParam) == IDC_ITEM_EDIT)
    {
      if (SendMessage(hWndItemsList, LB_GETSELCOUNT, 0, 0) == 1)
      {
        FAVITEM *lpFavItem;
        INT_PTR nResult;

        if (SendMessage(hWndItemsList, LB_GETSELITEMS, 1, (LPARAM)&nItem) != LB_ERR)
        {
          if (lpFavItem=StackGetFavouriteByIndex(&hFavStack, nItem + 1))
          {
            nResult=DialogBoxParamWide(hInstanceDLL, MAKEINTRESOURCEW(IDD_FAVEDIT), hDlg, (DLGPROC)FavEditDlgProc, (LPARAM)lpFavItem);

            if (nResult)
            {
              SendMessage(hWndItemsList, LB_DELETESTRING, nItem, 0);
              ListBox_InsertStringWide(hWndItemsList, nItem, bFavShowFileDlg?lpFavItem->wszFile:lpFavItem->wszName);
              SendMessage(hWndItemsList, LB_SETSEL, TRUE, nItem);
              bListChanged=TRUE;
            }
          }
        }
      }
    }
    else if (LOWORD(wParam) == IDC_ITEM_LIST)
    {
      if (HIWORD(wParam) == LBN_SELCHANGE)
      {
        int nCount;
        int nSelCount;

        nCount=(int)SendMessage(hWndItemsList, LB_GETCOUNT, 0, 0);
        nSelCount=(int)SendMessage(hWndItemsList, LB_GETSELCOUNT, 0, 0);

        xprintfW(wszBuffer, L"%d / %d", nSelCount, nCount);
        SetWindowTextWide(hWndStats, wszBuffer);
      }
      else if (HIWORD(wParam) == LBN_DBLCLK)
      {
        PostMessage(hDlg, WM_COMMAND, IDC_ITEM_OPEN, 0);
      }
    }
    else if (LOWORD(wParam) == IDC_ITEM_OPEN ||
             LOWORD(wParam) == IDCANCEL)
    {
      if (LOWORD(wParam) == IDC_ITEM_OPEN)
      {
        FAVITEM *lpFavItem;
        OPENDOCUMENTW od;
        int *lpSelItems;
        int nSelCount;
        int i;

        if (nSelCount=GetListBoxSelItems(hWndItemsList, &lpSelItems))
        {
          //Close dialog
          EndDialog(hDlg, 0);
          hWndFavDlg=NULL;

          for (i=0; i < nSelCount; ++i)
          {
            if (lpFavItem=StackGetFavouriteByIndex(&hFavStack, lpSelItems[i] + 1))
            {
              if (TranslateFileString(lpFavItem->wszFile, wszBuffer, BUFFER_SIZE))
              {
                od.pFile=wszBuffer;
                od.pWorkDir=NULL;
                od.dwFlags=OD_ADT_BINARY_ERROR|OD_ADT_REG_CODEPAGE;
                od.nCodePage=0;
                od.bBOM=0;
                SendMessage(hMainWnd, AKD_OPENDOCUMENTW, (WPARAM)NULL, (LPARAM)&od);
              }
            }
            else break;

            if (nMDI == WMD_SDI) break;
          }
          FreeListBoxSelItems(&lpSelItems);
        }
      }
      if (bListChanged || bFavShowFileDlg != bFavShowFile)
      {
        bFavShowFile=bFavShowFileDlg;
        dwSaveFlags|=OF_FAVTEXT;

        if (dwSaveFlags)
        {
          SaveOptions(dwSaveFlags);
          dwSaveFlags=0;
        }
      }
      ImageList_Destroy(hFavImageList);
      DestroyWindow(hFavToolbar);
      DestroyMenu(hMenuList);
      IconMenu_Free(hIconMenuList, NULL);

      if (hWndFavDlg)
      {
        EndDialog(hDlg, 0);
        hWndFavDlg=NULL;
      }
      return TRUE;
    }
  }

  //Dialog resize messages
  {
    DIALOGRESIZEMSG drsm={&drs[0], &rcFavMinMaxDialog, &rcFavCurrentDialog, DRM_PAINTSIZEGRIP, hDlg, uMsg, wParam, lParam};

    if (SendMessage(hMainWnd, AKD_DIALOGRESIZE, 0, (LPARAM)&drsm))
      dwSaveFlags|=OF_FAVRECT;
  }
  return FALSE;
}

LRESULT CALLBACK NewListBoxProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  if (uMsg == WM_GETDLGCODE)
  {
    MSG *msg=(MSG *)lParam;

    if (msg)
    {
      if (msg->message == WM_KEYDOWN)
      {
        if (msg->wParam == VK_RETURN)
        {
          return DLGC_WANTALLKEYS;
        }
      }
    }
  }

  if (uMsg == WM_KEYDOWN ||
      uMsg == WM_SYSKEYDOWN)
  {
    BOOL bAlt=FALSE;
    BOOL bShift=FALSE;
    BOOL bControl=FALSE;

    if (GetKeyState(VK_MENU) < 0)
      bAlt=TRUE;
    if (GetKeyState(VK_SHIFT) < 0)
      bShift=TRUE;
    if (GetKeyState(VK_CONTROL) < 0)
      bControl=TRUE;

    if (wParam == VK_RETURN)
    {
      if (!bAlt && !bShift && !bControl)
      {
        PostMessage(GetParent(hWnd), WM_COMMAND, IDC_ITEM_OPEN, 0);
        return TRUE;
      }
    }
    else if (wParam == VK_DELETE)
    {
      if (!bAlt && !bShift && !bControl)
      {
        PostMessage(GetParent(hWnd), WM_COMMAND, IDC_ITEM_DELETE, 0);
        return TRUE;
      }
    }
    else if (wParam == VK_DOWN)
    {
      if (bAlt && !bShift && !bControl)
      {
        PostMessage(GetParent(hWnd), WM_COMMAND, IDC_ITEM_MOVEDOWN, 0);
        return TRUE;
      }
    }
    else if (wParam == VK_UP)
    {
      if (bAlt && !bShift && !bControl)
      {
        PostMessage(GetParent(hWnd), WM_COMMAND, IDC_ITEM_MOVEUP, 0);
        return TRUE;
      }
    }
    else if (wParam == VK_F2)
    {
      if (!bAlt && !bShift && !bControl)
      {
        PostMessage(GetParent(hWnd), WM_COMMAND, IDC_ITEM_EDIT, 0);
        return TRUE;
      }
    }
  }

  if (bOldWindows)
    return CallWindowProcA(OldListBoxProc, hWnd, uMsg, wParam, lParam);
  else
    return CallWindowProcW(OldListBoxProc, hWnd, uMsg, wParam, lParam);
}

BOOL CALLBACK FavEditDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndName;
  static HWND hWndFile;
  static FAVITEM *fi;

  if (uMsg == WM_INITDIALOG)
  {
    fi=(FAVITEM *)lParam;

    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hIconFav);
    hWndName=GetDlgItem(hDlg, IDC_FAVNAME);
    hWndFile=GetDlgItem(hDlg, IDC_FAVFILE);

    if (hWndFavDlg)
      SetWindowTextWide(hDlg, GetLangStringW(wLangModule, STRID_FAVEDITING));
    else
      SetWindowTextWide(hDlg, GetLangStringW(wLangModule, STRID_FAVADDING));
    SetDlgItemTextWide(hDlg, IDC_FAVNAME_LABEL, GetLangStringW(wLangModule, STRID_FAVNAME));
    SetDlgItemTextWide(hDlg, IDC_FAVFILE_LABEL, GetLangStringW(wLangModule, STRID_FAVFILE));
    SetDlgItemTextWide(hDlg, IDOK, GetLangStringW(wLangModule, STRID_OK));
    SetDlgItemTextWide(hDlg, IDCANCEL, GetLangStringW(wLangModule, STRID_CANCEL));
    SetWindowTextWide(hWndName, fi->wszName);
    SetWindowTextWide(hWndFile, fi->wszFile);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDOK)
    {
      GetWindowTextWide(hWndName, fi->wszName, MAX_PATH);
      GetWindowTextWide(hWndFile, fi->wszFile, MAX_PATH);

      //Remove '=' from name to avoid save/read problems.
      xstrrepW(fi->wszName, -1, L"=", -1, L"", -1, TRUE, fi->wszName, NULL);

      EndDialog(hDlg, 1);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDCANCEL)
    {
      EndDialog(hDlg, 0);
      return TRUE;
    }
  }
  return FALSE;
}

void CreateFavStack(STACKFAV *hStack, const wchar_t *wpText)
{
  FAVITEM *lpFavItem=NULL;
  const wchar_t *wpName;
  const wchar_t *wpFile;
  int nNameLen;
  int nFileLen;

  if (wpText)
  {
    while (*wpText)
    {
      wpName=wpText;
      while (*wpText != L'=' && *wpText != L'\r' && *wpText != L'\n' && *wpText) ++wpText;

      if (*wpText == L'=')
      {
        nNameLen=(int)(wpText - wpName);
        wpFile=++wpText;

        while (*wpText != L'\r' && *wpText != L'\n') ++wpText;
        nFileLen=(int)(wpText - wpFile);

        if (lpFavItem=StackInsertFavourite(hStack))
        {
          xstrcpynW(lpFavItem->wszName, wpName, nNameLen + 1);
          xstrcpynW(lpFavItem->wszFile, wpFile, nFileLen + 1);
        }
        else break;
      }
      while (*wpText == L'\r' || *wpText == L'\n') ++wpText;
    }
  }
}

FAVITEM* StackInsertFavourite(STACKFAV *hStack)
{
  FAVITEM *lpElement=NULL;

  StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, -1, sizeof(FAVITEM));
  return lpElement;
}

FAVITEM* StackGetFavouriteByIndex(STACKFAV *hStack, int nIndex)
{
  FAVITEM *lpElement=NULL;

  StackGetElement((stack *)hStack->first, (stack *)hStack->last, (stack **)&lpElement, nIndex);
  return lpElement;
}

FAVITEM* StackGetFavouriteByFile(STACKFAV *hStack, const wchar_t *wpFile, int *nIndex)
{
  FAVITEM *lpElement;
  int nCount=0;

  for (lpElement=hStack->first; lpElement; lpElement=lpElement->next)
  {
    if (!xstrcmpiW(lpElement->wszFile, wpFile))
      break;
    ++nCount;
  }
  if (nIndex) *nIndex=nCount;
  return lpElement;
}

int StackSortFavourites(STACKFAV *hStack, int nUpDown, BOOL bShowFile)
{
  FAVITEM *tmp1;
  FAVITEM *tmp2;
  FAVITEM *tmpNext;
  int i;

  if (nUpDown != 1 && nUpDown != -1) return 1;

  for (tmp1=(FAVITEM *)hStack->first; tmp1; tmp1=tmpNext)
  {
    tmpNext=tmp1->next;

    for (tmp2=(FAVITEM *)hStack->first; tmp2 != tmp1; tmp2=tmp2->next)
    {
      if (bShowFile)
        i=xstrcmpiW(tmp2->wszFile, tmp1->wszFile);
      else
        i=xstrcmpiW(tmp2->wszName, tmp1->wszName);

      if (i == 0 || i == nUpDown)
      {
        StackMoveBefore((stack **)&hStack->first, (stack **)&hStack->last, (stack *)tmp1, (stack *)tmp2);
        break;
      }
    }
  }
  return 0;
}

void StackFreeFavourites(STACKFAV *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}

void FillFavouritesListBox(STACKFAV *hStack, HWND hWnd, BOOL bShowFile)
{
  FAVITEM *lpElement=(FAVITEM *)hStack->first;
  int nItem=0;

  SendMessage(hWnd, LB_RESETCONTENT, 0, 0);

  while (lpElement)
  {
    ListBox_InsertStringWide(hWnd, nItem++, bShowFile?lpElement->wszFile:lpElement->wszName);

    lpElement=lpElement->next;
  }
}

int MoveListBoxItem(STACKFAV *hStack, HWND hWnd, int nOldIndex, int nNewIndex)
{
  FAVITEM *lpElement;
  wchar_t *wpText;
  int nIndex=LB_ERR;
  int nTextLen;

  if ((nTextLen=(int)SendMessage(hWnd, LB_GETTEXTLEN, nOldIndex, 0)) != LB_ERR)
  {
    if (wpText=(wchar_t *)GlobalAlloc(GMEM_FIXED, (nTextLen + 1) * sizeof(wchar_t)))
    {
      if (!StackGetElement((stack *)hStack->first, (stack *)hStack->last, (stack **)&lpElement, nOldIndex + 1))
        StackMoveIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack *)lpElement, nNewIndex + 1);
      ListBox_GetTextWide(hWnd, nOldIndex, wpText);
      SendMessage(hWnd, LB_DELETESTRING, nOldIndex, 0);
      nIndex=ListBox_InsertStringWide(hWnd, nNewIndex, wpText);
      GlobalFree((HGLOBAL)wpText);
    }
  }
  return nIndex;
}

BOOL ShiftListBoxSelItems(STACKFAV *hStack, HWND hWnd, BOOL bMoveDown)
{
  int *lpSelItems;
  int nSelCount;
  int nMinIndex;
  int nMaxIndex;
  int nOldIndex=-1;
  int nNewIndex=-1;
  int i;
  BOOL bResult=FALSE;

  nMinIndex=0;
  nMaxIndex=(int)SendMessage(hWnd, LB_GETCOUNT, 0, 0) - 1;

  if (nSelCount=GetListBoxSelItems(hWnd, &lpSelItems))
  {
    if (!bMoveDown)
    {
      for (i=0; i < nSelCount; ++i)
      {
        if (lpSelItems[i] > nMinIndex)
        {
          if (nNewIndex == -1 && i > 0)
          {
            if (lpSelItems[i] - 1 <= lpSelItems[i - 1])
              continue;
          }
          nOldIndex=lpSelItems[i];
          nNewIndex=lpSelItems[i] - 1;

          MoveListBoxItem(hStack, hWnd, nOldIndex, nNewIndex);
          SendMessage(hWnd, LB_SETSEL, TRUE, nNewIndex);
          bResult=TRUE;
        }
      }
    }
    else
    {
      for (i=--nSelCount; i >= 0; --i)
      {
        if (lpSelItems[i] < nMaxIndex)
        {
          if (nNewIndex == -1 && i < nSelCount)
          {
            if (lpSelItems[i] + 1 >= lpSelItems[i + 1])
              continue;
          }
          nOldIndex=lpSelItems[i];
          nNewIndex=lpSelItems[i] + 1;

          MoveListBoxItem(hStack, hWnd, nOldIndex, nNewIndex);
          SendMessage(hWnd, LB_SETSEL, TRUE, nNewIndex);
          bResult=TRUE;
        }
      }
    }
    FreeListBoxSelItems(&lpSelItems);
  }
  return bResult;
}

int DeleteListBoxSelItems(STACKFAV *hStack, HWND hWnd)
{
  FAVITEM *lpElement;
  int *lpSelItems;
  int nCount;
  int nSelCount;
  int i;

  if (nSelCount=GetListBoxSelItems(hWnd, &lpSelItems))
  {
    for (i=nSelCount - 1; i >= 0; --i)
    {
      if (!StackGetElement((stack *)hStack->first, (stack *)hStack->last, (stack **)&lpElement, lpSelItems[i] + 1))
        StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)lpElement);
      SendMessage(hWnd, LB_DELETESTRING, lpSelItems[i], 0);
    }
    if (nSelCount == 1)
    {
      nCount=(int)SendMessage(hWnd, LB_GETCOUNT, 0, 0);
      if (nCount == lpSelItems[0])
        SendMessage(hWnd, LB_SETSEL, TRUE, lpSelItems[0] - 1);
      else
        SendMessage(hWnd, LB_SETSEL, TRUE, lpSelItems[0]);
    }
    else SetFocus(hWnd);

    FreeListBoxSelItems(&lpSelItems);
  }
  return nSelCount;
}

int DeleteListBoxOldItems(STACKFAV *hStack, HWND hWnd)
{
  FAVITEM *lpElement=(FAVITEM *)hStack->first;
  FAVITEM *lpElementNext;
  int nItem=0;
  int nDelCount=0;

  while (lpElement)
  {
    lpElementNext=lpElement->next;

    if (!FileExistsWide(lpElement->wszFile))
    {
      StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)lpElement);
      SendMessage(hWnd, LB_DELETESTRING, nItem, 0);
      ++nDelCount;
    }
    ++nItem;

    lpElement=lpElementNext;
  }
  return nDelCount;
}

int GetListBoxSelItems(HWND hWnd, int **lpSelItems)
{
  int nSelCount;

  if (lpSelItems)
  {
    nSelCount=(int)SendMessage(hWnd, LB_GETSELCOUNT, 0, 0);

    if (*lpSelItems=(int *)GlobalAlloc(GPTR, nSelCount * sizeof(int)))
    {
      return (int)SendMessage(hWnd, LB_GETSELITEMS, nSelCount, (LPARAM)*lpSelItems);
    }
  }
  return 0;
}

void FreeListBoxSelItems(int **lpSelItems)
{
  if (lpSelItems && *lpSelItems)
  {
    GlobalFree((HGLOBAL)*lpSelItems);
    *lpSelItems=NULL;
  }
}

BOOL CreateContextMenu(POPUPMENU *hMenuStack, const wchar_t *wpText, int nType)
{
  wchar_t wszMenuItem[MAX_PATH];
  wchar_t wszMethodName[MAX_PATH];
  wchar_t wszIconFile[MAX_PATH];
  STACKSUBMENU hSubmenuStack={0};
  SUBMENUITEM *lpSubmenuItem=NULL;
  SHOWSUBMENUITEM *lpShowSubmenuItem=NULL;
  MENUITEM *lpMenuItem=NULL;
  MAININDEXITEM *lpMainIndexItem=NULL;
  HMENU hMenu=NULL;
  HMENU hSubMenu;
  HMENU hParentMenu=NULL;
  const wchar_t *wpTextBegin=wpText;
  const wchar_t *wpCount=wpText;
  const wchar_t *wpLineBegin=wpText;
  const wchar_t *wpErrorBegin=wpText;
  const wchar_t *wpSetArg;
  DWORD dwAction=0;
  DWORD dwNewFlags;
  DWORD dwSetFlags=0;
  int nFlagCountNoSDI=0;
  int nFlagCountNoMDI=0;
  int nFlagCountNoPMDI=0;
  int nFlagCountNoFileExist=0;
  int nPrevSeparator=0;
  BOOL bMethod;
  BOOL bMainMenuParent;
  int nPlus;
  int nMinus;
  int nParentIndex=0;
  int nSubMenuIndex=0;
  int nSubMenuCodeItem=0;
  int nItem;
  INT_PTR nImageListIconIndex;
  int nFileIconIndex;
  int nMessageID;
  int i;

  if (nType == TYPE_MANUAL)
    nItem=IDM_MIN_MANUAL;
  else
    nItem=IDM_MIN_MENU;
  hMenuStack->bClearMainMenu=FALSE;

  if (wpCount)
  {
    hMenu=hSubMenu=CreatePopupMenu();

    if (hMenuStack->hIconMenu=IconMenu_Alloc(hMainWnd))
    {
      hMenuStack->hImageList=ImageList_Create(sizeIcon.cx, sizeIcon.cy, ILC_COLOR32|ILC_MASK, 0, 0);
      ImageList_SetBkColor(hMenuStack->hImageList, GetSysColor(COLOR_MENU));
    }

    while (*wpCount)
    {
      bMethod=FALSE;
      lpMainIndexItem=NULL;
      lpMenuItem=NULL;
      wszIconFile[0]=L'\0';
      nImageListIconIndex=-1;
      nFileIconIndex=-1;
      SkipComment(&wpCount);
      wpLineBegin=wpCount;

      if (*wpCount == L'{')
      {
        if (!IsFlagOn(dwSetFlags, CCMS_NOSDI|CCMS_NOMDI|CCMS_NOPMDI|CCMS_NOFILEEXIST))
        {
          wpErrorBegin=wpCount++;
          nMessageID=STRID_PARSEMSG_UNNAMEDSUBMENU;
          goto Error;
        }
      }
      NextString(wpCount, wszMenuItem, MAX_PATH, &wpCount, &nMinus);

      //Set options
      wpSetArg=wpLineBegin;

      if (!xstrcmpnW(L"SET(", wpSetArg, (UINT_PTR)-1))
      {
        dwNewFlags=(DWORD)xatoiW(wpSetArg + 4, &wpSetArg);

        if ((dwNewFlags & CCMS_NOSDI) && (dwSetFlags & CCMS_NOSDI))
          ++nFlagCountNoSDI;
        if ((dwNewFlags & CCMS_NOMDI) && (dwSetFlags & CCMS_NOMDI))
          ++nFlagCountNoMDI;
        if ((dwNewFlags & CCMS_NOPMDI) && (dwSetFlags & CCMS_NOPMDI))
          ++nFlagCountNoPMDI;
        if ((dwNewFlags & CCMS_NOFILEEXIST) && (dwSetFlags & CCMS_NOFILEEXIST))
          ++nFlagCountNoFileExist;
        if ((dwNewFlags & CCMS_NOFILEEXIST) && !(dwSetFlags & CCMS_NOFILEEXIST))
        {
          wchar_t wszPath[MAX_PATH];
          wchar_t *wpFileName;

          if (*wpSetArg == L',' && NextString(++wpSetArg, wszPath, MAX_PATH, &wpSetArg, NULL))
          {
            if (TranslateFileString(wszPath, wszBuffer, BUFFER_SIZE))
            {
              if (SearchPathWide(NULL, wszBuffer, NULL, MAX_PATH, wszPath, &wpFileName))
                dwNewFlags&=~CCMS_NOFILEEXIST;
            }
            while (*wpSetArg == L' ' || *wpSetArg == L'\t') ++wpSetArg;
            if (*wpSetArg == L')') ++wpSetArg;
            wpCount=wpSetArg;
          }
        }
        dwSetFlags|=dwNewFlags;

        if (IsFlagOn(dwSetFlags, CCMS_NOICONS))
          bUseIcons=FALSE;
        if (IsFlagOn(dwSetFlags, CCMS_NOSDI|CCMS_NOMDI|CCMS_NOPMDI|CCMS_NOFILEEXIST))
          continue;
        goto CheckSubmenuClose;
      }
      else if (!xstrcmpnW(L"UNSET(", wpSetArg, (UINT_PTR)-1))
      {
        dwNewFlags=(DWORD)xatoiW(wpSetArg + 6, &wpSetArg);

        if ((dwNewFlags & CCMS_NOSDI) && (dwSetFlags & CCMS_NOSDI))
        {
          if (nFlagCountNoSDI > 0 && --nFlagCountNoSDI >= 0)
            dwNewFlags&=~CCMS_NOSDI;
        }
        if ((dwNewFlags & CCMS_NOMDI) && (dwSetFlags & CCMS_NOMDI))
        {
          if (nFlagCountNoMDI > 0 && --nFlagCountNoMDI >= 0)
            dwNewFlags&=~CCMS_NOMDI;
        }
        if ((dwNewFlags & CCMS_NOPMDI) && (dwSetFlags & CCMS_NOPMDI))
        {
          if (nFlagCountNoPMDI > 0 && --nFlagCountNoPMDI >= 0)
            dwNewFlags&=~CCMS_NOPMDI;
        }
        if ((dwNewFlags & CCMS_NOFILEEXIST) && (dwSetFlags & CCMS_NOFILEEXIST))
        {
          if (nFlagCountNoFileExist > 0 && --nFlagCountNoFileExist >= 0)
            dwNewFlags&=~CCMS_NOFILEEXIST;
        }
        dwSetFlags&=~dwNewFlags;

        if (!IsFlagOn(dwSetFlags, CCMS_NOICONS))
          bUseIcons=TRUE;
        if (IsFlagOn(dwSetFlags, CCMS_NOSDI|CCMS_NOMDI|CCMS_NOPMDI|CCMS_NOFILEEXIST))
          continue;
        goto CheckSubmenuClose;
      }
      if (IsFlagOn(dwSetFlags, CCMS_NOSDI|CCMS_NOMDI|CCMS_NOPMDI|CCMS_NOFILEEXIST))
        continue;

      //Item name is special string
      if (!xstrcmpW(wszMenuItem, L"SEPARATOR"))
      {
        InsertMenuCommon(hMenuStack->hIconMenu, NULL, -1, 0, 0, hSubMenu, -1, MF_BYPOSITION|MF_SEPARATOR, IDM_SEPARATOR, NULL);
        ++nSubMenuIndex;
        ++nSubMenuCodeItem;
        nPrevSeparator=IDM_SEPARATOR;
        bMethod=TRUE;
      }
      else if (!xstrcmpW(wszMenuItem, L"SEPARATOR1"))
      {
        if (!nPrevSeparator && nSubMenuCodeItem)
        {
          InsertMenuCommon(hMenuStack->hIconMenu, NULL, -1, 0, 0, hSubMenu, -1, MF_BYPOSITION|MF_SEPARATOR, IDM_SEPARATOR1, NULL);
          ++nSubMenuIndex;
          ++nSubMenuCodeItem;
        }
        nPrevSeparator=IDM_SEPARATOR1;
        bMethod=TRUE;
      }
      else
      {
        nPrevSeparator=0;

        if (!xstrcmpW(wszMenuItem, L"EXPLORER"))
        {
          hMenuStack->hExplorerSubMenu=hSubMenu;
          hMenuStack->nExplorerFirstIndex=nSubMenuIndex;
          hMenuStack->nExplorerLastIndex=-1;
          ++nSubMenuCodeItem;
          bMethod=TRUE;
        }
        else if (!xstrcmpW(wszMenuItem, L"FAVOURITES"))
        {
          hMenuStack->hFavouritesSubMenu=hSubMenu;
          hMenuStack->nFavouritesFirstIndex=nSubMenuIndex;
          hMenuStack->nFavouritesLastIndex=-1;
          ++nSubMenuCodeItem;
          bMethod=TRUE;
        }
        else if (!xstrcmpW(wszMenuItem, L"RECENTFILES"))
        {
          hMenuStack->hRecentFilesSubMenu=hSubMenu;
          hMenuStack->nRecentFilesFirstIndex=nSubMenuIndex;
          hMenuStack->nRecentFilesLastIndex=-1;
          ++nSubMenuCodeItem;
          bMethod=TRUE;
        }
        else if (!xstrcmpW(wszMenuItem, L"LANGUAGES"))
        {
          hMenuStack->hLanguagesSubMenu=hSubMenu;
          hMenuStack->nLanguagesFirstIndex=nSubMenuIndex;
          hMenuStack->nLanguagesLastIndex=-1;
          ++nSubMenuCodeItem;
          bMethod=TRUE;
        }
        else if (!xstrcmpW(wszMenuItem, L"OPENCODEPAGES"))
        {
          hMenuStack->hOpenCodepagesSubMenu=hSubMenu;
          hMenuStack->nOpenCodepagesFirstIndex=nSubMenuIndex;
          hMenuStack->nOpenCodepagesLastIndex=-1;
          ++nSubMenuCodeItem;
          bMethod=TRUE;
        }
        else if (!xstrcmpW(wszMenuItem, L"SAVECODEPAGES"))
        {
          hMenuStack->hSaveCodepagesSubMenu=hSubMenu;
          hMenuStack->nSaveCodepagesFirstIndex=nSubMenuIndex;
          hMenuStack->nSaveCodepagesLastIndex=-1;
          ++nSubMenuCodeItem;
          bMethod=TRUE;
        }
        else if (!xstrcmpW(wszMenuItem, L"MDIDOCUMENTS"))
        {
          hMenuStack->hMdiDocumentsSubMenu=hSubMenu;
          ++nSubMenuCodeItem;
          bMethod=TRUE;
        }
        else if (!xstrcmpW(wszMenuItem, L"CLEAR"))
        {
          hMenuStack->bClearMainMenu=TRUE;
          bMethod=TRUE;
        }
        else
        {
          //Item name is normal string
          while (*wpCount == L' ' || *wpCount == L'\t') ++wpCount;

          if (*wpCount == L'+')
          {
            ++wpCount;
            nPlus=1;
          }
          else nPlus=0;

          //Parse methods
          for (;;)
          {
            while (*wpCount == L' ' || *wpCount == L'\t') ++wpCount;
            wpErrorBegin=wpCount;

            if (!GetMethodName(wpCount, wszMethodName, MAX_PATH, &wpCount))
              break;

            if (!xstrcmpiW(wszMethodName, L"Index") || !xstrcmpiW(wszMethodName, L"Break"))
            {
              if (!lpMainIndexItem)
              {
                if (!hParentMenu)
                {
                  if (!StackInsertIndex((stack **)&hMenuStack->hMainMenuIndexStack.first, (stack **)&hMenuStack->hMainMenuIndexStack.last, (stack **)&lpMainIndexItem, -1, sizeof(MAININDEXITEM)))
                  {
                    lpMainIndexItem->nMainMenuIndex=(int)xatoiW(wpCount, &wpCount);
                    if (!xstrcmpiW(wszMethodName, L"Break"))
                      lpMainIndexItem->bMenuBreak=TRUE;
                    if (*wpCount == L')') ++wpCount;
                  }
                }
                else
                {
                  nMessageID=STRID_PARSEMSG_NONPARENT;
                  goto Error;
                }
              }
              else
              {
                nMessageID=STRID_PARSEMSG_METHODALREADYDEFINED;
                goto Error;
              }
            }
            else if (!xstrcmpiW(wszMethodName, L"Icon"))
            {
              if (nImageListIconIndex == -1)
              {
                GetIconParameters(wpCount, wszIconFile, MAX_PATH, &nFileIconIndex, &wpCount);

                if (!IsFlagOn(dwSetFlags, CCMS_NOICONS))
                {
                  //Load icon optimization (part 1):
                  nImageListIconIndex=hMenuStack->nImageListIconCount;
                  ++hMenuStack->nImageListIconCount;
                }
              }
              else
              {
                nMessageID=STRID_PARSEMSG_METHODALREADYDEFINED;
                goto Error;
              }
            }
            else
            {
              //Actions
              if (!lpMenuItem)
              {
                if (!xstrcmpiW(wszMethodName, L"Command"))
                  dwAction=EXTACT_COMMAND;
                else if (!xstrcmpiW(wszMethodName, L"Call"))
                  dwAction=EXTACT_CALL;
                else if (!xstrcmpiW(wszMethodName, L"Exec"))
                  dwAction=EXTACT_EXEC;
                else if (!xstrcmpiW(wszMethodName, L"OpenFile"))
                  dwAction=EXTACT_OPENFILE;
                else if (!xstrcmpiW(wszMethodName, L"SaveFile"))
                  dwAction=EXTACT_SAVEFILE;
                else if (!xstrcmpiW(wszMethodName, L"Font"))
                  dwAction=EXTACT_FONT;
                else if (!xstrcmpiW(wszMethodName, L"Recode"))
                  dwAction=EXTACT_RECODE;
                else if (!xstrcmpiW(wszMethodName, L"Insert"))
                  dwAction=EXTACT_INSERT;
                else if (!xstrcmpiW(wszMethodName, L"Link"))
                  dwAction=EXTACT_LINK;
                else if (!xstrcmpiW(wszMethodName, L"Favourites"))
                  dwAction=EXTACT_FAVOURITE;
                else if (!xstrcmpiW(wszMethodName, L"Menu"))
                {
                  dwAction=EXTACT_MENU;
                  nMinus=1;
                }
                else dwAction=0;

                if (dwAction)
                {
                  if (!StackInsertIndex((stack **)&hMenuStack->hMenuItemStack.first, (stack **)&hMenuStack->hMenuItemStack.last, (stack **)&lpMenuItem, -1, sizeof(MENUITEM)))
                  {
                    lpMenuItem->bUpdateItem=!nMinus;
                    lpMenuItem->bAutoLoad=nPlus;
                    lpMenuItem->dwAction=dwAction;
                    lpMenuItem->nTextOffset=(int)(wpLineBegin - wpTextBegin);
                    lpMenuItem->nItem=nItem;
                    lpMenuItem->nMenuType=nType;
                    lpMenuItem->hSubMenu=hSubMenu;
                    lpMenuItem->nSubMenuIndex=nSubMenuIndex;
                    lpMenuItem->nFileIconIndex=-1;
                    lpMenuItem->nImageListIconIndex=-1;
                    ParseMethodParameters(&lpMenuItem->hParamStack, wpCount, &wpCount);

                    if (dwAction == EXTACT_COMMAND)
                    {
                      EXTPARAM *lpParameter;
                      int nCommand=0;

                      if (lpParameter=GetMethodParameter(&lpMenuItem->hParamStack, 1))
                        nCommand=(int)lpParameter->nNumber;

                      if (!*wszMenuItem)
                      {
                        if (GetMenuStringWide(hMainMenu, nCommand, wszMenuItem, MAX_PATH, MF_BYCOMMAND))
                        {
                          if (IsFlagOn(dwSetFlags, CCMS_NOSHORTCUT))
                          {
                            for (i=0; wszMenuItem[i]; ++i)
                            {
                              if (wszMenuItem[i] == L'\t')
                              {
                                wszMenuItem[i]=L'\0';
                                break;
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                  bMethod=TRUE;
                }
                else
                {
                  nMessageID=STRID_PARSEMSG_UNKNOWNMETHOD;
                  goto Error;
                }
              }
              else
              {
                nMessageID=STRID_PARSEMSG_METHODALREADYDEFINED;
                goto Error;
              }
            }
          }

          //Add menu item
          if (bMethod)
          {
            //Load icon optimization (part 2):
            if (nFileIconIndex != -1)
            {
              xstrcpynW(lpMenuItem->wszIconFile, wszIconFile, MAX_PATH);
              lpMenuItem->nFileIconIndex=nFileIconIndex;
              lpMenuItem->nImageListIconIndex=(int)nImageListIconIndex;
            }

            if (dwAction == EXTACT_MENU)
            {
              HMENU hManualMenu=NULL;
              BOOL bResult=TRUE;

              if (!hMenuManualStack.hPopupMenu && nType != TYPE_MANUAL)
                bResult=CreateContextMenu(&hMenuManualStack, wszManualText, TYPE_MANUAL);

              if (bResult && lpMenuItem->hParamStack.first)
              {
                if (!StackInsertIndex((stack **)&hMenuStack->hShowSubmenuStack.first, (stack **)&hMenuStack->hShowSubmenuStack.last, (stack **)&lpShowSubmenuItem, -1, sizeof(SHOWSUBMENUITEM)))
                {
                  lpShowSubmenuItem->lpMenuItem=lpMenuItem;
                  lpShowSubmenuItem->wpShowMenuName=lpMenuItem->hParamStack.first->wpString;
                  xstrcpynW(lpShowSubmenuItem->wszMenuItem, wszMenuItem, MAX_PATH);
                }
                hMenuStack->bLinkToManualMenu=TRUE;

                if (nType != TYPE_MANUAL)
                  hManualMenu=FindRootSubMenuByName(&hMenuManualStack, lpMenuItem->hParamStack.first->wpString);
              }

              //Add popup submenu
              if (!hParentMenu && nType == TYPE_MAIN)
              {
                //Avoid Win95 problem GetMenuString and MF_OWNERDRAW items
                AppendMenuWide(hSubMenu, MF_BYPOSITION|MF_POPUP, (UINT_PTR)hManualMenu, wszMenuItem);
              }
              else InsertMenuCommon(hMenuStack->hIconMenu, hMenuStack->hImageList, nImageListIconIndex, sizeIcon.cx, sizeIcon.cy, hSubMenu, -1, MF_BYPOSITION|MF_POPUP, (UINT_PTR)hManualMenu, wszMenuItem);
            }
            else
            {
              InsertMenuCommon(hMenuStack->hIconMenu, hMenuStack->hImageList, nImageListIconIndex, sizeIcon.cx, sizeIcon.cy, hSubMenu, -1, MF_BYPOSITION|MF_STRING, nItem, wszMenuItem);
            }
            ++nItem;
            ++nSubMenuIndex;
            ++nSubMenuCodeItem;
          }
        }
      }

      //Check for submenu open
      SkipComment(&wpCount);

      if (*wpCount == L'{')
      {
        if (!bMethod)
        {
          ++wpCount;

          if (!StackInsertIndex((stack **)&hSubmenuStack.first, (stack **)&hSubmenuStack.last, (stack **)&lpSubmenuItem, -1, sizeof(SUBMENUITEM)))
          {
            lpSubmenuItem->hSubMenu=hSubMenu;
            lpSubmenuItem->nSubMenuIndex=nSubMenuIndex;
            lpSubmenuItem->nSubMenuCodeItem=nSubMenuCodeItem;
          }
          else break;

          if (!hParentMenu && nType == TYPE_MAIN)
            bMainMenuParent=TRUE;
          else
            bMainMenuParent=FALSE;
          hParentMenu=hSubMenu;
          nParentIndex=nSubMenuIndex;
          hSubMenu=CreatePopupMenu();
          nSubMenuIndex=0;
          nSubMenuCodeItem=0;
          nPrevSeparator=0;

          //Load icon optimization (part 3):
          if (!StackInsertIndex((stack **)&hMenuStack->hMenuItemStack.first, (stack **)&hMenuStack->hMenuItemStack.last, (stack **)&lpMenuItem, -1, sizeof(MENUITEM)))
          {
            lpMenuItem->nTextOffset=(int)(wpLineBegin - wpTextBegin);
            lpMenuItem->nMenuType=nType;
            lpMenuItem->hSubMenu=hParentMenu;
            lpMenuItem->nSubMenuIndex=nParentIndex;
            lpMenuItem->nFileIconIndex=-1;
            lpMenuItem->nImageListIconIndex=-1;

            if (nFileIconIndex != -1)
            {
              xstrcpynW(lpMenuItem->wszIconFile, wszIconFile, MAX_PATH);
              lpMenuItem->nFileIconIndex=nFileIconIndex;
              lpMenuItem->nImageListIconIndex=(int)nImageListIconIndex;
            }
          }

          //Add popup submenu
          if (bMainMenuParent)
          {
            //Avoid Win95 problem GetMenuString and MF_OWNERDRAW items
            AppendMenuWide(hParentMenu, MF_BYPOSITION|MF_POPUP, (UINT_PTR)hSubMenu, wszMenuItem);
          }
          else InsertMenuCommon(hMenuStack->hIconMenu, hMenuStack->hImageList, nImageListIconIndex, sizeIcon.cx, sizeIcon.cy, hParentMenu, -1, MF_BYPOSITION|MF_POPUP, (UINT_PTR)hSubMenu, wszMenuItem);
        }
        else
        {
          wpErrorBegin=wpCount++;
          nMessageID=STRID_PARSEMSG_UNNAMEDSUBMENU;
          goto Error;
        }
      }
      else
      {
        if (!bMethod)
        {
          wpErrorBegin=wpLineBegin;
          nMessageID=STRID_PARSEMSG_NOMETHOD;
          goto Error;
        }
      }

      //Check for submenu close
      CheckSubmenuClose:
      for (;;)
      {
        SkipComment(&wpCount);

        if (*wpCount == L'}')
        {
          if (lpSubmenuItem=hSubmenuStack.last)
          {
            if (nPrevSeparator == IDM_SEPARATOR1)
              RemoveSeparator1(hMenuStack, hSubMenu);

            //Goto parent submenu
            hSubMenu=lpSubmenuItem->hSubMenu;
            nSubMenuIndex=lpSubmenuItem->nSubMenuIndex;
            nSubMenuCodeItem=lpSubmenuItem->nSubMenuCodeItem;
            StackDelete((stack **)&hSubmenuStack.first, (stack **)&hSubmenuStack.last, (stack *)lpSubmenuItem);

            if (lpSubmenuItem=hSubmenuStack.last)
              hParentMenu=lpSubmenuItem->hSubMenu;
            else
              hParentMenu=NULL;
            ++nSubMenuIndex;
            ++nSubMenuCodeItem;
            nPrevSeparator=0;
          }
          else
          {
            wpErrorBegin=wpCount++;
            nMessageID=STRID_PARSEMSG_NOOPENBRACKET;
            goto Error;
          }
          ++wpCount;
        }
        else break;
      }
    }
    if (nPrevSeparator == IDM_SEPARATOR1)
      RemoveSeparator1(hMenuStack, hSubMenu);
    StackClear((stack **)&hSubmenuStack.first, (stack **)&hSubmenuStack.last);

    //Load icon optimization (part 4):
    //Resize hImageList and fill it later with requested icons when submenu is about to open
    ImageList_SetImageCount(hMenuStack->hImageList, hMenuStack->nImageListIconCount);
  }
  hMenuStack->hPopupMenu=hMenu;

  //Resolve submenu links
  if (nType == TYPE_MANUAL)
  {
    HMENU hManualMenu;

    for (lpShowSubmenuItem=hMenuStack->hShowSubmenuStack.first; lpShowSubmenuItem; lpShowSubmenuItem=lpShowSubmenuItem->next)
    {
      if (hManualMenu=FindRootSubMenuByName(hMenuStack, lpShowSubmenuItem->wpShowMenuName))
      {
        ModifyMenuCommon(hMenuStack->hIconMenu, hMenuStack->hImageList, lpShowSubmenuItem->lpMenuItem->nImageListIconIndex, sizeIcon.cx, sizeIcon.cy, lpShowSubmenuItem->lpMenuItem->hSubMenu, lpShowSubmenuItem->lpMenuItem->nSubMenuIndex, MF_BYPOSITION|MF_POPUP, (UINT_PTR)hManualMenu, lpShowSubmenuItem->wszMenuItem);
      }
    }
  }
  return TRUE;

  Error:
  hMenuStack->hPopupMenu=hMenu;
  FreeContextMenu(hMenuStack);

  nExtMenuIndex=nType;
  crExtSetSel.cpMin=wpErrorBegin - wpText;
  crExtSetSel.cpMax=wpCount - wpText;
  bExtFocusEdit=TRUE;

  if (!hThread)
    hThread=CreateThread(NULL, 0, ThreadProc, NULL, 0, &dwThreadId);
  else
    PostMessage(hWndMainDlg, AKDLL_MENUINDEX, nExtMenuIndex, 0);

  xprintfW(wszBuffer, GetLangStringW(wLangModule, nMessageID), wszMenuItem, wszMethodName);
  MessageBoxW(hWndMainDlg?hWndMainDlg:NULL, wszBuffer, wszPluginTitle, MB_OK|MB_ICONERROR|(hWndMainDlg?0:MB_TOPMOST));
  return FALSE;
}

DWORD IsFlagOn(DWORD dwSetFlags, DWORD dwCheckFlags)
{
  if (!dwSetFlags) return 0;

  if ((dwCheckFlags & CCMS_NOSDI) && (dwSetFlags & CCMS_NOSDI) && nMDI == WMD_SDI)
    return CCMS_NOSDI;
  if ((dwCheckFlags & CCMS_NOMDI) && (dwSetFlags & CCMS_NOMDI) && nMDI == WMD_MDI)
    return CCMS_NOMDI;
  if ((dwCheckFlags & CCMS_NOPMDI) && (dwSetFlags & CCMS_NOPMDI) && nMDI == WMD_PMDI)
    return CCMS_NOPMDI;
  if ((dwCheckFlags & CCMS_NOSHORTCUT) && (dwSetFlags & CCMS_NOSHORTCUT))
    return CCMS_NOSHORTCUT;
  if ((dwCheckFlags & CCMS_NOICONS) && (dwSetFlags & CCMS_NOICONS))
    return CCMS_NOICONS;
  if ((dwCheckFlags & CCMS_NOFILEEXIST) && (dwSetFlags & CCMS_NOFILEEXIST))
    return CCMS_NOFILEEXIST;
  return 0;
}

void UpdateContextMenu(POPUPMENU *hMenuStack, int nType, HMENU hSubMenu)
{
  EDITINFO ei;
  MENUITEM *lpElement;
  EXTPARAM *lpParameter;
  BOOL bInitMenu=FALSE;

  Loop:
  lpElement=hMenuStack->hMenuItemStack.first;
  ei.hWndEdit=NULL;

  while (lpElement)
  {
    if (hSubMenu)
    {
      if (lpElement->hSubMenu == hSubMenu)
      {
        //Load icon optimization (part 5):
        if (lpElement->nFileIconIndex != -1)
        {
          HICON hIcon=NULL;

          if (!*lpElement->wszIconFile)
          {
            hIcon=(HICON)LoadImageA(hInstanceDLL, MAKEINTRESOURCEA(lpElement->nFileIconIndex + 100), IMAGE_ICON, sizeIcon.cx, sizeIcon.cy, 0);

            if (hIcon)
            {
              ImageList_ReplaceIcon(hMenuStack->hImageList, lpElement->nImageListIconIndex, hIcon);
              DestroyIcon(hIcon);
            }
          }
          else
          {
            wchar_t wszPath[MAX_PATH];
            wchar_t *wpFileName;

            if (TranslateFileString(lpElement->wszIconFile, wszPath, MAX_PATH))
            {
              if (SearchPathWide(NULL, wszPath, NULL, MAX_PATH, lpElement->wszIconFile, &wpFileName))
              {
                if (bBigIcons)
                  ExtractIconExWide(wszPath, lpElement->nFileIconIndex, &hIcon, NULL, 1);
                else
                  ExtractIconExWide(wszPath, lpElement->nFileIconIndex, NULL, &hIcon, 1);

                if (hIcon)
                {
                  ImageList_ReplaceIcon(hMenuStack->hImageList, lpElement->nImageListIconIndex, hIcon);
                  DestroyIcon(hIcon);
                }
              }
            }
          }

          if (!hIcon)
            ImageList_ReplaceIcon(hMenuStack->hImageList, lpElement->nImageListIconIndex, hIconHollow);
          lpElement->nFileIconIndex=-1;
        }
      }
    }
    else
    {
      if (lpElement->bUpdateItem)
      {
        if (lpElement->dwAction == EXTACT_COMMAND)
        {
          DWORD dwMenuState=0;
          int nCommand=0;

          if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
            nCommand=(int)lpParameter->nNumber;

          if (nCommand)
          {
            if (nCommand == IDM_VIEW_SPLIT_WINDOW_ALL ||
                nCommand == IDM_VIEW_SPLIT_WINDOW_WE ||
                nCommand == IDM_VIEW_SPLIT_WINDOW_NS)
            {
              if (!ei.hWndEdit)
                SendMessage(hMainWnd, AKD_GETEDITINFO, (WPARAM)NULL, (LPARAM)&ei);
              if (ei.hWndEdit)
              {
                if ((nCommand == IDM_VIEW_SPLIT_WINDOW_ALL && ei.hWndClone1 && ei.hWndClone2 && ei.hWndClone3) ||
                    (nCommand == IDM_VIEW_SPLIT_WINDOW_WE && ei.hWndClone1 && !ei.hWndClone2 && !ei.hWndClone3) ||
                    (nCommand == IDM_VIEW_SPLIT_WINDOW_NS && !ei.hWndClone1 && ei.hWndClone2 && !ei.hWndClone3))
                {
                  dwMenuState=MF_CHECKED;
                }
              }
            }
            else
            {
              if (!bInitMenu)
              {
                SendMessage(hMainWnd, WM_INITMENU, (WPARAM)hMainMenu, IMENU_EDIT|IMENU_CHECKS);
                bInitMenu=TRUE;
              }
              dwMenuState=GetMenuState(hMainMenu, nCommand, MF_BYCOMMAND);
            }
            if (dwMenuState != (DWORD)-1)
            {
              CheckMenuItem(hMenuStack->hPopupMenu, lpElement->nItem, MF_BYCOMMAND|dwMenuState);
              EnableMenuItem(hMenuStack->hPopupMenu, lpElement->nItem, MF_BYCOMMAND|dwMenuState);
            }
          }
        }
        else if (lpElement->dwAction == EXTACT_CALL)
        {
          PLUGINFUNCTION *pf;
          wchar_t *wpFunction=NULL;
          DWORD dwFlags;
          int nDllAction=0;
          BOOL bCoderMark;

          if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
            wpFunction=lpParameter->wpString;
          if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 2))
            nDllAction=(int)lpParameter->nNumber;

          if (wpFunction)
          {
            dwFlags=MF_UNCHECKED;
            if (pf=(PLUGINFUNCTION *)SendMessage(hMainWnd, AKD_DLLFINDW, (WPARAM)wpFunction, 0))
              if (pf->bRunning) dwFlags=MF_CHECKED;

            //Coder special processing
            if (!xstrcmpinW(L"Coder::", wpFunction, (UINT_PTR)-1))
            {
              if ((nDllAction == DLLA_CODER_SETEXTENSION || nDllAction == DLLA_CODER_SETVARTHEME) && !xstrcmpiW(wpFunction, L"Coder::Settings"))
                bCoderMark=FALSE;
              else if (nDllAction == DLLA_HIGHLIGHT_MARK && !xstrcmpiW(wpFunction, L"Coder::HighLight"))
                bCoderMark=TRUE;
              else
                bCoderMark=-1;

              if (bCoderMark != -1)
              {
                if ((pf=(PLUGINFUNCTION *)SendMessage(hMainWnd, AKD_DLLFINDW, (WPARAM)L"Coder::HighLight", 0)) && pf->bRunning)
                {
                  dwFlags=MF_CHECKED;
                }
                else if (bCoderMark == FALSE)
                {
                  if ((pf=(PLUGINFUNCTION *)SendMessage(hMainWnd, AKD_DLLFINDW, (WPARAM)L"Coder::CodeFold", 0)) && pf->bRunning)
                  {
                    dwFlags=MF_CHECKED;
                  }
                  else if (nDllAction == DLLA_CODER_SETEXTENSION)
                  {
                    if ((pf=(PLUGINFUNCTION *)SendMessage(hMainWnd, AKD_DLLFINDW, (WPARAM)L"Coder::AutoComplete", 0)) && pf->bRunning)
                      dwFlags=MF_CHECKED;
                  }
                }

                if (dwFlags == MF_CHECKED)
                {
                  PLUGINCALLSENDW pcs;
                  DLLEXTCODERCHECKALIAS decca;
                  DLLEXTCODERCHECKVARTHEME deccvt;
                  DLLEXTCODERCHECKMARK deccm;
                  wchar_t wszAlias[MAX_PATH];
                  BOOL bActive=FALSE;
                  BOOL bCheck=FALSE;

                  if (bCoderMark)
                  {
                    deccm.dwStructSize=sizeof(DLLEXTCODERCHECKMARK);
                    deccm.nAction=DLLA_HIGHLIGHT_CHECKMARK;
                    deccm.nMarkID=-1;
                    deccm.wpColorText=NULL;
                    deccm.wpColorBk=NULL;
                    deccm.lpbActive=&bActive;
                    if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 7))
                      deccm.nMarkID=(int)lpParameter->nNumber;
                    if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 3))
                      deccm.wpColorText=lpParameter->wpString;
                    if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 4))
                      deccm.wpColorBk=lpParameter->wpString;
                    pcs.lParam=(LPARAM)&deccm;
                    bCheck=TRUE;
                  }
                  else
                  {
                    if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 3))
                    {
                      if (nDllAction == DLLA_CODER_SETEXTENSION)
                      {
                        decca.dwStructSize=sizeof(DLLEXTCODERCHECKALIAS);
                        decca.nAction=DLLA_CODER_CHECKALIAS;

                        wszAlias[0]=L'.';
                        xstrcpynW(wszAlias + 1, lpParameter->wpString, MAX_PATH - 1);
                        if (!wszAlias[1]) wszAlias[0]=L'\0';
                        decca.wpAlias=wszAlias;
                        decca.lpbActive=&bActive;
                        pcs.lParam=(LPARAM)&decca;
                        bCheck=TRUE;
                      }
                      else
                      {
                        deccvt.dwStructSize=sizeof(DLLEXTCODERCHECKVARTHEME);
                        deccvt.nAction=DLLA_CODER_CHECKVARTHEME;
                        deccvt.wpVarTheme=lpParameter->wpString;
                        deccvt.lpbActive=&bActive;
                        pcs.lParam=(LPARAM)&deccvt;
                        bCheck=TRUE;
                      }
                    }
                  }

                  if (bCheck)
                  {
                    pcs.pFunction=wpFunction;
                    pcs.dwSupport=PDS_STRWIDE;
                    SendMessage(hMainWnd, AKD_DLLCALLW, 0, (LPARAM)&pcs);
                  }
                  if (bActive)
                    dwFlags=MF_CHECKED;
                  else
                    dwFlags=MF_UNCHECKED;
                }
              }
            }
            //SpecialChar special processing
            else if (!xstrcmpiW(L"SpecialChar::Settings", wpFunction))
            {
              if (nDllAction == DLLA_SPECIALCHAR_OLDSET)
              {
                dwFlags=MF_UNCHECKED;
                if (pf=(PLUGINFUNCTION *)SendMessage(hMainWnd, AKD_DLLFINDW, (WPARAM)L"SpecialChar::Main", 0))
                  if (pf->bRunning) dwFlags=MF_CHECKED;

                if (dwFlags == MF_CHECKED)
                {
                  PLUGINCALLSENDW pcs;
                  DLLEXTSPECIALCHAR desc;
                  wchar_t *wpSpecialChar=NULL;
                  BOOL bColorEnable=FALSE;
                  BOOL bSelColorEnable=FALSE;

                  if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 3))
                  {
                    wpSpecialChar=lpParameter->wpString;
                  }
                  if (wpSpecialChar)
                  {
                    desc.dwStructSize=sizeof(DLLEXTSPECIALCHAR);
                    desc.nAction=DLLA_SPECIALCHAR_OLDGET;
                    desc.pSpecialChar=(unsigned char *)wpSpecialChar;
                    desc.lpcrColor=NULL;
                    desc.lpcrSelColor=NULL;
                    desc.lpbColorEnable=&bColorEnable;
                    desc.lpbSelColorEnable=&bSelColorEnable;

                    pcs.pFunction=wpFunction;
                    pcs.lParam=(LPARAM)&desc;
                    pcs.dwSupport=PDS_STRWIDE;
                    SendMessage(hMainWnd, AKD_DLLCALLW, 0, (LPARAM)&pcs);
                  }
                  if (bColorEnable || bSelColorEnable)
                    dwFlags=MF_CHECKED;
                  else
                    dwFlags=MF_UNCHECKED;
                }
              }
            }
            //LineBoard special processing
            else if (!xstrcmpiW(L"LineBoard::Main", wpFunction))
            {
              if (nDllAction == DLLA_LINEBOARD_SETRULERHEIGHT)
              {
                dwFlags=MF_UNCHECKED;
                if (pf=(PLUGINFUNCTION *)SendMessage(hMainWnd, AKD_DLLFINDW, (WPARAM)L"LineBoard::Main", 0))
                  if (pf->bRunning) dwFlags=MF_CHECKED;

                if (dwFlags == MF_CHECKED)
                {
                  PLUGINCALLSENDW pcs;
                  DLLEXTLINEBOARD delb;
                  BOOL bRulerEnable=FALSE;
                  int nRulerHeight=FALSE;

                  delb.dwStructSize=sizeof(DLLEXTLINEBOARD);
                  delb.nAction=DLLA_LINEBOARD_GETRULERHEIGHT;
                  delb.lpnRulerHeight=&nRulerHeight;
                  delb.lpbRulerEnable=&bRulerEnable;

                  pcs.pFunction=wpFunction;
                  pcs.lParam=(LPARAM)&delb;
                  pcs.dwSupport=PDS_STRWIDE;
                  SendMessage(hMainWnd, AKD_DLLCALLW, 0, (LPARAM)&pcs);

                  if (bRulerEnable && nRulerHeight)
                    dwFlags=MF_CHECKED;
                  else
                    dwFlags=MF_UNCHECKED;
                }
              }
            }
            CheckMenuItem(hMenuStack->hPopupMenu, lpElement->nItem, dwFlags);
          }
        }
        else if (lpElement->dwAction == EXTACT_LINK)
        {
          if (nType == TYPE_URL)
            EnableMenuItem(hMenuStack->hPopupMenu, lpElement->nItem, MF_BYCOMMAND|MF_ENABLED);
          else
            EnableMenuItem(hMenuStack->hPopupMenu, lpElement->nItem, MF_BYCOMMAND|MF_GRAYED);
        }
        else if (lpElement->dwAction == EXTACT_FAVOURITE)
        {
          int nCmd=0;
          DWORD dwFlags=MF_GRAYED;

          if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
            nCmd=(int)lpParameter->nNumber;

          if (nCmd == FAV_ADDDIALOG ||
              nCmd == FAV_ADDSILENT ||
              nCmd == FAV_DELSILENT)
          {
            if (*wszCurrentFile)
            {
              if (StackGetFavouriteByFile(&hFavStack, wszCurrentFile, NULL))
              {
                if (nCmd == FAV_DELSILENT)
                  dwFlags=MF_ENABLED;
              }
              else
              {
                if (nCmd != FAV_DELSILENT)
                  dwFlags=MF_ENABLED;
              }
            }
            EnableMenuItem(hMenuStack->hPopupMenu, lpElement->nItem, MF_BYCOMMAND|dwFlags);
          }
        }
      }
    }
    lpElement=lpElement->next;
  }
  if (hMenuStack->bLinkToManualMenu && hMenuStack != &hMenuManualStack)
  {
    hMenuStack=&hMenuManualStack;
    goto Loop;
  }
}

int ShowContextMenu(POPUPMENU *hMenuStack, HWND hWnd, int nType, int x, int y, HMENU hMenu)
{
  DWORD dwFlags;
  int nCmd;

  if (!hMenu) hMenu=hMenuStack->hPopupMenu;
  UpdateContextMenu(hMenuStack, nType, NULL);
  lpCurrentMenuStack=hMenuStack;
  nCurrentMenuType=nType;

  dwFlags=TPM_RETURNCMD|TPM_LEFTBUTTON|TPM_RIGHTBUTTON;
  if (nType == TYPE_RECENTFILES)
    dwFlags|=TPM_RECURSE;
  else if (nType == TYPE_MANUAL)
    dwFlags&=~TPM_RIGHTBUTTON;
  nCmd=TrackPopupMenu(hMenu, dwFlags, x, y, 0, hWnd, NULL);

  lpCurrentMenuStack=NULL;
  nCurrentMenuType=TYPE_UNKNOWN;

  return nCmd;
}

void ShowUrlMenu(HWND hWnd, CHARRANGE64 *crUrl, int x, int y)
{
  MENUITEM *lpElement;
  EXTPARAM *lpParameter;
  GETTEXTRANGE gtr;
  CHARRANGE64 cr;
  int nCmd;
  int nLockScroll;
  BOOL bResult=TRUE;

  if (GetCurFile(wszCurrentFile, MAX_PATH))
  {
    if (!hMenuUrlStack.hPopupMenu)
      bResult=CreateContextMenu(&hMenuUrlStack, wszUrlText, TYPE_URL);

    if (bResult)
    {
      if (nCmd=ShowContextMenu(&hMenuUrlStack, hMainWnd, TYPE_URL, x, y, NULL))
      {
        if (lpElement=GetContextMenuItem(&hMenuUrlStack, nCmd))
        {
          if (lpElement->dwAction == EXTACT_LINK)
          {
            if (GetKeyState(VK_CONTROL) & 0x80)
            {
              ViewItemCode(lpElement);
            }
            else
            {
              int nCommand=0;

              if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
                nCommand=(int)lpParameter->nNumber;

              if (nCommand)
              {
                if (nCommand == URL_OPEN)
                {
                  gtr.cpMin=crUrl->cpMin;
                  gtr.cpMax=crUrl->cpMax;

                  if (SendMessage(hMainWnd, AKD_GETTEXTRANGEW, (LPARAM)hWnd, (WPARAM)&gtr))
                  {
                    ShellExecuteWide(hWnd, L"open", (wchar_t *)gtr.pText, NULL, NULL, SW_SHOWNORMAL);
                    SendMessage(hMainWnd, AKD_FREETEXT, 0, (WPARAM)gtr.pText);
                  }
                }
                else if (nCommand == URL_COPY)
                {
                  SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
                  if ((nLockScroll=(int)SendMessage(hWnd, AEM_LOCKSCROLL, (WPARAM)-1, 0)) == -1)
                    SendMessage(hWnd, AEM_LOCKSCROLL, SB_BOTH, TRUE);

                  SendMessage(hWnd, EM_EXGETSEL64, 0, (WPARAM)&cr);
                  SendMessage(hWnd, EM_EXSETSEL64, 0, (WPARAM)crUrl);
                  SendMessage(hMainWnd, AKD_COPY, (WPARAM)hWnd, 0);
                  SendMessage(hWnd, EM_EXSETSEL64, 0, (WPARAM)&cr);

                  if (nLockScroll == -1)
                    SendMessage(hWnd, AEM_LOCKSCROLL, SB_BOTH, FALSE);
                  SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
                  InvalidateRect(hWnd, NULL, FALSE);
                }
                else if (nCommand == URL_SELECT)
                {
                  SendMessage(hWnd, EM_EXSETSEL64, 0, (WPARAM)crUrl);
                }
                else if (nCommand == URL_CUT)
                {
                  SendMessage(hWnd, EM_EXSETSEL64, 0, (WPARAM)crUrl);
                  SendMessage(hMainWnd, WM_COMMAND, IDM_EDIT_CUT, 0);
                }
                else if (nCommand == URL_PASTE)
                {
                  SendMessage(hWnd, EM_EXSETSEL64, 0, (WPARAM)crUrl);
                  SendMessage(hMainWnd, WM_COMMAND, IDM_EDIT_PASTE, 0);
                }
                else if (nCommand == URL_DELETE)
                {
                  SendMessage(hWnd, EM_EXSETSEL64, 0, (WPARAM)crUrl);
                  SendMessage(hMainWnd, WM_COMMAND, IDM_EDIT_CLEAR, 0);
                }
              }
            }
          }
          else
          {
            gtr.cpMin=crUrl->cpMin;
            gtr.cpMax=crUrl->cpMax;
            wszUrlLink[0]=L'\0';

            if (SendMessage(hMainWnd, AKD_GETTEXTRANGEW, (LPARAM)hWnd, (WPARAM)&gtr))
            {
              xstrcpynW(wszUrlLink, (wchar_t *)gtr.pText, MAX_PATH);
              SendMessage(hMainWnd, AKD_FREETEXT, 0, (WPARAM)gtr.pText);
            }
            CallContextMenu(&hMenuUrlStack, nCmd);
            wszUrlLink[0]=L'\0';
          }
        }
      }
    }
  }
}

void InitMenuPopup(POPUPMENU *hMenuStack, int nMenuType)
{
  MENUITEM *lpElement;
  wchar_t wszItem[MAX_PATH];
  DWORD dwState;
  int nItemID;
  int nCountBefore;
  int nCountAfter;
  int i;

  if (nMenuType == INIT_EXPLORER)
  {
    if (xstrcmpiW(hMenuStack->wszExplorerFile, wszCurrentFile))
    {
      xstrcpynW(hMenuStack->wszExplorerFile, wszCurrentFile, MAX_PATH);
      GetExplorerMenu(&hMenuStack->pExplorerMenu, &hMenuStack->pExplorerSubMenu2, &hMenuStack->pExplorerSubMenu3, hMainWnd, hMenuStack->wszExplorerFile);

      //Clear explorer menu
      for (i=hMenuStack->nExplorerLastIndex; i >= hMenuStack->nExplorerFirstIndex; --i)
      {
        DeleteMenu(hMenuStack->hExplorerSubMenu, i, MF_BYPOSITION);
      }
      hMenuStack->nExplorerLastIndex=-1;

      if (hMenuStack->pExplorerMenu)
      {
        nCountBefore=GetMenuItemCount(hMenuStack->hExplorerSubMenu);
        hMenuStack->pExplorerMenu->lpVtbl->QueryContextMenu(hMenuStack->pExplorerMenu, hMenuStack->hExplorerSubMenu, hMenuStack->nExplorerFirstIndex, IDM_MIN_EXPLORER, IDM_MAX_EXPLORER, 0);
        nCountAfter=GetMenuItemCount(hMenuStack->hExplorerSubMenu);
        hMenuStack->nExplorerLastIndex=hMenuStack->nExplorerFirstIndex + (nCountAfter - nCountBefore) - 1;
      }
    }
  }
  else if (nMenuType == INIT_FAVOURITES)
  {
    FAVITEM *lpFavItem;
    FAVITEM *lpFavCurrent;

    //Clear favourites menu
    for (i=hMenuStack->nFavouritesLastIndex; i >= hMenuStack->nFavouritesFirstIndex; --i)
    {
      DeleteMenuCommon(hMenuStack->hIconMenu, hMenuStack->hFavouritesSubMenu, i, MF_BYPOSITION);
    }
    hMenuStack->nFavouritesLastIndex=-1;

    //Fill menu
    lpFavCurrent=StackGetFavouriteByFile(&hFavStack, wszCurrentFile, NULL);

    for (i=0, lpFavItem=hFavStack.first; lpFavItem; ++i, lpFavItem=lpFavItem->next)
    {
      InsertMenuCommon(hMenuStack->hIconMenu, NULL, -1, 0, 0, hMenuStack->hFavouritesSubMenu, hMenuStack->nFavouritesFirstIndex + i, MF_BYPOSITION|(lpFavCurrent == lpFavItem?MF_CHECKED:0), IDM_MIN_FAVOURITES + i, lpFavItem->wszName);
    }
    hMenuStack->nFavouritesLastIndex=hMenuStack->nFavouritesFirstIndex + i - 1;
  }
  else if (nMenuType == INIT_RECENTFILES)
  {
    //Clear recent files menu
    for (i=hMenuStack->nRecentFilesLastIndex; i >= hMenuStack->nRecentFilesFirstIndex; --i)
    {
      DeleteMenuCommon(hMenuStack->hIconMenu, hMenuStack->hRecentFilesSubMenu, i, MF_BYPOSITION);
    }
    hMenuStack->nRecentFilesLastIndex=-1;

    //Force update AkelPad's recent files list
    SendMessage(hMainWnd, WM_INITMENUPOPUP, (WPARAM)hMenuRecentFiles, 0);

    //Fill menu
    if (GetMenuItemID(hMenuRecentFiles, 0) != IDM_RECENT_FILES)
    {
      for (i=0; GetMenuStringWide(hMenuRecentFiles, i, wszItem, MAX_PATH, MF_BYPOSITION); ++i)
      {
        nItemID=GetMenuItemID(hMenuRecentFiles, i);
        InsertMenuCommon(hMenuStack->hIconMenu, NULL, -1, 0, 0, hMenuStack->hRecentFilesSubMenu, hMenuStack->nRecentFilesFirstIndex + i, MF_BYPOSITION, nItemID, wszItem);
      }
      hMenuStack->nRecentFilesLastIndex=hMenuStack->nRecentFilesFirstIndex + i - 1;
    }
  }
  else if (nMenuType == INIT_LANGUAGES)
  {
    //Clear languages menu
    for (i=hMenuStack->nLanguagesLastIndex; i >= hMenuStack->nLanguagesFirstIndex; --i)
    {
      DeleteMenuCommon(hMenuStack->hIconMenu, hMenuStack->hLanguagesSubMenu, i, MF_BYPOSITION);
    }
    hMenuStack->nLanguagesLastIndex=-1;

    //Force update AkelPad's language list
    SendMessage(hMainWnd, WM_INITMENUPOPUP, (WPARAM)hMenuLanguage, 0);

    //Fill menu
    if (GetMenuItemID(hMenuLanguage, 0) != IDM_LANGUAGE)
    {
      for (i=0; GetMenuStringWide(hMenuLanguage, i, wszItem, MAX_PATH, MF_BYPOSITION); ++i)
      {
        if ((dwState=GetMenuState(hMenuLanguage, i, MF_BYPOSITION)) != (DWORD)-1)
        {
          nItemID=GetMenuItemID(hMenuLanguage, i);
          InsertMenuCommon(hMenuStack->hIconMenu, NULL, -1, 0, 0, hMenuStack->hLanguagesSubMenu, hMenuStack->nLanguagesFirstIndex + i, MF_BYPOSITION|dwState, nItemID, wszItem);
        }
        else break;
      }
      hMenuStack->nLanguagesLastIndex=hMenuStack->nLanguagesFirstIndex + i - 1;
    }

    //Check internal language item
    if ((dwState=GetMenuState(hMenuLanguage, IDM_LANGUAGE, MF_BYCOMMAND)) != (DWORD)-1)
    {
      if (dwState & MF_CHECKED)
      {
        if (lpElement=GetCommandItem(hMenuStack, IDM_LANGUAGE))
          CheckMenuItem(hMenuStack->hLanguagesSubMenu, lpElement->nItem, MF_BYCOMMAND|dwState);
      }
    }
  }
  else if (nMenuType == INIT_OPENCODEPAGES)
  {
    //Clear codepages menu
    for (i=hMenuStack->nOpenCodepagesLastIndex; i >= hMenuStack->nOpenCodepagesFirstIndex; --i)
    {
      DeleteMenuCommon(hMenuStack->hIconMenu, hMenuStack->hOpenCodepagesSubMenu, i, MF_BYPOSITION);
    }
    hMenuStack->nOpenCodepagesLastIndex=-1;

    //Force update AkelPad's codepages list
    SendMessage(hMainWnd, WM_INITMENUPOPUP, (WPARAM)hPopupOpenCodepage, 0);

    //Fill menu
    for (i=0; GetMenuStringWide(hPopupOpenCodepage, i, wszItem, MAX_PATH, MF_BYPOSITION); ++i)
    {
      if ((dwState=GetMenuState(hPopupOpenCodepage, i, MF_BYPOSITION)) != (DWORD)-1)
      {
        nItemID=GetMenuItemID(hPopupOpenCodepage, i);
        InsertMenuCommon(hMenuStack->hIconMenu, NULL, -1, 0, 0, hMenuStack->hOpenCodepagesSubMenu, hMenuStack->nOpenCodepagesFirstIndex + i, MF_BYPOSITION|dwState, nItemID, wszItem);
      }
      else break;
    }
    hMenuStack->nOpenCodepagesLastIndex=hMenuStack->nOpenCodepagesFirstIndex + i - 1;
  }
  else if (nMenuType == INIT_SAVECODEPAGES)
  {
    //Clear codepages menu
    for (i=hMenuStack->nSaveCodepagesLastIndex; i >= hMenuStack->nSaveCodepagesFirstIndex; --i)
    {
      DeleteMenuCommon(hMenuStack->hIconMenu, hMenuStack->hSaveCodepagesSubMenu, i, MF_BYPOSITION);
    }
    hMenuStack->nSaveCodepagesLastIndex=-1;

    //Force update AkelPad's codepages list
    SendMessage(hMainWnd, WM_INITMENUPOPUP, (WPARAM)hPopupSaveCodepage, 0);

    //Fill menu
    for (i=0; GetMenuStringWide(hPopupSaveCodepage, i, wszItem, MAX_PATH, MF_BYPOSITION); ++i)
    {
      if ((dwState=GetMenuState(hPopupSaveCodepage, i, MF_BYPOSITION)) != (DWORD)-1)
      {
        nItemID=GetMenuItemID(hPopupSaveCodepage, i);
        InsertMenuCommon(hMenuStack->hIconMenu, NULL, -1, 0, 0, hMenuStack->hSaveCodepagesSubMenu, hMenuStack->nSaveCodepagesFirstIndex + i, MF_BYPOSITION|dwState, nItemID, wszItem);
      }
      else break;
    }
    hMenuStack->nSaveCodepagesLastIndex=hMenuStack->nSaveCodepagesFirstIndex + i - 1;
  }
}

HMENU InsertMainMenu(POPUPMENU *hMenuStack)
{
  MAININDEXITEM *lpMainIndexItem;
  wchar_t wszMenuItem[MAX_PATH];
  HMENU hMenu=hMainMenu;
  int nMainMenuIndex;
  int nMainMenuCount;
  int nSubMenuIndex=0;
  int nState;

  if (hMenuStack->bClearMainMenu)
  {
    hMenu=CreateMenu();
  }

  for (lpMainIndexItem=hMenuStack->hMainMenuIndexStack.first; lpMainIndexItem; lpMainIndexItem=lpMainIndexItem->next)
  {
    NextItem:
    if ((nState=GetMenuState(hMenuStack->hPopupMenu, nSubMenuIndex, MF_BYPOSITION)) != -1)
    {
      if (!(LOBYTE(nState) & MF_POPUP))
      {
        ++nSubMenuIndex;
        goto NextItem;
      }
    }
    else break;

    if (lpMainIndexItem->hSubMenu=GetSubMenu(hMenuStack->hPopupMenu, nSubMenuIndex))
    {
      if (GetMenuStringWide(hMenuStack->hPopupMenu, nSubMenuIndex, wszMenuItem, MAX_PATH, MF_BYPOSITION))
      {
        nMainMenuIndex=lpMainIndexItem->nMainMenuIndex;

        if (nMainMenuIndex < 0)
        {
          nMainMenuCount=GetSubMenuCount(hMenu);

          if (nMainMenuCount + nMainMenuIndex >= 0)
            nMainMenuIndex=nMainMenuCount + nMainMenuIndex + 1;
          else
            nMainMenuIndex=0;
        }
        else
        {
          if ((nState=GetMenuState(hMenu, 0, MF_BYPOSITION)) != -1)
          {
            if (LOBYTE(nState) & MF_BITMAP)
              ++nMainMenuIndex;
          }
        }
        InsertMenuWide(hMenu, nMainMenuIndex, MF_BYPOSITION|MF_POPUP|(lpMainIndexItem->bMenuBreak?MF_MENUBARBREAK:0), (UINT_PTR)lpMainIndexItem->hSubMenu, wszMenuItem);
      }
    }
    ++nSubMenuIndex;
  }
  return hMenu;
}

void SetMainMenu(POPUPMENU *hMenuStack, HMENU hMenu, BOOL bDrawMenuBar)
{
  if (hMenuStack->bClearMainMenu)
  {
    if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
      SetMenu(hMainWnd, hMenu);
    else if (nMDI == WMD_MDI)
      SendMessage(hMdiClient, WM_MDISETMENU, (WPARAM)hMenu, (LPARAM)hMenuStack->hMdiDocumentsSubMenu);

    if (hMenuStack->hMainMenu)
    {
      DestroyMenu(hMenuStack->hMainMenu);
      hMenuStack->hMainMenu=NULL;
    }
    hMenuStack->hMainMenu=hMenu;
  }
  DrawMenuBar(hMainWnd);
}

void UnsetMainMenu(POPUPMENU *hMenuStack)
{
  if (hMenuStack->bClearMainMenu)
  {
    if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
    {
      SetMenu(hMainWnd, hMainMenu);
    }
    else if (nMDI == WMD_MDI)
    {
      SendMessage(hMdiClient, WM_MDISETMENU, (WPARAM)hMainMenu, (LPARAM)hMenuWindow);
      DrawMenuBar(hMainWnd);
    }
  }
  else
  {
    //Detach submenus from main menu
    MAININDEXITEM *lpMainIndexItem;
    int nIndex;

    for (lpMainIndexItem=hMenuStack->hMainMenuIndexStack.last; lpMainIndexItem; lpMainIndexItem=lpMainIndexItem->prev)
    {
      if ((nIndex=GetSubMenuIndex(hMainMenu, lpMainIndexItem->hSubMenu)) != -1)
        RemoveMenu(hMainMenu, nIndex, MF_BYPOSITION);
    }
    DrawMenuBar(hMainWnd);
  }
}

void ShowMainMenu(BOOL bShow)
{
  HMENU hMenu;

  if (bShow)
  {
    if (hMenuMainHide)
    {
      SetMenu(hMainWnd, hMenuMainHide);
      hMenuMainHide=NULL;
    }
  }
  else
  {
    if (hMenu=GetMenu(hMainWnd))
    {
      hMenuMainHide=hMenu;
      SetMenu(hMainWnd, NULL);
    }
  }
}

MENUITEM* GetContextMenuItem(POPUPMENU *hMenuStack, int nItem)
{
  if (nItem >= IDM_MIN_MANUAL && nItem <= IDM_MAX_MENU)
  {
    MENUITEM *lpElement;

    for (lpElement=hMenuStack->hMenuItemStack.first; lpElement; lpElement=lpElement->next)
    {
      if (lpElement->nItem > nItem)
        return NULL;
      if (lpElement->nItem == nItem)
        break;
    }
    return lpElement;
  }
  return NULL;
}

MENUITEM* GetCommandItem(POPUPMENU *hMenuStack, int nCommand)
{
  MENUITEM *lpElement;
  EXTPARAM *lpParameter;

  for (lpElement=hMenuStack->hMenuItemStack.first; lpElement; lpElement=lpElement->next)
  {
    if (lpElement->dwAction == EXTACT_COMMAND)
    {
      if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
      {
        if (lpParameter->nNumber == nCommand)
          break;
      }
    }
  }
  return lpElement;
}

void ViewItemCode(MENUITEM *lpElement)
{
  nExtMenuIndex=lpElement->nMenuType;
  crExtSetSel.cpMin=lpElement->nTextOffset;
  crExtSetSel.cpMax=-1;
  bExtFocusEdit=TRUE;

  if (!hThread)
    hThread=CreateThread(NULL, 0, ThreadProc, NULL, 0, &dwThreadId);
  else
    PostMessage(hWndMainDlg, AKDLL_MENUINDEX, nExtMenuIndex, 0);
}

void CallContextMenu(POPUPMENU *hMenuStack, int nItem)
{
  if (nItem >= IDM_MIN_EXPLORER && nItem <= IDM_MAX_EXPLORER)
  {
    CMINVOKECOMMANDINFO ici;

    xmemset(&ici, 0, sizeof(CMINVOKECOMMANDINFO));
    ici.cbSize=sizeof(CMINVOKECOMMANDINFO);
    ici.hwnd=hMainWnd;
    ici.lpVerb=MAKEINTRESOURCEA(nItem - IDM_MIN_EXPLORER);
    ici.nShow=SW_SHOWNORMAL;

    if (hMenuStack->pExplorerMenu) hMenuStack->pExplorerMenu->lpVtbl->InvokeCommand(hMenuStack->pExplorerMenu, &ici);
  }
  else if (nItem >= IDM_MIN_FAVOURITES && nItem <= IDM_MAX_FAVOURITES)
  {
    FAVITEM *lpFavItem;

    if (lpFavItem=StackGetFavouriteByIndex(&hFavStack, (nItem - IDM_MIN_FAVOURITES) + 1))
    {
      OPENDOCUMENTW od;

      if (TranslateFileString(lpFavItem->wszFile, wszBuffer, BUFFER_SIZE))
      {
        od.pFile=wszBuffer;
        od.pWorkDir=NULL;
        od.dwFlags=OD_ADT_BINARY_ERROR|OD_ADT_REG_CODEPAGE;
        od.nCodePage=0;
        od.bBOM=0;
        SendMessage(hMainWnd, AKD_OPENDOCUMENTW, (WPARAM)NULL, (LPARAM)&od);
      }
    }
  }
  else if (nItem >= IDM_MIN_MANUAL && nItem <= IDM_MAX_MENU)
  {
    MENUITEM *lpElement;
    EXTPARAM *lpParameter;

    if (nItem <= IDM_MAX_MANUAL)
      hMenuStack=&hMenuManualStack;

    if (lpElement=GetContextMenuItem(hMenuStack, nItem))
    {
      if (GetKeyState(VK_CONTROL) & 0x80)
      {
        ViewItemCode(lpElement);
        return;
      }

      if (lpElement->dwAction == EXTACT_COMMAND)
      {
        int nCommand=0;
        LPARAM lParam=0;

        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
          nCommand=(int)lpParameter->nNumber;
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 2))
          lParam=lpParameter->nNumber;

        if (nCommand < IDM_MIN_EXPLORER || nCommand > IDM_MAX_MENU)
        {
          SendMessage(hMainWnd, WM_COMMAND, nCommand, lParam);
        }
      }
      else if (lpElement->dwAction == EXTACT_CALL)
      {
        PLUGINCALLSENDW pcs;
        PLUGINCALLPOSTW *pcp;
        wchar_t *wpFunction=lpElement->hParamStack.first->wpString;
        int nStructSize;
        BOOL bCall=FALSE;

        if (lpElement->bAutoLoad)
        {
          pcs.pFunction=wpFunction;
          pcs.lParam=0;
          pcs.dwSupport=PDS_GETSUPPORT;

          if (SendMessage(hMainWnd, AKD_DLLCALLW, 0, (LPARAM)&pcs) != UD_FAILED)
          {
            if (pcs.dwSupport & PDS_NOAUTOLOAD)
            {
              xprintfW(wszBuffer, GetLangStringW(wLangModule, STRID_AUTOLOAD), wpFunction);
              MessageBoxW(hMainWnd, wszBuffer, wszPluginTitle, MB_OK|MB_ICONEXCLAMATION);
            }
            else bCall=TRUE;
          }
        }
        else bCall=TRUE;

        if (bCall)
        {
          ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hMenuStack->hPopupMenu, nItem, wszUrlLink);

          if (nStructSize=StructMethodParameters(&lpElement->hParamStack, NULL))
          {
            if (pcp=(PLUGINCALLPOSTW *)GlobalAlloc(GPTR, sizeof(PLUGINCALLPOSTW) + nStructSize))
            {
              xstrcpynW(pcp->szFunction, wpFunction, MAX_PATH);
              if (nStructSize > (INT_PTR)sizeof(INT_PTR))
              {
                pcp->lParam=(LPARAM)((unsigned char *)pcp + sizeof(PLUGINCALLPOSTW));
                StructMethodParameters(&lpElement->hParamStack, (unsigned char *)pcp->lParam);
              }
              else pcp->lParam=0;

              PostMessage(hMainWnd, AKD_DLLCALLW, (lpElement->bAutoLoad?DLLCF_SWITCHAUTOLOAD|DLLCF_SAVEONEXIT:0), (LPARAM)pcp);
            }
          }
        }
      }
      else if (lpElement->dwAction == EXTACT_EXEC)
      {
        STARTUPINFOW si;
        PROCESS_INFORMATION pi;
        wchar_t *wpCmdLine=NULL;
        wchar_t *wpWorkDir=NULL;
        BOOL bWait=FALSE;

        ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hMenuStack->hPopupMenu, nItem, wszUrlLink);
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
          wpCmdLine=lpParameter->wpExpanded;
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 2))
          wpWorkDir=lpParameter->wpExpanded;
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 3))
          bWait=(BOOL)lpParameter->nNumber;

        if (wpCmdLine)
        {
          xmemset(&si, 0, sizeof(STARTUPINFOW));
          si.cb=sizeof(STARTUPINFOW);
          if (CreateProcessWide(NULL, wpCmdLine, NULL, NULL, FALSE, 0, NULL, (wpWorkDir && *wpWorkDir)?wpWorkDir:NULL, &si, &pi))
          {
            if (bWait)
            {
              WaitForSingleObject(pi.hProcess, INFINITE);
              //GetExitCodeProcess(pi.hProcess, &dwExit);
            }
            CloseHandle(pi.hProcess);
            CloseHandle(pi.hThread);
          }
        }
      }
      else if (lpElement->dwAction == EXTACT_OPENFILE ||
               lpElement->dwAction == EXTACT_SAVEFILE)
      {
        wchar_t *wpFile=NULL;
        int nCodePage=-1;
        BOOL bBOM=-1;

        ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hMenuStack->hPopupMenu, nItem, wszUrlLink);
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
          wpFile=lpParameter->wpExpanded;
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 2))
          nCodePage=(int)lpParameter->nNumber;
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 3))
          bBOM=(BOOL)lpParameter->nNumber;

        if (lpElement->dwAction == EXTACT_OPENFILE)
        {
          OPENDOCUMENTW od;

          od.dwFlags=OD_ADT_BINARY_ERROR;
          if (nCodePage == -1) od.dwFlags|=OD_ADT_DETECT_CODEPAGE;
          if (bBOM == -1) od.dwFlags|=OD_ADT_DETECT_BOM;

          od.pFile=wpFile;
          od.pWorkDir=NULL;
          od.nCodePage=nCodePage;
          od.bBOM=bBOM;
          SendMessage(hMainWnd, AKD_OPENDOCUMENTW, (WPARAM)NULL, (LPARAM)&od);
        }
        else if (lpElement->dwAction == EXTACT_SAVEFILE)
        {
          EDITINFO ei;
          SAVEDOCUMENTW sd;

          if (SendMessage(hMainWnd, AKD_GETEDITINFO, (WPARAM)NULL, (LPARAM)&ei))
          {
            if (nCodePage == -1)
              nCodePage=ei.nCodePage;
            if (bBOM == -1)
              bBOM=ei.bBOM;
            sd.pFile=wpFile;
            sd.nCodePage=nCodePage;
            sd.bBOM=bBOM;
            sd.dwFlags=SD_UPDATE;
            SendMessage(hMainWnd, AKD_SAVEDOCUMENTW, (WPARAM)NULL, (LPARAM)&sd);
          }
        }
      }
      else if (lpElement->dwAction == EXTACT_FONT)
      {
        EDITINFO ei;
        LOGFONTW lfNew;
        HDC hDC;
        wchar_t *wpFaceName=NULL;
        DWORD dwFontStyle=0;
        int nPointSize=0;

        ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hMenuStack->hPopupMenu, nItem, wszUrlLink);
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
          wpFaceName=lpParameter->wpExpanded;
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 2))
          dwFontStyle=(DWORD)lpParameter->nNumber;
        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 3))
          nPointSize=(int)lpParameter->nNumber;

        if (SendMessage(hMainWnd, AKD_GETEDITINFO, (WPARAM)NULL, (LPARAM)&ei))
        {
          if (SendMessage(hMainWnd, AKD_GETFONTW, (WPARAM)ei.hWndEdit, (LPARAM)&lfNew))
          {
            if (nPointSize)
            {
              if (hDC=GetDC(ei.hWndEdit))
              {
                lfNew.lfHeight=-MulDiv(nPointSize, GetDeviceCaps(hDC, LOGPIXELSY), 72);
                ReleaseDC(ei.hWndEdit, hDC);
              }
            }
            if (dwFontStyle != FS_NONE)
            {
              lfNew.lfWeight=(dwFontStyle == FS_FONTBOLD || dwFontStyle == FS_FONTBOLDITALIC)?FW_BOLD:FW_NORMAL;
              lfNew.lfItalic=(dwFontStyle == FS_FONTITALIC || dwFontStyle == FS_FONTBOLDITALIC)?TRUE:FALSE;
            }
            if (*wpFaceName != L'\0')
            {
              xstrcpynW(lfNew.lfFaceName, wpFaceName, LF_FACESIZE);
            }
            SendMessage(hMainWnd, AKD_SETFONTW, (WPARAM)ei.hWndEdit, (LPARAM)&lfNew);
          }
        }
      }
      else if (lpElement->dwAction == EXTACT_RECODE)
      {
        EDITINFO ei;
        TEXTRECODE tr={0};

        if (SendMessage(hMainWnd, AKD_GETEDITINFO, (WPARAM)NULL, (LPARAM)&ei))
        {
          if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
            tr.nCodePageFrom=(int)lpParameter->nNumber;
          if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 2))
            tr.nCodePageTo=(int)lpParameter->nNumber;

          SendMessage(hMainWnd, AKD_RECODESEL, (WPARAM)ei.hWndEdit, (LPARAM)&tr);
        }
      }
      else if (lpElement->dwAction == EXTACT_INSERT)
      {
        EDITINFO ei;
        wchar_t *wpText=NULL;
        wchar_t *wpUnescText=NULL;
        int nTextLen=-1;
        int nUnescTextLen;
        DWORD dwCaret=(DWORD)-1;
        BOOL bEscSequences=FALSE;

        if (SendMessage(hMainWnd, AKD_GETEDITINFO, (WPARAM)NULL, (LPARAM)&ei))
        {
          if (!ei.bReadOnly)
          {
            ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hMenuStack->hPopupMenu, nItem, wszUrlLink);
            if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
              wpText=lpParameter->wpExpanded;
            if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 2))
              bEscSequences=(BOOL)lpParameter->nNumber;

            if (bEscSequences)
            {
              if (nUnescTextLen=(int)TranslateEscapeString(ei.hWndEdit, wpText, NULL, NULL))
              {
                if (wpUnescText=(wchar_t *)GlobalAlloc(GPTR, nUnescTextLen * sizeof(wchar_t)))
                {
                  nTextLen=(int)TranslateEscapeString(ei.hWndEdit, wpText, wpUnescText, &dwCaret);
                  wpText=wpUnescText;
                }
              }
            }
            if (wpText)
            {
              AEREPLACESELW rs;
              AECHARINDEX ciInsertPos;

              rs.pText=wpText;
              rs.dwTextLen=(UINT_PTR)nTextLen;
              rs.nNewLine=AELB_ASINPUT;
              rs.dwFlags=AEREPT_COLUMNASIS;
              rs.ciInsertStart=&ciInsertPos;
              rs.ciInsertEnd=NULL;
              SendMessage(ei.hWndEdit, AEM_REPLACESELW, 0, (LPARAM)&rs);

              if (dwCaret != (DWORD)-1)
              {
                AEINDEXOFFSET io;

                io.ciCharIn=&ciInsertPos;
                io.ciCharOut=&ciInsertPos;
                io.nOffset=dwCaret;
                io.nNewLine=AELB_ASINPUT;
                SendMessage(ei.hWndEdit, AEM_INDEXOFFSET, 0, (LPARAM)&io);
                SendMessage(ei.hWndEdit, AEM_EXSETSEL, (WPARAM)&ciInsertPos, (LPARAM)&ciInsertPos);
              }
            }
            if (wpUnescText) GlobalFree((HGLOBAL)wpUnescText);
          }
        }
      }
      else if (lpElement->dwAction == EXTACT_FAVOURITE)
      {
        FAVITEM *lpFavItem;
        int nCmd=0;

        if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
          nCmd=(int)lpParameter->nNumber;

        if (nCmd == FAV_ADDDIALOG ||
            nCmd == FAV_ADDSILENT)
        {
          INT_PTR nResult;

          if (lpFavItem=StackInsertFavourite(&hFavStack))
          {
            xstrcpynW(lpFavItem->wszName, GetFileName(wszCurrentFile, -1), MAX_PATH);
            xstrcpynW(lpFavItem->wszFile, wszCurrentFile, MAX_PATH);

            //Remove '=' from name to avoid save/read problems.
            xstrrepW(lpFavItem->wszName, -1, L"=", -1, L"", -1, TRUE, lpFavItem->wszName, NULL);

            if (nCmd == FAV_ADDDIALOG)
            {
              nResult=DialogBoxParamWide(hInstanceDLL, MAKEINTRESOURCEW(IDD_FAVEDIT), hMainWnd, (DLGPROC)FavEditDlgProc, (LPARAM)lpFavItem);

              if (nResult)
                dwSaveFlags|=OF_FAVTEXT;
              else
                StackDelete((stack **)&hFavStack.first, (stack **)&hFavStack.last, (stack *)lpFavItem);
            }
            else dwSaveFlags|=OF_FAVTEXT;
          }
        }
        else if (nCmd == FAV_MANAGE)
        {
          DialogBoxWide(hInstanceDLL, MAKEINTRESOURCEW(IDD_FAVLIST), hMainWnd, (DLGPROC)FavListDlgProc);
        }
        else if (nCmd == FAV_DELSILENT)
        {
          if (*wszCurrentFile)
          {
            if (lpFavItem=StackGetFavouriteByFile(&hFavStack, wszCurrentFile, NULL))
            {
              StackDelete((stack **)&hFavStack.first, (stack **)&hFavStack.last, (stack *)lpFavItem);
              dwSaveFlags|=OF_FAVTEXT;
            }
          }
        }
      }
    }
  }
  else SendMessage(hMainWnd, WM_COMMAND, nItem, 0);
}

void FreeContextMenu(POPUPMENU *hMenuStack)
{
  MENUITEM *lpMenuItem;
  SHOWSUBMENUITEM *lpShowSubmenuItem;

  if (hMenuStack->hPopupMenu)
  {
    //Detach submenus
    for (lpShowSubmenuItem=hMenuStack->hShowSubmenuStack.last; lpShowSubmenuItem; lpShowSubmenuItem=lpShowSubmenuItem->prev)
    {
      RemoveMenu(lpShowSubmenuItem->lpMenuItem->hSubMenu, lpShowSubmenuItem->lpMenuItem->nSubMenuIndex, MF_BYPOSITION);
    }
    StackClear((stack **)&hMenuStack->hShowSubmenuStack.first, (stack **)&hMenuStack->hShowSubmenuStack.last);

    //Destroy popup menu
    DestroyMenu(hMenuStack->hPopupMenu);
    hMenuStack->hPopupMenu=NULL;
  }

  for (lpMenuItem=hMenuStack->hMenuItemStack.first; lpMenuItem; lpMenuItem=lpMenuItem->next)
  {
    FreeMethodParameters(&lpMenuItem->hParamStack);
  }
  StackClear((stack **)&hMenuStack->hMenuItemStack.first, (stack **)&hMenuStack->hMenuItemStack.last);

  StackClear((stack **)&hMenuStack->hMainMenuIndexStack.first, (stack **)&hMenuStack->hMainMenuIndexStack.last);

  if (hMenuStack->hIconMenu)
  {
    IconMenu_Free(hMenuStack->hIconMenu, NULL);
    hMenuStack->hIconMenu=NULL;
  }
  if (hMenuStack->hImageList)
  {
    ImageList_Destroy(hMenuStack->hImageList);
    hMenuStack->hImageList=NULL;
    hMenuStack->nImageListIconCount=0;
  }
  if (hMenuStack->hMainMenu)
  {
    DestroyMenu(hMenuStack->hMainMenu);
    hMenuStack->hMainMenu=NULL;
  }
  if (hMenuStack->hExplorerSubMenu)
  {
    hMenuStack->hExplorerSubMenu=NULL;
  }
  if (hMenuStack->hFavouritesSubMenu)
  {
    hMenuStack->hFavouritesSubMenu=NULL;
  }
  if (hMenuStack->hRecentFilesSubMenu)
  {
    hMenuStack->hRecentFilesSubMenu=NULL;
  }
  if (hMenuStack->hLanguagesSubMenu)
  {
    hMenuStack->hLanguagesSubMenu=NULL;
  }
  if (hMenuStack->hOpenCodepagesSubMenu)
  {
    hMenuStack->hOpenCodepagesSubMenu=NULL;
  }
  if (hMenuStack->hSaveCodepagesSubMenu)
  {
    hMenuStack->hSaveCodepagesSubMenu=NULL;
  }
  if (hMenuStack->hMdiDocumentsSubMenu)
  {
    hMenuStack->hMdiDocumentsSubMenu=NULL;
  }
  if (hMenuStack->pExplorerSubMenu3)
  {
//    hMenuStack->pExplorerSubMenu3->Release();
    hMenuStack->pExplorerSubMenu3=NULL;
  }
  if (hMenuStack->pExplorerSubMenu2)
  {
//    hMenuStack->pExplorerSubMenu2->Release();
    hMenuStack->pExplorerSubMenu2=NULL;
  }
  if (hMenuStack->pExplorerMenu)
  {
//    hMenuStack->pExplorerMenu->Release();
    hMenuStack->pExplorerMenu=NULL;
  }
  hMenuStack->wszExplorerFile[0]=L'\0';
}

BOOL InsertMenuCommon(HICONMENU hIconMenu, HIMAGELIST hImageList, INT_PTR nIconIndex, int nIconWidth, int nIconHeight, HMENU hMenu, int uPosition, UINT uFlags, UINT_PTR uIDNewItem, const wchar_t *lpNewItem)
{
  if (!bUseIcons)
  {
    if (uPosition == -1)
      return AppendMenuWide(hMenu, uFlags, uIDNewItem, lpNewItem);
    else
      return InsertMenuWide(hMenu, uPosition, uFlags, uIDNewItem, lpNewItem);
  }
  else return IconMenu_AddItemW(hIconMenu, hImageList, nIconIndex, nIconWidth, nIconHeight, hMenu, uPosition, uFlags, uIDNewItem, lpNewItem);
}

BOOL ModifyMenuCommon(HICONMENU hIconMenu, HIMAGELIST hImageList, INT_PTR nIconIndex, int nIconWidth, int nIconHeight, HMENU hMenu, int uPosition, UINT uFlags, UINT_PTR uIDNewItem, const wchar_t *lpNewItem)
{
  if (!bUseIcons)
    return ModifyMenuWide(hMenu, uPosition, uFlags, uIDNewItem, lpNewItem);
  else
    return IconMenu_ModifyItemW(hIconMenu, hImageList, nIconIndex, nIconWidth, nIconHeight, hMenu, uPosition, uFlags, uIDNewItem, lpNewItem);
}

BOOL DeleteMenuCommon(HICONMENU hIconMenu, HMENU hMenu, int uPosition, UINT uFlags)
{
  return IconMenu_DelItem(hIconMenu, hMenu, uPosition, uFlags);
}

BOOL GetExplorerMenu(LPCONTEXTMENU *pContextMenu, LPCONTEXTMENU2 *pContextSubMenu2, LPCONTEXTMENU3 *pContextSubMenu3, HWND hWnd, wchar_t *wpFile)
{
  LPSHELLFOLDER pDesktopFolder=NULL;
  LPSHELLFOLDER pFolder=NULL;
  LPMALLOC pMalloc=NULL;
  LPITEMIDLIST pidlFirst;
  LPITEMIDLIST pidlNext;
  LPITEMIDLIST pidlLast;
  USHORT nTemp;
  BOOL bResult=FALSE;

  *pContextMenu=NULL;
  *pContextSubMenu2=NULL;
  *pContextSubMenu3=NULL;

  if (SHGetDesktopFolder(&pDesktopFolder) == NOERROR)
  {
    pDesktopFolder->lpVtbl->ParseDisplayName(pDesktopFolder, hWnd, NULL, wpFile, NULL, &pidlFirst, NULL);

    if (pidlFirst)
    {
      pidlNext=pidlLast=pidlFirst;

      for (;;)
      {
        pidlNext=NextPIDL(pidlNext);
        if (pidlNext->mkid.cb == 0) break;
        pidlLast=pidlNext;
      }

      nTemp=pidlLast->mkid.cb;
      pidlLast->mkid.cb=0;
      pDesktopFolder->lpVtbl->BindToObject(pDesktopFolder, pidlFirst, NULL, &IID_IShellFolder, (void **)&pFolder);
      pidlLast->mkid.cb=nTemp;

      if (pFolder)
      {
        pFolder->lpVtbl->GetUIObjectOf(pFolder, hWnd, 1, (LPCITEMIDLIST *)&pidlLast, &IID_IContextMenu, NULL, (void **)pContextMenu);

        if (pContextMenu)
        {
          (*pContextMenu)->lpVtbl->QueryInterface(*pContextMenu, &IID_IContextMenu2, (void **)pContextSubMenu2);
          (*pContextMenu)->lpVtbl->QueryInterface(*pContextMenu, &IID_IContextMenu3, (void **)pContextSubMenu3);
          bResult=TRUE;
        }
      }
      else *pContextMenu=NULL;
    }

    //Release
    if (SHGetMalloc(&pMalloc) == NOERROR)
    {
      pMalloc->lpVtbl->Free(pMalloc, pidlFirst);
      pMalloc->lpVtbl->Release(pMalloc);
    }
    if (pFolder) pFolder->lpVtbl->Release(pFolder);
    if (pDesktopFolder) pDesktopFolder->lpVtbl->Release(pDesktopFolder);
  }
  return bResult;
}

LPITEMIDLIST NextPIDL(LPCITEMIDLIST pidl)
{
  unsigned char *pMem=(unsigned char *)pidl;

  pMem+=pidl->mkid.cb;

  return (LPITEMIDLIST)pMem;
}

int GetSubMenuIndex(HMENU hMenu, HMENU hSubMenu)
{
  HMENU hSubMenuCount;
  int i;

  for (i=0; hSubMenuCount=GetSubMenu(hMenu, i); ++i)
  {
    if (hSubMenuCount == hSubMenu) return i;
  }
  return -1;
}

int GetSubMenuCount(HMENU hMenu)
{
  int i;

  for (i=0; GetSubMenu(hMenu, i); ++i);
  return i;
}

HMENU FindRootSubMenuByName(POPUPMENU *hMenuStack, const wchar_t *wpSubMenuName)
{
  ICONMENUSUBMENU *lpSubMenu;
  ICONMENUITEM *lpMenuItem;
  HMENU hMenu=NULL;

  if (wpSubMenuName && *wpSubMenuName)
  {
    if (lpSubMenu=IconMenu_GetMenuByHandle(hMenuStack->hIconMenu, hMenuStack->hPopupMenu))
    {
      for (lpMenuItem=lpSubMenu->first; lpMenuItem; lpMenuItem=lpMenuItem->next)
      {
        if (!xstrcmpiW(wpSubMenuName, lpMenuItem->wszStr))
        {
          if (lpMenuItem->uFlags & MF_POPUP)
          {
            hMenu=(HMENU)lpMenuItem->dwItemID;
            break;
          }
        }
      }
    }
    if (!hMenu)
    {
      xprintfW(wszBuffer, GetLangStringW(wLangModule, STRID_NOSUBMENUNAME), wpSubMenuName);
      MessageBoxW(hMainWnd, wszBuffer, wszPluginTitle, MB_OK|MB_ICONERROR);
    }
  }
  return hMenu;
}

void RemoveSeparator1(POPUPMENU *hMenuStack, HMENU hSubMenu)
{
  //Remove IDM_SEPARATOR1 item from last place in submenu
  ICONMENUSUBMENU *lpSubMenu;

  if (lpSubMenu=IconMenu_GetMenuByHandle(hMenuStack->hIconMenu, hSubMenu))
  {
    if (lpSubMenu->last && lpSubMenu->last->dwItemID == IDM_SEPARATOR1)
      DeleteMenuCommon(hMenuStack->hIconMenu, hSubMenu, lpSubMenu->nItemCount - 1, MF_BYPOSITION);
  }
}

void GetEditPos(HWND hWnd, POINT *pt, int nType)
{
  RECT rc;

  if (nType == GEP_CARET)
  {
    if (SendMessage(hWnd, AEM_GETCARETPOS, (WPARAM)pt, 0))
    {
      pt->y+=(int)SendMessage(hWnd, AEM_GETCHARSIZE, AECS_HEIGHT, 0);
      ClientToScreen(hWnd, pt);
      return;
    }
  }
  else if (nType == GEP_CURSOR)
  {
    //GetCursorPos called at the end
  }
  else
  {
    if (GetWindowRect(hWnd, &rc))
    {
      if (nType == GEP_LEFTTOP)
      {
        pt->x=rc.left;
        pt->y=rc.top;
        return;
      }
      else if (nType == GEP_RIGHTTOP)
      {
        pt->x=rc.right;
        pt->y=rc.top;
        return;
      }
      else if (nType == GEP_RIGHTBOTTOM)
      {
        pt->x=rc.right;
        pt->y=rc.bottom;
        return;
      }
      else if (nType == GEP_LEFTBOTTOM)
      {
        pt->x=rc.left;
        pt->y=rc.bottom;
        return;
      }
    }
  }
  GetCursorPos(pt);
}

void ShowStandardEditMenu(HWND hWnd, HMENU hMenu, BOOL bMouse)
{
  POINT pt;
  CHARRANGE64 cr;
  int nCmd;

  if (!bMouse && SendMessage(hWnd, AEM_GETCARETPOS, (WPARAM)&pt, 0))
  {
    pt.y+=(int)SendMessage(hWnd, AEM_GETCHARSIZE, AECS_HEIGHT, 0);
    ClientToScreen(hWnd, &pt);
  }
  else GetCursorPos(&pt);

  SendMessage(hWnd, EM_EXGETSEL64, 0, (LPARAM)&cr);
  EnableMenuItem(hMenu, IDM_EDIT_UNDO, SendMessage(hWnd, EM_CANUNDO, 0, 0)?MF_ENABLED:MF_GRAYED);
  EnableMenuItem(hMenu, IDM_EDIT_REDO, SendMessage(hWnd, EM_CANREDO, 0, 0)?MF_ENABLED:MF_GRAYED);
  EnableMenuItem(hMenu, IDM_EDIT_PASTE, SendMessage(hWnd, EM_CANPASTE, 0, 0)?MF_ENABLED:MF_GRAYED);
  EnableMenuItem(hMenu, IDM_EDIT_CUT, (cr.cpMin < cr.cpMax)?MF_ENABLED:MF_GRAYED);
  EnableMenuItem(hMenu, IDM_EDIT_CLEAR, (cr.cpMin < cr.cpMax)?MF_ENABLED:MF_GRAYED);
  EnableMenuItem(hMenu, IDM_EDIT_COPY, (cr.cpMin < cr.cpMax)?MF_ENABLED:MF_GRAYED);

  nCmd=TrackPopupMenu(hMenu, TPM_RETURNCMD|TPM_NONOTIFY|TPM_LEFTBUTTON|TPM_RIGHTBUTTON, pt.x, pt.y, 0, hWnd, NULL);

  if (nCmd == IDM_EDIT_UNDO)
    SendMessage(hWnd, EM_UNDO, 0, 0);
  else if (nCmd == IDM_EDIT_REDO)
    SendMessage(hWnd, EM_REDO, 0, 0);
  else if (nCmd == IDM_EDIT_CUT)
    SendMessage(hWnd, WM_CUT, 0, 0);
  else if (nCmd == IDM_EDIT_COPY)
    SendMessage(hWnd, WM_COPY, 0, 0);
  else if (nCmd == IDM_EDIT_PASTE)
    SendMessage(hWnd, WM_PASTE, 0, 0);
  else if (nCmd == IDM_EDIT_CLEAR)
    SendMessage(hWnd, WM_CLEAR, 0, 0);
  else if (nCmd == IDM_EDIT_SELECTALL)
    SendMessage(hWnd, EM_SETSEL, 0, -1);
}

void ParseMethodParameters(STACKEXTPARAM *hParamStack, const wchar_t *wpText, const wchar_t **wppText)
{
  EXTPARAM *lpParameter;
  const wchar_t *wpParamBegin=wpText;
  const wchar_t *wpParamEnd;
  wchar_t *wpString;
  wchar_t wchStopChar;
  int nStringLen;

  MethodParameter:
  while (*wpParamBegin == L' ' || *wpParamBegin == L'\t') ++wpParamBegin;

  if (*wpParamBegin == L'\"' || *wpParamBegin == L'\'' || *wpParamBegin == L'`')
  {
    //String
    wchStopChar=*wpParamBegin++;
    nStringLen=0;

    for (wpParamEnd=wpParamBegin; *wpParamEnd != wchStopChar && *wpParamEnd != L'\0'; ++wpParamEnd)
      ++nStringLen;

    if (!StackInsertIndex((stack **)&hParamStack->first, (stack **)&hParamStack->last, (stack **)&lpParameter, -1, sizeof(EXTPARAM)))
    {
      ++hParamStack->nElements;

      if (lpParameter->wpString=(wchar_t *)GlobalAlloc(GPTR, (nStringLen + 1) * sizeof(wchar_t)))
      {
        lpParameter->dwType=EXTPARAM_CHAR;
        wpString=lpParameter->wpString;

        for (wpParamEnd=wpParamBegin; *wpParamEnd != wchStopChar && *wpParamEnd != L'\0'; ++wpParamEnd)
          *wpString++=*wpParamEnd;
        *wpString=L'\0';

        if (bOldWindows)
        {
          nStringLen=WideCharToMultiByte(CP_ACP, 0, lpParameter->wpString, -1, NULL, 0, NULL, NULL);
          if (lpParameter->pString=(char *)GlobalAlloc(GPTR, nStringLen))
            WideCharToMultiByte(CP_ACP, 0, lpParameter->wpString, -1, lpParameter->pString, nStringLen, NULL, NULL);
        }
      }
    }
  }
  else
  {
    //Number
    for (wpParamEnd=wpParamBegin; *wpParamEnd != L',' && *wpParamEnd != L')' && *wpParamEnd != L'\0'; ++wpParamEnd);

    if (!StackInsertIndex((stack **)&hParamStack->first, (stack **)&hParamStack->last, (stack **)&lpParameter, -1, sizeof(EXTPARAM)))
    {
      ++hParamStack->nElements;

      lpParameter->dwType=EXTPARAM_INT;
      lpParameter->nNumber=xatoiW(wpParamBegin, NULL);
    }
  }

  while (*wpParamEnd != L',' && *wpParamEnd != L')' && *wpParamEnd != L'\0')
    ++wpParamEnd;
  if (*wpParamEnd == L',')
  {
    wpParamBegin=++wpParamEnd;
    goto MethodParameter;
  }
  if (*wpParamEnd == L')')
    ++wpParamEnd;
  if (wppText) *wppText=wpParamEnd;
}

void ExpandMethodParameters(STACKEXTPARAM *hParamStack, const wchar_t *wpFile, const wchar_t *wpExeDir, HMENU hPopupMenu, int nMenuItem, const wchar_t *wpUrlLink)
{
  //%f -file, %d -file directory, %a -AkelPad directory, %% -%
  //%m -menu handle, %i -menu item id, %u -url string
  EXTPARAM *lpParameter;
  wchar_t *wszSource;
  wchar_t *wpSource;
  wchar_t *wszTarget;
  wchar_t *wpTarget;
  int nStringLen;
  int nTargetLen=0;

  for (lpParameter=hParamStack->first; lpParameter; lpParameter=lpParameter->next)
  {
    if (lpParameter->dwType == EXTPARAM_CHAR)
    {
      //Expand environment strings
      nStringLen=ExpandEnvironmentStringsWide(lpParameter->wpString, NULL, 0);

      if (wszSource=(wchar_t *)GlobalAlloc(GPTR, nStringLen * sizeof(wchar_t)))
      {
        ExpandEnvironmentStringsWide(lpParameter->wpString, wszSource, nStringLen);

        //Expand plugin variables
        wszTarget=NULL;
        wpTarget=NULL;

        for (;;)
        {
          for (wpSource=wszSource; *wpSource;)
          {
            if (*wpSource == L'%')
            {
              if (*++wpSource == L'%')
              {
                ++wpSource;
                if (wszTarget) *wpTarget=L'%';
                ++wpTarget;
              }
              else if (*wpSource == L'f' || *wpSource == L'F')
              {
                ++wpSource;
                if (*wpFile) wpTarget+=xstrcpyW(wszTarget?wpTarget:NULL, wpFile) - !wszTarget;
              }
              else if (*wpSource == L'd' || *wpSource == L'D')
              {
                ++wpSource;
                if (*wpFile) wpTarget+=GetFileDir(wpFile, -1, wszTarget?wpTarget:NULL, MAX_PATH) - !wszTarget;
              }
              else if (*wpSource == L'a' || *wpSource == L'A')
              {
                ++wpSource;
                wpTarget+=xstrcpyW(wszTarget?wpTarget:NULL, wpExeDir) - !wszTarget;
              }
              else if (*wpSource == L'm' || *wpSource == L'M')
              {
                ++wpSource;
                wpTarget+=(int)xitoaW((INT_PTR)hPopupMenu, wszTarget?wpTarget:NULL) - !wszTarget;
              }
              else if (*wpSource == L'i' || *wpSource == L'I')
              {
                ++wpSource;
                wpTarget+=(int)xitoaW(nMenuItem, wszTarget?wpTarget:NULL) - !wszTarget;
              }
              else if (*wpSource == L'u' || *wpSource == L'U')
              {
                ++wpSource;
                wpTarget+=xstrcpyW(wszTarget?wpTarget:NULL, wpUrlLink) - !wszTarget;
              }
            }
            else
            {
              if (wszTarget) *wpTarget=*wpSource;
              ++wpTarget;
              ++wpSource;
            }
          }

          if (!wszTarget)
          {
            //Allocate wszTarget and loop again
            if (wszTarget=(wchar_t *)GlobalAlloc(GPTR, (INT_PTR)(wpTarget + 1)))
            {
              nTargetLen=(int)((INT_PTR)wpTarget / sizeof(wchar_t));
              wpTarget=wszTarget;
            }
            else break;
          }
          else
          {
            *wpTarget=L'\0';
            break;
          }
        }

        if (wszTarget)
        {
          if (lpParameter->pExpanded)
          {
            GlobalFree((HGLOBAL)lpParameter->pExpanded);
            lpParameter->pExpanded=NULL;
          }
          if (lpParameter->wpExpanded)
          {
            GlobalFree((HGLOBAL)lpParameter->wpExpanded);
            lpParameter->wpExpanded=NULL;
          }
          lpParameter->wpExpanded=wszTarget;
          lpParameter->nExpandedUnicodeLen=nTargetLen;

          if (bOldWindows)
          {
            lpParameter->nExpandedAnsiLen=WideCharToMultiByte(CP_ACP, 0, lpParameter->wpExpanded, -1, NULL, 0, NULL, NULL);
            if (lpParameter->pExpanded=(char *)GlobalAlloc(GPTR, lpParameter->nExpandedAnsiLen))
              WideCharToMultiByte(CP_ACP, 0, lpParameter->wpExpanded, -1, lpParameter->pExpanded, lpParameter->nExpandedAnsiLen, NULL, NULL);
            if (lpParameter->nExpandedAnsiLen) --lpParameter->nExpandedAnsiLen;
          }
        }
        GlobalFree((HGLOBAL)wszSource);
      }
    }
  }
}

int StructMethodParameters(STACKEXTPARAM *hParamStack, unsigned char *lpStruct)
{
  EXTPARAM *lpParameter;
  int nElementOffset;
  int nStringOffset=0;

  if (hParamStack->nElements)
  {
    //nStringOffset is pointer to memory where first string will be copied
    nElementOffset=0;
    nStringOffset=hParamStack->nElements * sizeof(INT_PTR);

    //First element in structure is the size of the call parameters structure
    if (lpStruct) *((INT_PTR *)(lpStruct + nElementOffset))=nStringOffset;
    nElementOffset+=sizeof(INT_PTR);

    //Skip hParamStack->first equal to "Plugin::Function".
    for (lpParameter=hParamStack->first->next; lpParameter; lpParameter=lpParameter->next)
    {
      if (lpParameter->dwType == EXTPARAM_CHAR)
      {
        //Strings located after call parameters structure
        if (lpStruct) *((INT_PTR *)(lpStruct + nElementOffset))=(INT_PTR)(lpStruct + nStringOffset);

        if (bOldWindows)
        {
          if (lpStruct) xmemcpy(lpStruct + nStringOffset, lpParameter->pExpanded, lpParameter->nExpandedAnsiLen + 1);
          nStringOffset+=(lpParameter->nExpandedAnsiLen + 1);
        }
        else
        {
          if (lpStruct) xmemcpy(lpStruct + nStringOffset, lpParameter->wpExpanded, (lpParameter->nExpandedUnicodeLen + 1) * sizeof(wchar_t));
          nStringOffset+=(lpParameter->nExpandedUnicodeLen + 1) * sizeof(wchar_t);
        }
      }
      else
      {
        if (lpStruct) *((INT_PTR *)(lpStruct + nElementOffset))=lpParameter->nNumber;
      }
      nElementOffset+=sizeof(INT_PTR);
    }
  }
  return nStringOffset;
}

EXTPARAM* GetMethodParameter(STACKEXTPARAM *hParamStack, int nIndex)
{
  EXTPARAM *lpParameter;

  if (!StackGetElement((stack *)hParamStack->first, (stack *)hParamStack->last, (stack **)&lpParameter, nIndex))
    return lpParameter;
  return NULL;
}

void GetIconParameters(const wchar_t *wpText, wchar_t *wszIconFile, int nMaxIconFile, int *nIconIndex, const wchar_t **wppText)
{
  wchar_t wchStopChar;
  int i;

  wszIconFile[0]=L'\0';
  *nIconIndex=0;

  //File
  while (*wpText == L' ' || *wpText == L'\t') ++wpText;

  if (*wpText == L'\"' || *wpText == L'\'' || *wpText == L'`')
  {
    wchStopChar=*wpText++;

    for (i=0; i < nMaxIconFile && *wpText != wchStopChar && *wpText != L'\0'; ++i, ++wpText)
    {
      wszIconFile[i]=*wpText;
    }
    wszIconFile[i]=L'\0';

    while (*wpText != L',' && *wpText != L')' && *wpText != L'\0')
      ++wpText;
    if (*wpText == L',')
      ++wpText;
    else
      goto End;
  }

  //Index
  while (*wpText == L' ' || *wpText == L'\t') ++wpText;

  *nIconIndex=(int)xatoiW(wpText, NULL);

  while (*wpText != L',' && *wpText != L')' && *wpText != L'\0')
    ++wpText;
  if (*wpText == L',')
    ++wpText;
  else
    goto End;

  //End
  End:
  if (*wpText == L')')
    ++wpText;
  *wppText=wpText;
}

void FreeMethodParameters(STACKEXTPARAM *hParamStack)
{
  EXTPARAM *lpParameter;

  for (lpParameter=hParamStack->first; lpParameter; lpParameter=lpParameter->next)
  {
    if (lpParameter->dwType == EXTPARAM_CHAR)
    {
      if (lpParameter->pString) GlobalFree((HGLOBAL)lpParameter->pString);
      if (lpParameter->wpString) GlobalFree((HGLOBAL)lpParameter->wpString);
      if (lpParameter->pExpanded) GlobalFree((HGLOBAL)lpParameter->pExpanded);
      if (lpParameter->wpExpanded) GlobalFree((HGLOBAL)lpParameter->wpExpanded);
    }
  }
  StackClear((stack **)&hParamStack->first, (stack **)&hParamStack->last);
  hParamStack->nElements=0;
}

int GetMethodName(const wchar_t *wpText, wchar_t *wszStr, int nStrLen, const wchar_t **wppText)
{
  int i=0;

  while (*wpText == L'\"' || *wpText == L' ' || *wpText == L'\t') ++wpText;

  while (*wpText != L'(' && *wpText != L'\r' && *wpText != L'\0')
  {
    if (i < nStrLen) wszStr[i++]=*wpText;
    ++wpText;
  }
  wszStr[i]=L'\0';
  if (*wpText != L'\r' && *wpText != L'\0') ++wpText;
  if (wppText) *wppText=wpText;
  return i;
}

int NextString(const wchar_t *wpText, wchar_t *wszStr, int nStrLen, const wchar_t **wppText, int *nMinus)
{
  int i;

  while (*wpText == L' ' || *wpText == L'\t' || *wpText == L'\r' || *wpText == L'\n') ++wpText;

  if (*wpText == L'-')
  {
    if (nMinus) *nMinus=1;
    ++wpText;
  }
  else
  {
    if (nMinus) *nMinus=0;
  }

  if (*wpText == L'\"')
  {
    ++wpText;
    i=0;

    while (*wpText != L'\"' && *wpText != L'\0')
    {
      if (i < nStrLen) wszStr[i++]=*wpText;
      ++wpText;
    }
  }
  else
  {
    i=0;

    while (*wpText != L' ' && *wpText != L'\r' && *wpText != L'\0')
    {
      if (i < nStrLen) wszStr[i++]=*wpText;
      ++wpText;
    }
  }
  wszStr[i]=L'\0';
  if (*wpText != L'\r' && *wpText != L'\0') ++wpText;
  if (wppText) *wppText=wpText;
  return i;
}

BOOL SkipComment(const wchar_t **wpText)
{
  for (;;)
  {
    while (**wpText == L' ' || **wpText == L'\t' || **wpText == L'\r' || **wpText == L'\n') ++*wpText;

    if (**wpText == L';' || **wpText == L'#')
    {
      while (**wpText != L'\r' && **wpText != L'\0') ++*wpText;
    }
    else break;
  }
  if (**wpText == L'\0')
    return FALSE;
  return TRUE;
}

int GetFileDir(const wchar_t *wpFile, int nFileLen, wchar_t *wszFileDir, DWORD dwFileDirLen)
{
  const wchar_t *wpCount;

  if (nFileLen == -1) nFileLen=(int)xstrlenW(wpFile);
  if (wszFileDir) wszFileDir[0]=L'\0';

  for (wpCount=wpFile + nFileLen - 1; wpCount >= wpFile; --wpCount)
  {
    if (*wpCount == L'\\')
      return (int)xstrcpynW(wszFileDir, wpFile, min(dwFileDirLen, (DWORD)(wpCount - wpFile) + 1));
  }
  return 0;
}

INT_PTR TranslateEscapeString(HWND hWndEdit, const wchar_t *wpInput, wchar_t *wszOutput, DWORD *lpdwCaret)
{
  EDITINFO ei;
  const wchar_t *a=wpInput;
  wchar_t *b=wszOutput;
  wchar_t whex[5];
  int nDec;

  if (SendMessage(hMainWnd, AKD_GETEDITINFO, (WPARAM)hWndEdit, (LPARAM)&ei))
  {
    for (whex[4]=L'\0'; *a; ++a)
    {
      if (*a == L'\\')
      {
        if (*++a == L'\\')
        {
          if (wszOutput) *b=L'\\';
          ++b;
        }
        else if (*a == L'n')
        {
          if (ei.nNewLine == NEWLINE_MAC)
          {
            if (wszOutput) *b=L'\r';
            ++b;
          }
          else if (ei.nNewLine == NEWLINE_UNIX)
          {
            if (wszOutput) *b=L'\n';
            ++b;
          }
          else if (ei.nNewLine == NEWLINE_WIN)
          {
            if (wszOutput) *b=L'\r';
            ++b;
            if (wszOutput) *b=L'\n';
            ++b;
          }
        }
        else if (*a == L't')
        {
          if (wszOutput) *b=L'\t';
          ++b;
        }
        else if (*a == L'[')
        {
          while (*++a == L' ');

          do
          {
            whex[0]=*a;
            if (!*a) goto Error;
            whex[1]=*++a;
            if (!*a) goto Error;
            whex[2]=*++a;
            if (!*a) goto Error;
            whex[3]=*++a;
            if (!*a) goto Error;
            nDec=(int)hex2decW(whex, 4);
            if (nDec == -1) goto Error;
            while (*++a == L' ');

            if (wszOutput) *b=(wchar_t)nDec;
            ++b;
          }
          while (*a && *a != L']');

          if (!*a) goto Error;
        }
        else if (*a == L's')
        {
          CHARRANGE64 cr;

          SendMessage(hWndEdit, EM_EXGETSEL64, 0, (WPARAM)&cr);
          if (cr.cpMin != cr.cpMax)
          {
            if (wszOutput)
            {
              wchar_t *wpText;
              INT_PTR nTextLen=0;

              if (wpText=(wchar_t *)SendMessage(hMainWnd, AKD_GETSELTEXTW, (WPARAM)hWndEdit, (LPARAM)&nTextLen))
              {
                xmemcpy(b, wpText, nTextLen * sizeof(wchar_t));
                b+=nTextLen;
                SendMessage(hMainWnd, AKD_FREETEXT, 0, (LPARAM)wpText);
              }
            }
            else b+=cr.cpMax - cr.cpMin;
          }
        }
        else if (*a == L'|')
        {
          if (lpdwCaret) *lpdwCaret=(DWORD)(b - wszOutput);
        }
        else goto Error;
      }
      else
      {
        if (wszOutput) *b=*a;
        ++b;
      }
    }
    if (wszOutput)
      *b=L'\0';
    else
      ++b;
    return (b - wszOutput);
  }

  Error:
  if (wszOutput) *wszOutput=L'\0';
  return 0;
}

int TranslateFileString(const wchar_t *wpString, wchar_t *wszBuffer, int nBufferSize)
{
  //%a -AkelPad directory, %% -%
  wchar_t *wpExeDir=wszExeDir;
  wchar_t *wszSource;
  wchar_t *wpSource;
  wchar_t *wpTarget=wszBuffer;
  wchar_t *wpTargetMax=wszBuffer + (wszBuffer?nBufferSize:0x7FFFFFFF);
  int nStringLen;
  BOOL bStringStart=TRUE;

  //Expand environment strings
  nStringLen=ExpandEnvironmentStringsWide(wpString, NULL, 0);

  if (wszSource=(wchar_t *)GlobalAlloc(GPTR, nStringLen * sizeof(wchar_t)))
  {
    ExpandEnvironmentStringsWide(wpString, wszSource, nStringLen);

    //Expand plugin variables
    for (wpSource=wszSource; *wpSource && wpTarget < wpTargetMax;)
    {
      if (bStringStart && *wpSource == L'%')
      {
        if (*++wpSource == L'%')
        {
          ++wpSource;
          if (wszBuffer) *wpTarget=L'%';
          ++wpTarget;
        }
        else if (*wpSource == L'a' || *wpSource == L'A')
        {
          ++wpSource;
          wpTarget+=xstrcpynW(wszBuffer?wpTarget:NULL, wpExeDir, wpTargetMax - wpTarget) - !wszBuffer;
        }
      }
      else
      {
        if (*wpSource != L'\"' && *wpSource != L'\'' && *wpSource != L'`')
          bStringStart=FALSE;
        if (wszBuffer) *wpTarget=*wpSource;
        ++wpTarget;
        ++wpSource;
      }
    }
    if (wpTarget < wpTargetMax)
    {
      if (wszBuffer)
        *wpTarget=L'\0';
      else
        ++wpTarget;
    }
    GlobalFree((HGLOBAL)wszSource);
  }
  return (int)(wpTarget - wszBuffer);
}

int GetCurFile(wchar_t *wszFile, int nMaxFile)
{
  EDITINFO ei;

  if (SendMessage(hMainWnd, AKD_GETEDITINFO, (WPARAM)NULL, (LPARAM)&ei))
  {
    return (int)xstrcpynW(wszFile, ei.wszFile, nMaxFile) + 1;
  }
  wszFile[0]=L'\0';
  return 0;
}

const wchar_t* GetFileName(const wchar_t *wpFile, int nFileLen)
{
  const wchar_t *wpCount;

  if (nFileLen == -1) nFileLen=(int)xstrlenW(wpFile);

  for (wpCount=wpFile + nFileLen - 1; wpCount >= wpFile; --wpCount)
  {
    if (*wpCount == L'\\')
      return wpCount + 1;
  }
  return wpFile;
}

INT_PTR GetEditText(HWND hWnd, wchar_t **wpText)
{
  GETTEXTRANGE gtr;
  INT_PTR nTextLen;

  gtr.cpMin=0;
  gtr.cpMax=-1;

  if (nTextLen=SendMessage(hMainWnd, AKD_GETTEXTRANGEW, (WPARAM)hWnd, (LPARAM)&gtr))
    *wpText=(wchar_t *)gtr.pText;
  else
    *wpText=(wchar_t *)HeapAlloc(hHeap, HEAP_ZERO_MEMORY, sizeof(wchar_t));
  return nTextLen;
}

BOOL GetWindowPos(HWND hWnd, HWND hWndOwner, RECT *rc)
{
  if (GetWindowRect(hWnd, rc))
  {
    rc->right-=rc->left;
    rc->bottom-=rc->top;

    if (hWndOwner)
    {
      if (!ScreenToClient(hWndOwner, (POINT *)&rc->left))
        return FALSE;
    }
    return TRUE;
  }
  return FALSE;
}

DWORD ScrollCaret(HWND hWnd)
{
  AESCROLLTOPOINT stp;
  DWORD dwScrollResult;

  //Test scroll to caret
  stp.dwFlags=AESC_TEST|AESC_POINTCARET|AESC_OFFSETCHARX|AESC_OFFSETCHARY;
  stp.nOffsetX=1;
  stp.nOffsetY=0;
  dwScrollResult=(DWORD)SendMessage(hWnd, AEM_SCROLLTOPOINT, 0, (LPARAM)&stp);

  //Scroll to caret
  stp.dwFlags=AESC_POINTCARET;
  stp.nOffsetX=3;
  stp.nOffsetY=2;
  if (dwScrollResult & AECSE_SCROLLEDX)
    stp.dwFlags|=AESC_OFFSETRECTDIVX;
  if (dwScrollResult & AECSE_SCROLLEDY)
    stp.dwFlags|=AESC_OFFSETRECTDIVY;
  return (DWORD)SendMessage(hWnd, AEM_SCROLLTOPOINT, 0, (LPARAM)&stp);
}


//// Options

INT_PTR WideOption(HANDLE hOptions, const wchar_t *pOptionName, DWORD dwType, BYTE *lpData, DWORD dwData)
{
  PLUGINOPTIONW po;

  po.pOptionName=pOptionName;
  po.dwType=dwType;
  po.lpData=lpData;
  po.dwData=dwData;
  return SendMessage(hMainWnd, AKD_OPTIONW, (WPARAM)hOptions, (LPARAM)&po);
}

void ReadOptions(DWORD dwFlags)
{
  HANDLE hOptions;
  DWORD dwSize;

  if (hOptions=(HANDLE)SendMessage(hMainWnd, AKD_BEGINOPTIONSW, POB_READ, (LPARAM)wszPluginName))
  {
    //Manual menu
    if (dwSize=(DWORD)WideOption(hOptions, L"ManualMenuText", PO_BINARY, NULL, 0))
    {
      if (wszManualText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize))
      {
        WideOption(hOptions, L"ManualMenuText", PO_BINARY, (LPBYTE)wszManualText, dwSize);
      }
    }

    //Main menu
    if (dwSize=(DWORD)WideOption(hOptions, L"MainMenuText", PO_BINARY, NULL, 0))
    {
      if (wszMainText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize))
      {
        WideOption(hOptions, L"MainMenuText", PO_BINARY, (LPBYTE)wszMainText, dwSize);
      }
    }
    WideOption(hOptions, L"MainMenuEnable", PO_DWORD, (LPBYTE)&bMenuMainEnable, sizeof(DWORD));
    WideOption(hOptions, L"MainMenuHide", PO_DWORD, (LPBYTE)&bMenuMainHide, sizeof(DWORD));

    //Edit menu
    if (dwSize=(DWORD)WideOption(hOptions, L"EditMenuText", PO_BINARY, NULL, 0))
    {
      if (wszEditText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize))
      {
        WideOption(hOptions, L"EditMenuText", PO_BINARY, (LPBYTE)wszEditText, dwSize);
      }
    }
    WideOption(hOptions, L"EditMenuEnable", PO_DWORD, (LPBYTE)&bMenuEditEnable, sizeof(DWORD));

    //Tab menu
    if (dwSize=(DWORD)WideOption(hOptions, L"TabMenuText", PO_BINARY, NULL, 0))
    {
      if (wszTabText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize))
      {
        WideOption(hOptions, L"TabMenuText", PO_BINARY, (LPBYTE)wszTabText, dwSize);
      }
    }
    WideOption(hOptions, L"TabMenuEnable", PO_DWORD, (LPBYTE)&bMenuTabEnable, sizeof(DWORD));

    //Link menu
    if (dwSize=(DWORD)WideOption(hOptions, L"UrlMenuText", PO_BINARY, NULL, 0))
    {
      if (wszUrlText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize))
      {
        WideOption(hOptions, L"UrlMenuText", PO_BINARY, (LPBYTE)wszUrlText, dwSize);
      }
    }
    WideOption(hOptions, L"UrlMenuEnable", PO_DWORD, (LPBYTE)&bMenuUrlEnable, sizeof(DWORD));

    //Recent files menu
    if (dwSize=(DWORD)WideOption(hOptions, L"RecentFilesMenuText", PO_BINARY, NULL, 0))
    {
      if (wszRecentFilesText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize))
      {
        WideOption(hOptions, L"RecentFilesMenuText", PO_BINARY, (LPBYTE)wszRecentFilesText, dwSize);
      }
    }
    WideOption(hOptions, L"RecentFilesMenuEnable", PO_DWORD, (LPBYTE)&bMenuRecentFilesEnable, sizeof(DWORD));

    //Favourites
    if (dwSize=(DWORD)WideOption(hOptions, L"FavText", PO_BINARY, NULL, 0))
    {
      if (wszFavText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize))
      {
        WideOption(hOptions, L"FavText", PO_BINARY, (LPBYTE)wszFavText, dwSize);
      }
    }
    WideOption(hOptions, L"FavShowFile", PO_DWORD, (LPBYTE)&bFavShowFile, sizeof(DWORD));

    //Window rectangle
    WideOption(hOptions, L"WindowRect", PO_BINARY, (LPBYTE)&rcMainCurrentDialog, sizeof(RECT));
    WideOption(hOptions, L"FavRect", PO_BINARY, (LPBYTE)&rcFavCurrentDialog, sizeof(RECT));

    SendMessage(hMainWnd, AKD_ENDOPTIONS, (WPARAM)hOptions, 0);
  }

  //Default menus
  if (!wszManualText) wszManualText=GetDefaultMenu(STRID_DEFAULTMANUAL);
  if (!wszMainText) wszMainText=GetDefaultMenu(STRID_DEFAULTMAIN);
  if (!wszEditText) wszEditText=GetDefaultMenu(STRID_DEFAULTEDIT);
  if (!wszTabText) wszTabText=GetDefaultMenu(STRID_DEFAULTTAB);
  if (!wszUrlText) wszUrlText=GetDefaultMenu(STRID_DEFAULTURL);
  if (!wszRecentFilesText) wszRecentFilesText=GetDefaultMenu(STRID_DEFAULTRECENTFILES);
}

void SaveOptions(DWORD dwFlags)
{
  HANDLE hOptions;

  if (hOptions=(HANDLE)SendMessage(hMainWnd, AKD_BEGINOPTIONSW, POB_SAVE, (LPARAM)wszPluginName))
  {
    if (dwFlags & OF_MENUSTEXT)
    {
      WideOption(hOptions, L"ManualMenuText", PO_BINARY, (LPBYTE)wszManualText, ((int)xstrlenW(wszManualText) + 1) * sizeof(wchar_t));
      WideOption(hOptions, L"MainMenuEnable", PO_DWORD, (LPBYTE)&bMenuMainEnable, sizeof(DWORD));
      WideOption(hOptions, L"MainMenuHide", PO_DWORD, (LPBYTE)&bMenuMainHide, sizeof(DWORD));
      WideOption(hOptions, L"MainMenuText", PO_BINARY, (LPBYTE)wszMainText, ((int)xstrlenW(wszMainText) + 1) * sizeof(wchar_t));
      WideOption(hOptions, L"EditMenuEnable", PO_DWORD, (LPBYTE)&bMenuEditEnable, sizeof(DWORD));
      WideOption(hOptions, L"EditMenuText", PO_BINARY, (LPBYTE)wszEditText, ((int)xstrlenW(wszEditText) + 1) * sizeof(wchar_t));
      WideOption(hOptions, L"TabMenuEnable", PO_DWORD, (LPBYTE)&bMenuTabEnable, sizeof(DWORD));
      WideOption(hOptions, L"TabMenuText", PO_BINARY, (LPBYTE)wszTabText, ((int)xstrlenW(wszTabText) + 1) * sizeof(wchar_t));
      WideOption(hOptions, L"UrlMenuEnable", PO_DWORD, (LPBYTE)&bMenuUrlEnable, sizeof(DWORD));
      WideOption(hOptions, L"UrlMenuText", PO_BINARY, (LPBYTE)wszUrlText, ((int)xstrlenW(wszUrlText) + 1) * sizeof(wchar_t));
      WideOption(hOptions, L"RecentFilesMenuEnable", PO_DWORD, (LPBYTE)&bMenuRecentFilesEnable, sizeof(DWORD));
      WideOption(hOptions, L"RecentFilesMenuText", PO_BINARY, (LPBYTE)wszRecentFilesText, ((int)xstrlenW(wszRecentFilesText) + 1) * sizeof(wchar_t));
    }
    if (dwFlags & OF_MAINRECT)
    {
      WideOption(hOptions, L"WindowRect", PO_BINARY, (LPBYTE)&rcMainCurrentDialog, sizeof(RECT));
    }
    if (dwFlags & OF_FAVTEXT)
    {
      FAVITEM *lpFavItem;
      DWORD dwSize;

      for (dwSize=0, lpFavItem=hFavStack.first; lpFavItem; lpFavItem=lpFavItem->next)
      {
        dwSize+=(DWORD)xprintfW(NULL, L"%s=%s\r", lpFavItem->wszName, lpFavItem->wszFile) - 1;
      }
      if (wszFavText=(wchar_t *)HeapAlloc(hHeap, 0, (dwSize + 1) * sizeof(wchar_t)))
      {
        for (dwSize=0, lpFavItem=hFavStack.first; lpFavItem; lpFavItem=lpFavItem->next)
        {
          dwSize+=(DWORD)xprintfW(wszFavText + dwSize, L"%s=%s\r", lpFavItem->wszName, lpFavItem->wszFile);
        }
        WideOption(hOptions, L"FavText", PO_BINARY, (LPBYTE)wszFavText, ((int)xstrlenW(wszFavText) + 1) * sizeof(wchar_t));

        HeapFree(hHeap, 0, wszFavText);
        wszFavText=NULL;
      }
      WideOption(hOptions, L"FavShowFile", PO_DWORD, (LPBYTE)&bFavShowFile, sizeof(DWORD));
    }
    if (dwFlags & OF_FAVRECT)
    {
      WideOption(hOptions, L"FavRect", PO_BINARY, (LPBYTE)&rcFavCurrentDialog, sizeof(RECT));
    }

    SendMessage(hMainWnd, AKD_ENDOPTIONS, (WPARAM)hOptions, 0);
  }
}

wchar_t* GetDefaultMenu(int nStringID)
{
  wchar_t *wszMenuText;
  DWORD dwSize;

  dwSize=(DWORD)xprintfW(NULL, L"%s", GetLangStringW(wLangModule, nStringID));
  if (wszMenuText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize * sizeof(wchar_t)))
    xprintfW(wszMenuText, L"%s", GetLangStringW(wLangModule, nStringID));
  return wszMenuText;
}

const char* GetLangStringA(LANGID wLangID, int nStringID)
{
  static char szStringBuf[MAX_PATH];

  WideCharToMultiByte(CP_ACP, 0, GetLangStringW(wLangID, nStringID), -1, szStringBuf, MAX_PATH, NULL, NULL);
  return szStringBuf;
}

const wchar_t* GetLangStringW(LANGID wLangID, int nStringID)
{
  if (wLangID == LANG_RUSSIAN)
  {
    if (nStringID == STRID_MANUAL)
      return L"\x041C\x0435\x043D\x044E\x0020\x0043\x006F\x006E\x0074\x0065\x0078\x0074\x004D\x0065\x006E\x0075\x003A\x003A\x0053\x0068\x006F\x0077";
    if (nStringID == STRID_MENUMAIN)
      return L"\x0413\x043B\x0430\x0432\x043D\x043E\x0435\x0020\x043C\x0435\x043D\x044E";
    if (nStringID == STRID_MENUEDIT)
      return L"\x041C\x0435\x043D\x044E\x0020\x043E\x043A\x043D\x0430\x0020\x0440\x0435\x0434\x0430\x043A\x0442\x0438\x0440\x043E\x0432\x0430\x043D\x0438\x044F";
    if (nStringID == STRID_MENUTAB)
      return L"\x041C\x0435\x043D\x044E\x0020\x0432\x043A\x043B\x0430\x0434\x043E\x043A";
    if (nStringID == STRID_MENUURL)
      return L"\x041C\x0435\x043D\x044E\x0020\x0441\x0441\x044B\x043B\x043E\x043A";
    if (nStringID == STRID_MENURECENTFILES)
      return L"\x041C\x0435\x043D\x044E\x0020\x043F\x043E\x0441\x043B\x0435\x0434\x043D\x0438\x0445\x0020\x0444\x0430\x0439\x043B\x043E\x0432";
    if (nStringID == STRID_ENABLE)
      return L"\x0412\x043A\x043B\x044E\x0447\x0435\x043D\x043E";
    if (nStringID == STRID_HIDE)
      return L"\x0421\x043A\x0440\x044B\x0432\x0430\x0442\x044C";
    if (nStringID == STRID_SHOW)
      return L"\x041F\x043E\x043A\x0430\x0437\x0430\x0442\x044C";
    if (nStringID == STRID_AUTOLOAD)
      return L"\x0022\x0025\x0073\x0022\x0020\x043D\x0435\x0020\x043F\x043E\x0434\x0434\x0435\x0440\x0436\x0438\x0432\x0430\x0435\x0442\x0020\x0430\x0432\x0442\x043E\x0437\x0430\x0433\x0440\x0443\x0437\x043A\x0443\x0020\x0028\x0022\x002B\x0022\x0020\x043F\x0435\x0440\x0435\x0434\x0020\x0022\x0043\x0061\x006C\x006C\x0022\x0029\x002E";
    if (nStringID == STRID_NOSUBMENUINDEX)
      return L"ContextMenu::Show - \x043F\x043E\x0434\x043C\x0435\x043D\x044E\x0020\x0441\x0020\x0438\x043D\x0434\x0435\x043A\x0441\x043E\x043C %d \x043D\x0435\x0020\x043D\x0430\x0439\x0434\x0435\x043D\x043E.";
    if (nStringID == STRID_NOSUBMENUNAME)
      return L"ContextMenu::Show - \x043F\x043E\x0434\x043C\x0435\x043D\x044E\x0020\x0441\x0020\x0438\x043C\x0435\x043D\x0435\x043C \"%s\" \x043D\x0435\x0020\x043D\x0430\x0439\x0434\x0435\x043D\x043E.";
    if (nStringID == STRID_PARSEMSG_NONPARENT)
      return L"\x041C\x0435\x0442\x043E\x0434 Index() \x0438\x0441\x043F\x043E\x043B\x044C\x0437\x0443\x0435\x0442\x0441\x044F\x0020\x0442\x043E\x043B\x044C\x043A\x043E\x0020\x0434\x043B\x044F\x0020\x043A\x043E\x0440\x043D\x0435\x0432\x043E\x0433\x043E\x0020\x043F\x043E\x0434\x043C\x0435\x043D\x044E\x002E";
    if (nStringID == STRID_PARSEMSG_UNKNOWNMETHOD)
      return L"\x041D\x0435\x0438\x0437\x0432\x0435\x0441\x0442\x043D\x044B\x0439\x0020\x043C\x0435\x0442\x043E\x0434 \"%0.s%s\".";
    if (nStringID == STRID_PARSEMSG_METHODALREADYDEFINED)
      return L"\x041C\x0435\x0442\x043E\x0434\x0020\x0443\x0436\x0435\x0020\x043D\x0430\x0437\x043D\x0430\x0447\x0435\x043D\x002E";
    if (nStringID == STRID_PARSEMSG_NOMETHOD)
      return L"\x042D\x043B\x0435\x043C\x0435\x043D\x0442\x0020\x043D\x0435\x0020\x0438\x0441\x043F\x043E\x043B\x044C\x0437\x0443\x0435\x0442\x0020\x043C\x0435\x0442\x043E\x0434\x0430\x0020\x0434\x043B\x044F\x0020\x0432\x044B\x043F\x043E\x043B\x043D\x0435\x043D\x0438\x044F\x002E";
    if (nStringID == STRID_PARSEMSG_UNNAMEDSUBMENU)
      return L"\x041D\x0435\x0020\x043D\x0430\x0439\x0434\x0435\x043D\x0020\x0437\x0430\x0433\x043E\x043B\x043E\x0432\x043E\x043A\x0020\x043F\x043E\x0434\x043C\x0435\x043D\x044E\x002E";
    if (nStringID == STRID_PARSEMSG_NOOPENBRACKET)
      return L"\x041D\x0435\x0442\x0020\x043E\x0442\x043A\x0440\x044B\x0432\x0430\x044E\x0449\x0435\x0439\x0020\x0441\x043A\x043E\x0431\x043A\x0438\x002E";
    if (nStringID == STRID_MENU_OPEN)
      return L"\x041E\x0442\x043A\x0440\x044B\x0442\x044C\tEnter";
    if (nStringID == STRID_MENU_MOVEUP)
      return L"\x041F\x0435\x0440\x0435\x043C\x0435\x0441\x0442\x0438\x0442\x044C\x0020\x0432\x0432\x0435\x0440\x0445\tAlt+Up";
    if (nStringID == STRID_MENU_MOVEDOWN)
      return L"\x041F\x0435\x0440\x0435\x043C\x0435\x0441\x0442\x0438\x0442\x044C\x0020\x0432\x043D\x0438\x0437\tAlt+Down";
    if (nStringID == STRID_MENU_SORT)
      return L"\x0421\x043E\x0440\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C";
    if (nStringID == STRID_MENU_DELETE)
      return L"\x0423\x0434\x0430\x043B\x0438\x0442\x044C\tDelete";
    if (nStringID == STRID_MENU_DELETEOLD)
      return L"\x0423\x0434\x0430\x043B\x0438\x0442\x044C\x0020\x043D\x0435\x0441\x0443\x0449\x0435\x0441\x0442\x0432\x0443\x044E\x0449\x0438\x0435";
    if (nStringID == STRID_MENU_EDIT)
      return L"\x0420\x0435\x0434\x0430\x043A\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C...\tF2";
    if (nStringID == STRID_FAVOURITES)
      return L"\x0418\x0437\x0431\x0440\x0430\x043D\x043D\x043E\x0435";
    if (nStringID == STRID_SHOWFILE)
      return L"\x041F\x043E\x043A\x0430\x0437\x044B\x0432\x0430\x0442\x044C\x0020\x0444\x0430\x0439\x043B\x044B";
    if (nStringID == STRID_FAVADDING)
      return L"\x0414\x043E\x0431\x0430\x0432\x043B\x0435\x043D\x0438\x0435\x0020\x0432\x0020\x0438\x0437\x0431\x0440\x0430\x043D\x043D\x043E\x0435";
    if (nStringID == STRID_FAVEDITING)
      return L"\x0420\x0435\x0434\x0430\x043A\x0442\x0438\x0440\x043E\x0432\x0430\x043D\x0438\x0435";
    if (nStringID == STRID_FAVNAME)
      return L"\x0418\x043C\x044F:";
    if (nStringID == STRID_FAVFILE)
      return L"\x0424\x0430\x0439\x043B:";
    if (nStringID == STRID_PLUGIN)
      return L"%s \x043F\x043B\x0430\x0433\x0438\x043D";
    if (nStringID == STRID_OK)
      return L"\x004F\x004B";
    if (nStringID == STRID_CANCEL)
      return L"\x041E\x0442\x043C\x0435\x043D\x0430";
    if (nStringID == STRID_CLOSE)
      return L"\x0417\x0430\x043A\x0440\x044B\x0442\x044C";

    //Default manual menu
    if (nStringID == STRID_DEFAULTMANUAL)
      return L"\
\"CODER\"\r\
{\r\
    \"\x041F\x043E\x0434\x0441\x0432\x0435\x0442\x043A\x0430\x0020\x0441\x0438\x043D\x0442\x0430\x043A\x0441\x0438\x0441\x0430\" +Call(\"Coder::HighLight\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 0)\r\
    \"\x0421\x0432\x043E\x0440\x0430\x0447\x0438\x0432\x0430\x043D\x0438\x0435\x0020\x0431\x043B\x043E\x043A\x043E\x0432\" +Call(\"Coder::CodeFold\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 1)\r\
    \"\x0410\x0432\x0442\x043E\x0434\x043E\x043F\x043E\x043B\x043D\x0435\x043D\x0438\x0435\" +Call(\"Coder::AutoComplete\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 2)\r\
    SEPARATOR1\r\
    \"\x041E\x0431\x043D\x043E\x0432\x0438\x0442\x044C\x0020\x043A\x044D\x0448\" Call(\"Coder::Settings\", 2)\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"Coder::Settings\")\r\
}\r\
" L"\
\"MARK\"\r\
{\r\
    \"\x0411\x0438\x0440\x044E\x0437\x043E\x0432\x044B\x043C\" Call(\"Coder::HighLight\", 2, 0, \"#9BFFFF\", 1, 0, 11) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 6)\r\
    \"\x041E\x0440\x0430\x043D\x0436\x0435\x0432\x044B\x043C\" Call(\"Coder::HighLight\", 2, 0, \"#FFCD9B\", 1, 0, 12) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 7)\r\
    \"\x0416\x0435\x043B\x0442\x044B\x043C\" Call(\"Coder::HighLight\", 2, 0, \"#FFFF9B\", 1, 0, 13) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 8)\r\
    \"\x0424\x0438\x043E\x043B\x0435\x0442\x043E\x0432\x044B\x043C\" Call(\"Coder::HighLight\", 2, 0, \"#BE7DFF\", 1, 0, 14) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 9)\r\
    \"\x0417\x0435\x043B\x0435\x043D\x044B\x043C\" Call(\"Coder::HighLight\", 2, 0, \"#88E188\", 1, 0, 15) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 10)\r\
    SEPARATOR1\r\
    -\"\x0423\x0431\x0440\x0430\x0442\x044C\x0020\x0432\x0441\x0435\x0020\x043E\x0442\x043C\x0435\x0442\x043A\x0438\" Call(\"Coder::HighLight\", 3, 0) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 11)\r\
}\r\
" L"\
\"SYNTAXTHEME\"\r\
{\r\
    \"1\x0441\" Call(\"Coder::Settings\", 1, \"1s\")\r\
    \"Assembler\" Call(\"Coder::Settings\", 1, \"asm\")\r\
    \"AutoIt\" Call(\"Coder::Settings\", 1, \"au3\")\r\
    \"Bat\" Call(\"Coder::Settings\", 1, \"bat\")\r\
    \"C++\" Call(\"Coder::Settings\", 1, \"cpp\")\r\
    \"Sharp\" Call(\"Coder::Settings\", 1, \"cs\")\r\
    \"CSS\" Call(\"Coder::Settings\", 1, \"css\")\r\
    \"HTML\" Call(\"Coder::Settings\", 1, \"html\")\r\
    \"Ini\" Call(\"Coder::Settings\", 1, \"ini\")\r\
    \"Inno\" Call(\"Coder::Settings\", 1, \"iss\")\r\
    \"JScript\" Call(\"Coder::Settings\", 1, \"js\")\r\
" L"\
    \"Lua\" Call(\"Coder::Settings\", 1, \"lua\")\r\
    \"NSIS\" Call(\"Coder::Settings\", 1, \"nsi\")\r\
    \"Pascal\" Call(\"Coder::Settings\", 1, \"dpr\")\r\
    \"Perl\" Call(\"Coder::Settings\", 1, \"pl\")\r\
    \"PHP\" Call(\"Coder::Settings\", 1, \"php\")\r\
    \"Python\" Call(\"Coder::Settings\", 1, \"py\")\r\
    \"Resource\" Call(\"Coder::Settings\", 1, \"rc\")\r\
    \"SQL\" Call(\"Coder::Settings\", 1, \"sql\")\r\
    \"VBScript\" Call(\"Coder::Settings\", 1, \"vbs\")\r\
    \"XML\" Call(\"Coder::Settings\", 1, \"xml\")\r\
    SEPARATOR1\r\
    \"\x0411\x0435\x0437\x0020\x0442\x0435\x043C\x044B\" Call(\"Coder::Settings\", 1, \"?\")\r\
}\r\
" L"\
\"COLORTHEME\"\r\
{\r\
    \"Default\" Call(\"Coder::Settings\", 5, \"Default\")\r\
    SEPARATOR1\r\
    \"Active4D\" Call(\"Coder::Settings\", 5, \"Active4D\")\r\
    \"Bespin\" Call(\"Coder::Settings\", 5, \"Bespin\")\r\
    \"Cobalt\" Call(\"Coder::Settings\", 5, \"Cobalt\")\r\
    \"Dawn\" Call(\"Coder::Settings\", 5, \"Dawn\")\r\
    \"Earth\" Call(\"Coder::Settings\", 5, \"Earth\")\r\
    \"iPlastic\" Call(\"Coder::Settings\", 5, \"iPlastic\")\r\
    \"Lazy\" Call(\"Coder::Settings\", 5, \"Lazy\")\r\
    \"Mac Classic\" Call(\"Coder::Settings\", 5, \"Mac Classic\")\r\
    \"Monokai\" Call(\"Coder::Settings\", 5, \"Monokai\")\r\
    \"Solarized Light\" Call(\"Coder::Settings\", 5, \"Solarized Light\")\r\
    \"Solarized Dark\" Call(\"Coder::Settings\", 5, \"Solarized Dark\")\r\
    \"SpaceCadet\" Call(\"Coder::Settings\", 5, \"SpaceCadet\")\r\
    \"Sunburst\" Call(\"Coder::Settings\", 5, \"Sunburst\")\r\
    \"Twilight\" Call(\"Coder::Settings\", 5, \"Twilight\")\r\
    \"Zenburn\" Call(\"Coder::Settings\", 5, \"Zenburn\")\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"Coder::Settings\")\r\
}\r\
" L"\
\"XBRACKETS\"\r\
{\r\
    \"\x0412\x043A\x043B\x044E\x0447\x0438\x0442\x044C\" +Call(\"XBrackets::Main\")\r\
    SEPARATOR1\r\
    \"\x041F\x0435\x0440\x0435\x0439\x0442\x0438\x0020\x043A\x0020\x043F\x0430\x0440\x043D\x043E\x0439\x0020\x0441\x043A\x043E\x0431\x043A\x0435\" Call(\"XBrackets::GoToMatchingBracket\")\r\
    \"\x0412\x044B\x0434\x0435\x043B\x0438\x0442\x044C\x0020\x0434\x043E\x0020\x043F\x0430\x0440\x043D\x043E\x0439\x0020\x0441\x043A\x043E\x0431\x043A\x0438\" Call(\"XBrackets::SelToMatchingBracket\")\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"XBrackets::Settings\")\r\
}\r\
" L"\
\"SPELLCHECK\"\r\
{\r\
    \"\x0424\x043E\x043D\x043E\x0432\x0430\x044F\x0020\x043F\x0440\x043E\x0432\x0435\x0440\x043A\x0430\" +Call(\"SpellCheck::Background\")\r\
    SEPARATOR1\r\
    \"\x041F\x0440\x043E\x0432\x0435\x0440\x0438\x0442\x044C\x0020\x0434\x043E\x043A\x0443\x043C\x0435\x043D\x0442\" Call(\"SpellCheck::CheckDocument\")\r\
    \"\x041F\x0440\x043E\x0432\x0435\x0440\x0438\x0442\x044C\x0020\x0432\x044B\x0434\x0435\x043B\x0435\x043D\x0438\x0435\" Call(\"SpellCheck::CheckSelection\")\r\
    \"\x041F\x0440\x043E\x0432\x0435\x0440\x0438\x0442\x044C\x0020\x0441\x043B\x043E\x0432\x043E\" Call(\"SpellCheck::Suggest\")\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"SpellCheck::Settings\")\r\
}\r\
" L"\
\"SPECIALCHAR\"\r\
{\r\
    \"\x0412\x043A\x043B\x044E\x0447\x0438\x0442\x044C\" +Call(\"SpecialChar::Main\")\r\
    SEPARATOR1\r\
    \"\x041B\x0438\x043D\x0438\x044F\x0020\x043E\x0442\x0441\x0442\x0443\x043F\x0430\" Call(\"SpecialChar::Settings\", 1, \"8\", \"0\", \"0\", -1, -1)\r\
    \"\x0412\x0441\x0435\x0020\x0441\x0438\x043C\x0432\x043E\x043B\x044B\" Call(\"SpecialChar::Settings\", 1, \"1,2,3,9,7,4,5,6\", \"0\", \"0\", -3, -3)\r\
    SEPARATOR1\r\
    \"\x041F\x0440\x043E\x0431\x0435\x043B\x0020\x0438\x0020\x0422\x0430\x0431\x0443\x043B\x044F\x0446\x0438\x044F\" Call(\"SpecialChar::Settings\", 1, \"1,2\", \"0\", \"0\", -3, -3)\r\
    \"\x041D\x043E\x0432\x0430\x044F\x0020\x0441\x0442\x0440\x043E\x043A\x0430\x0020\x0438\x0020\x041A\x043E\x043D\x0435\x0446\x0020\x0444\x0430\x0439\x043B\x0430\" Call(\"SpecialChar::Settings\", 1, \"3,9\", \"0\", \"0\", -3, -3)\r\
    \"\x041F\x0435\x0440\x0435\x043D\x043E\x0441\x0020\x0441\x0442\x0440\x043E\x043A\x0438\" Call(\"SpecialChar::Settings\", 1, \"7\", \"0\", \"0\", -1, -1)\r\
    \"\x0412\x0435\x0440\x0442\x0422\x0430\x0431\x002C\x0020\x041F\x0440\x043E\x0433\x043E\x043D\x0020\x043B\x0438\x0441\x0442\x0430\x002C\x0020\x004E\x0075\x006C\x006C\" Call(\"SpecialChar::Settings\", 1, \"4,5,6\", \"0\", \"0\", -3, -3)\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"SpecialChar::Settings\")\r\
}\r\
" L"\
\"LINEBOARD\"\r\
{\r\
    \"\x0412\x043A\x043B\x044E\x0447\x0438\x0442\x044C\" +Call(\"LineBoard::Main\")\r\
    \"\x041B\x0438\x043D\x0435\x0439\x043A\x0430\" Call(\"LineBoard::Main\", 3, -1)\r\
    SEPARATOR1\r\
    -\"\x041C\x0435\x043D\x044E\x0020\x0437\x0430\x043A\x043B\x0430\x0434\x043E\x043A\" Call(\"LineBoard::Main\", 17)\r\
    -\"\x041F\x0435\x0440\x0435\x0439\x0442\x0438\x0020\x043A\x0020\x0441\x043B\x0435\x0434\x0443\x044E\x0449\x0435\x0439\x0020\x0437\x0430\x043A\x043B\x0430\x0434\x043A\x0435\" Call(\"LineBoard::Main\", 18)\r\
    -\"\x041F\x0435\x0440\x0435\x0439\x0442\x0438\x0020\x043A\x0020\x043F\x0440\x0435\x0434\x044B\x0434\x0443\x0449\x0435\x0439\x0020\x0437\x0430\x043A\x043B\x0430\x0434\x043A\x0435\" Call(\"LineBoard::Main\", 19)\r\
    SEPARATOR1\r\
    -\"\x0423\x0441\x0442\x0430\x043D\x043E\x0432\x0438\x0442\x044C\x0020\x0437\x0430\x043A\x043B\x0430\x0434\x043A\x0443\" Call(\"LineBoard::Main\", 15)\r\
    -\"\x0423\x0434\x0430\x043B\x0438\x0442\x044C\x0020\x0437\x0430\x043A\x043B\x0430\x0434\x043A\x0443\" Call(\"LineBoard::Main\", 16)\r\
    -\"\x0423\x0434\x0430\x043B\x0438\x0442\x044C\x0020\x0432\x0441\x0435\x0020\x0437\x0430\x043A\x043B\x0430\x0434\x043A\x0438\" Call(\"LineBoard::Main\", 14)\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"LineBoard::Settings\")\r\
}\r\
" L"\
\"CLIPBOARD\"\r\
{\r\
    \"\x0417\x0430\x0445\x0432\x0430\x0442\" +Call(\"Clipboard::Capture\")\r\
    \"\x0412\x0441\x0442\x0430\x0432\x043A\x0430\x0020\x0441\x0435\x0440\x0438\x0439\x043D\x043E\x0433\x043E\x0020\x043D\x043E\x043C\x0435\x0440\x0430\" Call(\"Clipboard::PasteSerial\")\r\
    SEPARATOR1\r\
    \"\x0410\x0432\x0442\x043E\x043C\x0430\x0442\x0438\x0447\x0435\x0441\x043A\x043E\x0435\x0020\x043A\x043E\x043F\x0438\x0440\x043E\x0432\x0430\x043D\x0438\x0435\x0020\x0432\x044B\x0434\x0435\x043B\x0435\x043D\x0438\x044F\" +Call(\"Clipboard::SelAutoCopy\")\r\
    \"\x0412\x0441\x0442\x0430\x0432\x0438\x0442\x044C\x0020\x0442\x0435\x043A\x0441\x0442\" Call(\"Clipboard::Paste\")\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"Clipboard::Settings\")\r\
}\r\
" L"\
\"SAVEFILE\"\r\
{\r\
    \"\x0410\x0432\x0442\x043E\x0441\x043E\x0445\x0440\x0430\x043D\x0435\x043D\x0438\x0435\" +Call(\"SaveFile::AutoSave\")\r\
    \"\x0421\x043E\x0445\x0440\x0430\x043D\x0435\x043D\x0438\x0435\x0020\x0431\x0435\x0437 BOM\" +Call(\"SaveFile::SaveNoBOM\")\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"SaveFile::Settings\")\r\
}\r\
" L"\
\"LOG\"\r\
{\r\
    \"\x0412\x0020\x0440\x0435\x0430\x043B\x044C\x043D\x043E\x043C\x0020\x0432\x0440\x0435\x043C\x0435\x043D\x0438\" Call(\"Log::Watch\")\r\
    \"\x041F\x0430\x043D\x0435\x043B\x044C\x0020\x0432\x044B\x0432\x043E\x0434\x0430\" Call(\"Log::Output\") Icon(\"%a\\AkelFiles\\Plugs\\Log.dll\", 1)\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"Log::Settings\")\r\
}\r\
" L"\
\"EXPLORE\"\r\
{\r\
    \"\x0412\x043A\x043B\x044E\x0447\x0438\x0442\x044C\" +Call(\"Explorer::Main\")\r\
    SEPARATOR1\r\
    -\"\x041A\x0020\x0442\x0435\x043A\x0443\x0449\x0435\x043C\x0443\x0020\x0444\x0430\x0439\x043B\x0443\" Call(\"Explorer::Main\", 1, \"\")\r\
    -\"\x041A\x043E\x0440\x0435\x043D\x044C\x0020\x0041\x006B\x0065\x006C\x0050\x0061\x0064\x0027\x0430\" Call(\"Explorer::Main\", 1, \"%a\")\r\
    -\"\x041F\x043B\x0430\x0433\x0438\x043D\x044B\x0020\x0041\x006B\x0065\x006C\x0050\x0061\x0064\x0027\x0430\" Call(\"Explorer::Main\", 1, \"%a\\AkelFiles\\Plugs\")\r\
    -\"\x0421\x043A\x0440\x0438\x043F\x0442\x044B\x0020\x0041\x006B\x0065\x006C\x0050\x0061\x0064\x0027\x0430\" Call(\"Explorer::Main\", 1, \"%a\\AkelFiles\\Plugs\\Scripts\")\r\
    SEPARATOR1\r\
    -\"Program Files\" Call(\"Explorer::Main\", 1, \"%ProgramFiles%\")\r\
}\r\
" L"\
\"QSEARCH\"\r\
{\r\
    \"\x0411\x044B\x0441\x0442\x0440\x043E\x0435\x0020\x043F\x0435\x0440\x0435\x043A\x043B\x044E\x0447\x0435\x043D\x0438\x0435\x0020\x0434\x0438\x0430\x043B\x043E\x0433\x043E\x0432\" +Call(\"QSearch::DialogSwitcher\") Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 34)\r\
}\r\
" L"\
\"SCRIPTS\"\r\
{\r\
    -\"\x041F\x043E\x0441\x043B\x0435\x0434\x043D\x0438\x0439\x0020\x0432\x044B\x0437\x0432\x0430\x043D\x043D\x044B\x0439\" Call(\"Scripts::Main\", 1, \"\")\r\
    SEPARATOR1\r\
    -\"\x041F\x043E\x0438\x0441\x043A\x002F\x0417\x0430\x043C\x0435\x043D\x0430\x0020\x0441\x0020\x0440\x0435\x0433\x0443\x043B\x044F\x0440\x043D\x044B\x043C\x0438\x0020\x0432\x044B\x0440\x0430\x0436\x0435\x043D\x0438\x044F\x043C\x0438\" Call(\"Scripts::Main\", 1, \"SearchReplace.js\")\r\
    -\"\x0424\x0438\x043B\x044C\x0442\x0440\x0020\x0441\x0442\x0440\x043E\x043A\x0020\x0441\x0020\x0440\x0435\x0433\x0443\x043B\x044F\x0440\x043D\x044B\x043C\x0438\x0020\x0432\x044B\x0440\x0430\x0436\x0435\x043D\x0438\x044F\x043C\x0438\" Call(\"Scripts::Main\", 1, \"LinesFilter.js\")\r\
    SEPARATOR1\r\
    -\"\x041F\x0440\x043E\x0432\x0435\x0440\x0438\x0442\x044C\x0020\x043E\x0440\x0444\x043E\x0433\x0440\x0430\x0444\x0438\x044E\" Call(\"Scripts::Main\", 1, \"SpellCheck.js\")\r\
    -\"\x0422\x0435\x043A\x0441\x0442\x043E\x0432\x044B\x0439\x0020\x043A\x0430\x043B\x044C\x043A\x0443\x043B\x044F\x0442\x043E\x0440\" Call(\"Scripts::Main\", 1, \"Calculator.js\")\r\
    SEPARATOR1\r\
" L"\
    -\"\x0418\x0441\x043F\x0440\x0430\x0432\x0438\x0442\x044C\x0020\x043D\x0430\x0431\x043E\x0440 En->Ru\" Call(\"Scripts::Main\", 1, \"Keyboard.js\", `-Type=Layout -Direction=En->Ru`)\r\
    -\"\x0418\x0441\x043F\x0440\x0430\x0432\x0438\x0442\x044C\x0020\x043D\x0430\x0431\x043E\x0440 Ru->En\" Call(\"Scripts::Main\", 1, \"Keyboard.js\", `-Type=Layout -Direction=Ru->En`)\r\
    -\"\x0422\x0440\x0430\x043D\x0441\x043B\x0438\x0442\x0435\x0440\x0430\x0446\x0438\x044F En->Ru\" Call(\"Scripts::Main\", 1, \"Keyboard.js\", `-Type=Translit -Direction=En->Ru`)\r\
    -\"\x0422\x0440\x0430\x043D\x0441\x043B\x0438\x0442\x0435\x0440\x0430\x0446\x0438\x044F Ru->En\" Call(\"Scripts::Main\", 1, \"Keyboard.js\", `-Type=Translit -Direction=Ru->En`)\r\
    SEPARATOR1\r\
    -\"\x0412\x0441\x0442\x0430\x0432\x0438\x0442\x044C\x0020\x0444\x0430\x0439\x043B\" Call(\"Scripts::Main\", 1, \"InsertFile.js\")\r\
    -\"\x041F\x0435\x0440\x0435\x0438\x043C\x0435\x043D\x043E\x0432\x0430\x0442\x044C\x0020\x0442\x0435\x043A\x0443\x0449\x0438\x0439\x0020\x0444\x0430\x0439\x043B\" Call(\"Scripts::Main\", 1, \"RenameFile.js\")\r\
    SEPARATOR1\r\
    -\"\x0421\x043A\x0440\x0438\x043F\x0442\x044B...\" Call(\"Scripts::Main\")\r\
}\r\
" L"\
\"SESSIONS\"\r\
{\r\
    \"\x0412\x043A\x043B\x044E\x0447\x0438\x0442\x044C\" +Call(\"Sessions::Main\", 10)\r\
    SEPARATOR1\r\
    -\"\x041E\x0442\x043A\x0440\x044B\x0442\x044C...\" Call(\"Sessions::Main\")\r\
}\r\
" L"\
\"TEMPLATES\"\r\
{\r\
    \"\x0412\x043A\x043B\x044E\x0447\x0438\x0442\x044C\" +Call(\"Templates::Main\")\r\
    SEPARATOR1\r\
    \"\x041E\x0442\x043A\x0440\x044B\x0442\x044C...\" Call(\"Templates::Open\")\r\
}\r\
" L"\
\"EXIT\"\r\
{\r\
    \"\x0412\x043A\x043B\x044E\x0447\x0438\x0442\x044C\" +Call(\"Exit::Main\")\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"Exit::Settings\")\r\
}\r\
" L"\
\"SMARTSEL\"\r\
{\r\
  \"\x0423\x043C\x043D\x0430\x044F\x0020\x043A\x043B\x0430\x0432\x0438\x0448\x0430 Home\" +Call(\"SmartSel::SmartHome\")\r\
  \"\x041E\x043F\x0446\x0438\x044F\x003A\x0020\x0438\x043D\x0432\x0435\x0440\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C\" +Call(\"SmartSel::altSmartHome\")\r\
  SEPARATOR1\r\
  \"\x0423\x043C\x043D\x0430\x044F\x0020\x043A\x043B\x0430\x0432\x0438\x0448\x0430 End\" +Call(\"SmartSel::SmartEnd\")\r\
  \"\x041E\x043F\x0446\x0438\x044F\x003A\x0020\x0438\x043D\x0432\x0435\x0440\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C\" +Call(\"SmartSel::altSmartEnd\")\r\
  SEPARATOR1\r\
  \"\x0423\x043C\x043D\x044B\x0435\x0020\x043A\x043B\x0430\x0432\x0438\x0448\x0438 Up/Down\" +Call(\"SmartSel::SmartUpDown\")\r\
  \"\x041E\x043F\x0446\x0438\x044F: +PageUp/PageDown\" +Call(\"SmartSel::altSmartUpDown\")\r\
  SEPARATOR1\r\
  \"\x0418\x0441\x043A\x043B\x044E\x0447\x0438\x0442\x044C\x0020\x0045\x004F\x004C\x0020\x0438\x0437\x0020\x0432\x044B\x0434\x0435\x043B\x0435\x043D\x0438\x044F\" +Call(\"SmartSel::NoSelEOL\")\r\
  \"\x041E\x043F\x0446\x0438\x044F\x003A\x0020\x0442\x043E\x043B\x044C\x043A\x043E\x0020\x043E\x0434\x043D\x0430\x0020\x0441\x0442\x0440\x043E\x043A\x0430\" +Call(\"SmartSel::altNoSelEOL\")\r\
  SEPARATOR1\r\
  \"\x0423\x043C\x043D\x0430\x044F\x0020\x043A\x043B\x0430\x0432\x0438\x0448\x0430 Backspace\" +Call(\"SmartSel::SmartBackspace\")\r\
}\r\
" L"\
\"HOTKEYS\"\r\
{\r\
    \"\x041A\x043B\x0430\x0432\x0438\x0448\x0430 Escape\" Menu(\"EXIT\") Icon(\"%a\\AkelFiles\\Plugs\\Exit.dll\", 0)\r\
    \"\x041D\x0430\x0432\x0438\x0433\x0430\x0446\x0438\x044F\" Menu(\"SMARTSEL\")\r\
}\r\
" L"\
\"FORMAT\"\r\
{\r\
    \"\x0421\x043E\x0440\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C\x0020\x0441\x0442\x0440\x043E\x043A\x0438\x0020\x043F\x043E\x0020\x0432\x043E\x0437\x0440\x0430\x0441\x0442\x0430\x043D\x0438\x044E\" Call(\"Format::LineSortStrAsc\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 0)\r\
    \"\x0421\x043E\x0440\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C\x0020\x0441\x0442\x0440\x043E\x043A\x0438\x0020\x043F\x043E\x0020\x0443\x0431\x044B\x0432\x0430\x043D\x0438\x044E\" Call(\"Format::LineSortStrDesc\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 1)\r\
    \"\x0421\x043E\x0440\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C\x0020\x0441\x0442\x0440\x043E\x043A\x0438\x0020\x043F\x043E\x0020\x0447\x0438\x0441\x043B\x043E\x0432\x043E\x043C\x0443\x0020\x0432\x043E\x0437\x0440\x0430\x0441\x0442\x0430\x043D\x0438\x044E\" Call(\"Format::LineSortIntAsc\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 2)\r\
    \"\x0421\x043E\x0440\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C\x0020\x0441\x0442\x0440\x043E\x043A\x0438\x0020\x043F\x043E\x0020\x0447\x0438\x0441\x043B\x043E\x0432\x043E\x043C\x0443\x0020\x0443\x0431\x044B\x0432\x0430\x043D\x0438\x044E\" Call(\"Format::LineSortIntDesc\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 3)\r\
    SEPARATOR1\r\
    \"\x041F\x043E\x043B\x0443\x0447\x0438\x0442\x044C\x0020\x0443\x043D\x0438\x043A\x0430\x043B\x044C\x043D\x044B\x0435\x0020\x0441\x0442\x0440\x043E\x043A\x0438\" Call(\"Format::LineGetUnique\")\r\
    \"\x041F\x043E\x043B\x0443\x0447\x0438\x0442\x044C\x0020\x0434\x0443\x0431\x043B\x0438\x0440\x0443\x044E\x0449\x0438\x0435\x0441\x044F\x0020\x0441\x0442\x0440\x043E\x043A\x0438\" Call(\"Format::LineGetDuplicates\")\r\
    \"\x0423\x0434\x0430\x043B\x0438\x0442\x044C\x0020\x0434\x0443\x0431\x043B\x0438\x0440\x0443\x044E\x0449\x0438\x0435\x0441\x044F\x0020\x0441\x0442\x0440\x043E\x043A\x0438\" Call(\"Format::LineRemoveDuplicates\")\r\
    SEPARATOR1\r\
" L"\
    \"\x0418\x043D\x0432\x0435\x0440\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C\x0020\x043F\x043E\x0440\x044F\x0434\x043E\x043A\x0020\x0441\x0442\x0440\x043E\x043A\" Call(\"Format::LineReverse\")\r\
    \"\x0412\x0441\x0442\x0430\x0432\x0438\x0442\x044C\x0020\x0440\x0430\x0437\x0440\x044B\x0432\x0020\x0441\x0442\x0440\x043E\x043A\x0438\x0020\x0432\x0020\x043C\x0435\x0441\x0442\x0430\x0445\x0020\x043F\x0435\x0440\x0435\x043D\x043E\x0441\x0430\" Call(\"Format::LineFixWrap\")\r\
    SEPARATOR1\r\
    \"\x0417\x0430\x0448\x0438\x0444\x0440\x043E\x0432\x0430\x0442\x044C\x0020\x0432\x044B\x0434\x0435\x043B\x0435\x043D\x0438\x0435\" Call(\"Format::Encrypt\")\r\
    \"\x0420\x0430\x0441\x0448\x0438\x0444\x0440\x043E\x0432\x0430\x0442\x044C\x0020\x0432\x044B\x0434\x0435\x043B\x0435\x043D\x0438\x0435\" Call(\"Format::Decrypt\")\r\
    SEPARATOR1\r\
    \"\x0418\x0437\x0432\x043B\x0435\x0447\x044C\x0020\x0441\x0441\x044B\x043B\x043A\x0438\x0020\x0438\x0437\x0020\x0048\x0054\x004D\x004C\x0020\x0442\x0435\x043A\x0441\x0442\x0430\" Call(\"Format::LinkExtract\")\r\
}\r\
" L"\
\"SCROLL\"\r\
{\r\
    \"\x0412\x0435\x0440\x0442\x0438\x043A\x0430\x043B\x044C\x043D\x0430\x044F\x0020\x0441\x0438\x043D\x0445\x0440\x043E\x043D\x0438\x0437\x0430\x0446\x0438\x044F\" Call(\"Scroll::SyncVert\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 1)\r\
    \"\x0413\x043E\x0440\x0438\x0437\x043E\x043D\x0442\x0430\x043B\x044C\x043D\x0430\x044F\x0020\x0441\x0438\x043D\x0445\x0440\x043E\x043D\x0438\x0437\x0430\x0446\x0438\x044F\" Call(\"Scroll::SyncHorz\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 0)\r\
    SEPARATOR1\r\
    \"\x0410\x0432\x0442\x043E\x043C\x0430\x0442\x0438\x0447\x0435\x0441\x043A\x0430\x044F\x0020\x043F\x0440\x043E\x043A\x0440\x0443\x0442\x043A\x0430\x0020\x0442\x0435\x043A\x0441\x0442\x0430\" +Call(\"Scroll::AutoScroll\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 2)\r\
    \"\x041E\x0431\x0440\x0430\x0431\x043E\x0442\x043A\x0430\x0020\x043D\x0435\x043F\x0440\x043E\x043A\x0440\x0443\x0447\x0438\x0432\x0430\x0435\x043C\x044B\x0445\x0020\x043E\x043F\x0435\x0440\x0430\x0446\x0438\x0439\" +Call(\"Scroll::NoScroll\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 3)\r\
    \"\x0412\x044B\x0440\x0430\x0432\x043D\x0438\x0432\x0430\x043D\x0438\x0435\x0020\x043A\x0430\x0440\x0435\x0442\x043A\x0438\" +Call(\"Scroll::AlignCaret\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 4)\r\
    \"\x0410\x0432\x0442\x043E\x043C\x0430\x0442\x0438\x0447\x0435\x0441\x043A\x0430\x044F\x0020\x043F\x0435\x0440\x0435\x0434\x0430\x0447\x0430\x0020\x0444\x043E\x043A\x0443\x0441\x0430\" +Call(\"Scroll::AutoFocus\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 5)\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"Scroll::Settings\")\r\
}\r\
" L"\
\"MINIMIZETOTRAY\"\r\
{\r\
    \"\x0421\x0432\x043E\x0440\x0430\x0447\x0438\x0432\x0430\x0442\x044C\x0020\x0432\x0020\x0442\x0440\x0435\x0439\x0020\x0432\x0441\x0435\x0433\x0434\x0430\" +Call(\"MinimizeToTray::Always\")\r\
}\r\
" L"\
\"SOUNDS\"\r\
{\r\
    \"\x0412\x043A\x043B\x044E\x0447\x0438\x0442\x044C\" +Call(\"Sounds::Main\")\r\
    SEPARATOR1\r\
    \"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C...\" Call(\"Sounds::Settings\")\r\
}\r\
" L"\
\"BBCODE\"\r\
{\r\
    \"URL=\" Insert(`[URL=\\|]\\s[/URL]`, 1)\r\
    \"QUOTE\" Insert(`[QUOTE]\\s[/QUOTE]`, 1)\r\
    \"QUOTE=\" Insert(`[QUOTE=\"\\|\"]\\s[/QUOTE]`, 1)\r\
    \"CODE\" Insert(`[CODE]\\s[/CODE]`, 1)\r\
    \"MORE=\" Insert(`[MORE=\"\\|\"]\\s[/MORE]`, 1)\r\
    \"IMG\" Insert(`[IMG]\\s[/IMG]`, 1)\r\
    \"b\" Insert(`[b]\\s[/b]`, 1)\r\
    \"i\" Insert(`[i]\\s[/i]`, 1)\r\
}\r";

    //Default favourites main menu
    if (nStringID == STRID_DEFAULTMAIN)
      return L"\
\"\x0026\x0418\x0437\x0431\x0440\x0430\x043D\x043D\x043E\x0435\" Index(3)\r\
{\r\
    \"\x0414\x043E\x0431\x0430\x0432\x0438\x0442\x044C\" Favourites(1) Icon(0)\r\
    \"\x0423\x043F\x0440\x0430\x0432\x043B\x0435\x043D\x0438\x0435...\" Favourites(3) Icon(1)\r\
    SEPARATOR1\r\
    FAVOURITES\r\
}\r\
" L"\
\"\x041F\x043B\x0430\x0026\x0433\x0438\x043D\x044B\" Index(-2)\r\
{\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Coder.dll\")\r\
        \"\x041F\x0440\x043E\x0433\x0440\x0430\x043C\x043C\x0438\x0440\x043E\x0432\x0430\x043D\x0438\x0435\" Menu(\"CODER\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 12)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\XBrackets.dll\")\r\
        \"\x041F\x0430\x0440\x043D\x044B\x0435\x0020\x0441\x043A\x043E\x0431\x043A\x0438\" Menu(\"XBRACKETS\") Icon(\"%a\\AkelFiles\\Plugs\\XBrackets.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\SpellCheck.dll\")\r\
        \"\x041F\x0440\x043E\x0432\x0435\x0440\x043A\x0430\x0020\x043E\x0440\x0444\x043E\x0433\x0440\x0430\x0444\x0438\x0438\" Menu(\"SPELLCHECK\") Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 35)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\SpecialChar.dll\")\r\
        \"\x0421\x043F\x0435\x0446\x0438\x0430\x043B\x044C\x043D\x044B\x0435\x0020\x0441\x0438\x043C\x0432\x043E\x043B\x044B\" Menu(\"SPECIALCHAR\") Icon(\"%a\\AkelFiles\\Plugs\\SpecialChar.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\LineBoard.dll\")\r\
        \"\x041D\x043E\x043C\x0435\x0440\x0430\x0020\x0441\x0442\x0440\x043E\x043A\x002C\x0020\x0437\x0430\x043A\x043B\x0430\x0434\x043A\x0438\" Menu(\"LINEBOARD\") Icon(\"%a\\AkelFiles\\Plugs\\LineBoard.dll\", 0)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\Clipboard.dll\")\r\
        \"\x0411\x0443\x0444\x0435\x0440\x0020\x043E\x0431\x043C\x0435\x043D\x0430\" Menu(\"CLIPBOARD\") Icon(\"%a\\AkelFiles\\Plugs\\Clipboard.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\SaveFile.dll\")\r\
        \"\x0421\x043E\x0445\x0440\x0430\x043D\x0435\x043D\x0438\x0435\x0020\x0444\x0430\x0439\x043B\x0430\" Menu(\"SAVEFILE\") Icon(\"%a\\AkelFiles\\Plugs\\SaveFile.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Log.dll\")\r\
        \"\x041F\x0440\x043E\x0441\x043C\x043E\x0442\x0440\x0020\x043B\x043E\x0433\x0430\" Menu(\"LOG\") Icon(\"%a\\AkelFiles\\Plugs\\Log.dll\", 0)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\ToolBar.dll\")\r\
        \"\x041F\x0430\x043D\x0435\x043B\x044C\x0020\x0438\x043D\x0441\x0442\x0440\x0443\x043C\x0435\x043D\x0442\x043E\x0432\" +Call(\"ToolBar::Main\")\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Explorer.dll\")\r\
        \"\x041F\x0430\x043D\x0435\x043B\x044C\x0020\x043F\x0440\x043E\x0432\x043E\x0434\x043D\x0438\x043A\x0430\" Menu(\"EXPLORE\") Icon(\"%a\\AkelFiles\\Plugs\\Explorer.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\QSearch.dll\")\r\
        \"\x041F\x0430\x043D\x0435\x043B\x044C\x0020\x043F\x043E\x0438\x0441\x043A\x0430\" +Call(\"QSearch::QSearch\") Icon(\"%a\\AkelFiles\\Plugs\\QSearch.dll\", 0)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\Scripts.dll\")\r\
        -\"\x0421\x043A\x0440\x0438\x043F\x0442\x044B...\" +Call(\"Scripts::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Scripts.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Macros.dll\")\r\
        -\"\x041C\x0430\x043A\x0440\x043E\x0441\x044B...\" +Call(\"Macros::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\RecentFiles.dll\")\r\
        -\"\x041F\x043E\x0441\x043B\x0435\x0434\x043D\x0438\x0435\x0020\x0444\x0430\x0439\x043B\x044B...\" Call(\"RecentFiles::Manage\") Icon(\"%a\\AkelFiles\\Plugs\\RecentFiles.dll\", 0)\r\
    UNSET(32)\r\
    SET(1)\r\
        # MDI/PMDI\r\
        SET(32, \"%a\\AkelFiles\\Plugs\\Sessions.dll\")\r\
            \"\x0421\x0435\x0441\x0441\x0438\x0438\" Menu(\"SESSIONS\") Icon(\"%a\\AkelFiles\\Plugs\\Sessions.dll\", 0)\r\
        UNSET(32)\r\
    UNSET(1)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Templates.dll\")\r\
        \"\x0428\x0430\x0431\x043B\x043E\x043D\x044B\" Menu(\"TEMPLATES\") Icon(\"%a\\AkelFiles\\Plugs\\Toolbar.dll\", 37)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\Hotkeys.dll\")\r\
        -\"\x0413\x043E\x0440\x044F\x0447\x0438\x0435\x0020\x043A\x043B\x0430\x0432\x0438\x0448\x0438...\" +Call(\"Hotkeys::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Hotkeys.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Exit.dll\")\r\
        \"\x041A\x043B\x0430\x0432\x0438\x0448\x0430 Escape\" Menu(\"EXIT\") Icon(\"%a\\AkelFiles\\Plugs\\Exit.dll\", 0)\r\
    UNSET(32)\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\SmartSel.dll\")\r\
        \"\x041D\x0430\x0432\x0438\x0433\x0430\x0446\x0438\x044F\" Menu(\"SMARTSEL\")\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\MinimizeToTray.dll\")\r\
        \"\x0421\x0432\x0435\x0440\x043D\x0443\x0442\x044C\x0020\x0432\x0020\x0442\x0440\x0435\x0439\" Call(\"MinimizeToTray::Now\") Icon(\"%a\\AkelFiles\\Plugs\\MinimizeToTray.dll\", 0)\r\
        \"\x0421\x0432\x043E\x0440\x0430\x0447\x0438\x0432\x0430\x0442\x044C\x0020\x0432\x0020\x0442\x0440\x0435\x0439\x0020\x0432\x0441\x0435\x0433\x0434\x0430\" +Call(\"MinimizeToTray::Always\")\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\QSearch.dll\")\r\
        \"\x0411\x044B\x0441\x0442\x0440\x043E\x0435\x0020\x043F\x0435\x0440\x0435\x043A\x043B\x044E\x0447\x0435\x043D\x0438\x0435\x0020\x0434\x0438\x0430\x043B\x043E\x0433\x043E\x0432\" +Call(\"QSearch::DialogSwitcher\") Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 34)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\Sounds.dll\")\r\
        \"\x0417\x0432\x0443\x043A\x043E\x0432\x043E\x0439\x0020\x043D\x0430\x0431\x043E\x0440\x0020\x0442\x0435\x043A\x0441\x0442\x0430\" Menu(\"SOUNDS\") Icon(\"%a\\AkelFiles\\Plugs\\Sounds.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Speech.dll\")\r\
        \"\x041C\x0430\x0448\x0438\x043D\x043D\x043E\x0435\x0020\x0447\x0442\x0435\x043D\x0438\x0435\x0020\x0442\x0435\x043A\x0441\x0442\x0430\" +Call(\"Speech::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Speech.dll\", 0)\r\
    UNSET(32)\r\
}\r";

    //Default edit menu
    if (nStringID == STRID_DEFAULTEDIT)
      return L"\
SET(8)\r\
    \"\" Command(4151) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 10)\r\
    \"\" Command(4152) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 11)\r\
    SEPARATOR1\r\
    \"\" Command(4153) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 7)\r\
    \"\" Command(4154) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 8)\r\
    \"\" Command(4155) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 9)\r\
    \"\" Command(4156) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 25)\r\
    SEPARATOR1\r\
    \"\" Command(4157)\r\
UNSET(8)\r\
" L"\
SEPARATOR1\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Scripts.dll\")\r\
    \"\x0421\x043A\x0440\x0438\x043F\x0442\x044B\" Menu(\"SCRIPTS\") Icon(\"%a\\AkelFiles\\Plugs\\Scripts.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Format.dll\")\r\
    \"\x041F\x0440\x0435\x043E\x0431\x0440\x0430\x0437\x043E\x0432\x0430\x0442\x044C\" Menu(\"FORMAT\")\r\
UNSET(32)\r\
\"BBCode\" Menu(\"BBCODE\")\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Scroll.dll\")\r\
    \"\x041F\x0440\x043E\x043A\x0440\x0443\x0442\x043A\x0430\" Menu(\"SCROLL\")\r\
UNSET(32)\r\
SEPARATOR1\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Coder.dll\")\r\
    \"\x041E\x0442\x043C\x0435\x0442\x0438\x0442\x044C\" Menu(\"MARK\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 0)\r\
    \"\x0421\x0438\x043D\x0442\x0430\x043A\x0441\x0438\x0447\x0435\x0441\x043A\x0430\x044F\x0020\x0442\x0435\x043C\x0430\" Menu(\"SYNTAXTHEME\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 4)\r\
    \"\x0426\x0432\x0435\x0442\x043E\x0432\x0430\x044F\x0020\x0442\x0435\x043C\x0430\" Menu(\"COLORTHEME\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 5)\r\
UNSET(32)\r\
SEPARATOR1\r\
SET(32, \"%a\\AkelFiles\\Plugs\\HexSel.dll\")\r\
    \"\x0048\x0065\x0078\x0020\x043A\x043E\x0434\" +Call(\"HexSel::Main\") Icon(\"%a\\AkelFiles\\Plugs\\HexSel.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Stats.dll\")\r\
    \"\x0421\x0442\x0430\x0442\x0438\x0441\x0442\x0438\x043A\x0430\" Call(\"Stats::Main\")\r\
UNSET(32)\r\
SEPARATOR1\r\
SET(32, \"%a\\AkelFiles\\Plugs\\FullScreen.dll\")\r\
    \"\x041F\x043E\x043B\x043D\x043E\x044D\x043A\x0440\x0430\x043D\x043D\x044B\x0439\x0020\x0440\x0435\x0436\x0438\x043C\" Call(\"FullScreen::Main\") Icon(\"%a\\AkelFiles\\Plugs\\FullScreen.dll\", 0)\r\
UNSET(32)\r";

    //Default tabs menu
    if (nStringID == STRID_DEFAULTTAB)
      return L"\
\"\x0412\x044B\x0431\x043E\x0440\x0020\x043E\x043A\x043D\x0430\" Command(4327)\r\
SEPARATOR1\r\
\"\x0417\x0430\x043A\x0440\x044B\x0442\x044C\" Command(4318)\r\
\"\x0417\x0430\x043A\x0440\x044B\x0442\x044C\x0020\x0432\x0441\x0435\" Command(4319)\r\
\"\x0417\x0430\x043A\x0440\x044B\x0442\x044C\x0020\x0432\x0441\x0435\x002C\x0020\x043A\x0440\x043E\x043C\x0435\x0020\x0430\x043A\x0442\x0438\x0432\x043D\x043E\x0439\" Command(4320)\r\
SEPARATOR1\r\
SET(4)\r\
    #\x0422\x043E\x043B\x044C\x043A\x043E\x0020\x0434\x043B\x044F\x0020\x004D\x0044\x0049\r\
    \"\x0026\x0413\x043E\x0440\x0438\x0437\x043E\x043D\x0442\x0430\x043B\x044C\x043D\x043E\" Command(4307)\r\
    \"\x0026\x0412\x0435\x0440\x0442\x0438\x043A\x0430\x043B\x044C\x043D\x043E\" Command(4308)\r\
    \"\x0026\x041A\x0430\x0441\x043A\x0430\x0434\x043E\x043C\" Command(4309)\r\
    SEPARATOR1\r\
UNSET(4)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Scripts.dll\")\r\
    -\"\x041A\x043E\x043F\x0438\x0440\x043E\x0432\x0430\x0442\x044C\x0020\x043F\x0443\x0442\x044C\" Call(\"Scripts::Main\", 1, \"EvalCmd.js\", `AkelPad.SetClipboardText(AkelPad.GetEditFile(0));`)\r\
UNSET(32)\r\
\"\x041C\x0435\x043D\x044E\x0020\x043F\x0440\x043E\x0432\x043E\x0434\x043D\x0438\x043A\x0430\"\r\
{\r\
    EXPLORER\r\
}\r";

    //Default URL menu
    if (nStringID == STRID_DEFAULTURL)
      return L"\
\"\x041E\x0442\x043A\x0440\x044B\x0442\x044C\" Link(1)\r\
\"\x041A\x043E\x043F\x0438\x0440\x043E\x0432\x0430\x0442\x044C\" Link(2)\r\
\"\x0412\x044B\x0434\x0435\x043B\x0438\x0442\x044C\" Link(3)\r\
SEPARATOR1\r\
\"\x0412\x044B\x0440\x0435\x0437\x0430\x0442\x044C\" Link(4)\r\
\"\x0412\x0441\x0442\x0430\x0432\x0438\x0442\x044C\" Link(5)\r\
\"\x0423\x0434\x0430\x043B\x0438\x0442\x044C\" Link(6)\r\
SEPARATOR1\r\
" L"\
SET(8)\r\
    \"\" Command(4151) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 10)\r\
    \"\" Command(4152) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 11)\r\
    SEPARATOR1\r\
    \"\" Command(4153) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 7)\r\
    \"\" Command(4154) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 8)\r\
    \"\" Command(4155) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 9)\r\
    \"\" Command(4156) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 25)\r\
    SEPARATOR1\r\
    \"\" Command(4157)\r\
UNSET(8)\r";

    //Default recent files menu
    if (nStringID == STRID_DEFAULTRECENTFILES)
      return L"\
EXPLORER\r";

  }
  else
  {
    if (nStringID == STRID_MANUAL)
      return L"ContextMenu::Show menu";
    if (nStringID == STRID_MENUMAIN)
      return L"Main menu";
    if (nStringID == STRID_MENUEDIT)
      return L"Edit menu";
    if (nStringID == STRID_MENUTAB)
      return L"Tab menu";
    if (nStringID == STRID_MENUURL)
      return L"URL menu";
    if (nStringID == STRID_MENURECENTFILES)
      return L"Recent files menu";
    if (nStringID == STRID_ENABLE)
      return L"Enable";
    if (nStringID == STRID_HIDE)
      return L"Hide";
    if (nStringID == STRID_SHOW)
      return L"Show";
    if (nStringID == STRID_AUTOLOAD)
      return L"\"%s\" doesn't support autoload (\"+\" before \"Call\").";
    if (nStringID == STRID_NOSUBMENUINDEX)
      return L"ContextMenu::Show - submenu with index %d not found.";
    if (nStringID == STRID_NOSUBMENUNAME)
      return L"ContextMenu::Show - submenu with name \"%s\" not found.";
    if (nStringID == STRID_PARSEMSG_NONPARENT)
      return L"Method Index () is used only for the root submenu.";
    if (nStringID == STRID_PARSEMSG_UNKNOWNMETHOD)
      return L"Unknown method \"%0.s%s\".";
    if (nStringID == STRID_PARSEMSG_METHODALREADYDEFINED)
      return L"Method already defined.";
    if (nStringID == STRID_PARSEMSG_NOMETHOD)
      return L"The element does not use method for execution.";
    if (nStringID == STRID_PARSEMSG_UNNAMEDSUBMENU)
      return L"Unable to find the title of a submenu.";
    if (nStringID == STRID_PARSEMSG_NOOPENBRACKET)
      return L"No opening bracket.";
    if (nStringID == STRID_MENU_OPEN)
      return L"Open\tEnter";
    if (nStringID == STRID_MENU_MOVEUP)
      return L"Move up\tAlt+Up";
    if (nStringID == STRID_MENU_MOVEDOWN)
      return L"Move down\tAlt+Down";
    if (nStringID == STRID_MENU_SORT)
      return L"Sort";
    if (nStringID == STRID_MENU_DELETE)
      return L"Delete\tDelete";
    if (nStringID == STRID_MENU_DELETEOLD)
      return L"Delete non-existent";
    if (nStringID == STRID_MENU_EDIT)
      return L"Edit...\tF2";
    if (nStringID == STRID_FAVOURITES)
      return L"Favourites";
    if (nStringID == STRID_SHOWFILE)
      return L"Show files";
    if (nStringID == STRID_FAVADDING)
      return L"Adding to favourites";
    if (nStringID == STRID_FAVEDITING)
      return L"Editing";
    if (nStringID == STRID_FAVNAME)
      return L"Name:";
    if (nStringID == STRID_FAVFILE)
      return L"File:";
    if (nStringID == STRID_PLUGIN)
      return L"%s plugin";
    if (nStringID == STRID_OK)
      return L"OK";
    if (nStringID == STRID_CANCEL)
      return L"Cancel";
    if (nStringID == STRID_CLOSE)
      return L"Close";

    //Default manual menu
    if (nStringID == STRID_DEFAULTMANUAL)
      return L"\
\"CODER\"\r\
{\r\
    \"Syntax highlighting\" +Call(\"Coder::HighLight\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 0)\r\
    \"Code folding\" +Call(\"Coder::CodeFold\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 1)\r\
    \"Autocompletion\" +Call(\"Coder::AutoComplete\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 2)\r\
    SEPARATOR1\r\
    \"Update cache\" Call(\"Coder::Settings\", 2)\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"Coder::Settings\")\r\
}\r\
" L"\
\"MARK\"\r\
{\r\
    \"Cyan\" Call(\"Coder::HighLight\", 2, 0, \"#9BFFFF\", 1, 0, 11) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 6)\r\
    \"Orange\" Call(\"Coder::HighLight\", 2, 0, \"#FFCD9B\", 1, 0, 12) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 7)\r\
    \"Yellow\" Call(\"Coder::HighLight\", 2, 0, \"#FFFF9B\", 1, 0, 13) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 8)\r\
    \"Violet\" Call(\"Coder::HighLight\", 2, 0, \"#BE7DFF\", 1, 0, 14) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 9)\r\
    \"Green\" Call(\"Coder::HighLight\", 2, 0, \"#88E188\", 1, 0, 15) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 10)\r\
    SEPARATOR1\r\
    -\"Remove all marks\" Call(\"Coder::HighLight\", 3, 0) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 11)\r\
}\r\
" L"\
\"SYNTAXTHEME\"\r\
{\r\
    \"1s\" Call(\"Coder::Settings\", 1, \"1s\")\r\
    \"Assembler\" Call(\"Coder::Settings\", 1, \"asm\")\r\
    \"AutoIt\" Call(\"Coder::Settings\", 1, \"au3\")\r\
    \"Bat\" Call(\"Coder::Settings\", 1, \"bat\")\r\
    \"C++\" Call(\"Coder::Settings\", 1, \"cpp\")\r\
    \"Sharp\" Call(\"Coder::Settings\", 1, \"cs\")\r\
    \"CSS\" Call(\"Coder::Settings\", 1, \"css\")\r\
    \"HTML\" Call(\"Coder::Settings\", 1, \"html\")\r\
    \"Ini\" Call(\"Coder::Settings\", 1, \"ini\")\r\
    \"Inno\" Call(\"Coder::Settings\", 1, \"iss\")\r\
    \"JScript\" Call(\"Coder::Settings\", 1, \"js\")\r\
" L"\
    \"Lua\" Call(\"Coder::Settings\", 1, \"lua\")\r\
    \"NSIS\" Call(\"Coder::Settings\", 1, \"nsi\")\r\
    \"Pascal\" Call(\"Coder::Settings\", 1, \"dpr\")\r\
    \"Perl\" Call(\"Coder::Settings\", 1, \"pl\")\r\
    \"PHP\" Call(\"Coder::Settings\", 1, \"php\")\r\
    \"Python\" Call(\"Coder::Settings\", 1, \"py\")\r\
    \"Resource\" Call(\"Coder::Settings\", 1, \"rc\")\r\
    \"SQL\" Call(\"Coder::Settings\", 1, \"sql\")\r\
    \"VBScript\" Call(\"Coder::Settings\", 1, \"vbs\")\r\
    \"XML\" Call(\"Coder::Settings\", 1, \"xml\")\r\
    SEPARATOR1\r\
    \"None\" Call(\"Coder::Settings\", 1, \"?\")\r\
}\r\
" L"\
\"COLORTHEME\"\r\
{\r\
    \"Default\" Call(\"Coder::Settings\", 5, \"Default\")\r\
    SEPARATOR1\r\
    \"Active4D\" Call(\"Coder::Settings\", 5, \"Active4D\")\r\
    \"Bespin\" Call(\"Coder::Settings\", 5, \"Bespin\")\r\
    \"Cobalt\" Call(\"Coder::Settings\", 5, \"Cobalt\")\r\
    \"Dawn\" Call(\"Coder::Settings\", 5, \"Dawn\")\r\
    \"Earth\" Call(\"Coder::Settings\", 5, \"Earth\")\r\
    \"iPlastic\" Call(\"Coder::Settings\", 5, \"iPlastic\")\r\
    \"Lazy\" Call(\"Coder::Settings\", 5, \"Lazy\")\r\
    \"Mac Classic\" Call(\"Coder::Settings\", 5, \"Mac Classic\")\r\
    \"Monokai\" Call(\"Coder::Settings\", 5, \"Monokai\")\r\
    \"Solarized Light\" Call(\"Coder::Settings\", 5, \"Solarized Light\")\r\
    \"Solarized Dark\" Call(\"Coder::Settings\", 5, \"Solarized Dark\")\r\
    \"SpaceCadet\" Call(\"Coder::Settings\", 5, \"SpaceCadet\")\r\
    \"Sunburst\" Call(\"Coder::Settings\", 5, \"Sunburst\")\r\
    \"Twilight\" Call(\"Coder::Settings\", 5, \"Twilight\")\r\
    \"Zenburn\" Call(\"Coder::Settings\", 5, \"Zenburn\")\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"Coder::Settings\")\r\
}\r\
" L"\
\"XBRACKETS\"\r\
{\r\
    \"Enable\" +Call(\"XBrackets::Main\")\r\
    SEPARATOR1\r\
    \"Go to matching bracket\" Call(\"XBrackets::GoToMatchingBracket\")\r\
    \"Select to matching bracket\" Call(\"XBrackets::SelToMatchingBracket\")\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"XBrackets::Settings\")\r\
}\r\
" L"\
\"SPELLCHECK\"\r\
{\r\
    \"Background check\" +Call(\"SpellCheck::Background\")\r\
    SEPARATOR1\r\
    \"Check document\" Call(\"SpellCheck::CheckDocument\")\r\
    \"Check selection\" Call(\"SpellCheck::CheckSelection\")\r\
    \"Check word\" Call(\"SpellCheck::Suggest\")\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"SpellCheck::Settings\")\r\
}\r\
" L"\
\"SPECIALCHAR\"\r\
{\r\
    \"Enable\" +Call(\"SpecialChar::Main\")\r\
    SEPARATOR1\r\
    \"Indent line\" Call(\"SpecialChar::Settings\", 1, \"8\", \"0\", \"0\", -1, -1)\r\
    \"All symbols\" Call(\"SpecialChar::Settings\", 1, \"1,2,3,9,7,4,5,6\", \"0\", \"0\", -3, -3)\r\
    SEPARATOR1\r\
    \"Space and Tabulation\" Call(\"SpecialChar::Settings\", 1, \"1,2\", \"0\", \"0\", -3, -3)\r\
    \"New line and End of line\" Call(\"SpecialChar::Settings\", 1, \"3,9\", \"0\", \"0\", -3, -3)\r\
    \"Wrap line\" Call(\"SpecialChar::Settings\", 1, \"7\", \"0\", \"0\", -1, -1)\r\
    \"VertTab, Form-feed, Null\" Call(\"SpecialChar::Settings\", 1, \"4,5,6\", \"0\", \"0\", -3, -3)\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"SpecialChar::Settings\")\r\
}\r\
" L"\
\"LINEBOARD\"\r\
{\r\
    \"Enable\" +Call(\"LineBoard::Main\")\r\
    \"Ruler\" Call(\"LineBoard::Main\", 3, -1)\r\
    SEPARATOR1\r\
    -\"Bookmark menu\" Call(\"LineBoard::Main\", 17)\r\
    -\"Go to next bookmark\" Call(\"LineBoard::Main\", 18)\r\
    -\"Go to previous bookmark\" Call(\"LineBoard::Main\", 19)\r\
    SEPARATOR1\r\
    -\"Set bookmark\" Call(\"LineBoard::Main\", 15)\r\
    -\"Delete bookmark\" Call(\"LineBoard::Main\", 16)\r\
    -\"Delete all bookmarks\" Call(\"LineBoard::Main\", 14)\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"LineBoard::Settings\")\r\
}\r\
" L"\
\"CLIPBOARD\"\r\
{\r\
    \"Capture\" +Call(\"Clipboard::Capture\")\r\
    \"Serial number insertion\" Call(\"Clipboard::PasteSerial\")\r\
    SEPARATOR1\r\
    \"Automatic copy selection\" +Call(\"Clipboard::SelAutoCopy\")\r\
    \"Paste text\" Call(\"Clipboard::Paste\")\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"Clipboard::Settings\")\r\
}\r\
" L"\
\"SAVEFILE\"\r\
{\r\
    \"Autosaving\" +Call(\"SaveFile::AutoSave\")\r\
    \"Saving without BOM\" +Call(\"SaveFile::SaveNoBOM\")\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"SaveFile::Settings\")\r\
}\r\
" L"\
\"LOG\"\r\
{\r\
    \"In real time\" Call(\"Log::Watch\")\r\
    \"Output panel\" Call(\"Log::Output\") Icon(\"%a\\AkelFiles\\Plugs\\Log.dll\", 1)\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"Log::Settings\")\r\
}\r\
" L"\
\"EXPLORE\"\r\
{\r\
    \"Enable\" +Call(\"Explorer::Main\")\r\
    SEPARATOR1\r\
    -\"Current file\" Call(\"Explorer::Main\", 1, \"\")\r\
    -\"AkelPad's root\" Call(\"Explorer::Main\", 1, \"%a\")\r\
    -\"AkelPad's plugins\" Call(\"Explorer::Main\", 1, \"%a\\AkelFiles\\Plugs\")\r\
    -\"AkelPad's scripts\" Call(\"Explorer::Main\", 1, \"%a\\AkelFiles\\Plugs\\Scripts\")\r\
    SEPARATOR1\r\
    -\"Program Files\" Call(\"Explorer::Main\", 1, \"%ProgramFiles%\")\r\
}\r\
" L"\
\"QSEARCH\"\r\
{\r\
    \"Quick dialogs switch\" +Call(\"QSearch::DialogSwitcher\") Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 34)\r\
}\r\
" L"\
\"SCRIPTS\"\r\
{\r\
    -\"Last called\" Call(\"Scripts::Main\", 1, \"\")\r\
    SEPARATOR1\r\
    -\"Search/Replace with regular expressions\" Call(\"Scripts::Main\", 1, \"SearchReplace.js\")\r\
    -\"Filter lines with regular expressions\" Call(\"Scripts::Main\", 1, \"LinesFilter.js\")\r\
    SEPARATOR1\r\
    -\"Spell check\" Call(\"Scripts::Main\", 1, \"SpellCheck.js\")\r\
    -\"Text calculator\" Call(\"Scripts::Main\", 1, \"Calculator.js\")\r\
    SEPARATOR1\r\
" L"\
    -\"Convert keyboard layout En->Ru\" Call(\"Scripts::Main\", 1, \"Keyboard.js\", `-Type=Layout -Direction=En->Ru`)\r\
    -\"Convert keyboard layout Ru->En\" Call(\"Scripts::Main\", 1, \"Keyboard.js\", `-Type=Layout -Direction=Ru->En`)\r\
    -\"Transliteration Latin->Cyrillic\" Call(\"Scripts::Main\", 1, \"Keyboard.js\", `-Type=Translit -Direction=En->Ru`)\r\
    -\"Transliteration Cyrillic->Latin\" Call(\"Scripts::Main\", 1, \"Keyboard.js\", `-Type=Translit -Direction=Ru->En`)\r\
    SEPARATOR1\r\
    -\"Insert file\" Call(\"Scripts::Main\", 1, \"InsertFile.js\")\r\
    -\"Rename current file\" Call(\"Scripts::Main\", 1, \"RenameFile.js\")\r\
    SEPARATOR1\r\
    -\"Scripts...\" Call(\"Scripts::Main\")\r\
}\r\
" L"\
\"SESSIONS\"\r\
{\r\
    \"Enable\" +Call(\"Sessions::Main\", 10)\r\
    SEPARATOR1\r\
    -\"Open...\" Call(\"Sessions::Main\")\r\
}\r\
" L"\
\"TEMPLATES\"\r\
{\r\
    \"Enable\" +Call(\"Templates::Main\")\r\
    SEPARATOR1\r\
    \"Open...\" Call(\"Templates::Open\")\r\
}\r\
" L"\
\"EXIT\"\r\
{\r\
    \"Enable\" +Call(\"Exit::Main\")\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"Exit::Settings\")\r\
}\r\
" L"\
\"SMARTSEL\"\r\
{\r\
  \"Smart Home\" +Call(\"SmartSel::SmartHome\")\r\
  \"Option: invert\" +Call(\"SmartSel::altSmartHome\")\r\
  SEPARATOR1\r\
  \"Smart End\" +Call(\"SmartSel::SmartEnd\")\r\
  \"Option: invert\" +Call(\"SmartSel::altSmartEnd\")\r\
  SEPARATOR1\r\
  \"Smart Up/Down\" +Call(\"SmartSel::SmartUpDown\")\r\
  \"Option: +PageUp/PageDown\" +Call(\"SmartSel::altSmartUpDown\")\r\
  SEPARATOR1\r\
  \"Exclude EOL from selection\" +Call(\"SmartSel::NoSelEOL\")\r\
  \"Option: only single line\" +Call(\"SmartSel::altNoSelEOL\")\r\
  SEPARATOR1\r\
  \"Smart Backspace\" +Call(\"SmartSel::SmartBackspace\")\r\
}\r\
" L"\
\"HOTKEYS\"\r\
{\r\
    \"Escape key\" Menu(\"EXIT\") Icon(\"%a\\AkelFiles\\Plugs\\Exit.dll\", 0)\r\
    \"Navigation\" Menu(\"SMARTSEL\")\r\
}\r\
" L"\
\"FORMAT\"\r\
{\r\
    \"Sort lines by string ascending\" Call(\"Format::LineSortStrAsc\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 0)\r\
    \"Sort lines by string descending\" Call(\"Format::LineSortStrDesc\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 1)\r\
    \"Sort lines by integer ascending\" Call(\"Format::LineSortIntAsc\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 2)\r\
    \"Sort lines by integer descending\" Call(\"Format::LineSortIntDesc\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 3)\r\
    SEPARATOR1\r\
    \"Extract unique lines\" Call(\"Format::LineGetUnique\")\r\
    \"Extract duplicate lines\" Call(\"Format::LineGetDuplicates\")\r\
    \"Remove duplicate lines\" Call(\"Format::LineRemoveDuplicates\")\r\
    SEPARATOR1\r\
" L"\
    \"Reverse line order\" Call(\"Format::LineReverse\")\r\
    \"Insert end-of-line character in wrap places\" Call(\"Format::LineFixWrap\")\r\
    SEPARATOR1\r\
    \"Encrypt selection\" Call(\"Format::Encrypt\")\r\
    \"Decrypt selection\" Call(\"Format::Decrypt\")\r\
    SEPARATOR1\r\
    \"Extract links from HTML text\" Call(\"Format::LinkExtract\")\r\
}\r\
" L"\
\"SCROLL\"\r\
{\r\
    \"Vertical synchronization\" Call(\"Scroll::SyncVert\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 1)\r\
    \"Horizontal synchronization\" Call(\"Scroll::SyncHorz\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 0)\r\
    SEPARATOR1\r\
    \"Autoscrolling\" +Call(\"Scroll::AutoScroll\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 2)\r\
    \"No scroll operations\" +Call(\"Scroll::NoScroll\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 3)\r\
    \"Align caret\" +Call(\"Scroll::AlignCaret\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 4)\r\
    \"Autofocus\" +Call(\"Scroll::AutoFocus\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 5)\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"Scroll::Settings\")\r\
}\r\
" L"\
\"MINIMIZETOTRAY\"\r\
{\r\
    \"Always minimize to tray\" +Call(\"MinimizeToTray::Always\")\r\
}\r\
" L"\
\"SOUNDS\"\r\
{\r\
    \"Enable\" +Call(\"Sounds::Main\")\r\
    SEPARATOR1\r\
    \"Settings...\" Call(\"Sounds::Settings\")\r\
}\r\
" L"\
\"BBCODE\"\r\
{\r\
    \"URL=\" Insert(`[URL=\\|]\\s[/URL]`, 1)\r\
    \"QUOTE\" Insert(`[QUOTE]\\s[/QUOTE]`, 1)\r\
    \"QUOTE=\" Insert(`[QUOTE=\"\\|\"]\\s[/QUOTE]`, 1)\r\
    \"CODE\" Insert(`[CODE]\\s[/CODE]`, 1)\r\
    \"MORE=\" Insert(`[MORE=\"\\|\"]\\s[/MORE]`, 1)\r\
    \"IMG\" Insert(`[IMG]\\s[/IMG]`, 1)\r\
    \"b\" Insert(`[b]\\s[/b]`, 1)\r\
    \"i\" Insert(`[i]\\s[/i]`, 1)\r\
}\r";

    //Default favourites main menu
    if (nStringID == STRID_DEFAULTMAIN)
      return L"\
\"F&avourites\" Index(3)\r\
{\r\
    \"Add\" Favourites(1) Icon(0)\r\
    \"Manage...\" Favourites(3) Icon(1)\r\
    SEPARATOR1\r\
    FAVOURITES\r\
}\r\
" L"\
\"&Plugins\" Index(-2)\r\
{\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Coder.dll\")\r\
        \"Programming\" Menu(\"CODER\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 12)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\XBrackets.dll\")\r\
        \"Brackets\" Menu(\"XBRACKETS\") Icon(\"%a\\AkelFiles\\Plugs\\XBrackets.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\SpellCheck.dll\")\r\
        \"Spell check\" Menu(\"SPELLCHECK\") Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 35)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\SpecialChar.dll\")\r\
        \"Special characters\" Menu(\"SPECIALCHAR\") Icon(\"%a\\AkelFiles\\Plugs\\SpecialChar.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\LineBoard.dll\")\r\
        \"Line numbers, bookmarks\" Menu(\"LINEBOARD\") Icon(\"%a\\AkelFiles\\Plugs\\LineBoard.dll\", 0)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\Clipboard.dll\")\r\
        \"Clipboard\" Menu(\"CLIPBOARD\") Icon(\"%a\\AkelFiles\\Plugs\\Clipboard.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\SaveFile.dll\")\r\
        \"File saving\" Menu(\"SAVEFILE\") Icon(\"%a\\AkelFiles\\Plugs\\SaveFile.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Log.dll\")\r\
        \"Log view\" Menu(\"LOG\") Icon(\"%a\\AkelFiles\\Plugs\\Log.dll\", 0)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\ToolBar.dll\")\r\
        \"Toolbar panel\" +Call(\"ToolBar::Main\")\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Explorer.dll\")\r\
        \"Explorer panel\" Menu(\"EXPLORE\") Icon(\"%a\\AkelFiles\\Plugs\\Explorer.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\QSearch.dll\")\r\
        \"Search panel\" +Call(\"QSearch::QSearch\") Icon(\"%a\\AkelFiles\\Plugs\\QSearch.dll\", 0)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\Scripts.dll\")\r\
        -\"Scripts...\" +Call(\"Scripts::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Scripts.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Macros.dll\")\r\
        -\"Macros...\" +Call(\"Macros::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\RecentFiles.dll\")\r\
        -\"Recent files...\" Call(\"RecentFiles::Manage\") Icon(\"%a\\AkelFiles\\Plugs\\RecentFiles.dll\", 0)\r\
    UNSET(32)\r\
    SET(1)\r\
        # MDI/PMDI\r\
        SET(32, \"%a\\AkelFiles\\Plugs\\Sessions.dll\")\r\
            \"Sessions...\" Menu(\"SESSIONS\") Icon(\"%a\\AkelFiles\\Plugs\\Sessions.dll\", 0)\r\
        UNSET(32)\r\
    UNSET(1)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Templates.dll\")\r\
        \"Templates\" Menu(\"TEMPLATES\") Icon(\"%a\\AkelFiles\\Plugs\\Toolbar.dll\", 37)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\Hotkeys.dll\")\r\
        -\"Hotkeys...\" +Call(\"Hotkeys::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Hotkeys.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Exit.dll\")\r\
        \"Escape key\" Menu(\"EXIT\") Icon(\"%a\\AkelFiles\\Plugs\\Exit.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\SmartSel.dll\")\r\
        \"Navigation\" Menu(\"SMARTSEL\")\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\MinimizeToTray.dll\")\r\
        \"Minimize to tray\" Call(\"MinimizeToTray::Now\") Icon(\"%a\\AkelFiles\\Plugs\\MinimizeToTray.dll\", 0)\r\
        \"Always minimize to tray\" +Call(\"MinimizeToTray::Always\")\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\QSearch.dll\")\r\
        \"Quick dialogs switch\" +Call(\"QSearch::DialogSwitcher\") Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 34)\r\
    UNSET(32)\r\
    SEPARATOR1\r\
" L"\
    SET(32, \"%a\\AkelFiles\\Plugs\\Sounds.dll\")\r\
        \"Sound typing\" Menu(\"SOUNDS\") Icon(\"%a\\AkelFiles\\Plugs\\Sounds.dll\", 0)\r\
    UNSET(32)\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Speech.dll\")\r\
        \"Machine reading\" +Call(\"Speech::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Speech.dll\", 0)\r\
    UNSET(32)\r\
}\r";

    //Default edit menu
    if (nStringID == STRID_DEFAULTEDIT)
      return L"\
SET(8)\r\
    \"\" Command(4151) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 10)\r\
    \"\" Command(4152) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 11)\r\
    SEPARATOR1\r\
    \"\" Command(4153) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 7)\r\
    \"\" Command(4154) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 8)\r\
    \"\" Command(4155) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 9)\r\
    \"\" Command(4156) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 25)\r\
    SEPARATOR1\r\
    \"\" Command(4157)\r\
UNSET(8)\r\
" L"\
SEPARATOR1\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Scripts.dll\")\r\
    \"Scripts\" Menu(\"SCRIPTS\") Icon(\"%a\\AkelFiles\\Plugs\\Scripts.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Format.dll\")\r\
    \"Format\" Menu(\"FORMAT\")\r\
UNSET(32)\r\
\"BBCode\" Menu(\"BBCODE\")\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Scroll.dll\")\r\
    \"Scroll\" Menu(\"SCROLL\")\r\
UNSET(32)\r\
SEPARATOR1\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Coder.dll\")\r\
    \"Mark\" Menu(\"MARK\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 0)\r\
    \"Syntax theme\" Menu(\"SYNTAXTHEME\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 4)\r\
    \"Color theme\" Menu(\"COLORTHEME\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 5)\r\
UNSET(32)\r\
SEPARATOR1\r\
SET(32, \"%a\\AkelFiles\\Plugs\\HexSel.dll\")\r\
    \"Hex code\" +Call(\"HexSel::Main\") Icon(\"%a\\AkelFiles\\Plugs\\HexSel.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Stats.dll\")\r\
    \"Statistics\" Call(\"Stats::Main\")\r\
UNSET(32)\r\
SEPARATOR1\r\
SET(32, \"%a\\AkelFiles\\Plugs\\FullScreen.dll\")\r\
    \"Fullscreen mode\" Call(\"FullScreen::Main\") Icon(\"%a\\AkelFiles\\Plugs\\FullScreen.dll\", 0)\r\
UNSET(32)\r";

    //Default tabs menu
    if (nStringID == STRID_DEFAULTTAB)
      return L"\
\"Window list\" Command(4327)\r\
SEPARATOR1\r\
\"Close\" Command(4318)\r\
\"Close all\" Command(4319)\r\
\"Close all but active\" Command(4320)\r\
SEPARATOR1\r\
SET(4)\r\
    #Only for MDI\r\
    \"Tile &Horizontal\" Command(4307)\r\
    \"Tile &Vertical\" Command(4308)\r\
    \"&Cascade\" Command(4309)\r\
    SEPARATOR1\r\
UNSET(4)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Scripts.dll\")\r\
    -\"Copy path\" Call(\"Scripts::Main\", 1, \"EvalCmd.js\", `AkelPad.SetClipboardText(AkelPad.GetEditFile(0));`)\r\
UNSET(32)\r\
\"Explorer menu\"\r\
{\r\
    EXPLORER\r\
}\r";

    //Default URL menu
    if (nStringID == STRID_DEFAULTURL)
      return L"\
\"Open\" Link(1)\r\
\"Copy\" Link(2)\r\
\"Select\" Link(3)\r\
SEPARATOR1\r\
\"Cut\" Link(4)\r\
\"Paste\" Link(5)\r\
\"Delete\" Link(6)\r\
SEPARATOR1\r\
" L"\
SET(8)\r\
    \"\" Command(4151) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 10)\r\
    \"\" Command(4152) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 11)\r\
    SEPARATOR1\r\
    \"\" Command(4153) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 7)\r\
    \"\" Command(4154) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 8)\r\
    \"\" Command(4155) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 9)\r\
    \"\" Command(4156) Icon(\"%a\\AkelFiles\\Plugs\\ToolBar.dll\", 25)\r\
    SEPARATOR1\r\
    \"\" Command(4157)\r\
UNSET(8)\r";

    //Default recent files menu
    if (nStringID == STRID_DEFAULTRECENTFILES)
      return L"\
EXPLORER\r";

  }
  return L"";
}

BOOL IsExtCallParamValid(LPARAM lParam, int nIndex)
{
  if (*((INT_PTR *)lParam) >= (INT_PTR)((nIndex + 1) * sizeof(INT_PTR)))
    return TRUE;
  return FALSE;
}

INT_PTR GetExtCallParam(LPARAM lParam, int nIndex)
{
  if (*((INT_PTR *)lParam) >= (INT_PTR)((nIndex + 1) * sizeof(INT_PTR)))
    return *(((INT_PTR *)lParam) + nIndex);
  return 0;
}

void InitCommon(PLUGINDATA *pd)
{
  bInitCommon=TRUE;
  hInstanceDLL=pd->hInstanceDLL;
  hMainWnd=pd->hMainWnd;
  hMdiClient=pd->hMdiClient;
  hMainMenu=pd->hMainMenu;
  hMenuRecentFiles=pd->hMenuRecentFiles;
  hMenuLanguage=pd->hMenuLanguage;
  hMainIcon=pd->hMainIcon;
  hGlobalAccel=pd->hGlobalAccel;
  bOldWindows=pd->bOldWindows;
  bAkelEdit=pd->bAkelEdit;
  nMDI=pd->nMDI;
  wLangModule=PRIMARYLANGID(pd->wLangModule);
  hMenuWindow=GetSubMenu(hMainMenu, MENU_MDI_POSITION);
  hPopupEdit=GetSubMenu(pd->hPopupMenu, MENU_POPUP_EDIT);
  hPopupCodepage=GetSubMenu(pd->hPopupMenu, MENU_POPUP_CODEPAGE);
  hPopupOpenCodepage=GetSubMenu(hPopupCodepage, MENU_POPUP_CODEPAGE_OPEN);
  hPopupSaveCodepage=GetSubMenu(hPopupCodepage, MENU_POPUP_CODEPAGE_SAVE);
  hHeap=GetProcessHeap();

  //Initialize WideFunc.h header
  WideInitialize();

  //Plugin name
  {
    int i;

    for (i=0; pd->wszFunction[i] != L':'; ++i)
      wszPluginName[i]=pd->wszFunction[i];
    wszPluginName[i]=L'\0';
  }
  xprintfW(wszPluginTitle, GetLangStringW(wLangModule, STRID_PLUGIN), wszPluginName);
  xstrcpynW(wszExeDir, pd->wszAkelDir, MAX_PATH);
  ReadOptions(0);

  CreateFavStack(&hFavStack, wszFavText);
  HeapFree(hHeap, 0, wszFavText);
  wszFavText=NULL;

  if (bBigIcons)
  {
    sizeIcon.cx=32;
    sizeIcon.cy=32;
  }
  else
  {
    sizeIcon.cx=16;
    sizeIcon.cy=16;
  }
}

void InitMain()
{
  bInitMain=TRUE;

  //SubClass
  NewMainProcData=NULL;
  SendMessage(hMainWnd, AKD_SETMAINPROC, (WPARAM)NewMainProc, (LPARAM)&NewMainProcData);

  if (nMDI == WMD_MDI)
  {
    NewFrameProcData=NULL;
    SendMessage(hMainWnd, AKD_SETFRAMEPROC, (WPARAM)NewFrameProc, (LPARAM)&NewFrameProcData);
  }

  hIconFav=LoadIconA(hInstanceDLL, MAKEINTRESOURCEA(IDI_ICON_FAVMANAGE));
  hIconHollow=LoadIconA(hInstanceDLL, MAKEINTRESOURCEA(IDI_ICON_HOLLOW));
}

void UninitMain()
{
  bInitMain=FALSE;

  if (hIconFav)
  {
    DestroyIcon(hIconFav);
    hIconFav=NULL;
  }
  if (hIconHollow)
  {
    DestroyIcon(hIconHollow);
    hIconHollow=NULL;
  }

  //Remove subclass
  if (NewMainProcData)
  {
    SendMessage(hMainWnd, AKD_SETMAINPROC, (WPARAM)NULL, (LPARAM)&NewMainProcData);
    NewMainProcData=NULL;
  }
  if (NewFrameProcData)
  {
    SendMessage(hMainWnd, AKD_SETFRAMEPROC, (WPARAM)NULL, (LPARAM)&NewFrameProcData);
    NewFrameProcData=NULL;
  }

  //Save OF_FAVTEXT
  if (dwSaveFlags)
  {
    SaveOptions(dwSaveFlags);
    dwSaveFlags=0;
  }

  if (wszManualText)
  {
    HeapFree(hHeap, 0, wszManualText);
    wszManualText=NULL;
  }
  if (wszMainText)
  {
    HeapFree(hHeap, 0, wszMainText);
    wszMainText=NULL;
  }
  if (wszEditText)
  {
    HeapFree(hHeap, 0, wszEditText);
    wszEditText=NULL;
  }
  if (wszTabText)
  {
    HeapFree(hHeap, 0, wszTabText);
    wszTabText=NULL;
  }
  if (wszUrlText)
  {
    HeapFree(hHeap, 0, wszUrlText);
    wszUrlText=NULL;
  }
  if (wszRecentFilesText)
  {
    HeapFree(hHeap, 0, wszRecentFilesText);
    wszRecentFilesText=NULL;
  }
  UnsetMainMenu(&hMenuMainStack);
  FreeContextMenu(&hMenuMainStack);
  FreeContextMenu(&hMenuEditStack);
  FreeContextMenu(&hMenuTabStack);
  FreeContextMenu(&hMenuUrlStack);
  FreeContextMenu(&hMenuRecentFilesStack);
  FreeContextMenu(&hMenuManualStack);

  StackFreeFavourites(&hFavStack);
}

//Entry point
BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
  if (fdwReason == DLL_PROCESS_ATTACH)
  {
  }
  else if (fdwReason == DLL_THREAD_ATTACH)
  {
  }
  else if (fdwReason == DLL_THREAD_DETACH)
  {
  }
  else if (fdwReason == DLL_PROCESS_DETACH)
  {
  }
  return TRUE;
}
