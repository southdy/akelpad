#define WIN32_LEAN_AND_MEAN
#define _WIN32_IE 0x0500
#include <windows.h>
#include <commctrl.h>
#include <richedit.h>
#include "AkelEdit.h"
#include "AkelDLL.h"
#include "StackFunc.h"
#include "StrFunc.h"
#include "WideFunc.h"
#include "Resources\Resource.h"


/*
//Include stack functions
#define StackGetElement
#define StackInsertAfter
#define StackInsertBefore
#define StackInsertIndex
#define StackDelete
#define StackClear
#define StackSize
#include "StackFunc.h"

//Include string functions
#define WideCharLower
#define xmemcpy
#define xmemset
#define xmemcmp
#define xstrcmpW
#define xstrcmpiW
#define xstrcmpnW
#define xstrlenW
#define xstrcpyW
#define xstrcpynW
#define xatoiW
#define xitoaA
#define xitoaW
#define xuitoaW
#define dec2hexW
#define hex2decW
#define xprintfW
#include "StrFunc.h"

//Include wide functions
#define CallWindowProcWide
#define CreateDialogWide
#define CreateProcessWide
#define CreateWindowExWide
#define DefWindowProcWide
#define DispatchMessageWide
#define ExpandEnvironmentStringsWide
#define ExtractIconExWide
#define GetMenuStringWide
#define GetMessageWide
#define GetWindowLongPtrWide
#define GetWindowTextWide
#define InsertMenuWide
#define IsDialogMessageWide
#define RegisterClassWide
#define SearchPathWide
#define SetDlgItemTextWide
#define SetWindowLongPtrWide
#define SetWindowTextWide
#define TranslateAcceleratorWide
#define UnregisterClassWide
#include "WideFunc.h"
//*/

//Defines
#define DLLA_TOOLBAR_ROWS      1

#define STRID_SETUP                          1
#define STRID_AUTOLOAD                       2
#define STRID_BIGICONS                       3
#define STRID_FLATBUTTONS                    4
#define STRID_16BIT                          5
#define STRID_32BIT                          6
#define STRID_SIDE                           7
#define STRID_ROWS                           8
#define STRID_PARSEMSG_UNKNOWNMETHOD         9
#define STRID_PARSEMSG_METHODALREADYDEFINED  10
#define STRID_PARSEMSG_NOMETHOD              11
#define STRID_PLUGIN                         12
#define STRID_OK                             13
#define STRID_CANCEL                         14
#define STRID_DEFAULTMENU                    15

#define AKDLL_RECREATE        (WM_USER + 100)
#define AKDLL_REFRESH         (WM_USER + 101)
#define AKDLL_SETUP           (WM_USER + 102)
#define AKDLL_SELTEXT         (WM_USER + 103)
#define AKDLL_FREEMESSAGELOOP (WM_USER + 150)

#define OF_LISTTEXT       0x1
#define OF_SETTINGS       0x2
#define OF_RECT           0x4

#define BUFFER_SIZE       1024

#define EXTACT_COMMAND    1
#define EXTACT_CALL       2
#define EXTACT_EXEC       3
#define EXTACT_OPENFILE   4
#define EXTACT_SAVEFILE   5
#define EXTACT_FONT       6
#define EXTACT_RECODE     7
#define EXTACT_INSERT     8
#define EXTACT_MENU       9

#define EXTPARAM_CHAR     1
#define EXTPARAM_INT      2

#define FS_NONE            0
#define FS_FONTNORMAL      1
#define FS_FONTBOLD        2
#define FS_FONTITALIC      3
#define FS_FONTBOLDITALIC  4

//CreateToolbarData SET() flags
#define CCMS_NOSDI       0x01
#define CCMS_NOMDI       0x02
#define CCMS_NOPMDI      0x04
#define CCMS_NOFILEEXIST 0x20

#define MAX_TOOLBARTEXT_SIZE  64000

//Toolbar side
#define TBSIDE_LEFT    1
#define TBSIDE_RIGHT   2
#define TBSIDE_TOP     3
#define TBSIDE_BOTTOM  4

//Toolbar side priority
#define TSP_TOPBOTTOM  1
#define TSP_LEFTRIGHT  2

#define IMENU_EDIT     0x00000001
#define IMENU_CHECKS   0x00000004

#define TOOLBARBACKGROUNDA   "ToolbarBG"
#define TOOLBARBACKGROUNDW  L"ToolbarBG"

#define ROWSHOW_UNCHANGE -2
#define ROWSHOW_INVERT   -1
#define ROWSHOW_OFF       0
#define ROWSHOW_ON        1

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

typedef struct _TOOLBARITEM {
  struct _TOOLBARITEM *next;
  struct _TOOLBARITEM *prev;
  BOOL bUpdateItem;
  BOOL bAutoLoad;
  DWORD dwAction;
  int nTextOffset;
  TBBUTTON tbb;
  int nButtonWidth;
  wchar_t wszButtonItem[MAX_PATH];
  STACKEXTPARAM hParamStack;
  STACKEXTPARAM hParamMenuName;
} TOOLBARITEM;

typedef struct {
  TOOLBARITEM *first;
  TOOLBARITEM *last;
  HIMAGELIST hImageList;
} TOOLBARDATA;

typedef struct _ROWITEM {
  struct _ROWITEM *next;
  struct _ROWITEM *prev;
  int nRow;
  TOOLBARITEM *lpFirstToolbarItem;
} ROWITEM;

typedef struct {
  ROWITEM *first;
  ROWITEM *last;
  int nElements;
} STACKROW;

//ContextMenu::Show external call
typedef struct {
  UINT_PTR dwStructSize;
  INT_PTR nAction;
  unsigned char *pPosX;
  unsigned char *pPosY;
  INT_PTR nMenuIndex;
  unsigned char *pMenuName;
} DLLEXTCONTEXTMENU;

#define DLLA_CONTEXTMENU_SHOWSUBMENU   1

//Functions prototypes
BOOL CALLBACK MainDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK NewEditDlgProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
DWORD WINAPI ThreadProc(LPVOID lpParameter);
LRESULT CALLBACK NewMainProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void CALLBACK NewMainProcRet(CWPRETSTRUCT *cwprs);
LRESULT CALLBACK ToolbarBGProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK NewToolbarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

BOOL CreateToolbarWindow();
BOOL CreateToolbarData(TOOLBARDATA *hToolbarData, const wchar_t *wpText);
DWORD IsFlagOn(DWORD dwSetFlags, DWORD dwCheckFlags);
int ParseRows(STACKROW *lpRowListStack);
ROWITEM* GetRow(STACKROW *lpRowListStack, int nRow);
TOOLBARITEM* GetFirstToolbarItemOfNextRow(ROWITEM *lpRowItem);
void FreeRows(STACKROW *lpRowListStack);
void FreeToolbarData(TOOLBARDATA *hToolbarData);
void SetToolbarButtons(TOOLBARDATA *hToolbarData);
void ClearToolbarButtons();
void UpdateToolbar(TOOLBARDATA *hToolbarData);
void ViewItemCode(TOOLBARITEM *lpButton);
void CallToolbar(TOOLBARDATA *hToolbarData, int nItem);
void CallContextMenuShow(TOOLBARITEM *lpButton, int nPosX, int nPosY);
void DestroyToolbarWindow(BOOL bDestroyBG);
TOOLBARITEM* StackInsertBeforeButton(TOOLBARDATA *hToolbarData, TOOLBARITEM *lpInsertBefore);
TOOLBARITEM* StackGetButtonByID(TOOLBARDATA *hToolbarData, int nItemID);
TOOLBARITEM* StackGetButtonByIndex(TOOLBARDATA *hToolbarData, int nIndex);
void StackFreeButton(TOOLBARDATA *hToolbarData);

void ParseMethodParameters(STACKEXTPARAM *hParamStack, const wchar_t *wpText, const wchar_t **wppText);
void ExpandMethodParameters(STACKEXTPARAM *hParamStack, const wchar_t *wpFile, const wchar_t *wpExeDir, HWND hToolbar, int nButtonID, const RECT *lprcButton);
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
INT_PTR GetEditText(HWND hWnd, wchar_t **Text);
void ShowStandardEditMenu(HWND hWnd, HMENU hMenu, BOOL bMouse);
DWORD ScrollCaret(HWND hWnd);

void ReadOptions(DWORD dwFlags);
void SaveOptions(DWORD dwFlags);
const char* GetLangStringA(LANGID wLangID, int nStringID);
const wchar_t* GetLangStringW(LANGID wLangID, int nStringID);
BOOL IsExtCallParamValid(LPARAM lParam, int nIndex);
INT_PTR GetExtCallParam(LPARAM lParam, int nIndex);
BOOL SetExtCallParam(LPARAM lParam, int nIndex, INT_PTR nData);
void InitCommon(PLUGINDATA *pd);
void InitMain();
void UninitMain();

//Global variables
char szBuffer[BUFFER_SIZE];
wchar_t wszBuffer[BUFFER_SIZE];
wchar_t wszPluginName[MAX_PATH];
wchar_t wszPluginTitle[MAX_PATH];
HANDLE hHeap;
HINSTANCE hInstanceEXE;
HINSTANCE hInstanceDLL;
HWND hMainWnd;
HWND hMdiClient;
HMENU hMainMenu;
HMENU hMenuRecentFiles;
HMENU hPopupEdit;
HICON hMainIcon;
HACCEL hGlobalAccel;
BOOL bOldWindows;
BOOL bAkelEdit;
int nMDI;
LANGID wLangModule;
BOOL bInitCommon=FALSE;
BOOL bInitMain=FALSE;
DWORD dwSaveFlags=0;
char szExeDir[MAX_PATH];
wchar_t wszExeDir[MAX_PATH];
char *szToolBarText=NULL;
wchar_t *wszToolBarText=NULL;
TOOLBARDATA hToolbarData={0};
wchar_t wszRowList[MAX_PATH]=L"";
STACKROW hRowListStack={0};
HWND hToolbarBG=NULL;
HWND hToolbar=NULL;
BOOL bBigIcons=FALSE;
BOOL bFlatButtons=TRUE;
int nIconsBit=32;
int nToolbarSide=TBSIDE_TOP;
int nSidePriority=TSP_TOPBOTTOM;
SIZE sizeToolbar={0};
SIZE sizeButtons={0};
CHARRANGE64 crExtSetSel={0};
HANDLE hThread=NULL;
DWORD dwThreadId;
HWND hWndMainDlg=NULL;
RECT rcMainMinMaxDialog={532, 174, 0, 0};
RECT rcMainCurrentDialog={0};
WNDPROC lpOldToolbarProc=NULL;
WNDPROC lpOldEditDlgProc=NULL;
WNDPROCDATA *NewMainProcData=NULL;
WNDPROCRETDATA *NewMainProcRetData=NULL;
WNDPROCRETDATA *NewFrameProcRetData=NULL;


//Identification
void __declspec(dllexport) DllAkelPadID(PLUGINVERSION *pv)
{
  pv->dwAkelDllVersion=AKELDLL;
  pv->dwExeMinVersion3x=MAKE_IDENTIFIER(-1, -1, -1, -1);
  pv->dwExeMinVersion4x=MAKE_IDENTIFIER(4, 8, 8, 0);
  pv->pPluginName="ToolBar";
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

    if (nAction == DLLA_TOOLBAR_ROWS)
    {
      unsigned char *pRows=NULL;
      wchar_t *wpRows;
      int nCompare;

      if (IsExtCallParamValid(pd->lParam, 2))
        pRows=(unsigned char *)GetExtCallParam(pd->lParam, 2);

      if (pRows)
      {
        if (pd->dwSupport & PDS_STRANSI)
          wpRows=AllocWide((char *)pRows);
        else
          wpRows=(wchar_t *)pRows;
        nCompare=xstrcmpW(wpRows, wszRowList);
        xstrcpynW(wszRowList, wpRows, MAX_PATH);
        if (pd->dwSupport & PDS_STRANSI)
          FreeWide(wpRows);

        if (nCompare)
        {
          dwSaveFlags|=OF_SETTINGS;

          if (bInitMain)
          {
            PostMessage(hToolbarBG, AKDLL_RECREATE, 0, 0);

            //If plugin already loaded, stay in memory and don't change active status
            if (pd->bInMemory) pd->nUnload=UD_NONUNLOAD_UNCHANGE;
            return;
          }
        }
      }
    }
  }

  //Is plugin already loaded?
  if (bInitMain)
  {
    UninitMain();

    SendMessage(hMainWnd, AKD_RESIZE, 0, 0);
  }
  else
  {
    InitMain();

    if (!pd->bOnStart)
    {
      CreateToolbarWindow();
      SendMessage(hMainWnd, AKD_RESIZE, 0, 0);
    }

    //Stay in memory, and show as active
    pd->nUnload=UD_NONUNLOAD_ACTIVE;
  }
}

BOOL CALLBACK MainDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndToolBarText;
  static HWND hWndBigIconsCheck;
  static HWND hWndFlatButtonsCheck;
  static HWND hWndBit16;
  static HWND hWndBit32;
  static HWND hWndSideLeft;
  static HWND hWndSideTop;
  static HWND hWndSideRight;
  static HWND hWndSideBottom;
  static HWND hWndSideLabel;
  static HWND hWndRowsLabel;
  static HWND hWndRowsEdit;
  static HWND hWndOK;
  static HWND hWndCancel;
  static DIALOGRESIZE drs[]={{&hWndToolBarText,      DRS_SIZE|DRS_X, 0},
                             {&hWndToolBarText,      DRS_SIZE|DRS_Y, 0},
                             {&hWndBigIconsCheck,    DRS_MOVE|DRS_Y, 0},
                             {&hWndFlatButtonsCheck, DRS_MOVE|DRS_Y, 0},
                             {&hWndBit16,            DRS_MOVE|DRS_Y, 0},
                             {&hWndBit32,            DRS_MOVE|DRS_Y, 0},
                             {&hWndSideLeft,         DRS_MOVE|DRS_Y, 0},
                             {&hWndSideTop,          DRS_MOVE|DRS_Y, 0},
                             {&hWndSideRight,        DRS_MOVE|DRS_Y, 0},
                             {&hWndSideBottom,       DRS_MOVE|DRS_Y, 0},
                             {&hWndSideLabel,        DRS_MOVE|DRS_Y, 0},
                             {&hWndRowsLabel,        DRS_MOVE|DRS_Y, 0},
                             {&hWndRowsEdit,         DRS_MOVE|DRS_Y, 0},
                             {&hWndOK,               DRS_MOVE|DRS_X, 0},
                             {&hWndOK,               DRS_MOVE|DRS_Y, 0},
                             {&hWndCancel,           DRS_MOVE|DRS_X, 0},
                             {&hWndCancel,           DRS_MOVE|DRS_Y, 0},
                             {0, 0, 0}};

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndToolBarText=GetDlgItem(hDlg, IDC_TOOLBARTEXT);
    hWndBigIconsCheck=GetDlgItem(hDlg, IDC_BIGICONS);
    hWndFlatButtonsCheck=GetDlgItem(hDlg, IDC_FLATBUTTONS);
    hWndBit16=GetDlgItem(hDlg, IDC_16BIT);
    hWndBit32=GetDlgItem(hDlg, IDC_32BIT);
    hWndSideLeft=GetDlgItem(hDlg, IDC_SIDELEFT);
    hWndSideTop=GetDlgItem(hDlg, IDC_SIDETOP);
    hWndSideRight=GetDlgItem(hDlg, IDC_SIDERIGHT);
    hWndSideBottom=GetDlgItem(hDlg, IDC_SIDEBOTTOM);
    hWndSideLabel=GetDlgItem(hDlg, IDC_SIDE_LABEL);
    hWndRowsLabel=GetDlgItem(hDlg, IDC_ROWS_LABEL);
    hWndRowsEdit=GetDlgItem(hDlg, IDC_ROWS);
    hWndOK=GetDlgItem(hDlg, IDOK);
    hWndCancel=GetDlgItem(hDlg, IDCANCEL);

    SetWindowTextWide(hDlg, wszPluginTitle);
    SetDlgItemTextWide(hDlg, IDC_BIGICONS, GetLangStringW(wLangModule, STRID_BIGICONS));
    SetDlgItemTextWide(hDlg, IDC_FLATBUTTONS, GetLangStringW(wLangModule, STRID_FLATBUTTONS));
    SetDlgItemTextWide(hDlg, IDC_16BIT, GetLangStringW(wLangModule, STRID_16BIT));
    SetDlgItemTextWide(hDlg, IDC_32BIT, GetLangStringW(wLangModule, STRID_32BIT));
    SetDlgItemTextWide(hDlg, IDC_SIDE_LABEL, GetLangStringW(wLangModule, STRID_SIDE));
    SetDlgItemTextWide(hDlg, IDC_ROWS_LABEL, GetLangStringW(wLangModule, STRID_ROWS));
    SetDlgItemTextWide(hDlg, IDOK, GetLangStringW(wLangModule, STRID_OK));
    SetDlgItemTextWide(hDlg, IDCANCEL, GetLangStringW(wLangModule, STRID_CANCEL));

    SendMessage(hWndToolBarText, AEM_SETEVENTMASK, 0, AENM_SELCHANGE|AENM_TEXTCHANGE|AENM_TEXTINSERT|AENM_TEXTDELETE|AENM_POINT);
    SendMessage(hWndToolBarText, EM_SETEVENTMASK, 0, ENM_SELCHANGE|ENM_CHANGE);
    SendMessage(hWndToolBarText, EM_EXLIMITTEXT, 0, MAX_TOOLBARTEXT_SIZE);
    SetWindowTextWide(hWndToolBarText, wszToolBarText);

    if (bBigIcons)
      SendMessage(hWndBigIconsCheck, BM_SETCHECK, BST_CHECKED, 0);
    if (bFlatButtons)
      SendMessage(hWndFlatButtonsCheck, BM_SETCHECK, BST_CHECKED, 0);
    if (nIconsBit == 16)
      SendMessage(hWndBit16, BM_SETCHECK, BST_CHECKED, 0);
    else
      SendMessage(hWndBit32, BM_SETCHECK, BST_CHECKED, 0);

    if (nToolbarSide == TBSIDE_LEFT)
      SendMessage(hWndSideLeft, BM_SETCHECK, BST_CHECKED, 0);
    else if (nToolbarSide == TBSIDE_TOP)
      SendMessage(hWndSideTop, BM_SETCHECK, BST_CHECKED, 0);
    else if (nToolbarSide == TBSIDE_RIGHT)
      SendMessage(hWndSideRight, BM_SETCHECK, BST_CHECKED, 0);
    else if (nToolbarSide == TBSIDE_BOTTOM)
      SendMessage(hWndSideBottom, BM_SETCHECK, BST_CHECKED, 0);

    SetWindowTextWide(hWndRowsEdit, wszRowList);

    lpOldEditDlgProc=(WNDPROC)GetWindowLongPtrWide(hWndToolBarText, GWLP_WNDPROC);
    SetWindowLongPtrWide(hWndToolBarText, GWLP_WNDPROC, (UINT_PTR)NewEditDlgProc);

    //Post AKDLL_SELTEXT because dialog size can be changed after AKD_DIALOGRESIZE
    PostMessage(hDlg, AKDLL_SELTEXT, 0, 0);
  }
  else if (uMsg == AKDLL_FREEMESSAGELOOP)
  {
    MSG msg;

    while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
    {
      if (hWndMainDlg && !IsDialogMessageWide(hWndMainDlg, &msg))
      {
        TranslateMessage(&msg);
        DispatchMessageWide(&msg);
      }
    }
  }
  else if (uMsg == AKDLL_SELTEXT)
  {
    if (crExtSetSel.cpMin || crExtSetSel.cpMax)
    {
      int nLine;
      int nLockScroll=0;

      if (crExtSetSel.cpMax == -1)
      {
        nLine=(int)SendMessage(hWndToolBarText, EM_EXLINEFROMCHAR, 0, crExtSetSel.cpMin);
        crExtSetSel.cpMax=SendMessage(hWndToolBarText, EM_LINEINDEX, nLine, 0) + SendMessage(hWndToolBarText, EM_LINELENGTH, crExtSetSel.cpMin, 0);
      }

      if (bAkelEdit)
      {
        if ((nLockScroll=(int)SendMessage(hWndToolBarText, AEM_LOCKSCROLL, (WPARAM)-1, 0)) == -1)
          SendMessage(hWndToolBarText, AEM_LOCKSCROLL, SB_BOTH, TRUE);
      }
      SendMessage(hWndToolBarText, EM_SETSEL, crExtSetSel.cpMax, crExtSetSel.cpMin);
      if (bAkelEdit && nLockScroll == -1)
      {
        SendMessage(hWndToolBarText, AEM_LOCKSCROLL, SB_BOTH, FALSE);
        ScrollCaret(hWndToolBarText);
      }
      crExtSetSel.cpMin=0;
      crExtSetSel.cpMax=0;
    }
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
    {
      BOOL bUpdate=FALSE;

      if (LOWORD(wParam) == IDOK)
      {
        TOOLBARDATA hTestStack={0};
        wchar_t *wszTest;
        int nValue;

        //Big icons
        if (SendMessage(hWndBigIconsCheck, BM_GETCHECK, 0, 0) != bBigIcons)
        {
          bBigIcons=!bBigIcons;
          bUpdate=TRUE;
        }

        //Flat buttons
        if (SendMessage(hWndFlatButtonsCheck, BM_GETCHECK, 0, 0) != bFlatButtons)
        {
          bFlatButtons=!bFlatButtons;

          nValue=(int)SendMessage(hToolbar, TB_GETSTYLE, 0, 0);
          if (bFlatButtons)
            SendMessage(hToolbar, TB_SETSTYLE, 0, nValue | TBSTYLE_FLAT);
          else
            SendMessage(hToolbar, TB_SETSTYLE, 0, nValue & ~TBSTYLE_FLAT);

          InvalidateRect(hToolbar, NULL, TRUE);
        }

        //Icons bit
        if (SendMessage(hWndBit16, BM_GETCHECK, 0, 0) == BST_CHECKED)
          nValue=16;
        else
          nValue=32;
        if (nIconsBit != nValue)
        {
          nIconsBit=nValue;
          bUpdate=TRUE;
        }

        //Toolbar side
        if (SendMessage(hWndSideLeft, BM_GETCHECK, 0, 0) == BST_CHECKED)
          nValue=TBSIDE_LEFT;
        else if (SendMessage(hWndSideTop, BM_GETCHECK, 0, 0) == BST_CHECKED)
          nValue=TBSIDE_TOP;
        else if (SendMessage(hWndSideRight, BM_GETCHECK, 0, 0) == BST_CHECKED)
          nValue=TBSIDE_RIGHT;
        else if (SendMessage(hWndSideBottom, BM_GETCHECK, 0, 0) == BST_CHECKED)
          nValue=TBSIDE_BOTTOM;
        if (nToolbarSide != nValue)
        {
          nToolbarSide=nValue;
          bUpdate=TRUE;
        }

        if (SendMessage(hWndRowsEdit, EM_GETMODIFY, 0, 0))
        {
          GetWindowTextWide(hWndRowsEdit, wszRowList, MAX_PATH);
          bUpdate=TRUE;
        }

        if (SendMessage(hWndToolBarText, EM_GETMODIFY, 0, 0))
        {
          GetEditText(hWndToolBarText, &wszTest);

          //Test for errors
          if (!CreateToolbarData(&hTestStack, wszTest))
          {
            FreeToolbarData(&hTestStack);
            HeapFree(hHeap, 0, wszTest);
            return FALSE;
          }

          //Success
          FreeToolbarData(&hToolbarData);
          HeapFree(hHeap, 0, wszToolBarText);
          wszToolBarText=wszTest;
          hToolbarData=hTestStack;
          bUpdate=TRUE;
        }
        dwSaveFlags|=OF_LISTTEXT|OF_SETTINGS;
      }
      if (dwSaveFlags)
      {
        SaveOptions(dwSaveFlags);
        dwSaveFlags=0;
      }

      DestroyWindow(hWndMainDlg);
      hWndMainDlg=NULL;
      if (bUpdate)
        PostMessage(hToolbarBG, AKDLL_RECREATE, 0, 0);

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

  //Dialog resize messages
  {
    DIALOGRESIZEMSG drsm={&drs[0], &rcMainMinMaxDialog, &rcMainCurrentDialog, DRM_PAINTSIZEGRIP, hDlg, uMsg, wParam, lParam};

    if (SendMessage(hMainWnd, AKD_DIALOGRESIZE, 0, (LPARAM)&drsm))
      dwSaveFlags|=OF_RECT;
  }

  return FALSE;
}

LRESULT CALLBACK NewEditDlgProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  LRESULT lResult;

  if (uMsg == WM_GETDLGCODE)
  {
    lResult=CallWindowProcWide(lpOldEditDlgProc, hWnd, uMsg, wParam, lParam);

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

  return CallWindowProcWide(lpOldEditDlgProc, hWnd, uMsg, wParam, lParam);
}

DWORD WINAPI ThreadProc(LPVOID lpParameter)
{
  MSG msg;

  OleInitialize(NULL);

  hWndMainDlg=CreateDialogWide(hInstanceDLL, MAKEINTRESOURCEW(IDD_SETUP), hMainWnd, (DLGPROC)MainDlgProc);

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
  if (uMsg == AKDN_MAIN_ONSTART_PRESHOW)
  {
    CreateToolbarWindow();
  }
  else if (uMsg == AKDN_SIZE_ONSTART)
  {
    NSIZE *ns=(NSIZE *)lParam;
    LRESULT lResult=0;
    BOOL bProcess=FALSE;

    if (hToolbarBG && hToolbar && IsWindowVisible(hToolbar))
    {
      if ((nSidePriority == TSP_TOPBOTTOM && (nToolbarSide == TBSIDE_LEFT ||
                                              nToolbarSide == TBSIDE_RIGHT)) ||
          (nSidePriority == TSP_LEFTRIGHT && (nToolbarSide == TBSIDE_TOP ||
                                              nToolbarSide == TBSIDE_BOTTOM)))
      {
        //Give more priority.
        lResult=NewMainProcData->NextProc(hWnd, uMsg, wParam, lParam);
        bProcess=TRUE;
      }

      if (nToolbarSide == TBSIDE_TOP)
      {
        MoveWindow(hToolbarBG, ns->rcCurrent.left, ns->rcCurrent.top, ns->rcCurrent.right, sizeToolbar.cy, TRUE);
        MoveWindow(hToolbar, 0, 0, ns->rcCurrent.right, sizeToolbar.cy, TRUE);
        ns->rcCurrent.top+=sizeToolbar.cy;
        ns->rcCurrent.bottom-=sizeToolbar.cy;
      }
      else if (nToolbarSide == TBSIDE_BOTTOM)
      {
        MoveWindow(hToolbarBG, ns->rcCurrent.left, ns->rcCurrent.top + ns->rcCurrent.bottom - sizeToolbar.cy, ns->rcCurrent.right, sizeToolbar.cy, TRUE);
        MoveWindow(hToolbar, 0, 0, ns->rcCurrent.right, sizeToolbar.cy, TRUE);
        ns->rcCurrent.bottom-=sizeToolbar.cy;
      }
      else if (nToolbarSide == TBSIDE_LEFT)
      {
        MoveWindow(hToolbarBG, ns->rcCurrent.left, ns->rcCurrent.top, sizeToolbar.cx, ns->rcCurrent.bottom, TRUE);
        MoveWindow(hToolbar, 0, 0, sizeToolbar.cx, ns->rcCurrent.bottom, TRUE);
        ns->rcCurrent.left+=sizeToolbar.cx;
        ns->rcCurrent.right-=sizeToolbar.cx;
      }
      else if (nToolbarSide == TBSIDE_RIGHT)
      {
        MoveWindow(hToolbarBG, ns->rcCurrent.left + ns->rcCurrent.right - sizeToolbar.cx, ns->rcCurrent.top, sizeToolbar.cx, ns->rcCurrent.bottom, TRUE);
        MoveWindow(hToolbar, 0, 0, sizeToolbar.cx, ns->rcCurrent.bottom, TRUE);
        ns->rcCurrent.right-=sizeToolbar.cx;
      }
      if (bProcess) return lResult;
    }
  }
  else if (uMsg == AKDN_MAIN_ONFINISH)
  {
    NewMainProcData->NextProc(hWnd, uMsg, wParam, lParam);
    if (hWndMainDlg) SendMessage(hWndMainDlg, WM_COMMAND, IDCANCEL, 0);
    UninitMain();
    return FALSE;
  }

  //Call next procedure
  return NewMainProcData->NextProc(hWnd, uMsg, wParam, lParam);
}

void CALLBACK NewMainProcRet(CWPRETSTRUCT *cwprs)
{
  if (cwprs->message == AKDN_DLLCALL ||
      cwprs->message == AKDN_DLLUNLOAD ||
      cwprs->message == AKDN_FRAME_NOWINDOWS ||
      (cwprs->message == AKDN_FRAME_ACTIVATE && !(cwprs->wParam & FWA_NOVISUPDATE)))
  {
    PostMessage(hToolbarBG, AKDLL_REFRESH, 0, 0);
  }
  else if (cwprs->message == WM_COMMAND)
  {
    if (cwprs->hwnd == hMainWnd)
    {
      PostMessage(hToolbarBG, AKDLL_REFRESH, 0, 0);
    }
  }
  else if (cwprs->message == WM_NOTIFY)
  {
    if (cwprs->wParam == ID_EDIT)
    {
      if (bAkelEdit)
      {
        if (((NMHDR *)cwprs->lParam)->code == AEN_MODIFY)
        {
          PostMessage(hToolbarBG, AKDLL_REFRESH, 0, 0);
        }
      }
      if (((NMHDR *)cwprs->lParam)->code == EN_SELCHANGE)
      {
        PostMessage(hToolbarBG, AKDLL_REFRESH, 0, 0);
      }
    }

    if (!bAkelEdit)
    {
      if (cwprs->wParam == ID_STATUS)
      {
        if (((NMHDR *)cwprs->lParam)->code == (UINT)NM_DBLCLK)
        {
          if (((NMMOUSE *)cwprs->lParam)->dwItemSpec == 3)
          {
            PostMessage(hToolbarBG, AKDLL_REFRESH, 0, 0);
          }
        }
      }
    }
  }

  //Call next procedure
  if (cwprs->hwnd == hMainWnd)
  {
    if (NewMainProcRetData->NextProc)
      NewMainProcRetData->NextProc(cwprs);
  }
  else
  {
    if (NewFrameProcRetData->NextProc)
      NewFrameProcRetData->NextProc(cwprs);
  }
}

LRESULT CALLBACK ToolbarBGProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  if (uMsg == AKDLL_RECREATE)
  {
    DestroyToolbarWindow(FALSE);
    CreateToolbarWindow();
    SendMessage(hMainWnd, AKD_RESIZE, 0, 0);
  }
  else if (uMsg == AKDLL_REFRESH)
  {
    if (hToolbar) UpdateToolbar(&hToolbarData);
  }
  else if (uMsg == AKDLL_SETUP)
  {
    if (!hThread)
    {
      hThread=CreateThread(NULL, 0, ThreadProc, NULL, 0, &dwThreadId);
    }
  }
  else if (uMsg == WM_ERASEBKGND)
  {
    return 0;
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (((NMHDR *)lParam)->code == TTN_GETDISPINFOA)
    {
      TOOLBARITEM *lpButton;

      if (lpButton=StackGetButtonByID(&hToolbarData, (int)((NMHDR *)lParam)->idFrom))
      {
        //WideCharToMultiByte(CP_ACP, 0, lpButton->wszButtonItem, -1, ((NMTTDISPINFOA *)lParam)->szText, 80, NULL, NULL);
        WideCharToMultiByte(CP_ACP, 0, lpButton->wszButtonItem, -1, szBuffer, BUFFER_SIZE, NULL, NULL);
        ((NMTTDISPINFOA *)lParam)->lpszText=szBuffer;
      }
    }
    else if (((NMHDR *)lParam)->code == TTN_GETDISPINFOW)
    {
      TOOLBARITEM *lpButton;

      if (lpButton=StackGetButtonByID(&hToolbarData, (int)((NMHDR *)lParam)->idFrom))
      {
        //xstrcpynW(((NMTTDISPINFOW *)lParam)->szText, lpButton->wszButtonItem, 80);
        ((NMTTDISPINFOW *)lParam)->lpszText=lpButton->wszButtonItem;
      }
    }
    else if (((NMHDR *)lParam)->code == TBN_DROPDOWN)
    {
      TOOLBARITEM *lpButton;
      RECT rcButton;

      if (lpButton=StackGetButtonByID(&hToolbarData, ((NMTOOLBARA *)lParam)->iItem))
      {
        SendMessage(hToolbar, TB_GETRECT, ((NMTOOLBARA *)lParam)->iItem, (LPARAM)&rcButton);
        ClientToScreen(hToolbar, (LPPOINT)&rcButton.left);
        ClientToScreen(hToolbar, (LPPOINT)&rcButton.right);

        if (!lpButton->hParamMenuName.first)
        {
          //IDM_FILE_OPEN
          TrackPopupMenu(hMenuRecentFiles, TPM_LEFTBUTTON, rcButton.left, rcButton.bottom, 0, hMainWnd, NULL);
        }
        else
        {
          //ContextMenu::Show
          CallContextMenuShow(lpButton, rcButton.left, rcButton.bottom);
        }
      }
    }
  }
  else if (uMsg == WM_COMMAND)
  {
    if ((HWND)lParam == hToolbar)
    {
      CallToolbar(&hToolbarData, LOWORD(wParam));
      return TRUE;
    }
  }

  return DefWindowProcWide(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK NewToolbarProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  if (uMsg == WM_LBUTTONDBLCLK)
  {
    POINT pt={LOWORD(lParam), HIWORD(lParam)};

    if (SendMessage(hToolbar, TB_HITTEST, 0, (LPARAM)&pt) < 0)
    {
      PostMessage(hToolbarBG, AKDLL_SETUP, 0, 0);
    }
  }
  else if (uMsg == WM_RBUTTONDOWN)
  {
    //When left and right button pressed, after releasing window is not accept any mouse buttons actions.
    return 0;
  }
  else if (uMsg == WM_RBUTTONUP)
  {
    TOOLBARITEM *lpButton;
    POINT pt;
    int nIndex;

    //Find mouse over item
    GetCursorPos(&pt);
    ScreenToClient(hToolbar, &pt);
    nIndex=(int)SendMessage(hToolbar, TB_HITTEST, 0, (LPARAM)&pt);
    if (nIndex < 0) nIndex=(0 - nIndex) - 1;

    lpButton=StackGetButtonByIndex(&hToolbarData, nIndex);
    ViewItemCode(lpButton);
    return 0;
  }
  return CallWindowProcWide(lpOldToolbarProc, hWnd, uMsg, wParam, lParam);
}

BOOL CreateToolbarWindow()
{
  WNDCLASSW wndclass;
  DWORD dwStyle=0;
  BOOL bResult=TRUE;

  if (!hToolbarBG)
  {
    wndclass.style        =CS_HREDRAW|CS_VREDRAW|CS_DBLCLKS;
    wndclass.lpfnWndProc  =ToolbarBGProc;
    wndclass.cbClsExtra   =0;
    wndclass.cbWndExtra   =DLGWINDOWEXTRA;
    wndclass.hInstance    =hInstanceDLL;
    wndclass.hIcon        =NULL;
    wndclass.hCursor      =LoadCursor(NULL, IDC_ARROW);
    wndclass.hbrBackground=(HBRUSH)(UINT_PTR)(COLOR_BTNFACE + 1);
    wndclass.lpszMenuName =NULL;
    wndclass.lpszClassName=TOOLBARBACKGROUNDW;
    if (!RegisterClassWide(&wndclass)) return FALSE;

    hToolbarBG=CreateWindowExWide(0,
                          TOOLBARBACKGROUNDW,
                          NULL,
                          WS_CHILD|WS_VISIBLE,
                          0, 0, 0, 0,
                          hMainWnd,
                          NULL,
                          hInstanceDLL,
                          NULL);
  }

  if (nToolbarSide == TBSIDE_LEFT)
    dwStyle=CCS_VERT|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE;
  else if (nToolbarSide == TBSIDE_TOP)
    dwStyle=CCS_TOP|/*CCS_NODIVIDER|*/CCS_NOPARENTALIGN|CCS_NORESIZE;
  else if (nToolbarSide == TBSIDE_RIGHT)
    dwStyle=CCS_VERT|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE;
  else if (nToolbarSide == TBSIDE_BOTTOM)
    dwStyle=CCS_BOTTOM|CCS_NODIVIDER|CCS_NOPARENTALIGN|CCS_NORESIZE;

  hToolbar=CreateWindowExWide(0,
                        L"ToolbarWindow32",
                        NULL,
                        WS_CHILD|WS_VISIBLE|TBSTYLE_TOOLTIPS|TBSTYLE_TRANSPARENT|
                        (bFlatButtons?TBSTYLE_FLAT:0)|dwStyle,
                        0, 0, 0, 0,
                        hToolbarBG,
                        (HMENU)(UINT_PTR)IDC_TOOLBAR,
                        hInstanceDLL,
                        NULL);

  lpOldToolbarProc=(WNDPROC)GetWindowLongPtrWide(hToolbar, GWLP_WNDPROC);
  SetWindowLongPtrWide(hToolbar, GWLP_WNDPROC, (UINT_PTR)NewToolbarProc);

  dwStyle=(DWORD)SendMessage(hToolbar, TB_GETEXTENDEDSTYLE, 0, 0);
  SendMessage(hToolbar, TB_SETEXTENDEDSTYLE, 0, dwStyle|TBSTYLE_EX_DRAWDDARROWS);
  SendMessage(hToolbar, TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);

  ParseRows(&hRowListStack);
  if (!CreateToolbarData(&hToolbarData, wszToolBarText))
    bResult=FALSE;

  SetToolbarButtons(&hToolbarData);
  UpdateToolbar(&hToolbarData);
  return bResult;
}

BOOL CreateToolbarData(TOOLBARDATA *hToolbarData, const wchar_t *wpText)
{
  ROWITEM *lpRowItem=NULL;
  TOOLBARITEM *lpButton;
  TOOLBARITEM *lpLastButton=NULL;
  TOOLBARITEM *lpNextRowItemFirstButton=NULL;
  EXTPARAM *lpParameter;
  STACKEXTPARAM hParamButton={0};
  STACKEXTPARAM hParamMenuName={0};
  wchar_t wszButtonItem[MAX_PATH];
  wchar_t wszMethodName[MAX_PATH];
  wchar_t wszIconFile[MAX_PATH];
  const wchar_t *wpTextBegin=wpText;
  const wchar_t *wpCount=wpText;
  const wchar_t *wpLineBegin;
  const wchar_t *wpErrorBegin=wpText;
  const wchar_t *wpSetArg;
  HICON hIcon;
  SIZE sizeIcon;
  DWORD dwAction;
  DWORD dwNewFlags;
  DWORD dwSetFlags=0;
  int nPlus;
  int nMinus;
  int nFileIconIndex;
  int nItemBitmap=0;
  int nItemCommand=0;
  int nMessageID;
  int nRow=1;
  BOOL bMethod;
  BOOL bPrevSeparator=FALSE;
  BOOL bInRow=TRUE;

  if (wpCount)
  {
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
    hToolbarData->hImageList=ImageList_Create(sizeIcon.cx, sizeIcon.cy, (nIconsBit == 16?ILC_COLOR16:ILC_COLOR32)|ILC_MASK, 0, 0);
    ImageList_SetBkColor(hToolbarData->hImageList, GetSysColor(COLOR_BTNFACE));

    if (hRowListStack.nElements)
    {
      if (lpRowItem=GetRow(&hRowListStack, nRow))
      {
        lpNextRowItemFirstButton=GetFirstToolbarItemOfNextRow(lpRowItem);
        bInRow=TRUE;
      }
      else
      {
        lpNextRowItemFirstButton=NULL;
        bInRow=FALSE;
      }
    }

    while (*wpCount)
    {
      bMethod=FALSE;
      lpButton=NULL;
      hIcon=NULL;
      FreeMethodParameters(&hParamMenuName);
      if (!SkipComment(&wpCount)) break;
      wpLineBegin=wpCount;
      NextString(wpCount, wszButtonItem, MAX_PATH, &wpCount, &nMinus);

      //Set options
      wpSetArg=wpLineBegin;

      if (!xstrcmpnW(L"SET(", wpSetArg, (UINT_PTR)-1))
      {
        dwNewFlags=(DWORD)xatoiW(wpSetArg + 4, &wpSetArg);

        if (dwNewFlags & CCMS_NOFILEEXIST)
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
        continue;
      }
      else if (!xstrcmpnW(L"UNSET(", wpSetArg, (UINT_PTR)-1))
      {
        dwSetFlags&=~xatoiW(wpSetArg + 6, NULL);
        continue;
      }
      if (IsFlagOn(dwSetFlags, CCMS_NOSDI|CCMS_NOMDI|CCMS_NOPMDI|CCMS_NOFILEEXIST))
      {
        //Next line
        while (*wpCount != L'\r' && *wpCount != L'\n') ++wpCount;
        continue;
      }

      if (!xstrcmpW(wszButtonItem, L"SEPARATOR") ||
          !xstrcmpW(wszButtonItem, L"SEPARATOR1"))
      {
        if (bInRow && (!bPrevSeparator || !xstrcmpW(wszButtonItem, L"SEPARATOR")))
        {
          if (lpButton=StackInsertBeforeButton(hToolbarData, lpNextRowItemFirstButton))
          {
            lpLastButton=lpButton;
            if (lpRowItem && !lpRowItem->lpFirstToolbarItem)
              lpRowItem->lpFirstToolbarItem=lpButton;
            lpButton->nTextOffset=(int)(wpLineBegin - wpTextBegin);

            lpButton->tbb.iBitmap=-1;
            lpButton->tbb.idCommand=0;
            lpButton->tbb.fsState=0;
            lpButton->tbb.fsStyle=TBSTYLE_SEP;
            lpButton->tbb.dwData=0;
            lpButton->tbb.iString=0;
          }
        }
        bPrevSeparator=TRUE;
        bMethod=TRUE;
      }
      else
      {
        bPrevSeparator=FALSE;

        if (!xstrcmpW(wszButtonItem, L"BREAK"))
        {
          if (lpLastButton)
            lpLastButton->tbb.fsState|=TBSTATE_WRAP;
          bMethod=TRUE;
          ++nRow;

          if (hRowListStack.nElements)
          {
            if (lpRowItem=GetRow(&hRowListStack, nRow))
            {
              lpNextRowItemFirstButton=GetFirstToolbarItemOfNextRow(lpRowItem);
              bInRow=TRUE;
            }
            else
            {
              lpNextRowItemFirstButton=NULL;
              bInRow=FALSE;
            }
          }
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

            if (!xstrcmpiW(wszMethodName, L"Icon"))
            {
              if (!hIcon)
              {
                GetIconParameters(wpCount, wszIconFile, MAX_PATH, &nFileIconIndex, &wpCount);

                if (bInRow)
                {
                  if (!*wszIconFile)
                  {
                    hIcon=(HICON)LoadImageA(hInstanceDLL, MAKEINTRESOURCEA(nFileIconIndex + 100), IMAGE_ICON, sizeIcon.cx, sizeIcon.cy, 0);

                    if (hIcon)
                    {
                      ImageList_AddIcon(hToolbarData->hImageList, hIcon);
                      DestroyIcon(hIcon);
                    }
                  }
                  else
                  {
                    wchar_t wszPath[MAX_PATH];
                    wchar_t *wpFileName;

                    if (TranslateFileString(wszIconFile, wszPath, MAX_PATH))
                    {
                      if (SearchPathWide(NULL, wszPath, NULL, MAX_PATH, wszIconFile, &wpFileName))
                      {
                        if (bBigIcons)
                          ExtractIconExWide(wszPath, nFileIconIndex, &hIcon, NULL, 1);
                        else
                          ExtractIconExWide(wszPath, nFileIconIndex, NULL, &hIcon, 1);

                        if (hIcon)
                        {
                          ImageList_AddIcon(hToolbarData->hImageList, hIcon);
                          DestroyIcon(hIcon);
                        }
                      }
                    }
                  }
                }
              }
              else
              {
                nMessageID=STRID_PARSEMSG_METHODALREADYDEFINED;
                goto Error;
              }
            }
            else if (!xstrcmpiW(wszMethodName, L"Menu"))
            {
              if (!hParamMenuName.first)
              {
                ParseMethodParameters(&hParamMenuName, wpCount, &wpCount);
                if (!bInRow) FreeMethodParameters(&hParamMenuName);
                bMethod=TRUE;
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
              if (!lpButton)
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
                else
                  dwAction=0;

                if (dwAction)
                {
                  if (bInRow)
                  {
                    if (lpButton=StackInsertBeforeButton(hToolbarData, lpNextRowItemFirstButton))
                    {
                      lpLastButton=lpButton;
                      if (lpRowItem && !lpRowItem->lpFirstToolbarItem)
                        lpRowItem->lpFirstToolbarItem=lpButton;
                      lpButton->bUpdateItem=!nMinus;
                      lpButton->bAutoLoad=nPlus;
                      lpButton->dwAction=dwAction;
                      lpButton->nTextOffset=(int)(wpLineBegin - wpTextBegin);

                      xstrcpynW(lpButton->wszButtonItem, wszButtonItem, MAX_PATH);
                      lpButton->tbb.iBitmap=-1;
                      lpButton->tbb.idCommand=++nItemCommand;
                      lpButton->tbb.fsState=TBSTATE_ENABLED;
                      lpButton->tbb.fsStyle=TBSTYLE_BUTTON;
                      lpButton->tbb.dwData=0;
                      lpButton->tbb.iString=0;
                      ParseMethodParameters(&lpButton->hParamStack, wpCount, &wpCount);

                      if (dwAction == EXTACT_COMMAND)
                      {
                        int nCommand=0;

                        if (lpParameter=GetMethodParameter(&lpButton->hParamStack, 1))
                          nCommand=(int)lpParameter->nNumber;

                        if (nCommand == IDM_FILE_OPEN)
                        {
                          lpButton->tbb.fsStyle|=TBSTYLE_DROPDOWN;
                        }
                        if (!*lpButton->wszButtonItem)
                        {
                          GetMenuStringWide(hMainMenu, nCommand, lpButton->wszButtonItem, MAX_PATH, MF_BYCOMMAND);
                        }
                      }
                    }
                  }
                  else
                  {
                    ParseMethodParameters(&hParamButton, wpCount, &wpCount);
                    FreeMethodParameters(&hParamButton);
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

          if (bMethod)
          {
            if (hParamMenuName.first)
            {
              if (!lpButton)
              {
                //Method "Menu()" without action
                if (lpButton=StackInsertBeforeButton(hToolbarData, lpNextRowItemFirstButton))
                {
                  lpLastButton=lpButton;
                  if (lpRowItem && !lpRowItem->lpFirstToolbarItem)
                    lpRowItem->lpFirstToolbarItem=lpButton;
                  lpButton->dwAction=EXTACT_MENU;
                  lpButton->nTextOffset=(int)(wpLineBegin - wpTextBegin);

                  xstrcpynW(lpButton->wszButtonItem, wszButtonItem, MAX_PATH);
                  lpButton->tbb.iBitmap=-1;
                  lpButton->tbb.idCommand=++nItemCommand;
                  lpButton->tbb.fsState=TBSTATE_ENABLED;
                  lpButton->tbb.fsStyle=TBSTYLE_BUTTON;
                  lpButton->tbb.dwData=0;
                  lpButton->tbb.iString=0;
                }
              }
              else lpButton->tbb.fsStyle|=TBSTYLE_DROPDOWN;

              lpButton->hParamMenuName=hParamMenuName;
              hParamMenuName.first=0;
              hParamMenuName.last=0;
              hParamMenuName.nElements=0;
            }
            if (hIcon)
            {
              lpButton->tbb.iBitmap=nItemBitmap++;
            }
          }
          else
          {
            wpErrorBegin=wpLineBegin;
            nMessageID=STRID_PARSEMSG_NOMETHOD;
            goto Error;
          }
        }
      }
    }
  }
  if (lpLastButton)
    lpLastButton->tbb.fsState|=TBSTATE_WRAP;
  if (hToolbarData->last)
    hToolbarData->last->tbb.fsState&=~TBSTATE_WRAP;
  return TRUE;

  Error:
  crExtSetSel.cpMin=wpErrorBegin - wpText;
  crExtSetSel.cpMax=wpCount - wpText;

  if (!hWndMainDlg)
    PostMessage(hToolbarBG, AKDLL_SETUP, 0, 0);
  else
    PostMessage(hWndMainDlg, AKDLL_SELTEXT, 0, 0);

  xprintfW(wszBuffer, GetLangStringW(wLangModule, nMessageID), wszButtonItem, wszMethodName);
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
  if ((dwCheckFlags & CCMS_NOFILEEXIST) && (dwSetFlags & CCMS_NOFILEEXIST))
    return CCMS_NOFILEEXIST;
  return 0;
}

int ParseRows(STACKROW *lpRowListStack)
{
  STACKROW hCurRowListStack;
  ROWITEM *lpRowItem;
  wchar_t *wpCount=wszRowList;
  int nRow;
  int nShow;
  BOOL bCurExist;

  xmemcpy(&hCurRowListStack, lpRowListStack, sizeof(STACKROW));
  xmemset(lpRowListStack, 0, sizeof(STACKROW));

  if (*wpCount)
  {
    while (nRow=(int)xatoiW(wpCount, (const wchar_t **)&wpCount))
    {
      nShow=ROWSHOW_ON;
      if (*wpCount == L'(')
      {
        nShow=(int)xatoiW(++wpCount, (const wchar_t **)&wpCount);
        if (*wpCount == L')') ++wpCount;
      }

      if (nShow != ROWSHOW_OFF)
      {
        if (!hCurRowListStack.nElements || GetRow(&hCurRowListStack, nRow))
          bCurExist=TRUE;
        else
          bCurExist=FALSE;

        if (nShow == ROWSHOW_ON ||
            (nShow == ROWSHOW_UNCHANGE && bCurExist) ||
            (nShow == ROWSHOW_INVERT && !bCurExist))
        {
          if (!GetRow(lpRowListStack, nRow))
          {
            if (!StackInsertIndex((stack **)&lpRowListStack->first, (stack **)&lpRowListStack->last, (stack **)&lpRowItem, -1, sizeof(ROWITEM)))
            {
              lpRowItem->nRow=nRow;
              ++lpRowListStack->nElements;
            }
          }
        }
      }

      if (*wpCount == L',')
        ++wpCount;
      else
        break;
    }
  }
  FreeRows(&hCurRowListStack);

  //Rebuild rows
  wpCount=wszRowList;
  for (lpRowItem=lpRowListStack->first; lpRowItem; lpRowItem=lpRowItem->next)
  {
    wpCount+=xprintfW(wpCount, L"%s%d", (lpRowItem == lpRowListStack->first?L"":L","), lpRowItem->nRow);
  }

  return lpRowListStack->nElements;
}

ROWITEM* GetRow(STACKROW *lpRowListStack, int nRow)
{
  ROWITEM *lpRowItem;

  for (lpRowItem=lpRowListStack->first; lpRowItem; lpRowItem=lpRowItem->next)
  {
    if (lpRowItem->nRow == nRow)
      break;
  }
  return lpRowItem;
}

TOOLBARITEM* GetFirstToolbarItemOfNextRow(ROWITEM *lpRowItem)
{
  for (lpRowItem=lpRowItem->next; lpRowItem; lpRowItem=lpRowItem->next)
  {
    if (lpRowItem->lpFirstToolbarItem)
      return lpRowItem->lpFirstToolbarItem;
  }
  return NULL;
}

void FreeRows(STACKROW *lpRowListStack)
{
  StackClear((stack **)&lpRowListStack->first, (stack **)&lpRowListStack->last);
  lpRowListStack->nElements=0;
}

void FreeToolbarData(TOOLBARDATA *hToolbarData)
{
  StackFreeButton(hToolbarData);
  ImageList_Destroy(hToolbarData->hImageList);
}

void SetToolbarButtons(TOOLBARDATA *hToolbarData)
{
  TOOLBARITEM *lpButton;
  int nArrorWidth;
  int nRowWidth;
  int nMaxRowWidth;

  if (bBigIcons)
  {
    sizeButtons.cx=40;
    sizeButtons.cy=38;
    sizeToolbar.cx=44;
    sizeToolbar.cy=44;
    nArrorWidth=10;
    SendMessage(hToolbar, TB_SETBITMAPSIZE, 0, MAKELONG(32, 32));
  }
  else
  {
    sizeButtons.cx=24;
    sizeButtons.cy=22;
    sizeToolbar.cx=28;
    sizeToolbar.cy=28;
    nArrorWidth=9;
    SendMessage(hToolbar, TB_SETBITMAPSIZE, 0, MAKELONG(16, 16));
  }
  nMaxRowWidth=sizeButtons.cx;
  SendMessage(hToolbar, TB_SETBUTTONSIZE, 0, MAKELONG(sizeButtons.cx, sizeButtons.cy));
  SendMessage(hToolbar, TB_SETIMAGELIST, 0, (LPARAM)hToolbarData->hImageList);

  for (lpButton=hToolbarData->first; lpButton; lpButton=lpButton->next)
  {
    lpButton->nButtonWidth=sizeButtons.cx;
    if (lpButton->tbb.fsStyle & TBSTYLE_DROPDOWN)
      lpButton->nButtonWidth+=nArrorWidth;
    nRowWidth=lpButton->nButtonWidth;

    ////Calculate vertical row width
    //if (nToolbarSide == TBSIDE_LEFT || nToolbarSide == TBSIDE_RIGHT)
    //{
    //  for (lpTmpButton=lpButton->prev; lpTmpButton; lpTmpButton=lpTmpButton->prev)
    //  {
    //    if (lpTmpButton->tbb.fsState & TBSTATE_WRAP)
    //      break;
    //    nRowWidth+=lpTmpButton->nButtonWidth;
    //  }
    //}

    if (lpButton->tbb.fsState & TBSTATE_WRAP)
    {
      if (nToolbarSide == TBSIDE_TOP || nToolbarSide == TBSIDE_BOTTOM)
        sizeToolbar.cy+=sizeButtons.cy + ((lpButton->tbb.fsStyle & TBSTYLE_SEP)?8:0);
    }
    else if (nToolbarSide == TBSIDE_LEFT || nToolbarSide == TBSIDE_RIGHT)
      lpButton->tbb.fsState|=TBSTATE_WRAP;

    nMaxRowWidth=max(nRowWidth, nMaxRowWidth);
    SendMessage(hToolbar, TB_ADDBUTTONS, 1, (LPARAM)&lpButton->tbb);
  }

  if (nToolbarSide == TBSIDE_LEFT || nToolbarSide == TBSIDE_RIGHT)
    sizeToolbar.cx=(sizeToolbar.cx - sizeButtons.cx) + nMaxRowWidth;
}

void ClearToolbarButtons()
{
  while (SendMessage(hToolbar, TB_DELETEBUTTON, 0, 0));
}

void UpdateToolbar(TOOLBARDATA *hToolbarData)
{
  EDITINFO ei;
  TOOLBARITEM *lpButton=(TOOLBARITEM *)hToolbarData->first;
  EXTPARAM *lpParameter;
  BOOL bInitMenu=FALSE;

  ei.hWndEdit=NULL;

  while (lpButton)
  {
    if (lpButton->bUpdateItem)
    {
      if (lpButton->dwAction == EXTACT_COMMAND)
      {
        DWORD dwMenuState=0;
        DWORD dwToolState;
        int nCommand=0;

        if (lpParameter=GetMethodParameter(&lpButton->hParamStack, 1))
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
            dwToolState=(DWORD)SendMessage(hToolbar, TB_GETSTATE, lpButton->tbb.idCommand, 0);

            if (!(dwMenuState & MF_GRAYED) == !(dwToolState & TBSTATE_ENABLED))
              SendMessage(hToolbar, TB_ENABLEBUTTON, lpButton->tbb.idCommand, (dwMenuState & MF_GRAYED)?FALSE:TRUE);
            if (!(dwMenuState & MF_CHECKED) != !(dwToolState & TBSTATE_CHECKED))
              SendMessage(hToolbar, TB_CHECKBUTTON, lpButton->tbb.idCommand, (dwMenuState & MF_CHECKED)?TRUE:FALSE);
          }
        }
      }
      else if (lpButton->dwAction == EXTACT_CALL)
      {
        PLUGINFUNCTION *pf;
        wchar_t *wpFunction=NULL;
        BOOL bChecked;

        if (lpParameter=GetMethodParameter(&lpButton->hParamStack, 1))
          wpFunction=lpParameter->wpString;

        if (wpFunction)
        {
          bChecked=FALSE;
          if (pf=(PLUGINFUNCTION *)SendMessage(hMainWnd, AKD_DLLFINDW, (WPARAM)wpFunction, 0))
            if (pf->bRunning) bChecked=TRUE;
          SendMessage(hToolbar, TB_CHECKBUTTON, lpButton->tbb.idCommand, bChecked);
        }
      }
    }
    lpButton=lpButton->next;
  }
}

void ViewItemCode(TOOLBARITEM *lpButton)
{
  if (lpButton)
  {
    crExtSetSel.cpMin=lpButton->nTextOffset;
    crExtSetSel.cpMax=-1;
  }
  if (!hWndMainDlg)
    PostMessage(hToolbarBG, AKDLL_SETUP, 0, 0);
  else
    PostMessage(hWndMainDlg, AKDLL_SELTEXT, 0, 0);
}

void CallToolbar(TOOLBARDATA *hToolbarData, int nItem)
{
  TOOLBARITEM *lpElement;
  EXTPARAM *lpParameter;
  wchar_t wszCurrentFile[MAX_PATH];
  RECT rcButton;
  int nButtonID;

  if (lpElement=StackGetButtonByID(hToolbarData, nItem))
  {
    if (GetKeyState(VK_CONTROL) & 0x80)
    {
      ViewItemCode(lpElement);
      return;
    }

    GetCurFile(wszCurrentFile, MAX_PATH);
    nButtonID=lpElement->tbb.idCommand;

    SendMessage(hToolbar, TB_GETRECT, lpElement->tbb.idCommand, (LPARAM)&rcButton);
    ClientToScreen(hToolbar, (LPPOINT)&rcButton.left);
    ClientToScreen(hToolbar, (LPPOINT)&rcButton.right);

    if (lpElement->dwAction == EXTACT_COMMAND)
    {
      int nCommand=0;
      LPARAM lParam=0;

      if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 1))
        nCommand=(int)lpParameter->nNumber;
      if (lpParameter=GetMethodParameter(&lpElement->hParamStack, 2))
        lParam=lpParameter->nNumber;

      if (nCommand)
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
        ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hToolbar, nButtonID, &rcButton);

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

      ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hToolbar, nButtonID, &rcButton);
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

      ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hToolbar, nButtonID, &rcButton);
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

      ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hToolbar, nButtonID, &rcButton);
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
          ExpandMethodParameters(&lpElement->hParamStack, wszCurrentFile, wszExeDir, hToolbar, nButtonID, &rcButton);
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
    else if (lpElement->dwAction == EXTACT_MENU)
    {
      CallContextMenuShow(lpElement, rcButton.left, rcButton.bottom);
    }
  }
}

void CallContextMenuShow(TOOLBARITEM *lpButton, int nPosX, int nPosY)
{
  if (lpButton->hParamMenuName.first)
  {
    PLUGINCALLSENDW pcs;
    DLLEXTCONTEXTMENU decm;
    unsigned char szLeft[32];
    unsigned char szBottom[32];
    unsigned char *pMenuName;

    if (bOldWindows)
    {
      xitoaA(nPosX, (char *)szLeft);
      xitoaA(nPosY, (char *)szBottom);
      pMenuName=(unsigned char *)lpButton->hParamMenuName.first->pString;
    }
    else
    {
      xitoaW(nPosX, (wchar_t *)szLeft);
      xitoaW(nPosY, (wchar_t *)szBottom);
      pMenuName=(unsigned char *)lpButton->hParamMenuName.first->wpString;
    }
    decm.dwStructSize=sizeof(DLLEXTCONTEXTMENU);
    decm.nAction=DLLA_CONTEXTMENU_SHOWSUBMENU;
    decm.pPosX=szLeft;
    decm.pPosY=szBottom;
    decm.nMenuIndex=-1;
    decm.pMenuName=pMenuName;

    pcs.pFunction=L"ContextMenu::Show";
    pcs.lParam=(LPARAM)&decm;
    pcs.dwSupport=0;
    SendMessage(hMainWnd, AKD_DLLCALLW, 0, (LPARAM)&pcs);
  }
}

void DestroyToolbarWindow(BOOL bDestroyBG)
{
  DestroyWindow(hToolbar);
  hToolbar=NULL;
  FreeToolbarData(&hToolbarData);

  if (bDestroyBG)
  {
    DestroyWindow(hToolbarBG);
    hToolbarBG=NULL;
    UnregisterClassWide(TOOLBARBACKGROUNDW, hInstanceDLL);
  }
}

TOOLBARITEM* StackInsertBeforeButton(TOOLBARDATA *hToolbarData, TOOLBARITEM *lpInsertBefore)
{
  TOOLBARITEM *lpElement=NULL;

  StackInsertBefore((stack **)&hToolbarData->first, (stack **)&hToolbarData->last, (stack *)lpInsertBefore, (stack **)&lpElement, sizeof(TOOLBARITEM));
  return lpElement;
}

TOOLBARITEM* StackGetButtonByID(TOOLBARDATA *hToolbarData, int nItemID)
{
  TOOLBARITEM *lpElement=(TOOLBARITEM *)hToolbarData->first;

  while (lpElement)
  {
    if (lpElement->tbb.idCommand == nItemID)
      return lpElement;

    lpElement=lpElement->next;
  }
  return NULL;
}

TOOLBARITEM* StackGetButtonByIndex(TOOLBARDATA *hToolbarData, int nIndex)
{
  TOOLBARITEM *lpElement=(TOOLBARITEM *)hToolbarData->first;
  int nCount=0;

  while (lpElement)
  {
    if (nCount++ == nIndex)
      return lpElement;

    lpElement=lpElement->next;
  }
  return NULL;
}

void StackFreeButton(TOOLBARDATA *hToolbarData)
{
  TOOLBARITEM *lpElement=(TOOLBARITEM *)hToolbarData->first;

  while (lpElement)
  {
    FreeMethodParameters(&lpElement->hParamStack);
    FreeMethodParameters(&lpElement->hParamMenuName);

    lpElement=lpElement->next;
  }
  StackClear((stack **)&hToolbarData->first, (stack **)&hToolbarData->last);
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

void ExpandMethodParameters(STACKEXTPARAM *hParamStack, const wchar_t *wpFile, const wchar_t *wpExeDir, HWND hToolbar, int nButtonID, const RECT *lprcButton)
{
  //%f -file, %d -file directory, %a -AkelPad directory, %% -%
  //%m -toolbar handle, %i -button command id, button screen coordinates: %bl -left, %bt -top, %br -right, %bb -bottom
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
                wpTarget+=(int)xitoaW((INT_PTR)hToolbar, wszTarget?wpTarget:NULL) - !wszTarget;
              }
              else if (*wpSource == L'i' || *wpSource == L'I')
              {
                ++wpSource;
                wpTarget+=(int)xitoaW(nButtonID, wszTarget?wpTarget:NULL) - !wszTarget;
              }
              else if (*wpSource == L'b' || *wpSource == L'B')
              {
                if (*++wpSource == L'l' || *wpSource == L'L')
                {
                  ++wpSource;
                  wpTarget+=(int)xitoaW(lprcButton->left, wszTarget?wpTarget:NULL) - !wszTarget;
                }
                else if (*wpSource == L't' || *wpSource == L'T')
                {
                  ++wpSource;
                  wpTarget+=(int)xitoaW(lprcButton->top, wszTarget?wpTarget:NULL) - !wszTarget;
                }
                else if (*wpSource == L'r' || *wpSource == L'R')
                {
                  ++wpSource;
                  wpTarget+=(int)xitoaW(lprcButton->right, wszTarget?wpTarget:NULL) - !wszTarget;
                }
                else if (*wpSource == L'b' || *wpSource == L'B')
                {
                  ++wpSource;
                  wpTarget+=(int)xitoaW(lprcButton->bottom, wszTarget?wpTarget:NULL) - !wszTarget;
                }
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
    dwSize=(DWORD)WideOption(hOptions, L"ToolBarText", PO_BINARY, NULL, 0);

    if (dwSize)
    {
      if (wszToolBarText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize + sizeof(wchar_t)))
      {
        WideOption(hOptions, L"ToolBarText", PO_BINARY, (LPBYTE)wszToolBarText, dwSize);
        wszToolBarText[dwSize / sizeof(wchar_t)]=L'\0';
      }
    }

    WideOption(hOptions, L"BigIcons", PO_DWORD, (LPBYTE)&bBigIcons, sizeof(DWORD));
    WideOption(hOptions, L"FlatButtons", PO_DWORD, (LPBYTE)&bFlatButtons, sizeof(DWORD));
    WideOption(hOptions, L"IconsBit", PO_DWORD, (LPBYTE)&nIconsBit, sizeof(DWORD));
    WideOption(hOptions, L"ToolbarSide", PO_DWORD, (LPBYTE)&nToolbarSide, sizeof(DWORD));
    WideOption(hOptions, L"SidePriority", PO_DWORD, (LPBYTE)&nSidePriority, sizeof(DWORD));
    WideOption(hOptions, L"RowList", PO_STRING, (LPBYTE)wszRowList, MAX_PATH * sizeof(wchar_t));
    WideOption(hOptions, L"WindowRect", PO_BINARY, (LPBYTE)&rcMainCurrentDialog, sizeof(RECT));

    SendMessage(hMainWnd, AKD_ENDOPTIONS, (WPARAM)hOptions, 0);
  }

  if (!wszToolBarText)
  {
    dwSize=(DWORD)xprintfW(NULL, L"%s", GetLangStringW(wLangModule, STRID_DEFAULTMENU));

    if (wszToolBarText=(wchar_t *)HeapAlloc(hHeap, 0, dwSize * sizeof(wchar_t)))
    {
      xprintfW(wszToolBarText, L"%s", GetLangStringW(wLangModule, STRID_DEFAULTMENU));
    }
  }
}

void SaveOptions(DWORD dwFlags)
{
  HANDLE hOptions;

  if (hOptions=(HANDLE)SendMessage(hMainWnd, AKD_BEGINOPTIONSW, POB_SAVE, (LPARAM)wszPluginName))
  {
    if (dwFlags & OF_LISTTEXT)
    {
      WideOption(hOptions, L"ToolBarText", PO_BINARY, (LPBYTE)wszToolBarText, ((int)xstrlenW(wszToolBarText) + 1) * sizeof(wchar_t));
    }
    if (dwFlags & OF_SETTINGS)
    {
      WideOption(hOptions, L"BigIcons", PO_DWORD, (LPBYTE)&bBigIcons, sizeof(DWORD));
      WideOption(hOptions, L"FlatButtons", PO_DWORD, (LPBYTE)&bFlatButtons, sizeof(DWORD));
      WideOption(hOptions, L"IconsBit", PO_DWORD, (LPBYTE)&nIconsBit, sizeof(DWORD));
      WideOption(hOptions, L"ToolbarSide", PO_DWORD, (LPBYTE)&nToolbarSide, sizeof(DWORD));
      WideOption(hOptions, L"SidePriority", PO_DWORD, (LPBYTE)&nSidePriority, sizeof(DWORD));
      WideOption(hOptions, L"RowList", PO_STRING, (LPBYTE)wszRowList, ((int)xstrlenW(wszRowList) + 1) * sizeof(wchar_t));
    }
    if (dwFlags & OF_RECT)
    {
      WideOption(hOptions, L"WindowRect", PO_BINARY, (LPBYTE)&rcMainCurrentDialog, sizeof(RECT));
    }

    SendMessage(hMainWnd, AKD_ENDOPTIONS, (WPARAM)hOptions, 0);
  }
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
    if (nStringID == STRID_SETUP)
      return L"\x041D\x0430\x0441\x0442\x0440\x043E\x0438\x0442\x044C\x002E\x002E\x002E";
    if (nStringID == STRID_AUTOLOAD)
      return L"\x0022\x0025\x0073\x0022\x0020\x043D\x0435\x0020\x043F\x043E\x0434\x0434\x0435\x0440\x0436\x0438\x0432\x0430\x0435\x0442\x0020\x0430\x0432\x0442\x043E\x0437\x0430\x0433\x0440\x0443\x0437\x043A\x0443\x0020\x0028\x0022\x002B\x0022\x0020\x043F\x0435\x0440\x0435\x0434\x0020\x0022\x0043\x0061\x006C\x006C\x0022\x0029\x002E";
    if (nStringID == STRID_BIGICONS)
      return L"\x0411\x043E\x043B\x044C\x0448\x0438\x0435\x0020\x0438\x043A\x043E\x043D\x043A\x0438";
    if (nStringID == STRID_FLATBUTTONS)
      return L"\x041F\x043B\x043E\x0441\x043A\x0438\x0435\x0020\x043A\x043D\x043E\x043F\x043A\x0438";
    if (nStringID == STRID_16BIT)
      return L"\x0031\x0036\x002D\x0431\x0438\x0442";
    if (nStringID == STRID_32BIT)
      return L"\x0033\x0032\x002D\x0431\x0438\x0442\x0430";
    if (nStringID == STRID_SIDE)
      return L"\x0421\x0442\x043E\x0440\x043E\x043D\x0430";
    if (nStringID == STRID_ROWS)
      return L"\x0420\x044F\x0434\x044B (1,2...)";
    if (nStringID == STRID_PARSEMSG_UNKNOWNMETHOD)
      return L"\x041D\x0435\x0438\x0437\x0432\x0435\x0441\x0442\x043D\x044B\x0439\x0020\x043C\x0435\x0442\x043E\x0434 \"%0.s%s\".";
    if (nStringID == STRID_PARSEMSG_METHODALREADYDEFINED)
      return L"\x041C\x0435\x0442\x043E\x0434\x0020\x0443\x0436\x0435\x0020\x043D\x0430\x0437\x043D\x0430\x0447\x0435\x043D\x002E";
    if (nStringID == STRID_PARSEMSG_NOMETHOD)
      return L"\x042D\x043B\x0435\x043C\x0435\x043D\x0442\x0020\x043D\x0435\x0020\x0438\x0441\x043F\x043E\x043B\x044C\x0437\x0443\x0435\x0442\x0020\x043C\x0435\x0442\x043E\x0434\x0430\x0020\x0434\x043B\x044F\x0020\x0432\x044B\x043F\x043E\x043B\x043D\x0435\x043D\x0438\x044F\x002E";
    if (nStringID == STRID_PLUGIN)
      return L"%s \x043F\x043B\x0430\x0433\x0438\x043D";
    if (nStringID == STRID_OK)
      return L"\x004F\x004B";
    if (nStringID == STRID_CANCEL)
      return L"\x041E\x0442\x043C\x0435\x043D\x0430";

    //File menu
    if (nStringID == STRID_DEFAULTMENU)
      return L"\
\"\" Command(4101) Icon(0)\r\
\"\" Command(4102) Icon(1)\r\
\"\" Command(4103) Icon(2)\r\
\"\" Command(4104) Icon(3)\r\
SEPARATOR1\r\
\"\" Command(4105) Icon(4)\r\
\"\" Command(4106) Icon(5)\r\
SET(1)\r\
    # MDI/PMDI\r\
    \"\" Command(4110) Icon(32)\r\
    \"\" Command(4111) Icon(33)\r\
UNSET(1)\r\
SEPARATOR1\r\
\"\" Command(4108) Icon(6)\r\
\"\" Command(4114) Icon(21)\r\
SEPARATOR1\r\
" L"\
\"\" Command(4153) Icon(7)\r\
\"\" Command(4154) Icon(8)\r\
\"\" Command(4155) Icon(9)\r\
\"\" Command(4156) Icon(25)\r\
SEPARATOR1\r\
\"\" Command(4151) Icon(10)\r\
\"\" Command(4152) Icon(11)\r\
SEPARATOR1\r\
\"\" Command(4158) Icon(12)\r\
\"\" Command(4161) Icon(13)\r\
\"\" Command(4163) Icon(14)\r\
\"\" Command(4183) Icon(26)\r\
SEPARATOR1\r\
" L"\
\"\" Command(4201) Icon(27)\r\
-\"\x0423\x0432\x0435\x043B\x0438\x0447\x0438\x0442\x044C\x0020\x0448\x0440\x0438\x0444\x0442\" Command(4204) Icon(28)\r\
-\"\x0423\x043C\x0435\x043D\x044C\x0448\x0438\x0442\x044C\x0020\x0448\x0440\x0438\x0444\x0442\" Command(4205) Icon(29)\r\
\"\" Command(4202) Icon(30)\r\
SEPARATOR1\r\
\"\" Command(4216) Icon(20)\r\
\"\" Command(4209) Icon(16)\r\
\"\x0420\x0430\x0437\x0434\x0435\x043B\x0438\x0442\x044C\x0020\x043D\x0430\x0020\x0034\x0020\x0447\x0430\x0441\x0442\x0438\" Command(4212) Icon(22)\r\
\"\x0420\x0430\x0437\x0434\x0435\x043B\x0438\x0442\x044C\x0020\x0432\x0435\x0440\x0442\x0438\x043A\x0430\x043B\x044C\x043D\x043E\" Command(4213) Icon(23)\r\
\"\x0420\x0430\x0437\x0434\x0435\x043B\x0438\x0442\x044C\x0020\x0433\x043E\x0440\x0438\x0437\x043E\x043D\x0442\x0430\x043B\x044C\x043D\x043E\" Command(4214) Icon(24)\r\
\"\" Command(4210) Icon(15)\r\
SEPARATOR1\r\
" L"\
\"\" Command(4251) Icon(17)\r\
\"\" Command(4259) Icon(18)\r\
\"\" Command(4260) Icon(19)\r\
SEPARATOR1\r\
" L"\
\"\x0417\x0430\x043F\x0443\x0441\x0442\x0438\x0442\x044C\x0020\x0431\x043B\x043E\x043A\x043D\x043E\x0442\" Exec(\"notepad.exe\") Icon(\"notepad.exe\")\r\
SET(32, \"%a\\AkelFiles\\AkelUpdater.exe\")\r\
    \"\x0417\x0430\x043F\x0443\x0441\x0442\x0438\x0442\x044C\x0020\x043E\x0431\x043D\x043E\x0432\x043B\x0435\x043D\x0438\x0435\" Exec(\"%a\\AkelFiles\\AkelUpdater.exe\") Icon(\"%a\\AkelFiles\\AkelUpdater.exe\")\r\
UNSET(32)\r\
" L"\
\r\
SEPARATOR1\r\
BREAK\r\
\"\x0413\x043B\x0430\x0432\x043D\x043E\x0435\x0020\x043C\x0435\x043D\x044E\" Call(\"ContextMenu::Show\", 2, \"%bl\", \"%bb\") Icon(38)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Coder.dll\")\r\
    \"\x041F\x0440\x043E\x0433\x0440\x0430\x043C\x043C\x0438\x0440\x043E\x0432\x0430\x043D\x0438\x0435\" Menu(\"CODER\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 12)\r\
    \"\x041E\x0442\x043C\x0435\x0442\x0438\x0442\x044C\" Menu(\"MARK\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 0)\r\
    \"\x0421\x0438\x043D\x0442\x0430\x043A\x0441\x0438\x0447\x0435\x0441\x043A\x0430\x044F\x0020\x0442\x0435\x043C\x0430\" Menu(\"SYNTAXTHEME\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 4)\r\
    \"\x0426\x0432\x0435\x0442\x043E\x0432\x0430\x044F\x0020\x0442\x0435\x043C\x0430\" Menu(\"COLORTHEME\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 5)\r\
    -\"\x041F\x0430\x043D\x0435\x043B\x044C CodeFold\" Call(\"Coder::CodeFold\", 1) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 3)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\XBrackets.dll\")\r\
    \"\x041F\x0430\x0440\x043D\x044B\x0435\x0020\x0441\x043A\x043E\x0431\x043A\x0438\" +Call(\"XBrackets::Main\") Menu(\"XBRACKETS\") Icon(\"%a\\AkelFiles\\Plugs\\XBrackets.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\SpellCheck.dll\")\r\
    \"\x041F\x0440\x043E\x0432\x0435\x0440\x043A\x0430\x0020\x043E\x0440\x0444\x043E\x0433\x0440\x0430\x0444\x0438\x0438\" +Call(\"SpellCheck::Background\") Menu(\"SPELLCHECK\") Icon(35)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\SpecialChar.dll\")\r\
    \"\x0421\x043F\x0435\x0446\x0438\x0430\x043B\x044C\x043D\x044B\x0435\x0020\x0441\x0438\x043C\x0432\x043E\x043B\x044B\" +Call(\"SpecialChar::Main\") Menu(\"SPECIALCHAR\") Icon(\"%a\\AkelFiles\\Plugs\\SpecialChar.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\LineBoard.dll\")\r\
    \"\x041D\x043E\x043C\x0435\x0440\x0430\x0020\x0441\x0442\x0440\x043E\x043A\x002C\x0020\x0437\x0430\x043A\x043B\x0430\x0434\x043A\x0438\" +Call(\"LineBoard::Main\") Menu(\"LINEBOARD\") Icon(\"%a\\AkelFiles\\Plugs\\LineBoard.dll\", 0)\r\
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
    \"\x041F\x0440\x043E\x0441\x043C\x043E\x0442\x0440\x0020\x043B\x043E\x0433\x0430\" Call(\"Log::Watch\") Menu(\"LOG\") Icon(\"%a\\AkelFiles\\Plugs\\Log.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Explorer.dll\")\r\
    \"\x041F\x0430\x043D\x0435\x043B\x044C\x0020\x043F\x0440\x043E\x0432\x043E\x0434\x043D\x0438\x043A\x0430\" +Call(\"Explorer::Main\") Menu(\"EXPLORE\") Icon(\"%a\\AkelFiles\\Plugs\\Explorer.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\QSearch.dll\")\r\
    \"\x041F\x0430\x043D\x0435\x043B\x044C\x0020\x043F\x043E\x0438\x0441\x043A\x0430\" +Call(\"QSearch::QSearch\") Menu(\"QSEARCH\") Icon(\"%a\\AkelFiles\\Plugs\\QSearch.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Macros.dll\")\r\
    -\"\x041C\x0430\x043A\x0440\x043E\x0441\x044B...\" Call(\"Macros::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 0)\r\
    -\"\x0417\x0430\x043F\x0438\x0441\x0430\x0442\x044C\" Call(\"Macros::Main\", 2, \"%m\", \"%i\") Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 1)\r\
    -\"\x0412\x043E\x0441\x043F\x0440\x043E\x0438\x0437\x0432\x0435\x0441\x0442\x0438\x0020\x043E\x0434\x0438\x043D\x0020\x0440\x0430\x0437\" Call(\"Macros::Main\", 1, \"\", 1) Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 3)\r\
    -\"\x0412\x043E\x0441\x043F\x0440\x043E\x0438\x0437\x0432\x0435\x0441\x0442\x0438\x0020\x0434\x043E\x0020\x043A\x043E\x043D\x0446\x0430\" Call(\"Macros::Main\", 3, \"%m\", \"%i\") Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 4)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Scripts.dll\")\r\
    -\"\x0421\x043A\x0440\x0438\x043F\x0442\x044B...\" +Call(\"Scripts::Main\") Menu(\"SCRIPTS\") Icon(\"%a\\AkelFiles\\Plugs\\Scripts.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\RecentFiles.dll\")\r\
    -\"\x041F\x043E\x0441\x043B\x0435\x0434\x043D\x0438\x0435\x0020\x0444\x0430\x0439\x043B\x044B...\" Call(\"RecentFiles::Manage\") Icon(\"%a\\AkelFiles\\Plugs\\RecentFiles.dll\", 0)\r\
UNSET(32)\r\
SET(1)\r\
    # MDI/PMDI\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Sessions.dll\")\r\
        -\"\x0421\x0435\x0441\x0441\x0438\x0438...\" Call(\"Sessions::Main\") Menu(\"SESSIONS\") Icon(\"%a\\AkelFiles\\Plugs\\Sessions.dll\", 0)\r\
    UNSET(32)\r\
UNSET(1)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Templates.dll\")\r\
    -\"\x0428\x0430\x0431\x043B\x043E\x043D\x044B...\" Call(\"Templates::Open\") Menu(\"TEMPLATES\") Icon(37)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Format.dll\")\r\
    \"\x0421\x043E\x0440\x0442\x0438\x0440\x043E\x0432\x0430\x0442\x044C\x0020\x0441\x0442\x0440\x043E\x043A\x0438\x0020\x043F\x043E\x0020\x0432\x043E\x0437\x0440\x0430\x0441\x0442\x0430\x043D\x0438\x044E\" Call(\"Format::LineSortStrAsc\") Menu(\"FORMAT\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Scroll.dll\")\r\
    \"\x0412\x0435\x0440\x0442\x0438\x043A\x0430\x043B\x044C\x043D\x0430\x044F\x0020\x0441\x0438\x043D\x0445\x0440\x043E\x043D\x0438\x0437\x0430\x0446\x0438\x044F\" Call(\"Scroll::SyncVert\") Menu(\"SCROLL\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 1)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\HexSel.dll\")\r\
    \"Hex \x043A\x043E\x0434\" +Call(\"HexSel::Main\") Icon(\"%a\\AkelFiles\\Plugs\\HexSel.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Hotkeys.dll\")\r\
    -\"\x0413\x043E\x0440\x044F\x0447\x0438\x0435\x0020\x043A\x043B\x0430\x0432\x0438\x0448\x0438...\" +Call(\"Hotkeys::Main\") Menu(\"HOTKEYS\") Icon(\"%a\\AkelFiles\\Plugs\\Hotkeys.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\MinimizeToTray.dll\")\r\
    -\"\x0421\x0432\x0435\x0440\x043D\x0443\x0442\x044C\x0020\x0432\x0020\x0442\x0440\x0435\x0439\" Call(\"MinimizeToTray::Now\") Menu(\"MINIMIZETOTRAY\") Icon(\"%a\\AkelFiles\\Plugs\\MinimizeToTray.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Sounds.dll\")\r\
    \"\x0417\x0432\x0443\x043A\x043E\x0432\x043E\x0439\x0020\x043D\x0430\x0431\x043E\x0440\x0020\x0442\x0435\x043A\x0441\x0442\x0430\" +Call(\"Sounds::Main\") Menu(\"SOUNDS\") Icon(\"%a\\AkelFiles\\Plugs\\Sounds.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Speech.dll\")\r\
    \"\x041C\x0430\x0448\x0438\x043D\x043D\x043E\x0435\x0020\x0447\x0442\x0435\x043D\x0438\x0435\x0020\x0442\x0435\x043A\x0441\x0442\x0430\" +Call(\"Speech::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Speech.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r";

  }
  else
  {
    if (nStringID == STRID_SETUP)
      return L"Customize...";
    if (nStringID == STRID_AUTOLOAD)
      return L"\"%s\" doesn't support autoload (\"+\" before \"Call\").";
    if (nStringID == STRID_BIGICONS)
      return L"Big icons";
    if (nStringID == STRID_FLATBUTTONS)
      return L"Flat buttons";
    if (nStringID == STRID_16BIT)
      return L"16-bit";
    if (nStringID == STRID_32BIT)
      return L"32-bit";
    if (nStringID == STRID_SIDE)
      return L"Side";
    if (nStringID == STRID_ROWS)
      return L"Rows (1,2...)";
    if (nStringID == STRID_PARSEMSG_UNKNOWNMETHOD)
      return L"Unknown method \"%0.s%s\".";
    if (nStringID == STRID_PARSEMSG_METHODALREADYDEFINED)
      return L"Method already defined.";
    if (nStringID == STRID_PARSEMSG_NOMETHOD)
      return L"The element does not use method for execution.";
    if (nStringID == STRID_PLUGIN)
      return L"%s plugin";
    if (nStringID == STRID_OK)
      return L"OK";
    if (nStringID == STRID_CANCEL)
      return L"Cancel";

    //File menu
    if (nStringID == STRID_DEFAULTMENU)
      return L"\
\"\" Command(4101) Icon(0)\r\
\"\" Command(4102) Icon(1)\r\
\"\" Command(4103) Icon(2)\r\
\"\" Command(4104) Icon(3)\r\
SEPARATOR1\r\
\"\" Command(4105) Icon(4)\r\
\"\" Command(4106) Icon(5)\r\
SET(1)\r\
    # MDI/PMDI\r\
    \"\" Command(4110) Icon(32)\r\
    \"\" Command(4111) Icon(33)\r\
UNSET(1)\r\
SEPARATOR1\r\
\"\" Command(4108) Icon(6)\r\
\"\" Command(4114) Icon(21)\r\
SEPARATOR1\r\
" L"\
\"\" Command(4153) Icon(7)\r\
\"\" Command(4154) Icon(8)\r\
\"\" Command(4155) Icon(9)\r\
\"\" Command(4156) Icon(25)\r\
SEPARATOR1\r\
\"\" Command(4151) Icon(10)\r\
\"\" Command(4152) Icon(11)\r\
SEPARATOR1\r\
\"\" Command(4158) Icon(12)\r\
\"\" Command(4161) Icon(13)\r\
\"\" Command(4163) Icon(14)\r\
\"\" Command(4183) Icon(26)\r\
SEPARATOR1\r\
" L"\
\"\" Command(4201) Icon(27)\r\
-\"Increase font\" Command(4204) Icon(28)\r\
-\"Decrease font\" Command(4205) Icon(29)\r\
\"\" Command(4202) Icon(30)\r\
SEPARATOR1\r\
\"\" Command(4216) Icon(20)\r\
\"\" Command(4209) Icon(16)\r\
\"Split into four panes\" Command(4212) Icon(22)\r\
\"Split vertically\" Command(4213) Icon(23)\r\
\"Split horizontally\" Command(4214) Icon(24)\r\
\"\" Command(4210) Icon(15)\r\
SEPARATOR1\r\
" L"\
\"\" Command(4251) Icon(17)\r\
\"\" Command(4259) Icon(18)\r\
\"\" Command(4260) Icon(19)\r\
SEPARATOR1\r\
" L"\
\"Run Notepad\" Exec(\"notepad.exe\") Icon(\"notepad.exe\")\r\
SET(32, \"%a\\AkelFiles\\AkelUpdater.exe\")\r\
    \"Run AkelUpdater\" Exec(\"%a\\AkelFiles\\AkelUpdater.exe\") Icon(\"%a\\AkelFiles\\AkelUpdater.exe\")\r\
UNSET(32)\r\
" L"\
\r\
SEPARATOR1\r\
BREAK\r\
\"Main menu\" Call(\"ContextMenu::Show\", 2, \"%bl\", \"%bb\") Icon(38)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Coder.dll\")\r\
    \"Programming\" Menu(\"CODER\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 12)\r\
    \"Mark\" Menu(\"MARK\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 0)\r\
    \"Syntax theme\" Menu(\"SYNTAXTHEME\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 4)\r\
    \"Color theme\" Menu(\"COLORTHEME\") Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 5)\r\
    -\"CodeFold panel\" Call(\"Coder::CodeFold\", 1) Icon(\"%a\\AkelFiles\\Plugs\\Coder.dll\", 3)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\XBrackets.dll\")\r\
    \"Brackets\" +Call(\"XBrackets::Main\") Menu(\"XBRACKETS\") Icon(\"%a\\AkelFiles\\Plugs\\XBrackets.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\SpellCheck.dll\")\r\
    \"Spell check\" +Call(\"SpellCheck::Background\") Menu(\"SPELLCHECK\") Icon(35)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\SpecialChar.dll\")\r\
    \"Special characters\" +Call(\"SpecialChar::Main\") Menu(\"SPECIALCHAR\") Icon(\"%a\\AkelFiles\\Plugs\\SpecialChar.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\LineBoard.dll\")\r\
    \"Line numbers, bookmarks\" +Call(\"LineBoard::Main\") Menu(\"LINEBOARD\") Icon(\"%a\\AkelFiles\\Plugs\\LineBoard.dll\", 0)\r\
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
    \"Log view\" Call(\"Log::Watch\") Menu(\"LOG\") Icon(\"%a\\AkelFiles\\Plugs\\Log.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Explorer.dll\")\r\
    \"Explorer panel\" +Call(\"Explorer::Main\") Menu(\"EXPLORE\") Icon(\"%a\\AkelFiles\\Plugs\\Explorer.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\QSearch.dll\")\r\
    \"Search panel\" +Call(\"QSearch::QSearch\") Menu(\"QSEARCH\") Icon(\"%a\\AkelFiles\\Plugs\\QSearch.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Macros.dll\")\r\
    -\"Macros...\" Call(\"Macros::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 0)\r\
    -\"Record\" Call(\"Macros::Main\", 2, \"%m\", \"%i\") Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 1)\r\
    -\"Play once\" Call(\"Macros::Main\", 1, \"\", 1) Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 3)\r\
    -\"Play to the end\" Call(\"Macros::Main\", 3, \"%m\", \"%i\") Icon(\"%a\\AkelFiles\\Plugs\\Macros.dll\", 4)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Scripts.dll\")\r\
    -\"Scripts...\" +Call(\"Scripts::Main\") Menu(\"SCRIPTS\") Icon(\"%a\\AkelFiles\\Plugs\\Scripts.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\RecentFiles.dll\")\r\
   -\"Recent files...\" Call(\"RecentFiles::Manage\") Icon(\"%a\\AkelFiles\\Plugs\\RecentFiles.dll\", 0)\r\
UNSET(32)\r\
SET(1)\r\
    # MDI/PMDI\r\
    SET(32, \"%a\\AkelFiles\\Plugs\\Sessions.dll\")\r\
        -\"Sessions...\" Call(\"Sessions::Main\") Menu(\"SESSIONS\") Icon(\"%a\\AkelFiles\\Plugs\\Sessions.dll\", 0)\r\
    UNSET(32)\r\
UNSET(1)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Templates.dll\")\r\
    -\"Templates...\" Call(\"Templates::Open\") Menu(\"TEMPLATES\") Icon(37)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Format.dll\")\r\
    \"Sort lines by string ascending\" Call(\"Format::LineSortStrAsc\") Menu(\"FORMAT\") Icon(\"%a\\AkelFiles\\Plugs\\Format.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Scroll.dll\")\r\
    \"Vertical synchronization\" Call(\"Scroll::SyncVert\") Menu(\"SCROLL\") Icon(\"%a\\AkelFiles\\Plugs\\Scroll.dll\", 1)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\HexSel.dll\")\r\
    \"Hex code\" +Call(\"HexSel::Main\") Icon(\"%a\\AkelFiles\\Plugs\\HexSel.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Hotkeys.dll\")\r\
    -\"Hotkeys...\" +Call(\"Hotkeys::Main\") Menu(\"HOTKEYS\") Icon(\"%a\\AkelFiles\\Plugs\\Hotkeys.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\MinimizeToTray.dll\")\r\
    -\"Minimize to tray\" Call(\"MinimizeToTray::Now\") Menu(\"MINIMIZETOTRAY\") Icon(\"%a\\AkelFiles\\Plugs\\MinimizeToTray.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r\
" L"\
SET(32, \"%a\\AkelFiles\\Plugs\\Sounds.dll\")\r\
    \"Sound typing\" +Call(\"Sounds::Main\") Menu(\"SOUNDS\") Icon(\"%a\\AkelFiles\\Plugs\\Sounds.dll\", 0)\r\
UNSET(32)\r\
SET(32, \"%a\\AkelFiles\\Plugs\\Speech.dll\")\r\
    \"Machine reading\" +Call(\"Speech::Main\") Icon(\"%a\\AkelFiles\\Plugs\\Speech.dll\", 0)\r\
UNSET(32)\r\
SEPARATOR1\r";

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

BOOL SetExtCallParam(LPARAM lParam, int nIndex, INT_PTR nData)
{
  if (*((INT_PTR *)lParam) >= (INT_PTR)((nIndex + 1) * sizeof(INT_PTR)))
  {
    *(((INT_PTR *)lParam) + nIndex)=nData;
    return TRUE;
  }
  return FALSE;
}

void InitCommon(PLUGINDATA *pd)
{
  bInitCommon=TRUE;
  hInstanceDLL=pd->hInstanceDLL;
  hInstanceEXE=pd->hInstanceEXE;
  hMainWnd=pd->hMainWnd;
  hMdiClient=pd->hMdiClient;
  hMainMenu=pd->hMainMenu;
  hMenuRecentFiles=pd->hMenuRecentFiles;
  hPopupEdit=GetSubMenu(pd->hPopupMenu, MENU_POPUP_EDIT);
  hMainIcon=pd->hMainIcon;
  hGlobalAccel=pd->hGlobalAccel;
  bOldWindows=pd->bOldWindows;
  bAkelEdit=pd->bAkelEdit;
  nMDI=pd->nMDI;
  wLangModule=PRIMARYLANGID(pd->wLangModule);
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
}

void InitMain()
{
  bInitMain=TRUE;

  //SubClass
  NewMainProcData=NULL;
  SendMessage(hMainWnd, AKD_SETMAINPROC, (WPARAM)NewMainProc, (LPARAM)&NewMainProcData);

  NewMainProcRetData=NULL;
  SendMessage(hMainWnd, AKD_SETMAINPROCRET, (WPARAM)NewMainProcRet, (LPARAM)&NewMainProcRetData);

  if (nMDI == WMD_MDI)
  {
    NewFrameProcRetData=NULL;
    SendMessage(hMainWnd, AKD_SETFRAMEPROCRET, (WPARAM)NewMainProcRet, (LPARAM)&NewFrameProcRetData);
  }
}

void UninitMain()
{
  bInitMain=FALSE;

  //Save options
  if (dwSaveFlags)
  {
    SaveOptions(dwSaveFlags);
    dwSaveFlags=0;
  }

  //Remove subclass
  if (NewMainProcData)
  {
    SendMessage(hMainWnd, AKD_SETMAINPROC, (WPARAM)NULL, (LPARAM)&NewMainProcData);
    NewMainProcData=NULL;
  }
  if (NewMainProcRetData)
  {
    SendMessage(hMainWnd, AKD_SETMAINPROCRET, (WPARAM)NULL, (LPARAM)&NewMainProcRetData);
    NewMainProcRetData=NULL;
  }
  if (NewFrameProcRetData)
  {
    SendMessage(hMainWnd, AKD_SETFRAMEPROCRET, (WPARAM)NULL, (LPARAM)&NewFrameProcRetData);
    NewFrameProcRetData=NULL;
  }

  //Destroy toolbar
  DestroyToolbarWindow(TRUE);
  FreeRows(&hRowListStack);

  if (wszToolBarText)
  {
    HeapFree(hHeap, 0, wszToolBarText);
    wszToolBarText=NULL;
  }
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
