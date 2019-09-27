#define WIN32_LEAN_AND_MEAN
#define WINVER 0x0500
#define _WIN32_IE 0x0500
#include <windows.h>
#include <stddef.h>
#include <commdlg.h>
#include <commctrl.h>
#include <shellapi.h>
#include <shlobj.h>
#include <olectl.h>
#include <aclapi.h>
#include <richedit.h>
#include "AkelEdit\StackFunc.h"
#include "AkelEdit\StrFunc.h"
#include "AkelEdit\x64Func.h"
#include "WideFunc.h"
#include "AkelEdit\AkelBuild.h"
#include "AkelEdit\RegExpFunc.h"
#include "AkelFiles\Langs\Resources\resource.h"
#include "AkelFiles\Langs\Resources\version.h"
#include "AkelFiles\Plugs\AkelDLL\AkelDLL.h"
#include "AkelPad.h"
#include "Edit.h"


//Process
extern HANDLE hHeap;
extern HANDLE hMutex;
extern DWORD dwProcessId;

//WinMain
extern MAINCREATE mc;
extern HINSTANCE hInstance;
extern DWORD dwCmdShow;

//Identification
extern DWORD dwExeVersion;
extern BOOL bOldWindows;
extern BOOL bOldRichEdit;
extern BOOL bOldComctl32;
extern BOOL bAkelEdit;
extern BOOL bWindowsNT4;
extern INT_PTR nUniqueID;

//Buffers
extern unsigned char pcTranslateBuffer[TRANSLATE_BUFFER_SIZE];
extern wchar_t wszMsg[BUFFER_SIZE];
extern char buf[BUFFER_SIZE];
extern wchar_t wbuf[BUFFER_SIZE];
extern char buf2[BUFFER_SIZE];
extern wchar_t wbuf2[BUFFER_SIZE];

//Command line
extern wchar_t wszCmdLine[COMMANDLINE_SIZE];
extern wchar_t wszCmdArg[COMMANDARG_SIZE];
extern wchar_t *wpCmdLineBegin;
extern int nCmdLineBeginLen;
extern wchar_t *wpCmdLineEnd;
extern int nCmdLineEndLen;
extern BOOL bCmdLineChanged;
extern const wchar_t *wpCmdLine;
extern wchar_t *wpCmdParamsStart;
extern wchar_t *wpCmdParamsEnd;
extern DWORD dwCmdLineOptions;
extern BOOL bCmdLineQuitAsEnd;

//Language
extern HMODULE hLangLib;
extern DWORD dwLangSystem;
extern DWORD dwLangModule;
extern int nLangModuleCount;

//Procedures
extern HSTACK hMainProcStack;
extern HSTACK hMainProcRetStack;
extern HSTACK hEditProcStack;
extern HSTACK hEditProcRetStack;
extern HSTACK hFrameProcStack;
extern HSTACK hFrameProcRetStack;
extern WNDPROC lpfnMainProc;
extern WNDPROCRET lpfnMainProcRet;
extern WNDPROC lpfnFrameProc;
extern WNDPROCRET lpfnFrameProcRet;
extern WNDPROC lpfnEditProc;
extern WNDPROCRET lpfnEditProcRet;

//Plugins
extern HSTACK hPluginsStack;
extern HSTACK hPluginListStack;
extern HSTACK hHandlesStack;
extern RECT rcPluginsMinMaxDialog;
extern BOOL bSavePluginsStackOnExit;
extern WNDPROC lpOldHotkeyInputProc;
extern wchar_t wszLastFunction[MAX_PATH];
extern int nLastFunctionIndex;

//INI
extern INIFILE hIniFile;
extern wchar_t wszIniFile[MAX_PATH];

//Main settings
extern MAINOPTIONS moInit;
extern MAINOPTIONS moCur;
extern BOOL bSaveManual;

//Main Window
extern HWND hMainWnd;
extern HWND hDummyWindow;
extern DWORD dwLastMainSizeType;
extern DWORD dwMouseCapture;
extern HACCEL hGlobalAccel;
extern HACCEL hMainAccel;
extern HICON hMainIcon;
extern HCURSOR hCursorDragMove;
extern HCURSOR hCursorHandOpen;
extern HCURSOR hCursorHandClose;
extern HMENU hMainMenu;
extern HMENU hPopupMenu;
extern HMENU hPopupEdit;
extern HMENU hPopupView;
extern HMENU hPopupCodepage;
extern HMENU hPopupOpenCodepage;
extern HMENU hPopupSaveCodepage;
extern HMENU hPopupHeadline;
extern HMENU hMenuRecentFiles;
extern HMENU hMenuLanguage;
extern HMENU hMenuWindow;
extern BOOL bMenuPopupCodepage;
extern BOOL bMenuRecentFiles;
extern BOOL bMenuLanguage;
extern BOOL bMainStarting;
extern BOOL bMainCheckIdle;
extern int nMainOnFinish;
extern BOOL bEditOnFinish;
extern BOOL bFirstTabOnFinish;

//Status window
extern STATUSSTATE ssStatus;
extern STACKSTATUSPART hStatusStack;
extern HWND hStatus;
extern HWND hProgress;
extern int nStatusHeight;
extern int nProgressWidth;

//Clones
extern HCURSOR hCursorSizeWE;
extern HCURSOR hCursorSizeNS;
extern HCURSOR hCursorSizeALL;
extern HCURSOR hCursorClone;
extern int nLastSplit;

//Docks
extern HDOCK hDocksStack;
extern NSIZE nsSize;
extern WNDPROC lpOldCloseButtonProc;

//Owner-drawn buttons
extern HSTACK hButtonDrawStack;

//Codepages
extern int *lpCodepageList;
extern int nCodepageListLen;
extern BOOL bCodepageListChanged;
extern int *lpCodepageTable;
extern int nCodepageTableCount;
extern int nAnsiCodePage;
extern int nOemCodePage;

//Recent files
extern RECENTFILESTACK hRecentFilesStack;

//Open/Save document
extern OPENFILENAME_2000W *ofnStruct;
extern wchar_t *wszOfnFileList;
extern wchar_t wszOfnStream[MAX_PATH];
extern wchar_t wszFileFilter[MAX_PATH];
extern int nFileFilterLen;
extern BOOL bAutodetect;
extern BOOL bSaveDlg;
extern DWORD dwOfnFlags;
extern BOOL bOfnBOM;
extern int nOfnCodePage;
extern int nOfnStreamOffset;
extern POINT64 ptDocumentPos;
extern FILESTREAMDATA *lpStreamInData;
extern HWND hOfnDlgCombo;
extern HWND hOfnDlgEdit;
extern WNDPROC lpOldFilePreviewProc;
extern WNDPROC lpOldFileParentProc;

//AkelAdmin
extern wchar_t wszAkelAdminExe[MAX_PATH];
extern wchar_t wszAkelAdminPipe[32];
extern BOOL bPipeInitAkelAdmin;
extern BOOL bSetSecurity;
extern HICON hIconShieldAkelAdmin;

//MessageBox dialog
extern HWND hDlgMsgBox;

//Modeless dialog
extern HWND hDlgModeless;
extern int nModelessType;

//Recode dialog
extern RECT rcRecodeMinMaxDialog;

//Find/Replace dialog
extern RECT rcFindAndReplaceDlg;
extern wchar_t *wszFindText;
extern wchar_t *wszReplaceText;
extern int nFindTextLen;
extern int nReplaceTextLen;
extern BOOL bNoSearchFinishMsg;
extern WORD wLastReplaceButtonID;
extern WNDPROC lpOldComboboxEdit;

//Go to line dialog
extern RECT rcGotoDlg;
extern DWORD dwGotoType;

//Options dialog
extern HHOOK hHookPropertySheet;
extern HWND hPropertyTab;
extern int nPropertyStartPage;
extern BOOL bOptionsSave;
extern BOOL bOptionsRestart;

//Font/Color
extern HSTACK hFontsStack;
extern HSTACK hThemesStack;
extern HSTACK hBkImagesStack;
extern COLORREF crCustColors[16];
extern RECT rcColorsMinMaxDialog;
extern AECOLORS aecDefault;

//Print
extern HWND hWndPreviewEdit;
extern HWND hWndPreviewDlg;
extern HWND hWndZoomEdit;
extern HSTACK hPreviewAllPagesStack;
extern HSTACK hPreviewSelPagesStack;
extern RECT rcPreviewDialog;
extern RECT rcPreviewWindow;
extern RECT rcPreviewPaper;
extern RECT rcPreviewZoomed;
extern POINT ptPreviewScroll;
extern DWORD dwPreviewShowDialog;
extern int lpZoom[];
extern int nPreviewZoomMaxIndex;
extern int nPreviewZoomValue;
extern int nPreviewZoomPercent;
extern int nPreviewCharHeight;
extern int nPreviewAveCharWidth;
extern int nPreviewPageCur;
extern int nPreviewAllPageSum;
extern int nPreviewSelPageSum;
extern BOOL bPreviewSelection;
extern BOOL bPreviewRedrawLock;
extern HHOOK hHookKeys;
extern AEPRINT prn;
extern PRINTINFO prninfo;

//Zooming factor
extern POINT ptScale;

//Edit state
extern AECHARRANGE crCurSel;
extern AECHARINDEX ciCurCaret;
extern int nLoopCase;
extern DWORD dwWordBreakDefault;
extern BOOL bRecentCaretMsg;
extern BOOL bCheckingWriteTime;
extern WNDPROC lpOldEditProc;

//Execute
extern char szExeFile[MAX_PATH];
extern wchar_t wszExeFile[MAX_PATH];
extern int nExeFileLen;
extern char szExeDir[MAX_PATH];
extern wchar_t wszExeDir[MAX_PATH];
extern int nExeDirLen;

//Mdi
extern HSTACK hFramesStack;
extern FRAMEDATA fdInit;
extern FRAMEDATA fdDefault;
extern FRAMEDATA *lpFrameCurrent;
extern FRAMEDATA *lpFramePrevious;
extern int nMDI;
extern HWND hMdiClient;
extern BOOL bMdiMaximize;
extern BOOL bMdiNoWindows;
extern HWND hTab;
extern UINT_PTR dwTabOpenTimer;
extern int nTabOpenItem;
extern int nDocumentsCount;
extern int nDocumentsModified;
extern int nDocumentIndex;
extern DWORD dwChangedPrompt;
extern STACKASSOCICON hIconsStack;
extern HIMAGELIST hImageList;
extern HICON hIconEmpty;
extern BOOL bTabPressing;
extern BOOL bFrameActivating;
extern DWORD dwMdiFrameActivating;
extern RECT rcMdiListMinMaxDialog;
extern WNDPROC lpOldMdiClientProc;
extern WNDPROC lpOldTabProc;
extern FRAMEDATA *lpWndFrame;

//GetProcAddress
extern HMONITOR (WINAPI *MonitorFromPointPtr)(POINT, DWORD);
extern BOOL (WINAPI *GetMonitorInfoAPtr)(HMONITOR, LPMONITORINFO);
extern DWORD (WINAPI *SetSecurityInfoPtr)(HANDLE, SE_OBJECT_TYPE, SECURITY_INFORMATION, PSID, PSID, PACL, PACL);
extern DWORD (WINAPI *SetEntriesInAclWPtr)(ULONG, PEXPLICIT_ACCESSW, PACL, PACL *);
extern BOOL (WINAPI *ShellExecuteExWPtr)(LPSHELLEXECUTEINFOW);


//// Init function

HANDLE CreateEditWindow(HWND hWndParent, HWND hWndEditPMDI)
{
  CREATESTRUCTW cs;
  RECT rcRect;
  HANDLE hResult;

  GetClientRect(hWndParent, &rcRect);

  if (moCur.dwPaintOptions & PAINT_STATICEDGE)
    cs.dwExStyle=WS_EX_STATICEDGE;
  else if (moCur.dwPaintOptions & PAINT_NOEDGE)
    cs.dwExStyle=0;
  else
    cs.dwExStyle=WS_EX_CLIENTEDGE;
  if (moCur.bRichEditClass)
    cs.lpszClass=bOldWindows ? AES_RICHEDIT20A_UNICODE : AES_RICHEDIT20W;
  else
    cs.lpszClass=bOldWindows?L"AkelEditA":L"AkelEditW";
  cs.lpszName=NULL;
  cs.style=WS_CHILD|WS_VISIBLE|WS_HSCROLL|WS_VSCROLL|ES_MULTILINE|
           ((moCur.dwPaintOptions & PAINT_HIDESEL)?0:ES_NOHIDESEL)|
           ((moCur.dwPaintOptions & PAINT_HIDENOSCROLL)?0:ES_DISABLENOSCROLL)|
           ((moCur.dwEditStyle & EDS_GLOBALUNDO)?ES_GLOBALUNDO:0)|
           ((moCur.dwEditStyle & EDS_HEAPSERIALIZE)?ES_HEAPSERIALIZE:0);

  cs.x=0;
  cs.y=0;
  cs.cx=rcRect.right;
  cs.cy=rcRect.bottom;
  if ((nMDI == WMD_SDI || nMDI == WMD_PMDI) && moCur.bStatusBar && IsWindowVisible(hStatus))
    cs.cy-=nStatusHeight;
  cs.hwndParent=hWndParent;
  cs.hMenu=(HMENU)(UINT_PTR)ID_EDIT;
  cs.hInstance=hInstance;
  cs.lpCreateParams=NULL;

  if (hWndEditPMDI)
  {
    hResult=(HANDLE)SendMessage(hWndEditPMDI, AEM_CREATEDOCUMENT, 0, (LPARAM)&cs);
  }
  else
  {
    if (hResult=(HANDLE)CreateWindowExWide(cs.dwExStyle, cs.lpszClass, cs.lpszName, cs.style, cs.x, cs.y, cs.cx, cs.cy, cs.hwndParent, cs.hMenu, cs.hInstance, cs.lpCreateParams))
    {
      if (!aecDefault.dwFlags)
      {
        AECOLORS aec;

        //Get default colors
        aecDefault.dwFlags=AECLR_ALL;
        SendMessage((HWND)hResult, AEM_GETCOLORS, 0, (LPARAM)&aecDefault);
        if (!fdDefault.aec.dwFlags)
          xmemcpy(&fdDefault.aec, &aecDefault, sizeof(AECOLORS));

        //Standard theme
        API_LoadStringW(hLangLib, STR_STANDARDTHEME, wbuf, BUFFER_SIZE);
        StackThemeAdd(&hThemesStack, wbuf, &aecDefault, L"", EDIT_BKIMAGEALPHA, 1);

        if (hThemesStack.first == hThemesStack.last)
        {
          //Notepad++ theme
          aec.crCaret=RGB(0x80, 0x00, 0xFF);
          aec.crBasicText=RGB(0x00, 0x00, 0x00);
          aec.crBasicBk=RGB(0xFF, 0xFF, 0xFF);
          aec.crSelText=RGB(0x00, 0x00, 0x00);
          aec.crSelBk=RGB(0xC0, 0xC0, 0xC0);
          aec.crActiveLineText=RGB(0x00, 0x00, 0x00);
          aec.crActiveLineBk=RGB(0xE8, 0xE8, 0xFF);
          aec.crUrlText=RGB(0x00, 0x00, 0xFF);
          aec.crActiveColumn=RGB(0xE8, 0xE8, 0xFF);
          aec.crColumnMarker=RGB(0xC0, 0xC0, 0xC0);
          aec.crUrlCursorText=RGB(0x00, 0x00, 0x98);
          aec.crUrlVisitText=RGB(0x00, 0x00, 0x98);
          aec.crActiveLineBorder=RGB(0xDF, 0xDF, 0xF5);
          aec.crAltLineText=RGB(0x00, 0x00, 0x00);
          aec.crAltLineBk=RGB(0xF9, 0xF9, 0xF9);
          aec.crAltLineBorder=RGB(0xEF, 0xEF, 0xEF);
          StackThemeAdd(&hThemesStack, L"Notepad++", &aec, L"", EDIT_BKIMAGEALPHA, -1);
        }
      }
      lpOldEditProc=(WNDPROC)GetWindowLongPtrWide((HWND)hResult, GWLP_WNDPROC);
      SetWindowLongPtrWide((HWND)hResult, GWLP_WNDPROC, (UINT_PTR)CommonEditProc);
    }
  }
  return hResult;
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void SetEditWindowSettings(FRAMEDATA *lpFrame)
{
  DWORD dwOptions;
  DWORD dwOptionsEx;

  if (lpFrame->dwLockInherit & LI_FONT)
    xmemcpy(&lpFrame->lf, &fdDefault.lf, sizeof(LOGFONTW));
  if ((lpFrame->dwLockInherit & LI_COLORS) || !lpFrame->aec.dwFlags)
    xmemcpy(&lpFrame->aec, &fdDefault.aec, sizeof(AECOLORS));
  if (lpFrame->dwLockInherit & LI_BKIMAGE)
  {
    xstrcpynW(lpFrame->wszBkImageFile, fdDefault.wszBkImageFile, MAX_PATH);
    lpFrame->nBkImageAlpha=fdDefault.nBkImageAlpha;
  }
  if (lpFrame->dwLockInherit & LI_WRAP)
    lpFrame->ei.bWordWrap=fdDefault.ei.bWordWrap;
  lpFrame->dwLockInherit=0;

  //Set settings
  DoViewReadOnly(lpFrame, lpFrame->ei.bReadOnly, TRUE);
  SetEditNotify(lpFrame->ei.hWndEdit);

  dwOptions=AECO_NOCOLUMNPASTEHOTKEY;
  dwOptionsEx=0;
  if (lpFrame->bDetailedUndo)
    dwOptions|=AECO_DETAILEDUNDO;
  if (lpFrame->dwCaretOptions & CO_CARETOUTEDGE)
    dwOptions|=AECO_CARETOUTEDGE;
  if (lpFrame->dwCaretOptions & CO_CARETVERTLINE)
    dwOptions|=AECO_ACTIVECOLUMN;
  if (lpFrame->dwCaretOptions & CO_CARETACTIVELINE)
    dwOptions|=AECO_ACTIVELINE;
  if (lpFrame->dwCaretOptions & CO_CARETACTIVELINEBORDER)
    dwOptions|=AECO_ACTIVELINEBORDER;
  if (lpFrame->bAltLineBorder)
    dwOptions|=AECO_ALTLINEBORDER;
  if (!(lpFrame->dwMouseOptions & MO_LEFTMARGINSELECTION))
    dwOptions|=AECO_NOMARGINSEL;
  if (lpFrame->dwMouseOptions & MO_RICHEDITMOUSE)
    dwOptions|=AECO_LBUTTONUPCONTINUECAPTURE;
  if (!(lpFrame->dwMouseOptions & MO_MOUSEDRAGGING))
    dwOptions|=AECO_DISABLEDRAG;
  if (lpFrame->dwMouseOptions & MO_RCLICKMOVECARET)
    dwOptions|=AECO_RBUTTONDOWNMOVECARET;
  if (lpFrame->dwMouseOptions & MO_NONEWLINEMOUSESELECT)
    dwOptions|=AECO_NONEWLINEMOUSESELECT;
  if (lpFrame->dwMouseOptions & MO_MARGINSELUNWRAPLINE)
    dwOptions|=AECO_MARGINSELUNWRAPLINE;
  if (lpFrame->dwMouseOptions & MO_MBUTTONDOWNNOSCROLL)
    dwOptions|=AECO_MBUTTONDOWNNOSCROLL;
  if (lpFrame->dwMouseOptions & MO_INVERTHORZWHEEL)
    dwOptionsEx|=AECOE_INVERTHORZWHEEL;
  if (lpFrame->dwMouseOptions & MO_INVERTVERTWHEEL)
    dwOptionsEx|=AECOE_INVERTVERTWHEEL;
  if (moCur.dwPaintOptions & PAINT_PAINTGROUP)
    dwOptions|=AECO_PAINTGROUP;
  if (moCur.dwPaintOptions & PAINT_NONEWLINEDRAW)
    dwOptions|=AECO_NONEWLINEDRAW;
  if (moCur.dwPaintOptions & PAINT_ENTIRENEWLINEDRAW)
    dwOptions|=AECO_ENTIRENEWLINEDRAW;
  if (moCur.dwPaintOptions & PAINT_NOMARKERAFTERLASTLINE)
    dwOptions|=AECO_NOMARKERAFTERLASTLINE;
  SendMessage(lpFrame->ei.hWndEdit, AEM_SETOPTIONS, AECOOP_OR, dwOptions);
  if (dwOptionsEx)
    SendMessage(lpFrame->ei.hWndEdit, AEM_EXSETOPTIONS, AECOOP_OR, dwOptionsEx);

  //Font
  SetChosenFont(lpFrame->ei.hWndEdit, &lpFrame->lf);
  UpdateMappedPrintWidth(lpFrame);

  //Colors
  SendMessage(lpFrame->ei.hWndEdit, AEM_SETCOLORS, 0, (LPARAM)&lpFrame->aec);
  SetBkImage(lpFrame, lpFrame->wszBkImageFile, lpFrame->nBkImageAlpha);

  SendMessage(lpFrame->ei.hWndEdit, AEM_SETUNDOLIMIT, (WPARAM)lpFrame->nUndoLimit, 0);
  SetMargins(lpFrame->ei.hWndEdit, &lpFrame->rcEditMargins, NULL);
  SetTabStops(lpFrame->ei.hWndEdit, lpFrame->nTabStopSize, FALSE);
  DoViewWordWrap(lpFrame, lpFrame->ei.bWordWrap, TRUE);
  SetNewLineStatus(lpFrame, lpFrame->ei.nNewLine, AENL_INPUT);

  if (lpFrame->dwMarker)
  {
    SetMarker(lpFrame, lpFrame->dwMarker);
  }
  if (lpFrame->dwLineGap)
  {
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETLINEGAP, lpFrame->dwLineGap, 0);
  }
  if (lpFrame->nCaretWidth != 1)
  {
    POINT pt;

    SendMessage(lpFrame->ei.hWndEdit, AEM_GETCARETWIDTH, 0, (LPARAM)&pt);
    pt.x=lpFrame->nCaretWidth;
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETCARETWIDTH, 0, (LPARAM)&pt);
  }
  if (lpFrame->dwAltLineSkip && lpFrame->dwAltLineFill)
  {
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETALTLINE, MAKELONG(lpFrame->dwAltLineSkip, lpFrame->dwAltLineFill), 0);
  }
  if (lpFrame->bShowURL)
  {
    SendMessage(lpFrame->ei.hWndEdit, AEM_EXSETOPTIONS, AECOOP_OR, AECOE_DETECTURL);

    if (lpFrame->bUrlPrefixesEnable)
      SetUrlPrefixes(lpFrame->ei.hWndEdit, lpFrame->wszUrlPrefixes);
    if (lpFrame->bUrlDelimitersEnable)
    {
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLLEFTDELIMITERS, 0, (LPARAM)lpFrame->wszUrlLeftDelimiters);
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLRIGHTDELIMITERS, 0, (LPARAM)lpFrame->wszUrlRightDelimiters);
    }
  }
  if (dwWordBreakDefault == (DWORD)-1)
    dwWordBreakDefault=(DWORD)SendMessage(lpFrame->ei.hWndEdit, AEM_GETWORDBREAK, 0, 0);
  if (lpFrame->bWordDelimitersEnable)
  {
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDBREAK, moCur.dwWordBreakCustom, 0);
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDDELIMITERS, 0, (LPARAM)lpFrame->wszWordDelimiters);
  }
  if (lpFrame->bWrapDelimitersEnable)
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETWRAPDELIMITERS, 0, (LPARAM)lpFrame->wszWrapDelimiters);

  UpdateStatusUser(lpFrame, CSB_FONTPOINT|CSB_TABSIZE|CSB_MARKER|CSB_CAPSLOCK|CSB_NUMLOCK);
}

void SetEditNotify(HWND hWnd)
{
  SendMessage(hWnd, AEM_SETEVENTMASK, 0, AENM_SCROLL|AENM_PROGRESS|AENM_MODIFY|AENM_SELCHANGE|AENM_TEXTCHANGE|AENM_TEXTINSERT|AENM_TEXTDELETE|AENM_DRAGDROP|AENM_POINT|AENM_LINK|AENM_MARKER);
  SendMessage(hWnd, EM_SETEVENTMASK, 0, ENM_SELCHANGE|ENM_CHANGE|ENM_LINK);
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void ResizeEditWindow(FRAMEDATA *lpFrame, DWORD dwFlags)
{
  RECT *lprcEditWindow;

  if (nMDI == WMD_MDI)
    lprcEditWindow=&lpFrame->rcEditWindow;
  else
    lprcEditWindow=&fdDefault.rcEditWindow;
  UpdateWindow(hStatus);

  if (lpFrame->ei.hWndMaster)
  {
    RECT rc;

    lpFrame->rcMasterWindow.left=lprcEditWindow->left;
    lpFrame->rcMasterWindow.top=lprcEditWindow->top;

    if (lpFrame->ei.hWndClone1)
    {
      lpFrame->rcMasterWindow.right=max(lpFrame->rcMasterWindow.right, 40);
      lpFrame->rcMasterWindow.right=min(lpFrame->rcMasterWindow.right, lprcEditWindow->right - 40);
      lpFrame->rcMasterWindow.right=max(lpFrame->rcMasterWindow.right, 0);
    }
    else lpFrame->rcMasterWindow.right=lprcEditWindow->right;

    if (lpFrame->ei.hWndClone2)
    {
      lpFrame->rcMasterWindow.bottom=max(lpFrame->rcMasterWindow.bottom, 40);
      lpFrame->rcMasterWindow.bottom=min(lpFrame->rcMasterWindow.bottom, lprcEditWindow->bottom - 40);
      lpFrame->rcMasterWindow.bottom=max(lpFrame->rcMasterWindow.bottom, 0);
    }
    else lpFrame->rcMasterWindow.bottom=lprcEditWindow->bottom;

    if (lpFrame->ei.hWndMaster)
    {
      rc=lpFrame->rcMasterWindow;

      if (!(dwFlags & REW_TEST))
      {
        MoveWindow(lpFrame->ei.hWndMaster, rc.left, rc.top, rc.right, rc.bottom, (dwFlags & REW_NOREDRAW)?FALSE:TRUE);
      }
      else
      {
        ClientToScreenRect(lpFrame->hWndEditParent, &rc);
        DrawMovingRect(&rc);
      }
    }
    if (lpFrame->ei.hWndClone1)
    {
      rc.left=lpFrame->rcMasterWindow.left + lpFrame->rcMasterWindow.right;
      rc.top=lpFrame->rcMasterWindow.top;
      rc.right=lprcEditWindow->right - lpFrame->rcMasterWindow.right;
      rc.bottom=lpFrame->rcMasterWindow.bottom;

      if (!(dwFlags & REW_TEST))
      {
        MoveWindow(lpFrame->ei.hWndClone1, rc.left, rc.top, rc.right, rc.bottom, (dwFlags & REW_NOREDRAW)?FALSE:TRUE);
      }
      else
      {
        ClientToScreenRect(lpFrame->hWndEditParent, &rc);
        DrawMovingRect(&rc);
      }
    }
    if (lpFrame->ei.hWndClone2)
    {
      rc.left=lpFrame->rcMasterWindow.left;
      rc.top=lpFrame->rcMasterWindow.top + lpFrame->rcMasterWindow.bottom;
      rc.right=lpFrame->rcMasterWindow.right;
      rc.bottom=lprcEditWindow->bottom - lpFrame->rcMasterWindow.bottom;

      if (!(dwFlags & REW_TEST))
      {
        MoveWindow(lpFrame->ei.hWndClone2, rc.left, rc.top, rc.right, rc.bottom, (dwFlags & REW_NOREDRAW)?FALSE:TRUE);
      }
      else
      {
        ClientToScreenRect(lpFrame->hWndEditParent, &rc);
        DrawMovingRect(&rc);
      }
    }
    if (lpFrame->ei.hWndClone3)
    {
      rc.left=lpFrame->rcMasterWindow.left + lpFrame->rcMasterWindow.right;
      rc.top=lpFrame->rcMasterWindow.top + lpFrame->rcMasterWindow.bottom;
      rc.right=lprcEditWindow->right - lpFrame->rcMasterWindow.right;
      rc.bottom=lprcEditWindow->bottom - lpFrame->rcMasterWindow.bottom;

      if (!(dwFlags & REW_TEST))
      {
        MoveWindow(lpFrame->ei.hWndClone3, rc.left, rc.top, rc.right, rc.bottom, (dwFlags & REW_NOREDRAW)?FALSE:TRUE);
      }
      else
      {
        ClientToScreenRect(lpFrame->hWndEditParent, &rc);
        DrawMovingRect(&rc);
      }
    }
  }
  else
  {
    if (!(dwFlags & REW_TEST))
    {
      if (lpFrame->ei.hWndEdit)
        MoveWindow(lpFrame->ei.hWndEdit, lprcEditWindow->left, lprcEditWindow->top, lprcEditWindow->right, lprcEditWindow->bottom, (dwFlags & REW_NOREDRAW)?FALSE:TRUE);
    }
  }

  //Remember current frame edit RECT
  if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
    lpFrame->rcEditWindow=fdDefault.rcEditWindow;
}

FRAMEDATA* GetFrameDataFromEditWindow(HWND hWndEditCtrl)
{
  if (lpFrameCurrent->ei.hWndEdit)
  {
    if (!hWndEditCtrl || hWndEditCtrl == lpFrameCurrent->ei.hWndEdit)
      return lpFrameCurrent;
    if (lpFrameCurrent->ei.hWndMaster)
    {
      if (hWndEditCtrl == lpFrameCurrent->ei.hWndMaster ||
          hWndEditCtrl == lpFrameCurrent->ei.hWndClone1 ||
          hWndEditCtrl == lpFrameCurrent->ei.hWndClone2 ||
          hWndEditCtrl == lpFrameCurrent->ei.hWndClone3)
        return lpFrameCurrent;
    }
    if (nMDI == WMD_MDI)
    {
      HWND hWndParent;

      if (GetWindowLongPtrWide(hWndEditCtrl, GWLP_ID) == ID_EDIT)
      {
        if (hWndParent=GetParent(hWndEditCtrl))
        {
          return (FRAMEDATA *)GetWindowLongPtrWide(hWndParent, GWLP_USERDATA);
        }
      }
    }
  }
  return NULL;
}

FRAMEDATA* GetFrameDataFromEditDocument(AEHDOC hDocEditHandle)
{
  if (lpFrameCurrent->ei.hDocEdit)
  {
    if (!hDocEditHandle || hDocEditHandle == lpFrameCurrent->ei.hDocEdit)
      return lpFrameCurrent;
    if (lpFrameCurrent->ei.hDocMaster)
    {
      if (hDocEditHandle == lpFrameCurrent->ei.hDocMaster ||
          hDocEditHandle == lpFrameCurrent->ei.hDocClone1 ||
          hDocEditHandle == lpFrameCurrent->ei.hDocClone2 ||
          hDocEditHandle == lpFrameCurrent->ei.hDocClone3)
        return lpFrameCurrent;
    }
    return StackFrameGetByHandle(&hFramesStack, hDocEditHandle);
  }
  return NULL;
}

FRAMEDATA* CreateFrameData(HWND hWndEditParent, FRAMEDATA *lpFrameSource)
{
  FRAMEDATA *lpFrame;

  if (lpFrame=StackFrameInsert(&hFramesStack))
  {
    //Initialize settings
    CopyFrameData(lpFrame, lpFrameSource);
    lpFrame->hWndEditParent=hWndEditParent;
  }
  return lpFrame;
}

void CopyFrameData(FRAMEDATA *lpFrameTarget, FRAMEDATA *lpFrameSource)
{
  //Copy settings from lpFrameSource
  xmemcpy(((LPBYTE)lpFrameTarget) + sizeof(HSTACK), ((LPBYTE)lpFrameSource) + sizeof(HSTACK), offsetof(FRAMEDATA, lpEditProc) - sizeof(HSTACK));

  //Initialize own settings
  lpFrameTarget->hWndEditParent=NULL;
  lpFrameTarget->ei.hWndEdit=NULL;
  lpFrameTarget->ei.hDocEdit=NULL;
  lpFrameTarget->ei.pFile=bOldWindows?(LPBYTE)lpFrameTarget->szFile:(LPBYTE)lpFrameTarget->wszFile;
  lpFrameTarget->ei.szFile=lpFrameTarget->szFile;
  lpFrameTarget->ei.wszFile=lpFrameTarget->wszFile;
  lpFrameTarget->ei.nCodePage=moCur.nNewFileCodePage;
  lpFrameTarget->ei.bBOM=moCur.bNewFileBOM;
  lpFrameTarget->ei.nNewLine=moCur.nNewFileNewLine;
  lpFrameTarget->ei.bModified=FALSE;
  lpFrameTarget->ei.bReadOnly=FALSE;
  //lpFrameTarget->ei.bWordWrap=lpFrameSource->ei.bWordWrap;
  lpFrameTarget->ei.bOvertypeMode=FALSE;
  lpFrameTarget->ei.hWndMaster=NULL;
  lpFrameTarget->ei.hDocMaster=NULL;
  lpFrameTarget->ei.hWndClone1=NULL;
  lpFrameTarget->ei.hDocClone1=NULL;
  lpFrameTarget->ei.hWndClone2=NULL;
  lpFrameTarget->ei.hDocClone2=NULL;
  lpFrameTarget->ei.hWndClone3=NULL;
  lpFrameTarget->ei.hDocClone3=NULL;
  lpFrameTarget->szFile[0]='\0';
  lpFrameTarget->wszFile[0]=L'\0';
  lpFrameTarget->nFileLen=0;
  lpFrameTarget->nStreamOffset=0;
  lpFrameTarget->hIcon=hIconEmpty;
  lpFrameTarget->nIconIndex=0;

  lpFrameTarget->dwInputLocale=(HKL)(UINT_PTR)-1;
}

void CopyFrameSettings(FRAMEDATA *lpFrameTarget, FRAMEDATA *lpFrameSource)
{
  xmemcpy(&lpFrameTarget->lf, &lpFrameSource->lf, offsetof(FRAMEDATA, lpEditProc) - offsetof(FRAMEDATA, lf));
  lpFrameTarget->ei.bWordWrap=lpFrameSource->ei.bWordWrap;
  xstrcpynA(lpFrameTarget->szFile, lpFrameSource->szFile, MAX_PATH);
  xstrcpynW(lpFrameTarget->wszFile, lpFrameSource->wszFile, MAX_PATH);
  lpFrameTarget->nFileLen=lpFrameSource->nFileLen;
  lpFrameTarget->hIcon=lpFrameSource->hIcon;
  lpFrameTarget->nIconIndex=lpFrameSource->nIconIndex;
}

void SaveFrameData(FRAMEDATA *lpFrame)
{
  if (nMDI)
  {
    //Remember keyboard layout
    if (moCur.dwKeybLayoutOptions & KLO_REMEMBERLAYOUT)
    {
      lpFrame->dwInputLocale=(HKL)GetKeyboardLayout(0);
    }
  }
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void RestoreFrameData(FRAMEDATA *lpFrame, DWORD dwFlagsPMDI)
{
  if (nMDI == WMD_PMDI)
  {
    DWORD dwSetDataFlags=0;

    if (dwFlagsPMDI & FWA_NOVISUPDATE)
      dwSetDataFlags|=AESWD_NOALL;
    else if (dwFlagsPMDI & FWA_NOUPDATEEDIT)
      dwSetDataFlags|=AESWD_NOREDRAW;

    if (lpFrame->ei.hDocMaster)
    {
      SendMessage(lpFrame->ei.hWndMaster, AEM_SETDOCUMENT, (WPARAM)lpFrame->ei.hDocMaster, dwSetDataFlags);
      if (lpFrame->ei.hDocClone1)
        SendMessage(lpFrame->ei.hWndClone1, AEM_SETDOCUMENT, (WPARAM)lpFrame->ei.hDocClone1, dwSetDataFlags);
      if (lpFrame->ei.hDocClone2)
        SendMessage(lpFrame->ei.hWndClone2, AEM_SETDOCUMENT, (WPARAM)lpFrame->ei.hDocClone2, dwSetDataFlags);
      if (lpFrame->ei.hDocClone3)
        SendMessage(lpFrame->ei.hWndClone3, AEM_SETDOCUMENT, (WPARAM)lpFrame->ei.hDocClone3, dwSetDataFlags);
    }
    else if (lpFrame->ei.hDocEdit)
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETDOCUMENT, (WPARAM)lpFrame->ei.hDocEdit, dwSetDataFlags);

    //Edit window focus in other frame can be different
    if (GetFocus() != lpFrame->ei.hWndEdit)
      SetFocus(lpFrame->ei.hWndEdit);

    //If window size has been changed, update virtual window according to current window size
    if (xmemcmp(&lpFrame->rcEditWindow, &fdDefault.rcEditWindow, sizeof(RECT)))
    {
      SendMessage(lpFrame->ei.hWndEdit, AEM_UPDATESIZE, 0, 0);
      lpFrame->rcEditWindow=fdDefault.rcEditWindow;
    }

    if (!(dwFlagsPMDI & FWA_NOVISUPDATE))
    {
      SplitVisUpdate(lpFrame);
      ResizeEditWindow(lpFrame, (dwFlagsPMDI & FWA_NOUPDATEEDIT)?REW_NOREDRAW:0);
    }
  }
  //Update selection to set valid globals: crCurSel, ciCurCaret and nSelSubtract
  SetSelectionStatus(lpFrame->ei.hDocEdit, lpFrame->ei.hWndEdit, NULL, NULL);
  SetCodePageStatus(NULL, lpFrame->ei.nCodePage, lpFrame->ei.bBOM);
  SetNewLineStatus(NULL, lpFrame->ei.nNewLine, 0);
  SetModifyStatus(NULL, lpFrame->ei.bModified);
  SetOvertypeStatus(NULL, lpFrame->ei.bOvertypeMode);

  if (nMDI)
  {
    //Activate keyboard layout
    if (moCur.dwKeybLayoutOptions & KLO_REMEMBERLAYOUT)
    {
      ActivateKeyboard(lpFrame->dwInputLocale);
    }

    //Update tabs
    if (!bTabPressing)
    {
      int nTabItem;

      if ((nTabItem=GetTabItemFromParam(hTab, (LPARAM)lpFrame)) != -1)
      {
        nDocumentIndex=nTabItem;
        SendMessage(hTab, TCM_SETCURSEL, nTabItem, 0);
        UpdateTabs(hTab);
      }
    }

    //Set text for all parts
    SetTextStatusUser(lpFrame, (DWORD)-1);
  }
}

BOOL CreateFrameWindow(RECT *rcRectMDI)
{
  BOOL bResult=FALSE;

  if (nMDI == WMD_MDI)
  {
    DWORD dwStyle;

    if (lpFrameCurrent->hWndEditParent) SendMessage(hMdiClient, WM_MDIGETACTIVE, 0, (LPARAM)&bMdiMaximize);
    if (bMdiMaximize == -1)
      dwStyle=moCur.dwMdiStyle;
    else
      dwStyle=!bMdiMaximize?0:WS_MAXIMIZE;

    dwMdiFrameActivating=FWA_NOTIFY_CREATE;
    CreateMDIWindowWide(APP_MDI_CLASSW, L"", dwStyle, rcRectMDI?(DWORD)rcRectMDI->left:(DWORD)CW_USEDEFAULT, rcRectMDI?(DWORD)rcRectMDI->top:(DWORD)CW_USEDEFAULT, rcRectMDI?(DWORD)rcRectMDI->right:(DWORD)CW_USEDEFAULT, rcRectMDI?(DWORD)rcRectMDI->bottom:(DWORD)CW_USEDEFAULT, hMdiClient, hInstance, 0);
    bResult=TRUE;
  }
  else if (nMDI == WMD_PMDI)
  {
    FRAMEDATA *lpFrame;

    if (lpFrame=CreateFrameData(hMainWnd, lpFrameCurrent))
    {
      //Create virtual document. Procedure and window aren't changed.
      lpFrame->ei.hDocEdit=(AEHDOC)CreateEditWindow(lpFrame->hWndEditParent, fdDefault.ei.hWndEdit);
      lpFrame->lpEditProc=fdDefault.lpEditProc;
      lpFrame->ei.hWndEdit=fdDefault.ei.hWndEdit;

      InsertTabItem(hTab, (moCur.dwTabOptionsMDI & TAB_ADD_AFTERCURRENT)?nDocumentIndex + 1:-1, (LPARAM)lpFrame);

      ActivateFrameWindow(lpFrame, FWA_NOTIFY_CREATE);
      SetEditWindowSettings(lpFrameCurrent);
      SendMessage(hMainWnd, AKDN_EDIT_ONSTART, (WPARAM)lpFrameCurrent->ei.hWndEdit, (LPARAM)lpFrameCurrent->ei.hDocEdit);

      //Update status
      ++nDocumentsCount;
      UpdateStatusUser(lpFrameCurrent, CSB_DOCUMENTSCOUNT|CSB_DOCUMENTSMODIFIED|CSB_DOCUMENTSSAVED);

      bResult=TRUE;
    }
  }
  return bResult;
}

FRAMEDATA* ActivateFrameWindow(FRAMEDATA *lpFrame, DWORD dwFlags)
{
  FRAMEDATA *lpFrameLostFocus=lpFrameCurrent;

  bFrameActivating=TRUE;

  if (lpFrameCurrent != lpFrame)
  {
    //Save deactivated frame data
    if (lpFrameCurrent == &fdDefault)
      lpFramePrevious=NULL;
    else
      lpFramePrevious=lpFrameCurrent;
    SendMessage(hMainWnd, AKDN_FRAME_DEACTIVATE, dwFlags, (LPARAM)lpFramePrevious);

    if (nMDI == WMD_MDI)
    {
      //Activate frame
      dwMdiFrameActivating=dwFlags;
      SendMessage(hMdiClient, WM_MDIACTIVATE, (WPARAM)lpFrame->hWndEditParent, 0);
    }
    else if (nMDI == WMD_PMDI)
    {
      if (!(dwFlags & FWA_NOUPDATEORDER))
      {
        //Move item to the end of stack, to use access order later.
        StackFrameMove(&hFramesStack, lpFrame, -1);
      }

      if (lpFrameCurrent->ei.hDocEdit)
        SaveFrameData(lpFrameCurrent);

      //Restore activated frame data
      lpFrameCurrent=lpFrame;
      RestoreFrameData(lpFrameCurrent, dwFlags);

      //Set caption of main window
      if (lpFrameCurrent->wszFile[0] || lpFrameCurrent->ei.bModified)
      {
        xprintfW(wbuf, L"%s - [%s%s]", APP_MAIN_TITLEW, lpFrameCurrent->wszFile, ((moCur.dwShowModify & SM_FRAMETITLE_MDI) && lpFrameCurrent->ei.bModified)?L" *":L"");
        SetWindowTextWide(hMainWnd, wbuf);
      }
      else SetWindowTextWide(hMainWnd, APP_MAIN_TITLEW);

      SendMessage(hMainWnd, AKDN_FRAME_ACTIVATE, dwFlags, (LPARAM)lpFrameCurrent);
    }
  }
  bFrameActivating=FALSE;

  return lpFrameLostFocus;
}

FRAMEDATA* ActivateNextFrameWindow(FRAMEDATA *lpFrame, BOOL bPrev)
{
  if (moCur.dwTabOptionsMDI & TAB_SWITCH_RIGHTLEFT)
    lpFrame=GetNextTabFrame(lpFrame, bPrev);
  else if (moCur.dwTabOptionsMDI & TAB_SWITCH_NEXTPREV)
    lpFrame=StackFrameGetNext(&hFramesStack, lpFrame, bPrev);

  if (lpFrame)
    ActivateFrameWindow(lpFrame, FWA_NOUPDATEORDER);
  return lpFrameCurrent;
}

FRAMEDATA* GetNextTabFrame(FRAMEDATA *lpFrame, BOOL bPrev)
{
  int nTabItem;
  int nItemCount;

  if ((nTabItem=GetTabItemFromParam(hTab, (LPARAM)lpFrame)) != -1)
  {
    nItemCount=nDocumentsCount - 1;

    if (bPrev)
    {
      if (nTabItem - 1 < 0)
        nTabItem=nItemCount;
      else
        --nTabItem;
    }
    else
    {
      if (nTabItem + 1 > nItemCount)
        nTabItem=0;
      else
        ++nTabItem;
    }
    return (FRAMEDATA *)GetTabParamFromItem(hTab, nTabItem);
  }
  return NULL;
}

int DestroyFrameWindow(FRAMEDATA *lpFrame)
{
  if (lpFrame->ei.hWndEdit)
  {
    if (nMDI == WMD_MDI)
    {
      if (SendMessage(hMdiClient, WM_MDIDESTROY, (WPARAM)lpFrame->hWndEditParent, 0))
        return FWDE_ABORT;
    }
    else if (nMDI == WMD_PMDI)
    {
      FRAMEDATA *lpFrameToActivate;
      int nTabItem;
      DWORD dwPrompt=dwChangedPrompt;

      //Get previous frame
      if (lpFrame == lpFrameCurrent)
      {
        if (lpFramePrevious)
          lpFrameToActivate=lpFramePrevious;
        else
          lpFrameToActivate=StackFrameGetNext(&hFramesStack, lpFrame, TRUE);
      }
      else lpFrameToActivate=lpFrameCurrent;

      //Is save prompt required
      if ((dwChangedPrompt & PROMPT_NONE) || !lpFrame->ei.bModified || (moCur.bSilentCloseEmptyMDI && !lpFrame->ei.wszFile[0] && !GetTextLength(lpFrame->ei.hWndEdit)))
        dwPrompt|=PROMPT_NONE;

      //Activate frame
      if (dwPrompt & PROMPT_NONE)
      {
        bEditOnFinish=TRUE;
        ActivateFrameWindow(lpFrame, FWA_NOTIFY_BEFOREDESTROY|FWA_NOVISUPDATE);
        bEditOnFinish=FALSE;
      }
      else ActivateFrameWindow(lpFrame, FWA_NOTIFY_BEFOREDESTROY);

      if (lpFrame == lpFrameToActivate)
      {
        //Don't destroy last tab
        if (!CloseDocument(dwPrompt)) return FWDE_ABORT;
        SendMessage(hMainWnd, AKDN_FRAME_NOWINDOWS, 0, 0);

        //Save frame settings
        if (!nMainOnFinish || bFirstTabOnFinish)
        {
          bFirstTabOnFinish=FALSE;
          CopyFrameSettings(&fdDefault, lpFrame);
        }
        return FWDE_LASTTAB;
      }
      else
      {
        //Ask if document unsaved
        if (!SaveChanged(dwPrompt)) return FWDE_ABORT;
        RecentFilesSaveFile(lpFrame);
      }

      if ((nTabItem=GetTabItemFromParam(hTab, (LPARAM)lpFrame)) != -1)
      {
        //Update status
        if (lpFrame->ei.bModified)
          --nDocumentsModified;
        --nDocumentsCount;
        UpdateStatusUser(lpFrame, CSB_DOCUMENTSCOUNT|CSB_DOCUMENTSMODIFIED|CSB_DOCUMENTSSAVED);

        SendMessage(hMainWnd, AKDN_FRAME_DESTROY, 0, (LPARAM)lpFrame);

        //Destroy active document
        SplitDestroy(lpFrame, CN_CLONE1|CN_CLONE2|CN_CLONE3);
        SendMessage(hMainWnd, AKDN_EDIT_ONFINISH, (WPARAM)lpFrame->ei.hWndEdit, (LPARAM)lpFrame->ei.hDocEdit);
        SendMessage(lpFrame->ei.hWndEdit, AEM_SETDOCUMENT, (WPARAM)fdDefault.ei.hDocEdit, AESWD_NOREDRAW);
        SendMessage(lpFrame->ei.hWndEdit, AEM_DELETEDOCUMENT, (WPARAM)lpFrame->ei.hDocEdit, 0);

        //Save frame settings
        if (!nMainOnFinish || bFirstTabOnFinish)
        {
          bFirstTabOnFinish=FALSE;
          CopyFrameSettings(&fdDefault, lpFrame);
        }
        StackFrameDelete(&hFramesStack, lpFrame);

        //Remove tab item
        DeleteTabItem(hTab, nTabItem);

        //Activate previous window
        if (lpFrameToActivate)
          ActivateFrameWindow(lpFrameToActivate, FWA_NOTIFY_AFTERDESTROY);
      }
    }
    return FWDE_SUCCESS;
  }
  return FWDE_NOWINDOW;
}

BOOL FrameNoWindows()
{
  if (nMDI == WMD_MDI)
  {
    if (!lpFrameCurrent->hWndEditParent)
      return TRUE;
  }
  else
  {
    if (lpFrameCurrent == (FRAMEDATA *)hFramesStack.first &&
        lpFrameCurrent == (FRAMEDATA *)hFramesStack.last &&
        !lpFrameCurrent->ei.bModified &&
        !lpFrameCurrent->ei.wszFile[0])
    {
      return TRUE;
    }
  }
  return FALSE;
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void SplitCreate(FRAMEDATA *lpFrame, DWORD dwFlags)
{
  if (!lpFrame->ei.hWndMaster)
  {
    lpFrame->ei.hWndMaster=lpFrame->ei.hWndEdit;
    lpFrame->ei.hDocMaster=lpFrame->ei.hDocEdit;
    if (!fdDefault.ei.hWndMaster)
    {
      fdDefault.ei.hWndMaster=fdDefault.ei.hWndEdit;
      fdDefault.ei.hDocMaster=fdDefault.ei.hDocEdit;
    }

    if (dwFlags & CN_CLONE1)
    {
      if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
      {
        //Create real edit window for clone #1 if not created before
        if (!fdDefault.ei.hWndClone1)
        {
          fdDefault.ei.hWndClone1=(HWND)CreateEditWindow(hMainWnd, NULL);
          fdDefault.ei.hDocClone1=(AEHDOC)SendMessage(fdDefault.ei.hWndClone1, AEM_GETDOCUMENT, 0, 0);
        }
        lpFrame->ei.hWndClone1=fdDefault.ei.hWndClone1;

        if (nMDI == WMD_SDI)
        {
          lpFrame->ei.hDocClone1=fdDefault.ei.hDocClone1;
        }
        else if (nMDI == WMD_PMDI)
        {
          //Create and assign virtual document
          lpFrame->ei.hDocClone1=(AEHDOC)CreateEditWindow(lpFrame->hWndEditParent, lpFrame->ei.hWndClone1);
          SendMessage(lpFrame->ei.hWndClone1, AEM_SETDOCUMENT, (WPARAM)lpFrame->ei.hDocClone1, AESWD_NOREDRAW);
        }
      }
      else if (nMDI == WMD_MDI)
      {
        lpFrame->ei.hWndClone1=(HWND)CreateEditWindow(lpFrame->hWndEditParent, NULL);
        lpFrame->ei.hDocClone1=(AEHDOC)SendMessage(lpFrame->ei.hWndClone1, AEM_GETDOCUMENT, 0, 0);
      }
      SendMessage(lpFrame->ei.hWndClone1, AEM_SETERASERECT, 0, (LPARAM)NULL);
      SendMessage(lpFrame->ei.hWndClone1, AEM_SETRECT, AERC_MARGINS|AERC_UPDATE, (LPARAM)&lpFrame->rcEditMargins);
      SendMessage(lpFrame->ei.hWndMaster, AEM_ADDCLONE, (WPARAM)lpFrame->ei.hWndClone1, 0);
      SendMessage(hMainWnd, AKDN_EDIT_ONSTART, (WPARAM)lpFrame->ei.hWndClone1, (LPARAM)lpFrame->ei.hDocClone1);
    }
    if (dwFlags & CN_CLONE2)
    {
      if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
      {
        //Create real edit window for clone #2 if not created before
        if (!fdDefault.ei.hWndClone2)
        {
          fdDefault.ei.hWndClone2=(HWND)CreateEditWindow(hMainWnd, NULL);
          fdDefault.ei.hDocClone2=(AEHDOC)SendMessage(fdDefault.ei.hWndClone2, AEM_GETDOCUMENT, 0, 0);
        }
        lpFrame->ei.hWndClone2=fdDefault.ei.hWndClone2;

        if (nMDI == WMD_SDI)
        {
          lpFrame->ei.hDocClone2=fdDefault.ei.hDocClone2;
        }
        else if (nMDI == WMD_PMDI)
        {
          //Create and assign virtual document
          lpFrame->ei.hDocClone2=(AEHDOC)CreateEditWindow(lpFrame->hWndEditParent, lpFrame->ei.hWndClone2);
          SendMessage(lpFrame->ei.hWndClone2, AEM_SETDOCUMENT, (WPARAM)lpFrame->ei.hDocClone2, AESWD_NOREDRAW);
        }
      }
      else if (nMDI == WMD_MDI)
      {
        lpFrame->ei.hWndClone2=(HWND)CreateEditWindow(lpFrame->hWndEditParent, NULL);
        lpFrame->ei.hDocClone2=(AEHDOC)SendMessage(lpFrame->ei.hWndClone2, AEM_GETDOCUMENT, 0, 0);
      }
      SendMessage(lpFrame->ei.hWndClone2, AEM_SETERASERECT, 0, (LPARAM)NULL);
      SendMessage(lpFrame->ei.hWndClone2, AEM_SETRECT, AERC_MARGINS|AERC_UPDATE, (LPARAM)&lpFrame->rcEditMargins);
      SendMessage(lpFrame->ei.hWndMaster, AEM_ADDCLONE, (WPARAM)lpFrame->ei.hWndClone2, 0);
      SendMessage(hMainWnd, AKDN_EDIT_ONSTART, (WPARAM)lpFrame->ei.hWndClone2, (LPARAM)lpFrame->ei.hDocClone2);
    }
    if (dwFlags & CN_CLONE3)
    {
      if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
      {
        //Create real edit window for clone #3 if not created before
        if (!fdDefault.ei.hWndClone3)
        {
          fdDefault.ei.hWndClone3=(HWND)CreateEditWindow(hMainWnd, NULL);
          fdDefault.ei.hDocClone3=(AEHDOC)SendMessage(fdDefault.ei.hWndClone3, AEM_GETDOCUMENT, 0, 0);
        }
        lpFrame->ei.hWndClone3=fdDefault.ei.hWndClone3;

        if (nMDI == WMD_SDI)
        {
          lpFrame->ei.hDocClone3=fdDefault.ei.hDocClone3;
        }
        else if (nMDI == WMD_PMDI)
        {
          //Create and assign virtual document
          lpFrame->ei.hDocClone3=(AEHDOC)CreateEditWindow(lpFrame->hWndEditParent, lpFrame->ei.hWndClone3);
          SendMessage(lpFrame->ei.hWndClone3, AEM_SETDOCUMENT, (WPARAM)lpFrame->ei.hDocClone3, AESWD_NOREDRAW);
        }
      }
      else if (nMDI == WMD_MDI)
      {
        lpFrame->ei.hWndClone3=(HWND)CreateEditWindow(lpFrame->hWndEditParent, NULL);
        lpFrame->ei.hDocClone3=(AEHDOC)SendMessage(lpFrame->ei.hWndClone3, AEM_GETDOCUMENT, 0, 0);
      }
      SendMessage(lpFrame->ei.hWndClone3, AEM_SETERASERECT, 0, (LPARAM)NULL);
      SendMessage(lpFrame->ei.hWndClone3, AEM_SETRECT, AERC_MARGINS|AERC_UPDATE, (LPARAM)&lpFrame->rcEditMargins);
      SendMessage(lpFrame->ei.hWndMaster, AEM_ADDCLONE, (WPARAM)lpFrame->ei.hWndClone3, 0);
      SendMessage(hMainWnd, AKDN_EDIT_ONSTART, (WPARAM)lpFrame->ei.hWndClone3, (LPARAM)lpFrame->ei.hDocClone3);
    }
    SplitVisUpdate(lpFrame);
  }
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void SplitDestroy(FRAMEDATA *lpFrame, DWORD dwFlags)
{
  bEditOnFinish=TRUE;

  if (lpFrame->ei.hWndMaster)
  {
    if (dwFlags & CN_CLONE1)
    {
      if (lpFrame->ei.hWndClone1)
      {
        if (lpFrame != &fdDefault)
          SendMessage(hMainWnd, AKDN_EDIT_ONFINISH, (WPARAM)lpFrame->ei.hWndClone1, (LPARAM)lpFrame->ei.hDocClone1);

        if (nMDI == WMD_MDI || lpFrame == &fdDefault)
        {
          DestroyWindow(lpFrame->ei.hWndClone1);
        }
        else if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
        {
          if (fdDefault.ei.hDocClone1 != lpFrame->ei.hDocClone1)
          {
            //WMD_PMDI
            SendMessage(lpFrame->ei.hWndClone1, AEM_SETDOCUMENT, (WPARAM)fdDefault.ei.hDocClone1, AESWD_NOREDRAW);
            SendMessage(lpFrame->ei.hWndClone1, AEM_DELETEDOCUMENT, (WPARAM)lpFrame->ei.hDocClone1, 0);
          }
          SendMessage(lpFrame->ei.hWndMaster, AEM_DELCLONE, (WPARAM)lpFrame->ei.hWndClone1, 0);
        }

        if (lpFrame->ei.hWndEdit && lpFrame->ei.hWndClone1 == lpFrame->ei.hWndEdit)
        {
          lpFrame->ei.hWndEdit=lpFrame->ei.hWndMaster;
          lpFrame->ei.hDocEdit=lpFrame->ei.hDocMaster;
        }
        lpFrame->ei.hWndClone1=NULL;
        lpFrame->ei.hDocClone1=NULL;
      }
    }
    if (dwFlags & CN_CLONE2)
    {
      if (lpFrame->ei.hWndClone2)
      {
        if (lpFrame != &fdDefault)
          SendMessage(hMainWnd, AKDN_EDIT_ONFINISH, (WPARAM)lpFrame->ei.hWndClone2, (LPARAM)lpFrame->ei.hDocClone2);

        if (nMDI == WMD_MDI || lpFrame == &fdDefault)
        {
          DestroyWindow(lpFrame->ei.hWndClone2);
        }
        else if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
        {
          if (fdDefault.ei.hDocClone2 != lpFrame->ei.hDocClone2)
          {
            //WMD_PMDI
            SendMessage(lpFrame->ei.hWndClone2, AEM_SETDOCUMENT, (WPARAM)fdDefault.ei.hDocClone2, AESWD_NOREDRAW);
            SendMessage(lpFrame->ei.hWndClone2, AEM_DELETEDOCUMENT, (WPARAM)lpFrame->ei.hDocClone2, 0);
          }
          SendMessage(lpFrame->ei.hWndMaster, AEM_DELCLONE, (WPARAM)lpFrame->ei.hWndClone2, 0);
        }

        if (lpFrame->ei.hWndEdit && lpFrame->ei.hWndClone2 == lpFrame->ei.hWndEdit)
        {
          lpFrame->ei.hWndEdit=lpFrame->ei.hWndMaster;
          lpFrame->ei.hDocEdit=lpFrame->ei.hDocMaster;
        }
        lpFrame->ei.hWndClone2=NULL;
        lpFrame->ei.hDocClone2=NULL;
      }
    }
    if (dwFlags & CN_CLONE3)
    {
      if (lpFrame->ei.hWndClone3)
      {
        if (lpFrame != &fdDefault)
          SendMessage(hMainWnd, AKDN_EDIT_ONFINISH, (WPARAM)lpFrame->ei.hWndClone3, (LPARAM)lpFrame->ei.hDocClone3);

        if (nMDI == WMD_MDI || lpFrame == &fdDefault)
        {
          DestroyWindow(lpFrame->ei.hWndClone3);
        }
        else if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
        {
          if (fdDefault.ei.hDocClone3 != lpFrame->ei.hDocClone3)
          {
            //WMD_PMDI
            SendMessage(lpFrame->ei.hWndClone3, AEM_SETDOCUMENT, (WPARAM)fdDefault.ei.hDocClone3, AESWD_NOREDRAW);
            SendMessage(lpFrame->ei.hWndClone3, AEM_DELETEDOCUMENT, (WPARAM)lpFrame->ei.hDocClone3, 0);
          }
          SendMessage(lpFrame->ei.hWndMaster, AEM_DELCLONE, (WPARAM)lpFrame->ei.hWndClone3, 0);
        }

        if (lpFrame->ei.hWndEdit && lpFrame->ei.hWndClone3 == lpFrame->ei.hWndEdit)
        {
          lpFrame->ei.hWndEdit=lpFrame->ei.hWndMaster;
          lpFrame->ei.hDocEdit=lpFrame->ei.hDocMaster;
        }
        lpFrame->ei.hWndClone3=NULL;
        lpFrame->ei.hDocClone3=NULL;
      }
    }
    if (!lpFrame->ei.hWndClone1 && !lpFrame->ei.hWndClone2 && !lpFrame->ei.hWndClone3)
    {
      lpFrameCurrent->ei.hWndMaster=NULL;
      lpFrameCurrent->ei.hDocMaster=NULL;
    }
    if (nMainOnFinish != MOF_DESTROY)
    {
      SplitVisUpdate(lpFrame);
      SetSelectionStatus(lpFrame->ei.hDocEdit, lpFrame->ei.hWndEdit, NULL, NULL);
    }
  }
  bEditOnFinish=FALSE;
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void SplitVisUpdate(FRAMEDATA *lpFrame)
{
  if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
  {
    if (fdDefault.ei.hWndMaster)
    {
      if (fdDefault.ei.hWndClone1)
      {
        if (IsWindowVisible(fdDefault.ei.hWndClone1) == !lpFrame->ei.hDocClone1)
          ShowWindow(fdDefault.ei.hWndClone1, lpFrame->ei.hDocClone1?SW_SHOW:SW_HIDE);
      }
      if (fdDefault.ei.hWndClone2)
      {
        if (IsWindowVisible(fdDefault.ei.hWndClone2) == !lpFrame->ei.hDocClone2)
          ShowWindow(fdDefault.ei.hWndClone2, lpFrame->ei.hDocClone2?SW_SHOW:SW_HIDE);
      }
      if (fdDefault.ei.hWndClone3)
      {
        if (IsWindowVisible(fdDefault.ei.hWndClone3) == !lpFrame->ei.hDocClone3)
          ShowWindow(fdDefault.ei.hWndClone3, lpFrame->ei.hDocClone3?SW_SHOW:SW_HIDE);
      }
    }
  }
}


//// Menu functions

BOOL DoFileNew()
{
  if (nMDI)
    return CreateFrameWindow(NULL);
  else
    return CloseDocument(0);
}

BOOL CloseDocument(DWORD dwPrompt)
{
  if (!SaveChanged(dwPrompt)) return FALSE;

  RecentFilesSaveFile(lpFrameCurrent);
  SendMessage(hMainWnd, AKDN_EDIT_ONCLOSE, (WPARAM)lpFrameCurrent->ei.hWndEdit, (LPARAM)lpFrameCurrent->ei.hDocEdit);

  SetWindowTextWide(lpFrameCurrent->ei.hWndEdit, L"");
  lpFrameCurrent->szFile[0]='\0';
  lpFrameCurrent->wszFile[0]=L'\0';
  lpFrameCurrent->nFileLen=0;
  lpFrameCurrent->nStreamOffset=0;
  SetNewLineStatus(lpFrameCurrent, moCur.nNewFileNewLine, AENL_INPUT);
  SetModifyStatus(lpFrameCurrent, FALSE);
  SetCodePageStatus(lpFrameCurrent, moCur.nNewFileCodePage, moCur.bNewFileBOM);
  UpdateTitle(lpFrameCurrent);

  return TRUE;
}

HWND DoFileNewWindow(DWORD dwAddFlags)
{
  STARTUPINFOW si;
  PROCESS_INFORMATION pi;
  HWND hWndFriend=0;

  if (!GetModuleFileNameWide(hInstance, wbuf, MAX_PATH))
    return 0;

  xmemset(&si, 0, sizeof(STARTUPINFOW));
  si.cb=sizeof(STARTUPINFOW);
  si.dwFlags=STARTF_USESHOWWINDOW|dwAddFlags;
  si.wShowWindow=SW_SHOWNORMAL;

  if (CreateProcessWide(wbuf, NULL, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi))
  {
    WaitForInputIdle(pi.hProcess, INFINITE);
    EnumThreadWindows(pi.dwThreadId, EnumThreadWindowsProc, (LPARAM)&hWndFriend);
    CloseHandle(pi.hProcess);
    CloseHandle(pi.hThread);
    return hWndFriend;
  }
  return 0;
}

BOOL CALLBACK EnumThreadWindowsProc(HWND hWnd, LPARAM lParam)
{
  HWND *hWndFriend=(HWND *)lParam;

  GetClassNameWide(hWnd, wbuf, BUFFER_SIZE);
  if (!xstrcmpW(wbuf, APP_MAIN_CLASSW))
  {
    *hWndFriend=hWnd;
    return FALSE;
  }
  return TRUE;
}

BOOL DoFileOpen()
{
  OPENFILENAME_2000W ofn;
  wchar_t wszOpenFile[MAX_PATH];
  wchar_t wszOpenDir[MAX_PATH];
  const wchar_t *wpFileName;
  DIALOGCODEPAGE dc;
  int nFileLen;
  int nOpen;
  BOOL bShowPlacesBarInit;
  BOOL bSingleFile=TRUE;
  BOOL bResult=FALSE;

  if (nMDI == WMD_SDI && !SaveChanged(0)) return FALSE;

  if (wszOfnFileList=AllocWideStr(OPENFILELIST_SIZE))
  {
    FileDialog:
    bSaveDlg=FALSE;
    dc.nCodePage=-1;
    dc.bBOM=-1;
    bShowPlacesBarInit=moCur.bShowPlacesBar;

    //Initial directory
    nFileLen=(int)xstrcpynW(wszOpenFile, lpFrameCurrent->wszFile, MAX_PATH);
    if (*wszOpenFile)
    {
      GetFileDir(wszOpenFile, nFileLen, wszOpenDir, MAX_PATH);
      if (!DirExistsWide(wszOpenDir))
        xstrcpynW(wszOpenDir, moCur.wszLastDir, MAX_PATH);
    }
    else xstrcpynW(wszOpenDir, moCur.wszLastDir, MAX_PATH);

    //Initial file
    wpFileName=GetFileName(wszOpenFile, nFileLen);
    xstrcpynW(wszOfnFileList, wpFileName, MAX_PATH);
    if (lpFrameCurrent->nStreamOffset)
    {
      nOfnStreamOffset=lpFrameCurrent->nStreamOffset - (int)(wpFileName - wszOpenFile);
      wszOfnFileList[nOfnStreamOffset]=L'\0';
    }
    else nOfnStreamOffset=0;

    //Show dialog
    ofnStruct=&ofn;
    xmemset(&ofn, 0, sizeof(OPENFILENAME_2000W));
    wszOfnStream[0]=L'\0';
    ofn.lStructSize    =(moCur.bShowPlacesBar && !bOldWindows && !bWindowsNT4)?sizeof(OPENFILENAME_2000W):sizeof(OPENFILENAMEW);
    ofn.lCustData      =(LPARAM)&dc;
    ofn.hwndOwner      =hMainWnd;
    ofn.hInstance      =hLangLib;
    ofn.lpstrFile      =wszOfnFileList;
    ofn.lpstrFilter    =wszFileFilter;
    ofn.nFilterIndex   =2;
    ofn.nMaxFile       =OPENFILELIST_SIZE;
    ofn.lpstrInitialDir=wszOpenDir;
    ofn.lpstrDefExt    =NULL;
    ofn.Flags          =(nMDI?OFN_ALLOWMULTISELECT:0)|OFN_HIDEREADONLY|OFN_PATHMUSTEXIST|OFN_EXPLORER|OFN_ENABLEHOOK|OFN_ENABLETEMPLATE|OFN_ENABLESIZING|OFN_OVERWRITEPROMPT;
    ofn.lpfnHook       =(LPOFNHOOKPROC)FileDlgProc;
    ofn.lpTemplateName =MAKEINTRESOURCEW(IDD_OFN);

    bResult=GetOpenFileNameWide((OPENFILENAMEW *)&ofn);

    if (bResult)
    {
      //GetCurrentDirectoryWide(MAX_PATH, moCur.wszLastDir);
      SetCurrentDirectoryWide(wszExeDir);

      if (nMDI)
      {
        wchar_t wszFile[MAX_PATH];
        wchar_t wszString[MAX_PATH];
        wchar_t *wpFile=wszOfnFileList + xstrlenW(wszOfnFileList) + 1;
        MSG msg;
        int nFiles=0;
        int nFileCount=0;

        if (*wpFile)
        {
          //Multiple files selected
          bSingleFile=FALSE;
          if (*(wpFile - 2) == L'\\') *(wpFile - 2)=L'\0';
          xstrcpynW(moCur.wszLastDir, wszOfnFileList, MAX_PATH);

          //Get files count
          if (moCur.bStatusBar)
          {
            API_LoadStringW(hLangLib, STR_COUNT, wszString, MAX_PATH);
            xarraysizeW(wszOfnFileList, &nFiles);

            //First element in array is directory.
            --nFiles;
          }

          do
          {
            if (IsPathFull(wpFile))
              nFileLen=(int)xstrcpynW(wszFile, wpFile, MAX_PATH);  //.lnk target
            else
            {
              nFileLen=(int)xstrlenW(wpFile);
              xprintfW(wszFile, L"%s\\%s", wszOfnFileList, wpFile);
            }
            nOpen=OpenDocument(NULL, wszFile, dwOfnFlags|(*(wpFile + nFileLen + 1)?OD_MULTIFILE:0), nOfnCodePage, bOfnBOM);
            if (nOpen != EOD_SUCCESS && nOpen != EOD_MSGNO && nOpen != EOD_WINDOW_EXIST)
            {
              bResult=FALSE;
              break;
            }

            //Status update
            if (moCur.bStatusBar)
            {
              xprintfW(wbuf, wszString, ++nFileCount, nFiles);
              StatusBar_SetTextWide(hStatus, SBP_MODIFY, wbuf);

              while (PeekMessageWide(&msg, hStatus, 0, 0, PM_REMOVE))
              {
                TranslateMessage(&msg);
                DispatchMessageWide(&msg);
              }
            }

            //Win7: prevent system from mark program as hanged
            PeekMessageWide(&msg, hMainWnd, 0, 0, PM_NOREMOVE);
          }
          while (*(wpFile+=nFileLen + 1));

          if (moCur.bStatusBar)
            StatusBar_SetTextWide(hStatus, SBP_MODIFY, L"");
        }
      }

      if (bSingleFile)
      {
        GetFileDir(wszOfnFileList, -1, moCur.wszLastDir, MAX_PATH);

        if (wszOfnStream[0])
        {
          nFileLen=(int)xstrlenW(wszOfnFileList);
          if (wszOfnFileList[nFileLen - 1] == L'.')
            wszOfnFileList[--nFileLen]=L'\0';
          xprintfW(wszOfnFileList + nFileLen, L":%.%ds", sizeof(wszOfnFileList) / sizeof(wchar_t) - nFileLen - 1,  wszOfnStream);
        }
        nOpen=OpenDocument(NULL, wszOfnFileList, dwOfnFlags, nOfnCodePage, bOfnBOM);
        if (nOpen != EOD_SUCCESS && nOpen != EOD_MSGNO && nOpen != EOD_WINDOW_EXIST)
          bResult=FALSE;
      }
    }
    else if (moCur.bShowPlacesBar != bShowPlacesBarInit)
    {
      //Restart dialog
      goto FileDialog;
    }
    FreeWideStr(wszOfnFileList);
  }
  return bResult;
}

int DoFileReopenAs(DWORD dwFlags, int nCodePage, BOOL bBOM)
{
  int nResult=EOD_SUCCESS;
  int nAnswer=0;

  if (!lpFrameCurrent->wszFile[0])
  {
    if (dwFlags & OD_ADT_DETECT_CODEPAGE)
      nCodePage=lpFrameCurrent->ei.nCodePage;
    if (dwFlags & OD_ADT_DETECT_BOM)
      bBOM=IsCodePageUnicode(nCodePage);

    SetCodePageStatus(lpFrameCurrent, nCodePage, bBOM);
    return nResult;
  }

  if (lpFrameCurrent->ei.bModified)
  {
    API_LoadStringW(hLangLib, MSG_FILE_WILL_BE_REOPENED, wszMsg, BUFFER_SIZE);
    nAnswer=API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OKCANCEL|MB_ICONEXCLAMATION);
  }
  if (!lpFrameCurrent->ei.bModified || nAnswer == IDOK)
  {
    nResult=OpenDocument(NULL, lpFrameCurrent->wszFile, dwFlags|OD_REOPEN, nCodePage, bBOM);
  }
  return nResult;
}

BOOL DoFileSave()
{
  if (!lpFrameCurrent->ei.bModified && lpFrameCurrent->wszFile[0] && FileExistsWide(lpFrameCurrent->wszFile)) return TRUE;
  if (!lpFrameCurrent->wszFile[0]) return DoFileSaveAs(-1, -1);
  return !SaveDocument(NULL, lpFrameCurrent->wszFile, lpFrameCurrent->ei.nCodePage, lpFrameCurrent->ei.bBOM, SD_UPDATE);
}

BOOL SaveChanged(DWORD dwPrompt)
{
  if (!(dwPrompt & PROMPT_NONE) && lpFrameCurrent->ei.bModified)
  {
    BUTTONMESSAGEBOX bmb[]={{IDC_MESSAGEBOX_YES,     MAKEINTRESOURCEW(STR_MESSAGEBOX_YES),     BMB_DEFAULT},
                            {IDC_MESSAGEBOX_NO,      MAKEINTRESOURCEW(STR_MESSAGEBOX_NO),      0},
                            {IDC_MESSAGEBOX_NOTOALL, MAKEINTRESOURCEW(STR_MESSAGEBOX_NOTOALL), 0},
                            {IDCANCEL,               MAKEINTRESOURCEW(STR_MESSAGEBOX_CANCEL),  0},
                            {0, 0, 0}};
    int nChoice;

    if (nDocumentsModified <= 1)
    {
      if (dwPrompt & PROMPT_NOTOALLBUTTON)
        bmb[2].dwFlags|=BMB_DISABLED;
    }
    API_LoadStringW(hLangLib, MSG_DOCUMENT_CHANGED, wszMsg, MAX_PATH);
    if (dwPrompt & PROMPT_NOTOALLBUTTON)
      nChoice=MessageBoxCustom(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_ICONEXCLAMATION, NULL, &bmb[0]);
    else
      nChoice=API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_YESNOCANCEL|MB_ICONEXCLAMATION);

    if (nChoice == IDC_MESSAGEBOX_YES || nChoice == IDYES)
    {
      if (!DoFileSave()) return FALSE;
    }
    else if (nChoice == IDC_MESSAGEBOX_NOTOALL)
    {
      dwChangedPrompt|=PROMPT_NONE;
    }
    else if (nChoice == IDCANCEL)
    {
      return FALSE;
    }
  }
  return TRUE;
}

BOOL DoFileSaveAs(int nDialogCodePage, BOOL bDialogBOM)
{
  OPENFILENAME_2000W ofn;
  wchar_t wszSaveFile[MAX_PATH];
  wchar_t wszSaveDir[MAX_PATH];
  const wchar_t *wpFileName;
  DIALOGCODEPAGE dc;
  int nFileLen;
  BOOL bShowPlacesBarInit;
  BOOL bResult=FALSE;

  if (wszOfnFileList=AllocWideStr(MAX_PATH))
  {
    FileDialog:
    bSaveDlg=TRUE;
    dc.nCodePage=nDialogCodePage;
    dc.bBOM=bDialogBOM;
    bShowPlacesBarInit=moCur.bShowPlacesBar;

    //Initial directory
    nFileLen=(int)xstrcpynW(wszSaveFile, lpFrameCurrent->wszFile, MAX_PATH);
    if (*wszSaveFile)
    {
      GetFileDir(wszSaveFile, nFileLen, wszSaveDir, MAX_PATH);
      if (!DirExistsWide(wszSaveDir))
        xstrcpynW(wszSaveDir, moCur.wszLastDir, MAX_PATH);
    }
    else xstrcpynW(wszSaveDir, moCur.wszLastDir, MAX_PATH);

    //Initial file
    wpFileName=GetFileName(wszSaveFile, nFileLen);
    xstrcpynW(wszOfnFileList, wpFileName, MAX_PATH);
    if (lpFrameCurrent->nStreamOffset)
    {
      nOfnStreamOffset=lpFrameCurrent->nStreamOffset - (int)(wpFileName - wszSaveFile);
      wszOfnFileList[nOfnStreamOffset]=L'\0';
    }
    else
    {
      if (*wszSaveFile && !GetFileExt(wszSaveFile, nFileLen))
      {
        //Append dot to filename for file without extention
        //to avoid appending default extention during "Save as".
        xprintfW(wszOfnFileList, L"%s.", wpFileName);
      }
      nOfnStreamOffset=0;
    }

    ofnStruct=&ofn;
    xmemset(&ofn, 0, sizeof(OPENFILENAME_2000W));
    wszOfnStream[0]=L'\0';
    ofn.lStructSize    =(moCur.bShowPlacesBar && !bOldWindows && !bWindowsNT4)?sizeof(OPENFILENAME_2000W):sizeof(OPENFILENAMEW);
    ofn.lCustData      =(LPARAM)&dc;
    ofn.hwndOwner      =hMainWnd;
    ofn.hInstance      =hLangLib;
    ofn.lpstrFile      =wszOfnFileList;
    ofn.lpstrFilter    =wszFileFilter;
    ofn.nFilterIndex   =2;
    ofn.nMaxFile       =MAX_PATH;
    ofn.lpstrInitialDir=wszSaveDir;
    ofn.lpstrDefExt    =moCur.wszDefaultSaveExt;
    ofn.Flags          =OFN_HIDEREADONLY|OFN_PATHMUSTEXIST|OFN_EXPLORER|OFN_ENABLEHOOK|OFN_ENABLETEMPLATE|OFN_ENABLESIZING|OFN_OVERWRITEPROMPT|OFN_NOTESTFILECREATE;
    ofn.lpfnHook       =(LPOFNHOOKPROC)FileDlgProc;
    ofn.lpTemplateName =MAKEINTRESOURCEW(IDD_OFN);

    bResult=GetSaveFileNameWide((OPENFILENAMEW *)&ofn);

    if (bResult)
    {
      GetCurrentDirectoryWide(MAX_PATH, moCur.wszLastDir);
      SetCurrentDirectoryWide(wszExeDir);

      if (wszOfnStream[0])
      {
        nFileLen=(int)xstrlenW(wszOfnFileList);
        if (wszOfnFileList[nFileLen - 1] == L'.')
          wszOfnFileList[--nFileLen]=L'\0';
        xprintfW(wszOfnFileList + nFileLen, L":%.%ds", sizeof(wszOfnFileList) / sizeof(wchar_t) - nFileLen - 1,  wszOfnStream);
      }
      if (SaveDocument(NULL, wszOfnFileList, nOfnCodePage, bOfnBOM, SD_UPDATE) != ESD_SUCCESS)
        bResult=FALSE;
    }
    else if (moCur.bShowPlacesBar != bShowPlacesBarInit)
    {
      //Restart dialog
      goto FileDialog;
    }
    FreeWideStr(wszOfnFileList);
  }
  return bResult;
}

BOOL DoFileSaveAllAs()
{
  return (BOOL)API_DialogBox(hLangLib, MAKEINTRESOURCEW(IDD_SAVEALLAS), hMainWnd, (DLGPROC)SaveAllAsDlgProc);
}

BOOL DoFilePageSetup(HWND hWndOwner)
{
  BOOL bResult;

  if (bOldWindows)
  {
    PAGESETUPDLGA psdPageA;

    xmemset(&psdPageA, 0, sizeof(PAGESETUPDLGA));
    psdPageA.lStructSize        =sizeof(PAGESETUPDLGA);
    psdPageA.hwndOwner          =hWndOwner;
    psdPageA.lpfnPageSetupHook  =(LPPAGESETUPHOOK)PrintPageSetupDlgProc;
    psdPageA.Flags              =prninfo.dwPageFlags|PSD_MARGINS|PSD_ENABLEPAGESETUPHOOK;
    psdPageA.rtMargin           =prninfo.rtMargin;
    psdPageA.hDevMode           =prninfo.hDevMode;
    psdPageA.hDevNames          =prninfo.hDevNames;

    if (bResult=PageSetupDlgA(&psdPageA))
    {
      prninfo.rtMargin=psdPageA.rtMargin;
      moCur.rcPrintMargins=prninfo.rtMargin;
    }
    prninfo.hDevMode=psdPageA.hDevMode;
    prninfo.hDevNames=psdPageA.hDevNames;
  }
  else
  {
    PAGESETUPDLGW psdPageW;

    xmemset(&psdPageW, 0, sizeof(PAGESETUPDLGW));
    psdPageW.lStructSize        =sizeof(PAGESETUPDLGW);
    psdPageW.hwndOwner          =hWndOwner;
    psdPageW.lpfnPageSetupHook  =(LPPAGESETUPHOOK)PrintPageSetupDlgProc;
    psdPageW.Flags              =prninfo.dwPageFlags|PSD_MARGINS|PSD_ENABLEPAGESETUPHOOK;
    psdPageW.rtMargin           =prninfo.rtMargin;
    psdPageW.hDevMode           =prninfo.hDevMode;
    psdPageW.hDevNames          =prninfo.hDevNames;

    if (bResult=PageSetupDlgW(&psdPageW))
    {
      prninfo.rtMargin=psdPageW.rtMargin;
      moCur.rcPrintMargins=prninfo.rtMargin;
    }
    prninfo.hDevMode=psdPageW.hDevMode;
    prninfo.hDevNames=psdPageW.hDevNames;
  }
  UpdateMappedPrintWidth(lpFrameCurrent);
  UpdateStatusUser(lpFrameCurrent, CSB_MARKER);
  return bResult;
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
int DoFilePrint(FRAMEDATA *lpFrame, BOOL bSilent)
{
  int nResult=0;

  //Set print dialog settings
  if (!AEC_IndexCompare(&crCurSel.ciMin, &crCurSel.ciMax))
    prninfo.dwPrintFlags|=PD_NOSELECTION;
  else
    prninfo.dwPrintFlags&=~PD_NOSELECTION;

  //Choose printer
  if (!GetPrinter(hMainWnd, &prninfo, bSilent))
    return 0;

  if (prninfo.hDC)
  {
    prn.hPrinterDC=prninfo.hDC;
    prn.rcMargins=prninfo.rtMargin;

    nResult=PrintDocument(lpFrame->ei.hWndEdit, &prn, (bOldWindows?PRND_ANSI:0)|PRND_REALPRINT|(prninfo.dwPrintFlags & PD_SELECTION?PRND_SELECTION:PRND_ALLTEXT), 0);

    DeleteDC(prninfo.hDC);
    prninfo.hDC=NULL;
    prninfo.dwPrintFlags&=~PD_SELECTION;
    prninfo.dwPrintFlags&=~PD_PAGENUMS;
    prninfo.dwPrintFlags&=~PD_PRINTTOFILE;
  }
  UpdateMappedPrintWidth(lpFrame);
  UpdateStatusUser(lpFrame, CSB_MARKER);
  return nResult;
}

void DoFilePreview(HWND hWnd)
{
  hWndPreviewEdit=hWnd;
  API_DialogBox(hLangLib, MAKEINTRESOURCEW(IDD_PRINTPREVIEW), hMainWnd, (DLGPROC)PreviewDlgProc);
  hWndPreviewEdit=NULL;
}

void DoEditUndo(HWND hWnd)
{
  SendMessage(hWnd, AEM_UNDO, 0, 0);
}

void DoEditRedo(HWND hWnd)
{
  SendMessage(hWnd, AEM_REDO, 0, 0);
}

void DoEditCut(HWND hWnd)
{
  if (IsReadOnly(hWnd)) return;

  SendMessage(hWnd, AEM_CUT, 0, 0);
}

void DoEditCopy(HWND hWnd)
{
  SendMessage(hWnd, AEM_COPY, 0, 0);
}

BOOL DoEditPaste(HWND hWnd, DWORD dwFlags)
{
  CHARRANGE64 cr;

  if (IsReadOnly(hWnd)) return FALSE;

  if (dwFlags & PASTE_AFTER)
    SendMessage(hWnd, EM_EXGETSEL64, 0, (LPARAM)&cr);

  if (SendMessage(hWnd, AEM_PASTE, 0, (LPARAM)dwFlags))
  {
    if (dwFlags & PASTE_AFTER)
      SendMessage(hWnd, EM_SETSEL, (WPARAM)cr.cpMin, (LPARAM)cr.cpMin);
    return TRUE;
  }
  return FALSE;
}

void DoEditClear(HWND hWnd)
{
  if (IsReadOnly(hWnd)) return;

  ReplaceSelW(hWnd, L"", -1, AELB_ASINPUT, AEREPT_COLUMNASIS, NULL, NULL);
}

void DoEditSelectAll(HWND hWnd)
{
  AECHARRANGE cr;

  SendMessage(hWnd, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&cr.ciMin);
  SendMessage(hWnd, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&cr.ciMax);
  SetSel(hWnd, &cr, 0, &cr.ciMax);
}

void DoEditInsertChar()
{
  ShellExecuteA(hMainWnd, "open", "charmap.exe", NULL, NULL, SW_SHOWDEFAULT);
}

void DoEditInsertDate(HWND hWnd)
{
  if (IsReadOnly(hWnd)) return;

  GetTimeString(moCur.wszDateInsertFormat, wbuf, TRUE);
  ReplaceSelW(hWnd, wbuf, -1, AELB_ASINPUT, AEREPT_COLUMNASIS, NULL, NULL);
}

void DoEditRecode()
{
  if (!hDlgModeless)
  {
    nModelessType=MLT_RECODE;

    hDlgModeless=API_CreateDialog(hLangLib, MAKEINTRESOURCEW(IDD_RECODE), hMainWnd, (DLGPROC)RecodeDlgProc);

    if (hDlgModeless)
    {
      ShowWindow(hDlgModeless, SW_SHOW);
    }
  }
}

BOOL DoEditModifyStringInSelection(HWND hWnd, int nAction, const wchar_t *wpString)
{
  AECHARRANGE crRange;
  AETEXTRANGEW tr;
  wchar_t *wszRange=NULL;
  INT_PTR nRangeLen;
  INT_PTR nStringLenAll;
  INT_PTR nBufferLen;
  int nStringLen;
  int nStringBytes;
  int nFirstLine=0;
  INT_PTR a=0;
  INT_PTR b=0;
  INT_PTR i;
  BOOL bColumnSel;
  BOOL bCaretAtStart=FALSE;
  BOOL bResult=FALSE;

  if (!(nAction & STRSEL_CHECK) && IsReadOnly(hWnd))
    return FALSE;
  if (!AEC_IndexCompare(&crCurSel.ciMin, &ciCurCaret))
    bCaretAtStart=TRUE;
  crRange.ciMin=crCurSel.ciMin;
  crRange.ciMax=crCurSel.ciMax;

  if (!(nAction & STRSEL_MULTILINE) || crRange.ciMin.nLine != crRange.ciMax.nLine)
  {
    if (nAction & STRSEL_CHECK) return TRUE;

    if (!(bColumnSel=(BOOL)SendMessage(hWnd, AEM_GETCOLUMNSEL, 0, 0)) && (nAction & STRSEL_FULLLINE))
    {
      SendMessage(hWnd, AEM_GETINDEX, AEGI_WRAPLINEBEGIN, (LPARAM)&crRange.ciMin);
      if (!AEC_IsFirstCharInLine(&crRange.ciMax))
      {
        SendMessage(hWnd, AEM_GETINDEX, AEGI_WRAPLINEEND, (LPARAM)&crRange.ciMax);
        //SendMessage(hWnd, AEM_GETINDEX, AEGI_NEXTLINE, (LPARAM)&crRange.ciMax);
      }
    }

    if (nRangeLen=IndexSubtract(hWnd, &crRange.ciMax, &crRange.ciMin, AELB_ASIS, bColumnSel))
    {
      nStringLen=(int)xstrlenW(wpString);
      nStringBytes=nStringLen * sizeof(wchar_t);

      //Save scroll
      nFirstLine=SaveLineScroll(hWnd);

      if (!bColumnSel && (nAction & STRSEL_FULLLINE))
        SetSel(hWnd, &crRange, AESELT_LOCKSCROLL, NULL);

      if (nAction & STRSEL_INSERT)
      {
        nStringLenAll=(crRange.ciMax.nLine - crRange.ciMin.nLine + 1) * nStringLen;
        nBufferLen=nRangeLen + nStringLenAll;

        if (wszRange=AllocWideStr(nBufferLen + 1))
        {
          tr.cr.ciMin=crRange.ciMin;
          tr.cr.ciMax=crRange.ciMax;
          tr.bColumnSel=bColumnSel;
          tr.pBuffer=wszRange + nStringLenAll;
          tr.dwBufferMax=(UINT_PTR)-1;
          tr.nNewLine=AELB_ASIS;
          tr.bFillSpaces=TRUE;
          SendMessage(hWnd, AEM_GETTEXTRANGEW, 0, (LPARAM)&tr);
          b=nStringLenAll;

          goto InsertLineBeginning;

          while (b < nBufferLen)
          {
            if (wszRange[b] == L'\r' && wszRange[b + 1] == L'\r' && wszRange[b + 2] == L'\n')
            {
              wszRange[a++]=wszRange[b++];
              wszRange[a++]=wszRange[b++];
              wszRange[a++]=wszRange[b++];
            }
            else if (wszRange[b] == L'\r' && wszRange[b + 1] == L'\n')
            {
              wszRange[a++]=wszRange[b++];
              wszRange[a++]=wszRange[b++];
            }
            else if (wszRange[b] == L'\r')
            {
              wszRange[a++]=wszRange[b++];
            }
            else if (wszRange[b] == L'\n')
            {
              wszRange[a++]=wszRange[b++];
            }
            else
            {
              wszRange[a++]=wszRange[b++];
              continue;
            }

            InsertLineBeginning:
            if (b < nBufferLen || bColumnSel)
            {
              xmemcpy(wszRange + a, wpString, nStringBytes);
              a+=nStringLen;
            }
          }
          wszRange[a]=L'\0';
          bResult=TRUE;
        }
      }
      else if (nAction & STRSEL_DELETE)
      {
        if (wszRange=AllocWideStr(nRangeLen + 1))
        {
          tr.cr.ciMin=crRange.ciMin;
          tr.cr.ciMax=crRange.ciMax;
          tr.bColumnSel=bColumnSel;
          tr.pBuffer=wszRange;
          tr.dwBufferMax=(UINT_PTR)-1;
          tr.nNewLine=AELB_ASIS;
          tr.bFillSpaces=TRUE;
          SendMessage(hWnd, AEM_GETTEXTRANGEW, 0, (LPARAM)&tr);

          if (nAction & STRSEL_ALLSPACES)
          {
            while (b < nRangeLen)
            {
              if (wszRange[b] == L' ' || wszRange[b] == L'\t')
                ++b;
              else
                wszRange[a++]=wszRange[b++];
            }
          }
          else
          {
            goto DeleteLineBeginning;

            while (b < nRangeLen)
            {
              if (wszRange[b] == L'\r' && wszRange[b + 1] == L'\r' && wszRange[b + 2] == L'\n')
              {
                wszRange[a++]=wszRange[b++];
                wszRange[a++]=wszRange[b++];
                wszRange[a++]=wszRange[b++];
              }
              else if (wszRange[b] == L'\r' && wszRange[b + 1] == L'\n')
              {
                wszRange[a++]=wszRange[b++];
                wszRange[a++]=wszRange[b++];
              }
              else if (wszRange[b] == L'\r')
              {
                wszRange[a++]=wszRange[b++];
              }
              else if (wszRange[b] == L'\n')
              {
                wszRange[a++]=wszRange[b++];
              }
              else
              {
                wszRange[a++]=wszRange[b++];
                continue;
              }

              DeleteLineBeginning:
              if (b < nRangeLen)
              {
                if (nAction & STRSEL_LEADTAB)
                {
                  if (wszRange[b] == L'\t')
                    ++b;
                  else
                    for (i=0; i < lpFrameCurrent->nTabStopSize && wszRange[b] == L' '; ++i, ++b);
                }
                else if (nAction & STRSEL_LEADSPACE)
                {
                  if (wszRange[b] == L' ' || wszRange[b] == L'\t')
                    ++b;
                }
                else
                {
                  if (!xmemcmp(wszRange + b, wpString, nStringBytes))
                    b+=nStringLen;
                }
              }
            }
          }
          wszRange[a]=L'\0';
          bResult=TRUE;
        }
      }

      if (bResult)
      {
        ReplaceSelW(hWnd, wszRange, a, AELB_ASINPUT, (bColumnSel?AEREPT_COLUMNON:0)|AEREPT_LOCKSCROLL, &crRange.ciMin, &crRange.ciMax);
        SetSel(hWnd, &crRange, (bColumnSel?AESELT_COLUMNON:0)|AESELT_LOCKSCROLL, bCaretAtStart?&crRange.ciMin:&crRange.ciMax);
      }

      //Restore scroll
      RestoreLineScroll(hWnd, nFirstLine);

      FreeWideStr(wszRange);
      return bResult;
    }
  }
  if (nAction & STRSEL_CHECK) return FALSE;

  return TRUE;
}

BOOL DoEditDeleteFirstCharW(HWND hWnd)
{
  AECHARRANGE crRange;
  wchar_t *wszRange;
  INT_PTR nRangeLen;
  INT_PTR a;
  INT_PTR b;
  int nFirstLine=0;
  BOOL bDelete;
  BOOL bCaretAtStart=FALSE;

  if (IsReadOnly(hWnd))
    return FALSE;
  if (!AEC_IndexCompare(&crCurSel.ciMin, &ciCurCaret))
    bCaretAtStart=TRUE;
  crRange.ciMin=crCurSel.ciMin;
  crRange.ciMax=crCurSel.ciMax;

  if (nRangeLen=ExGetRangeTextW(hWnd, &crRange.ciMin, &crRange.ciMax, -1, &wszRange, AELB_ASIS, TRUE))
  {
    bDelete=TRUE;
    a=0, b=0;

    while (b < nRangeLen)
    {
      while (wszRange[b] == L'\r' || wszRange[b] == L'\n')
      {
        bDelete=TRUE;
        wszRange[a++]=wszRange[b++];
      }
      if (bDelete)
      {
        bDelete=FALSE;
        b++;
      }
      else wszRange[a++]=wszRange[b++];
    }
    wszRange[a]=L'\0';

    //Save scroll
    nFirstLine=SaveLineScroll(hWnd);

    ReplaceSelW(hWnd, wszRange, a, AELB_ASINPUT, AEREPT_COLUMNASIS|AEREPT_LOCKSCROLL, &crRange.ciMin, &crRange.ciMax);
    SetSel(hWnd, &crRange, AESELT_COLUMNASIS|AESELT_LOCKSCROLL, bCaretAtStart?&crRange.ciMin:&crRange.ciMax);

    //Restore scroll
    RestoreLineScroll(hWnd, nFirstLine);

    FreeText(wszRange);
    return TRUE;
  }
  SendMessage(hWnd, WM_KEYDOWN, VK_BACK, 0);
  SendMessage(hWnd, WM_KEYUP, VK_BACK, 0);
  return FALSE;
}

BOOL DoEditDeleteTrailingWhitespacesW(HWND hWnd)
{
  AECHARRANGE crRange;
  AECHARINDEX ciInitialCaret=ciCurCaret;
  wchar_t *wszRange;
  INT_PTR nRangeLen;
  INT_PTR nNewRangeLen;
  wchar_t *wpRead;
  wchar_t *wpWrite;
  wchar_t *wpMaxRead;
  int nFirstLine=0;
  BOOL bSelection;
  BOOL bCaretAtStart=FALSE;
  BOOL bResult=FALSE;

  if (IsReadOnly(hWnd))
    return FALSE;
  if (!AEC_IndexCompare(&crCurSel.ciMin, &ciCurCaret))
    bCaretAtStart=TRUE;

  //Save scroll
  nFirstLine=SaveLineScroll(hWnd);

  if (!AEC_IndexCompare(&crCurSel.ciMin, &crCurSel.ciMax))
  {
    SendMessage(hWnd, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&crRange.ciMin);
    SendMessage(hWnd, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&crRange.ciMax);
    bSelection=FALSE;
  }
  else
  {
    crRange.ciMin=crCurSel.ciMin;
    crRange.ciMax=crCurSel.ciMax;
    bSelection=TRUE;
  }

  if (nRangeLen=ExGetRangeTextW(hWnd, &crRange.ciMin, &crRange.ciMax, -1, &wszRange, AELB_ASIS, TRUE))
  {
    wpMaxRead=wszRange + nRangeLen;

    for (wpWrite=wszRange, wpRead=wszRange; wpRead < wpMaxRead; *wpWrite++=*wpRead++)
    {
      if (*wpRead == L'\r' || *wpRead == L'\n')
      {
        while (--wpWrite >= wszRange && (*wpWrite == L'\t' || *wpWrite == L' '));
        ++wpWrite;
      }
    }
    while (--wpWrite >= wszRange && (*wpWrite == L'\t' || *wpWrite == L' '));
    ++wpWrite;
    *wpWrite=L'\0';
    nNewRangeLen=wpWrite - wszRange;

    if (nRangeLen != nNewRangeLen)
    {
      if (!bSelection)
        SetSel(hWnd, &crRange, AESELT_LOCKSCROLL, &crRange.ciMax);

      ReplaceSelW(hWnd, wszRange, nNewRangeLen, AELB_ASINPUT, AEREPT_COLUMNASIS|AEREPT_LOCKSCROLL, &crRange.ciMin, &crRange.ciMax);

      //Update selection
      if (!bSelection)
      {
        SendMessage(hWnd, AEM_INDEXUPDATE, 0, (LPARAM)&ciInitialCaret);
        crRange.ciMin=ciInitialCaret;
        crRange.ciMax=ciInitialCaret;
      }
      SetSel(hWnd, &crRange, AESELT_COLUMNASIS|AESELT_LOCKSCROLL, bCaretAtStart?&crRange.ciMin:&crRange.ciMax);
      bResult=TRUE;
    }
    FreeText(wszRange);
  }

  //Restore scroll
  RestoreLineScroll(hWnd, nFirstLine);

  return bResult;
}

BOOL DoEditChangeCaseW(HWND hWnd, int nCase, BOOL bSelCurWord)
{
  AECHARRANGE crRange;
  AECHARINDEX ciInitialCaret=ciCurCaret;
  wchar_t *wszRange;
  int nFirstLine=0;
  INT_PTR nRangeLen;
  BOOL bSelection;
  BOOL bCaretAtStart=FALSE;
  BOOL bResult=FALSE;

  if (IsReadOnly(hWnd))
    return FALSE;
  if (!AEC_IndexCompare(&crCurSel.ciMin, &ciCurCaret))
    bCaretAtStart=TRUE;

  //Save scroll
  nFirstLine=SaveLineScroll(hWnd);

  if (!AEC_IndexCompare(&crCurSel.ciMin, &crCurSel.ciMax))
  {
    if (bSelCurWord)
    {
      crRange.ciMin=ciCurCaret;
      crRange.ciMax=ciCurCaret;

      if (!SendMessage(hWnd, AEM_ISDELIMITER, AEDLM_WORD|AEDLM_PREVCHAR, (LPARAM)&crRange.ciMin))
        SendMessage(hWnd, AEM_GETPREVBREAK, AEWB_LEFTWORDSTART, (LPARAM)&crRange.ciMin);
      if (!SendMessage(hWnd, AEM_ISDELIMITER, AEDLM_WORD, (LPARAM)&crRange.ciMax))
        SendMessage(hWnd, AEM_GETNEXTBREAK, AEWB_RIGHTWORDEND, (LPARAM)&crRange.ciMax);
    }
    else
    {
      SendMessage(hWnd, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&crRange.ciMin);
      SendMessage(hWnd, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&crRange.ciMax);
    }
    bSelection=FALSE;
  }
  else
  {
    crRange.ciMin=crCurSel.ciMin;
    crRange.ciMax=crCurSel.ciMax;
    bSelection=TRUE;
  }

  if (nRangeLen=ExGetRangeTextW(hWnd, &crRange.ciMin, &crRange.ciMax, -1, &wszRange, AELB_ASIS, TRUE))
  {
    ConvertCase(wszRange, nRangeLen, nCase);

    if (!bSelection)
      SetSel(hWnd, &crRange, AESELT_LOCKSCROLL, &crRange.ciMax);

    ReplaceSelW(hWnd, wszRange, -1, AELB_ASINPUT, AEREPT_COLUMNASIS|AEREPT_LOCKSCROLL, &crRange.ciMin, &crRange.ciMax);

    //Update selection
    if (!bSelection)
    {
      SendMessage(hWnd, AEM_INDEXUPDATE, 0, (LPARAM)&ciInitialCaret);
      crRange.ciMin=ciInitialCaret;
      crRange.ciMax=ciInitialCaret;
    }
    SetSel(hWnd, &crRange, AESELT_COLUMNASIS|AESELT_LOCKSCROLL, bCaretAtStart?&crRange.ciMin:&crRange.ciMax);

    FreeText(wszRange);
    bResult=TRUE;
  }

  //Restore scroll
  RestoreLineScroll(hWnd, nFirstLine);

  return bResult;
}

void ConvertCase(wchar_t *wszText, INT_PTR nTextLen, int nCase)
{
  wchar_t *wpText=wszText;
  const wchar_t *wpTextMax;

  if (nTextLen == -1)
    nTextLen=xstrlenW(wpText);
  wpTextMax=wpText + nTextLen;

  if (nCase == SCT_UPPERCASE)
  {
    for (; wpText < wpTextMax; ++wpText)
      *wpText=WideCharUpper(*wpText);
  }
  else if (nCase == SCT_LOWERCASE)
  {
    for (; wpText < wpTextMax; ++wpText)
      *wpText=WideCharLower(*wpText);
  }
  else if (nCase == SCT_SENTENCECASE)
  {
    while (wpText < wpTextMax)
    {
      while (wpText < wpTextMax && (AKD_wcschr(lpFrameCurrent->wszWordDelimiters, *wpText) || AKD_wcschr(STR_SENTENCE_DELIMITERSW, *wpText)))
      {
        ++wpText;
      }
      if (wpText < wpTextMax)
      {
        *wpText=WideCharUpper(*wpText);
        ++wpText;
      }
      while (wpText < wpTextMax && !AKD_wcschr(STR_SENTENCE_DELIMITERSW, *wpText))
      {
        *wpText=WideCharLower(*wpText);
        ++wpText;
      }
    }
  }
  else if (nCase == SCT_TITLECASE)
  {
    while (wpText < wpTextMax)
    {
      while (wpText < wpTextMax && AKD_wcschr(lpFrameCurrent->wszWordDelimiters, *wpText))
      {
        ++wpText;
      }
      if (wpText < wpTextMax)
      {
        *wpText=WideCharUpper(*wpText);
        ++wpText;
      }
      while (wpText < wpTextMax && !AKD_wcschr(lpFrameCurrent->wszWordDelimiters, *wpText))
      {
        *wpText=WideCharLower(*wpText);
        ++wpText;
      }
    }
  }
  else if (nCase == SCT_INVERTCASE)
  {
    for (; wpText < wpTextMax; ++wpText)
    {
      if (*wpText == WideCharLower(*wpText))
        *wpText=WideCharUpper(*wpText);
      else
        *wpText=WideCharLower(*wpText);
    }
  }
}

int DetectSelCase(HWND hWnd)
{
  AECHARRANGE crSel;
  DWORD dwCaseType=DC_UPPERCASE|DC_LOWERCASE|DC_SENTENCECASE|DC_TITLECASE;
  int nChar;
  wchar_t wchChar;
  BOOL bColumnSel;
  BOOL bStartSentence=TRUE;
  BOOL bStartTitle=TRUE;

  if (!GetSel(hWnd, &crSel, &bColumnSel, NULL) || bColumnSel)
    return SCT_NONE;
  AEC_ValidCharInLine(&crSel.ciMin);
  AEC_ValidCharInLine(&crSel.ciMax);

  for (; dwCaseType && AEC_IndexCompare(&crSel.ciMin, &crSel.ciMax) < 0; AEC_NextChar(&crSel.ciMin))
  {
    nChar=AEC_CharAtIndex(&crSel.ciMin);
    if (nChar < 0) nChar=L'\n';
    wchChar=(wchar_t)nChar;

    if (dwCaseType & DC_UPPERCASE)
    {
      if (wchChar != WideCharUpper(wchChar))
        dwCaseType&=~DC_UPPERCASE;
    }
    if (dwCaseType & DC_LOWERCASE)
    {
      if (wchChar != WideCharLower(wchChar))
        dwCaseType&=~DC_LOWERCASE;
    }
    if (dwCaseType & DC_SENTENCECASE)
    {
      if (AKD_wcschr(STR_SENTENCE_DELIMITERSW, wchChar))
      {
        bStartSentence=TRUE;
      }
      else if (bStartSentence)
      {
        if (!AKD_wcschr(lpFrameCurrent->wszWordDelimiters, wchChar))
        {
          if (wchChar != WideCharUpper(wchChar))
            dwCaseType&=~DC_SENTENCECASE;
          else
            bStartSentence=FALSE;
        }
      }
      else if (wchChar != WideCharLower(wchChar))
        dwCaseType&=~DC_SENTENCECASE;
    }
    if (dwCaseType & DC_TITLECASE)
    {
      if (AKD_wcschr(lpFrameCurrent->wszWordDelimiters, wchChar))
      {
        bStartTitle=TRUE;
      }
      else if (bStartTitle)
      {
        if (wchChar != WideCharUpper(wchChar))
          dwCaseType&=~DC_TITLECASE;
        else
          bStartTitle=FALSE;
      }
      else if (wchChar != WideCharLower(wchChar))
        dwCaseType&=~DC_TITLECASE;
    }
  }

  if (dwCaseType & DC_UPPERCASE)
    return SCT_UPPERCASE;
  if (dwCaseType & DC_LOWERCASE)
    return SCT_LOWERCASE;
  if (dwCaseType & DC_SENTENCECASE)
    return SCT_SENTENCECASE;
  if (dwCaseType & DC_TITLECASE)
    return SCT_TITLECASE;
  return SCT_NONE;
}

void DoEditFind()
{
  if (!hDlgModeless)
  {
    nModelessType=MLT_FIND;

    if (!nMDI)
      moCur.dwSearchOptions&=~FRF_ALLFILES;
    if (!AEC_IndexCompare(&crCurSel.ciMin, &crCurSel.ciMax))
      moCur.dwSearchOptions&=~FRF_SELECTION;

    hDlgModeless=API_CreateDialog(hLangLib, MAKEINTRESOURCEW(IDD_FIND), hMainWnd, (DLGPROC)FindAndReplaceDlgProc);

    if (hDlgModeless)
      ShowWindow(hDlgModeless, SW_SHOW);
  }
  else if (nModelessType == MLT_FIND)
  {
    if (!SendMessage(hDlgModeless, AKDLG_PUTFIND, 0, 0))
      BringWindowToTop(hDlgModeless);
    else
      SetFocus(hDlgModeless);
  }
}

INT_PTR DoEditFindNextDown(FRAMEDATA *lpFrame)
{
  DWORD dwFlags=(moCur.dwSearchOptions & ~FRF_UP & ~FRF_SELECTION & ~FRF_ALLFILES & ~FRF_BEGINNING) | FRF_DOWN;

  if (wszFindText)
    return TextFindW(lpFrame, dwFlags, wszFindText, nFindTextLen);
  DoEditFind();
  return 0;
}

INT_PTR DoEditFindNextUp(FRAMEDATA *lpFrame)
{
  DWORD dwFlags=(moCur.dwSearchOptions & ~FRF_DOWN & ~FRF_SELECTION & ~FRF_ALLFILES & ~FRF_BEGINNING) | FRF_UP;

  if (wszFindText)
    return TextFindW(lpFrame, dwFlags, wszFindText, nFindTextLen);
  DoEditFind();
  return 0;
}

void DoEditReplace()
{
  if (!hDlgModeless)
  {
    nModelessType=MLT_REPLACE;

    if (!nMDI)
      moCur.dwSearchOptions&=~FRF_ALLFILES;
    if (!AEC_IndexCompare(&crCurSel.ciMin, &crCurSel.ciMax))
      moCur.dwSearchOptions&=~FRF_SELECTION;

    hDlgModeless=API_CreateDialog(hLangLib, MAKEINTRESOURCEW(IDD_REPLACE), hMainWnd, (DLGPROC)FindAndReplaceDlgProc);

    if (hDlgModeless)
      ShowWindow(hDlgModeless, SW_SHOW);
  }
  else if (nModelessType == MLT_REPLACE)
  {
    if (!SendMessage(hDlgModeless, AKDLG_PUTFIND, 0, 0))
      BringWindowToTop(hDlgModeless);
    else
      SetFocus(hDlgModeless);
  }
}

void DoEditGoTo()
{
  if (!hDlgModeless)
  {
    nModelessType=MLT_GOTO;

    hDlgModeless=API_CreateDialog(hLangLib, MAKEINTRESOURCEW(IDD_GOTO), hMainWnd, (DLGPROC)GoToDlgProc);

    if (hDlgModeless)
    {
      ShowWindow(hDlgModeless, SW_SHOW);
    }
  }
}

BOOL DoViewFont(HWND hWndOwner, LOGFONTW *lfFont)
{
  if (bOldWindows)
  {
    CHOOSEFONTA cfA;
    LOGFONTA lfTmpA;

    xmemset(&cfA, 0, sizeof(CHOOSEFONTA));
    LogFontWtoA(lfFont, &lfTmpA);
    cfA.lStructSize=sizeof(CHOOSEFONTA);
    cfA.Flags      =CF_FORCEFONTEXIST|CF_INITTOLOGFONTSTRUCT|CF_SCREENFONTS;
    cfA.hwndOwner  =hWndOwner;
    cfA.lpLogFont  =&lfTmpA;
    if (!ChooseFontA(&cfA))
      return FALSE;
    LogFontAtoW(&lfTmpA, lfFont);
  }
  else
  {
    CHOOSEFONTW cfW;
    LOGFONTW lfTmpW;

    xmemset(&cfW, 0, sizeof(CHOOSEFONTW));
    xmemcpy(&lfTmpW, lfFont, sizeof(LOGFONTW));
    cfW.lStructSize=sizeof(CHOOSEFONTW);
    cfW.Flags      =CF_FORCEFONTEXIST|CF_INITTOLOGFONTSTRUCT|CF_SCREENFONTS;
    cfW.hwndOwner  =hWndOwner;
    cfW.lpLogFont  =&lfTmpW;
    if (!ChooseFontW(&cfW))
      return FALSE;
    xmemcpy(lfFont, &lfTmpW, sizeof(LOGFONTW));
  }
  lfFont->lfHeight=-mod(lfFont->lfHeight);

  //User pressed OK, so unlock font inheriting.
  //if (lpFrameCurrent->dwLockInherit & LI_FONT)
  //  lpFrameCurrent->dwLockInherit&=~LI_FONT;

  return TRUE;
}

BOOL DoViewColors()
{
  return (BOOL)API_DialogBox(hLangLib, MAKEINTRESOURCEW(IDD_COLORS), hMainWnd, (DLGPROC)ColorsDlgProc);
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void DoViewFontSize(FRAMEDATA *lpFrame, int nAction)
{
  if (nAction == FONTSIZE_INCREASE)
  {
    if (lpFrame->lf.lfHeight <= -1)
      lpFrame->lf.lfHeight-=1;
  }
  else if (nAction == FONTSIZE_DECREASE)
  {
    if (lpFrame->lf.lfHeight < -1)
      lpFrame->lf.lfHeight+=1;
  }
  else if (nAction == FONTSIZE_RESTORE)
  {
    lpFrame->lf.lfHeight=fdDefault.lf.lfHeight;
  }
  SetChosenFont(lpFrame->ei.hWndEdit, &lpFrame->lf);
  UpdateMappedPrintWidth(lpFrame);
  UpdateStatusUser(lpFrame, CSB_FONTPOINT|CSB_MARKER);
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void DoViewReadOnly(FRAMEDATA *lpFrame, BOOL bState, BOOL bFirst)
{
  CheckMenuItem(hMainMenu, IDM_VIEW_READONLY, bState?MF_CHECKED:MF_UNCHECKED);
  if (bFirst != TRUE && bState == lpFrame->ei.bReadOnly) return;
  lpFrame->ei.bReadOnly=bState;

  SendMessage(lpFrame->ei.hWndEdit, AEM_SETOPTIONS, lpFrame->ei.bReadOnly?AECOOP_OR:AECOOP_XOR, AECO_READONLY);
  SetModifyStatus(lpFrame, lpFrame->ei.bModified);
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void DoViewWordWrap(FRAMEDATA *lpFrame, BOOL bState, BOOL bFirst)
{
  CheckMenuItem(hMainMenu, IDM_VIEW_WORDWRAP, bState?MF_CHECKED:MF_UNCHECKED);
  if (bFirst != TRUE && bState == lpFrame->ei.bWordWrap) return;
  lpFrame->ei.bWordWrap=bState;

  if (lpFrame->ei.bWordWrap)
  {
    UpdateShowHScroll(lpFrame);
    SetWordWrap(lpFrame, lpFrame->dwWrapType, lpFrame->dwWrapLimit);
  }
  else
  {
    SetWordWrap(lpFrame, AEWW_NONE, 0);
    UpdateShowHScroll(lpFrame);
  }
}

void DoViewSplitWindow(BOOL bState, WPARAM wParam)
{
  CheckMenuItem(hMainMenu, IDM_VIEW_SPLIT_WINDOW_ONOFF, bState?MF_CHECKED:MF_UNCHECKED);

  if (bState)
  {
    RECT *lprcEditWindow;

    if (wParam == IDM_VIEW_SPLIT_WINDOW_WE)
    {
      SplitCreate(lpFrameCurrent, CN_CLONE1);
    }
    else if (wParam == IDM_VIEW_SPLIT_WINDOW_NS)
    {
      SplitCreate(lpFrameCurrent, CN_CLONE2);
    }
    else if (wParam == IDM_VIEW_SPLIT_WINDOW_ALL)
    {
      SplitCreate(lpFrameCurrent, CN_CLONE1|CN_CLONE2|CN_CLONE3);
    }

    if (nMDI == WMD_MDI)
      lprcEditWindow=&lpFrameCurrent->rcEditWindow;
    else
      lprcEditWindow=&fdDefault.rcEditWindow;
    lpFrameCurrent->rcMasterWindow.left=lprcEditWindow->left;
    lpFrameCurrent->rcMasterWindow.top=lprcEditWindow->top;
    lpFrameCurrent->rcMasterWindow.right=lprcEditWindow->right / 2;
    lpFrameCurrent->rcMasterWindow.bottom=lprcEditWindow->bottom / 2;
  }
  else
  {
    //Destroy
    SplitDestroy(lpFrameCurrent, CN_CLONE1|CN_CLONE2|CN_CLONE3);
    SetFocus(lpFrameCurrent->ei.hWndEdit);
  }

  if (lpFrameCurrent->ei.bWordWrap)
    UpdateShowHScroll(lpFrameCurrent);
  ResizeEditWindow(lpFrameCurrent, 0);
}

void DoViewOnTop(BOOL bState, BOOL bFirst)
{
  CheckMenuItem(hMainMenu, IDM_VIEW_ONTOP, bState?MF_CHECKED:MF_UNCHECKED);
  if (bFirst != TRUE && bState == moCur.bOnTop) return;
  moCur.bOnTop=bState;

  if (moCur.bOnTop)
    SetWindowPos(hMainWnd, (HWND)(UINT_PTR)-1/*HWND_TOPMOST*/, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
  else
    SetWindowPos(hMainWnd, (HWND)(UINT_PTR)-2/*HWND_NOTOPMOST*/, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
}

void DoViewShowStatusBar(BOOL bState, BOOL bFirst)
{
  CheckMenuItem(hMainMenu, IDM_VIEW_SHOW_STATUSBAR, bState?MF_CHECKED:MF_UNCHECKED);
  if (bFirst != TRUE && bState == moCur.bStatusBar) return;
  moCur.bStatusBar=bState;

  ShowWindow(hStatus, moCur.bStatusBar?SW_SHOW:SW_HIDE);
  UpdateSize();
}

BOOL DoSettingsExec()
{
  STARTUPINFOW si;
  PROCESS_INFORMATION pi;
  wchar_t *wszExecuteCommandExp;
  wchar_t *wszExecuteDirectoryExp;
  int nCommandLen;
  int nWorkDirLen;
  BOOL bResult=FALSE;

  nCommandLen=TranslateFileString(moCur.wszExecuteCommand, NULL, 0);
  nWorkDirLen=TranslateFileString(moCur.wszExecuteDirectory, NULL, 0);

  if (wszExecuteCommandExp=AllocWideStr(nCommandLen + 1))
  {
    if (wszExecuteDirectoryExp=AllocWideStr(nWorkDirLen + 1))
    {
      TranslateFileString(moCur.wszExecuteCommand, wszExecuteCommandExp, nCommandLen + 1);
      TranslateFileString(moCur.wszExecuteDirectory, wszExecuteDirectoryExp, nWorkDirLen + 1);

      if (*wszExecuteCommandExp)
      {
        xmemset(&si, 0, sizeof(STARTUPINFOW));
        si.cb=sizeof(STARTUPINFOW);
        if (CreateProcessWide(NULL, wszExecuteCommandExp, NULL, NULL, FALSE, 0, NULL, (wszExecuteDirectoryExp && *wszExecuteDirectoryExp)?wszExecuteDirectoryExp:NULL, &si, &pi))
        {
          bResult=TRUE;
          CloseHandle(pi.hProcess);
          CloseHandle(pi.hThread);
        }
        else
        {
          API_LoadStringW(hLangLib, MSG_ERROR_RUN, wszMsg, BUFFER_SIZE);
          API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
        }
      }
      FreeWideStr(wszExecuteDirectoryExp);
    }
    FreeWideStr(wszExecuteCommandExp);
  }
  return bResult;
}

void DoSettingsKeepSpace(BOOL bState)
{
  CheckMenuItem(hMainMenu, IDM_OPTIONS_KEEPSPACE, bState?MF_CHECKED:MF_UNCHECKED);
  moCur.bKeepSpace=bState;
}

void DoSettingsWatchFile(BOOL bState)
{
  CheckMenuItem(hMainMenu, IDM_OPTIONS_WATCHFILE, bState?MF_CHECKED:MF_UNCHECKED);
  moCur.bWatchFile=bState;
  if (moCur.bWatchFile) SetFocus(hMainWnd);
}

void DoSettingsSaveTime(BOOL bState)
{
  CheckMenuItem(hMainMenu, IDM_OPTIONS_SAVETIME, bState?MF_CHECKED:MF_UNCHECKED);
  moCur.bSaveTime=bState;
}

void DoSettingsSingleOpenFile(BOOL bState)
{
  CheckMenuItem(hMainMenu, IDM_OPTIONS_SINGLEOPEN_FILE, bState?MF_CHECKED:MF_UNCHECKED);
  moCur.bSingleOpenFile=bState;
}

void DoSettingsSingleOpenProgram(DWORD dwState)
{
  CheckMenuItem(hMainMenu, IDM_OPTIONS_SINGLEOPEN_PROGRAM, (dwState & SOP_ON)?MF_CHECKED:MF_UNCHECKED);
  moCur.dwSingleOpenProgram=dwState;
}

void DoSettingsPlugins()
{
  API_DialogBox(hLangLib, MAKEINTRESOURCEW(IDD_PLUGINS), hMainWnd, (DLGPROC)PluginsDlgProc);
}

void DoSettingsOptions()
{
  PROPSHEETHEADERW psh;
  PROPSHEETPAGEW psp[5];
  DWORD dwInitKeybLayoutOptions=moCur.dwKeybLayoutOptions;

  bOptionsSave=FALSE;
  bOptionsRestart=FALSE;

  xmemset(&psp, 0, sizeof(psp));
  psp[0].dwSize      =sizeof(PROPSHEETPAGEW);
  psp[0].dwFlags     =PSP_DEFAULT;
  psp[0].hInstance   =hLangLib;
  psp[0].pszTemplate =MAKEINTRESOURCEW(IDD_OPTIONS_GENERAL);
  psp[0].pfnDlgProc  =(DLGPROC)OptionsGeneralDlgProc;
  psp[1].dwSize      =sizeof(PROPSHEETPAGEW);
  psp[1].dwFlags     =PSP_DEFAULT;
  psp[1].hInstance   =hLangLib;
  psp[1].pszTemplate =MAKEINTRESOURCEW(IDD_OPTIONS_REGISTRY);
  psp[1].pfnDlgProc  =(DLGPROC)OptionsRegistryDlgProc;
  psp[2].dwSize      =sizeof(PROPSHEETPAGEW);
  psp[2].dwFlags     =PSP_DEFAULT;
  psp[2].hInstance   =hLangLib;
  psp[2].pszTemplate =MAKEINTRESOURCEW(IDD_OPTIONS_EDITOR1);
  psp[2].pfnDlgProc  =(DLGPROC)OptionsEditor1DlgProc;
  psp[3].dwSize      =sizeof(PROPSHEETPAGEW);
  psp[3].dwFlags     =PSP_DEFAULT;
  psp[3].hInstance   =hLangLib;
  psp[3].pszTemplate =MAKEINTRESOURCEW(IDD_OPTIONS_EDITOR2);
  psp[3].pfnDlgProc  =(DLGPROC)OptionsEditor2DlgProc;
  psp[4].dwSize      =sizeof(PROPSHEETPAGEW);
  psp[4].dwFlags     =PSP_DEFAULT;
  psp[4].hInstance   =hLangLib;
  psp[4].pszTemplate =MAKEINTRESOURCEW(IDD_OPTIONS_ADVANCED);
  psp[4].pfnDlgProc  =(DLGPROC)OptionsAdvancedDlgProc;

  API_LoadStringW(hLangLib, STR_OPTIONS, wbuf, BUFFER_SIZE);
  xmemset(&psh, 0, sizeof(PROPSHEETHEADERW));
  psh.pszCaption  =wbuf;
  psh.dwSize      =sizeof(PROPSHEETHEADERW);
  psh.dwFlags     =PSH_PROPSHEETPAGE|PSH_NOAPPLYNOW|PSH_USEICONID|PSH_USECALLBACK;
  psh.hwndParent  =hMainWnd;
  psh.hInstance   =hLangLib;
  psh.pszIcon     =MAKEINTRESOURCEW(IDI_ICON_MAIN);
  psh.nPages      =sizeof(psp) / sizeof(PROPSHEETPAGEW);
  psh.nStartPage  =nPropertyStartPage;
  psh.ppsp        =&psp[0];
  psh.pfnCallback =PropSheetProc;

  hHookPropertySheet=SetWindowsHookEx(WH_CBT, CBTPropertySheetProc, NULL, GetCurrentThreadId());
  PropertySheetWide(&psh);

  if (dwInitKeybLayoutOptions != moCur.dwKeybLayoutOptions)
  {
    //Run after property dialog closed, because SwitchLayout require focus
    if (moCur.dwKeybLayoutOptions & KLO_SWITCHLAYOUT)
      SwitchLayout(lpFrameCurrent->ei.hWndEdit, &ciCurCaret);
  }

  if (bOptionsSave)
  {
    //Save to INI
    SaveOptions(&moCur, lpFrameCurrent, SS_INI, TRUE);
    SaveThemes(SS_INI);
    StackPluginSave(&hPluginsStack, SS_INI);

    //Save to registry
    SaveOptions(&moCur, lpFrameCurrent, SS_REGISTRY, TRUE);
    SaveThemes(SS_REGISTRY);
    StackPluginSave(&hPluginsStack, SS_REGISTRY);
  }

  if (bOptionsRestart)
  {
    API_LoadStringW(hLangLib, MSG_RESTART_PROGRAM, wszMsg, BUFFER_SIZE);
    API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
  }
}

void DoWindowTabView(DWORD dwNewView, BOOL bFirst)
{
  DWORD dwOldView=moCur.dwTabOptionsMDI;
  DWORD dwStyle;
  int nCommand=0;

  if (dwNewView & TAB_VIEW_TOP)
  {
    nCommand=IDM_WINDOW_TABVIEW_TOP;
    dwNewView=TAB_VIEW_TOP;
  }
  else if (dwNewView & TAB_VIEW_BOTTOM)
  {
    nCommand=IDM_WINDOW_TABVIEW_BOTTOM;
    dwNewView=TAB_VIEW_BOTTOM;
  }
  else if (dwNewView & TAB_VIEW_NONE)
  {
    nCommand=IDM_WINDOW_TABVIEW_NONE;
    dwNewView=TAB_VIEW_NONE;
  }
  if (dwOldView & TAB_VIEW_TOP)
    dwOldView=TAB_VIEW_TOP;
  else if (dwOldView & TAB_VIEW_BOTTOM)
    dwOldView=TAB_VIEW_BOTTOM;
  else if (dwOldView & TAB_VIEW_NONE)
    dwOldView=TAB_VIEW_NONE;
  CheckMenuRadioItem(hMainMenu, IDM_WINDOW_TABVIEW_TOP, IDM_WINDOW_TABVIEW_NONE, nCommand, MF_BYCOMMAND);
  EnableMenuItem(hMainMenu, IDM_WINDOW_TABTYPE_STANDARD, !(dwNewView & TAB_VIEW_NONE)?MF_ENABLED:MF_GRAYED);
  EnableMenuItem(hMainMenu, IDM_WINDOW_TABTYPE_BUTTONS, !(dwNewView & TAB_VIEW_NONE)?MF_ENABLED:MF_GRAYED);
  if (!bOldComctl32) EnableMenuItem(hMainMenu, IDM_WINDOW_TABTYPE_FLATBUTTONS, !(dwNewView & TAB_VIEW_NONE)?MF_ENABLED:MF_GRAYED);

  if (bFirst != TRUE && dwNewView == dwOldView) return;
  moCur.dwTabOptionsMDI=moCur.dwTabOptionsMDI & ~TAB_VIEW_TOP & ~TAB_VIEW_BOTTOM & ~TAB_VIEW_NONE;
  moCur.dwTabOptionsMDI|=dwNewView;

  if (moCur.dwTabOptionsMDI & TAB_VIEW_TOP)
  {
    dwStyle=(DWORD)GetWindowLongPtrWide(hTab, GWL_STYLE);
    SetWindowLongPtrWide(hTab, GWL_STYLE, dwStyle & ~TCS_BOTTOM);
  }
  else if (moCur.dwTabOptionsMDI & TAB_VIEW_BOTTOM)
  {
    dwStyle=(DWORD)GetWindowLongPtrWide(hTab, GWL_STYLE);
    SetWindowLongPtrWide(hTab, GWL_STYLE, dwStyle|TCS_BOTTOM);
  }
  ShowWindow(hTab, !(moCur.dwTabOptionsMDI & TAB_VIEW_NONE)?SW_SHOW:SW_HIDE);
  UpdateSize();
}

void DoWindowTabType(DWORD dwNewType, BOOL bFirst)
{
  DWORD dwOldType=moCur.dwTabOptionsMDI;
  DWORD dwCurStyle;
  int nCommand=0;

  if (dwNewType & TAB_TYPE_STANDARD)
  {
    nCommand=IDM_WINDOW_TABTYPE_STANDARD;
    dwNewType=TAB_TYPE_STANDARD;
  }
  else if (dwNewType & TAB_TYPE_BUTTONS)
  {
    nCommand=IDM_WINDOW_TABTYPE_BUTTONS;
    dwNewType=TAB_TYPE_BUTTONS;
  }
  else if (dwNewType & TAB_TYPE_FLATBUTTONS)
  {
    nCommand=IDM_WINDOW_TABTYPE_FLATBUTTONS;
    dwNewType=TAB_TYPE_FLATBUTTONS;
  }
  if (dwOldType & TAB_TYPE_STANDARD)
    dwOldType=TAB_TYPE_STANDARD;
  else if (dwOldType & TAB_TYPE_BUTTONS)
    dwOldType=TAB_TYPE_BUTTONS;
  else if (dwOldType & TAB_TYPE_FLATBUTTONS)
    dwOldType=TAB_TYPE_FLATBUTTONS;
  CheckMenuRadioItem(hMainMenu, IDM_WINDOW_TABTYPE_STANDARD, IDM_WINDOW_TABTYPE_FLATBUTTONS, nCommand, MF_BYCOMMAND);

  if (bFirst != TRUE && dwNewType == dwOldType) return;
  moCur.dwTabOptionsMDI=moCur.dwTabOptionsMDI & ~TAB_TYPE_STANDARD & ~TAB_TYPE_BUTTONS & ~TAB_TYPE_FLATBUTTONS;
  moCur.dwTabOptionsMDI|=dwNewType;
  dwCurStyle=(DWORD)GetWindowLongPtrWide(hTab, GWL_STYLE);

  if (!(dwCurStyle & TCS_BUTTONS) == !(moCur.dwTabOptionsMDI & TAB_TYPE_STANDARD))
  {
    API_LoadStringW(hLangLib, MSG_RESTART_PROGRAM, wszMsg, BUFFER_SIZE);
    API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
  }
  else
  {
    if (moCur.dwTabOptionsMDI & TAB_TYPE_STANDARD)
    {
      SetWindowLongPtrWide(hTab, GWL_STYLE, (dwCurStyle | TCS_TABS) & ~TCS_BUTTONS & ~TCS_FLATBUTTONS);
      SendMessage(hTab, TCM_SETITEMSIZE, 0, MAKELPARAM(TAB_WIDTH, TAB_HEIGHT - 4));
    }
    else if (moCur.dwTabOptionsMDI & TAB_TYPE_BUTTONS)
    {
      SetWindowLongPtrWide(hTab, GWL_STYLE, (dwCurStyle | TCS_BUTTONS) & ~TCS_TABS & ~TCS_FLATBUTTONS);
      SendMessage(hTab, TCM_SETITEMSIZE, 0, MAKELPARAM(TAB_WIDTH, TAB_HEIGHT));
    }
    else if (moCur.dwTabOptionsMDI & TAB_TYPE_FLATBUTTONS)
    {
      SetWindowLongPtrWide(hTab, GWL_STYLE, (dwCurStyle | TCS_BUTTONS | TCS_FLATBUTTONS) & ~TCS_TABS);
      SendMessage(hTab, TCM_SETITEMSIZE, 0, MAKELPARAM(TAB_WIDTH, TAB_HEIGHT));
    }
  }
}

void DoWindowTabSwitch(DWORD dwNewSwitch, BOOL bFirst)
{
  DWORD dwOldSwitch=moCur.dwTabOptionsMDI;
  int nCommand=0;

  if (dwNewSwitch & TAB_SWITCH_NEXTPREV)
  {
    nCommand=IDM_WINDOW_TABSWITCH_NEXTPREV;
    dwNewSwitch=TAB_SWITCH_NEXTPREV;
  }
  else if (dwNewSwitch & TAB_SWITCH_RIGHTLEFT)
  {
    nCommand=IDM_WINDOW_TABSWITCH_RIGHTLEFT;
    dwNewSwitch=TAB_SWITCH_RIGHTLEFT;
  }
  if (dwOldSwitch & TAB_SWITCH_NEXTPREV)
    dwOldSwitch=TAB_SWITCH_NEXTPREV;
  else if (dwOldSwitch & TAB_SWITCH_RIGHTLEFT)
    dwOldSwitch=TAB_SWITCH_RIGHTLEFT;
  CheckMenuRadioItem(hMainMenu, IDM_WINDOW_TABSWITCH_NEXTPREV, IDM_WINDOW_TABSWITCH_RIGHTLEFT, nCommand, MF_BYCOMMAND);

  if (bFirst != TRUE && dwNewSwitch == dwOldSwitch) return;
  moCur.dwTabOptionsMDI=moCur.dwTabOptionsMDI & ~TAB_SWITCH_RIGHTLEFT & ~TAB_SWITCH_NEXTPREV;
  moCur.dwTabOptionsMDI|=dwNewSwitch;
}

void DoWindowSelectWindow()
{
  API_DialogBox(hLangLib, MAKEINTRESOURCEW(IDD_MDILIST), hMainWnd, (DLGPROC)MdiListDlgProc);
}

void DoHelpAbout()
{
  API_DialogBox(hLangLib, MAKEINTRESOURCEW(IDD_ABOUT), hMainWnd, (DLGPROC)AboutDlgProc);
}

void DoNonMenuDelLine(HWND hWnd, BOOL bAllWrap)
{
  AECHARRANGE cr;
  POINT64 ptGlobal;
  INT_PTR nCaretHorzIndent;

  if (IsReadOnly(hWnd)) return;

  if (bAllWrap)
  {
    AEC_WrapLineBeginEx(&crCurSel.ciMin, &cr.ciMin);
    AEC_WrapLineEndEx(&crCurSel.ciMax, &cr.ciMax);
  }
  else
  {
    cr.ciMin=crCurSel.ciMin;
    cr.ciMax=crCurSel.ciMax;
    cr.ciMin.nCharInLine=0;
  }
  if (!AEC_NextLineEx(&cr.ciMax, &cr.ciMax))
  {
    //Last line
    cr.ciMax.nCharInLine=cr.ciMax.lpLine->nLineLen;
    AEC_PrevLineEx(&cr.ciMin, &cr.ciMin);
  }
  nCaretHorzIndent=SendMessage(hWnd, AEM_GETCARETHORZINDENT, 0, 0);

  SetSel(hWnd, &cr, AESELT_LOCKSCROLL|AESELT_LOCKUPDATE|AESELT_LOCKCARET, NULL);
  ReplaceSelW(hWnd, L"", -1, AELB_ASINPUT, 0, NULL, NULL);

  SendMessage(hWnd, AEM_GETCARETPOS, 0, (WPARAM)&ptGlobal);
  ptGlobal.x=nCaretHorzIndent;
  SendMessage(hWnd, AEM_CHARFROMGLOBALPOS, (WPARAM)&ptGlobal, (LPARAM)&cr.ciMin);
  cr.ciMax=cr.ciMin;
  SetSel(hWnd, &cr, 0, NULL);
  SendMessage(hWnd, AEM_SETCARETHORZINDENT, nCaretHorzIndent, 0);
}

BOOL DoNonMenuSelJumpCaret(HWND hWnd)
{
  AECHARINDEX *lpciNewCaret=&crCurSel.ciMin;

  if (AEC_IndexCompare(&crCurSel.ciMin, &crCurSel.ciMax))
  {
    if (!AEC_IndexCompare(&crCurSel.ciMin, &ciCurCaret))
      lpciNewCaret=&crCurSel.ciMax;

    SetSel(hWnd, &crCurSel, AESELT_COLUMNASIS|AESELT_LOCKSCROLL, lpciNewCaret);
    if (!ScrollCaret(hWnd))
    {
      //Post WM_PAINT if necessary
      if (!GetUpdateRect(hWnd, NULL, FALSE))
      {
        RECT rc;

        rc.left=0;
        rc.top=0;
        rc.right=1;
        rc.bottom=1;
        InvalidateRect(hWnd, &rc, FALSE);
      }
    }
  }
  return (lpciNewCaret == &crCurSel.ciMin);
}


//// INI functions

BOOL OpenIniA(INIFILE *hIniFile, const char *pFile, BOOL bCreate)
{
  HANDLE hFile;
  DWORD dwFlags=0;

  if (!bCreate)
  {
    if (FileExistsAnsi(pFile))
      dwFlags=OPEN_EXISTING;
    else
      return FALSE;
  }
  else dwFlags=OPEN_ALWAYS;

  if ((hFile=API_CreateFileA(pFile, GENERIC_READ, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, dwFlags, FILE_ATTRIBUTE_NORMAL|FILE_FLAG_SEQUENTIAL_SCAN, NULL)) != INVALID_HANDLE_VALUE)
  {
    ReadIni(hIniFile, hFile);
    CloseHandle(hFile);
    return TRUE;
  }
  return FALSE;
}

BOOL OpenIniW(INIFILE *hIniFile, const wchar_t *wpFile, BOOL bCreate)
{
  HANDLE hFile;
  DWORD dwFlags=0;

  if (!bCreate)
  {
    if (FileExistsWide(wpFile))
      dwFlags=OPEN_EXISTING;
    else
      return FALSE;
  }
  else dwFlags=OPEN_ALWAYS;

  if ((hFile=API_CreateFileW(wpFile, GENERIC_READ, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, dwFlags, FILE_ATTRIBUTE_NORMAL|FILE_FLAG_SEQUENTIAL_SCAN, NULL)) != INVALID_HANDLE_VALUE)
  {
    ReadIni(hIniFile, hFile);
    CloseHandle(hFile);
    return TRUE;
  }
  return FALSE;
}

BOOL SaveIniA(INIFILE *hIniFile, const char *pFile)
{
  HANDLE hFile;
  DWORD dwAttr;
  BOOL bResult=FALSE;

  dwAttr=GetFileAttributesA(pFile);

  if (dwAttr == INVALID_FILE_ATTRIBUTES || !(dwAttr & FILE_ATTRIBUTE_READONLY))
  {
    if ((hFile=API_CreateFileA(pFile, GENERIC_WRITE, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL)) != INVALID_HANDLE_VALUE)
    {
      bResult=WriteIni(hIniFile, hFile);
      CloseHandle(hFile);
    }
  }
  return bResult;
}

BOOL SaveIniW(INIFILE *hIniFile, const wchar_t *wpFile)
{
  HANDLE hFile;
  DWORD dwAttr;
  BOOL bResult=FALSE;

  dwAttr=GetFileAttributesWide(wpFile);

  if (dwAttr == INVALID_FILE_ATTRIBUTES || !(dwAttr & FILE_ATTRIBUTE_READONLY))
  {
    if ((hFile=API_CreateFileW(wpFile, GENERIC_WRITE, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL)) != INVALID_HANDLE_VALUE)
    {
      bResult=WriteIni(hIniFile, hFile);
      CloseHandle(hFile);
    }
  }
  return bResult;
}

BOOL ReadIni(INIFILE *hIniFile, HANDLE hFile)
{
  INISECTION *lpIniSection=NULL;
  INIKEY *lpIniKey=NULL;
  wchar_t *wszText=NULL;
  wchar_t *wpText;
  wchar_t *wpSection;
  wchar_t *wpKey;
  wchar_t *wpString;
  wchar_t *wpLastChar;
  UINT_PTR dwFileSize;
  UINT_PTR dwBytesRead;
  UINT_PTR dwUnicodeLen;
  int nSectionLen;
  int nKeyLen;
  int nStringLen;

  if ((dwFileSize=GetFileSize64(hFile)) != (UINT_PTR)-1)
  {
    if (dwFileSize >= sizeof(wchar_t))
    {
      if (wszText=wpText=(wchar_t *)API_HeapAlloc(hHeap, 0, dwFileSize + sizeof(wchar_t)))
      {
        if (ReadFile64(hFile, wpText, dwFileSize, &dwBytesRead, NULL))
        {
          dwUnicodeLen=dwBytesRead / sizeof(wchar_t);
          wpLastChar=wpText + dwUnicodeLen;
          wpText[dwUnicodeLen++]=L'\0';
          if (wszText[0] == 0xFEFF)
          {
            ++wpText;
            --dwUnicodeLen;
          }

          while (*wpText == L' ' || *wpText == L'\t' || *wpText == L'\r' || *wpText == L'\n')
            ++wpText;

          while (wpText < wpLastChar)
          {
            //Section
            if (*wpText == L'[')
            {
              for (wpSection=++wpText; wpText < wpLastChar && *wpText != L']' && *wpText != L'\r' && *wpText != L'\n'; ++wpText);

              if (*wpText == L']')
              {
                nSectionLen=(int)(wpText - wpSection);
                *wpText++=L'\0';

                if (!StackInsertIndex((stack **)&hIniFile->first, (stack **)&hIniFile->last, (stack **)&lpIniSection, -1, sizeof(INISECTION)))
                {
                  lpIniSection->hIniFile=(HANDLE)hIniFile;
                  lpIniSection->nSectionBytes=(nSectionLen + 1) * sizeof(wchar_t);
                  if (lpIniSection->wszSection=(wchar_t *)API_HeapAlloc(hHeap, 0, lpIniSection->nSectionBytes))
                    xmemcpy(lpIniSection->wszSection, wpSection, lpIniSection->nSectionBytes);

                  lpIniSection->first=0;
                  lpIniSection->last=0;
                }
                else goto Error;
              }
              else goto Error;
            }
            else goto Error;

            //Keys and strings
            for (;;)
            {
              while (*wpText == L' ' || *wpText == L'\t' || *wpText == L'\r' || *wpText == L'\n')
                ++wpText;
              if (wpText >= wpLastChar || *wpText == L'[') break;

              for (wpKey=wpText; wpText < wpLastChar && *wpText != L'=' && *wpText != L'\r' && *wpText != L'\n'; ++wpText);

              if (*wpText == L'=')
              {
                nKeyLen=(int)(wpText - wpKey);
                *wpText++=L'\0';

                for (wpString=wpText; wpText < wpLastChar && *wpText != L'\r' && *wpText != L'\n'; ++wpText);
                nStringLen=(int)(wpText - wpString);
                if (*wpText) *wpText++=L'\0';

                if (!StackInsertIndex((stack **)&lpIniSection->first, (stack **)&lpIniSection->last, (stack **)&lpIniKey, -1, sizeof(INIKEY)))
                {
                  lpIniKey->nKeyBytes=(nKeyLen + 1) * sizeof(wchar_t);
                  if (lpIniKey->wszKey=(wchar_t *)API_HeapAlloc(hHeap, 0, lpIniKey->nKeyBytes))
                    xmemcpy(lpIniKey->wszKey, wpKey, lpIniKey->nKeyBytes);

                  lpIniKey->nStringBytes=(nStringLen + 1) * sizeof(wchar_t);
                  if (lpIniKey->wszString=(wchar_t *)API_HeapAlloc(hHeap, 0, lpIniKey->nStringBytes))
                    xmemcpy(lpIniKey->wszString, wpString, lpIniKey->nStringBytes);
                }
                else goto Error;
              }
              else goto Error;
            }
          }
          API_HeapFree(hHeap, 0, (LPVOID)wszText);
          return TRUE;
        }
      }
    }
  }
  Error:
  if (wszText) API_HeapFree(hHeap, 0, (LPVOID)wszText);
  return FALSE;
}

BOOL WriteIni(INIFILE *hIniFile, HANDLE hFile)
{
  INISECTION *lpIniSection=(INISECTION *)hIniFile->first;
  INIKEY *lpIniKey=NULL;
  UINT_PTR dwBytesWritten;

  if (!API_WriteFile(hFile, "\xFF\xFE", 2, &dwBytesWritten, NULL))
    return FALSE;

  while (lpIniSection)
  {
    //Section
    if (!API_WriteFile(hFile, L"[", 2, &dwBytesWritten, NULL))
      return FALSE;
    if (!API_WriteFile(hFile, lpIniSection->wszSection, lpIniSection->nSectionBytes - sizeof(wchar_t), &dwBytesWritten, NULL))
      return FALSE;
    if (!API_WriteFile(hFile, L"]\r\n", 6, &dwBytesWritten, NULL))
      return FALSE;

    //Keys and strings
    lpIniKey=(INIKEY *)lpIniSection->first;

    while (lpIniKey)
    {
      if (!API_WriteFile(hFile, lpIniKey->wszKey, lpIniKey->nKeyBytes - sizeof(wchar_t), &dwBytesWritten, NULL))
        return FALSE;
      if (!API_WriteFile(hFile, L"=", 2, &dwBytesWritten, NULL))
        return FALSE;
      if (!API_WriteFile(hFile, lpIniKey->wszString, lpIniKey->nStringBytes - sizeof(wchar_t), &dwBytesWritten, NULL))
        return FALSE;
      if (!API_WriteFile(hFile, L"\r\n", 4, &dwBytesWritten, NULL))
        return FALSE;
      lpIniKey=lpIniKey->next;
    }
    if (lpIniSection->next)
    {
      if (!API_WriteFile(hFile, L"\r\n", 4, &dwBytesWritten, NULL))
        return FALSE;
    }
    lpIniSection=lpIniSection->next;
  }
  return TRUE;
}

INISECTION* StackOpenIniSectionA(INIFILE *hIniFile, const char *pSection, int nSectionLen, BOOL bCreate)
{
  INISECTION *lpIniSection=NULL;
  wchar_t *wpSection;

  if (wpSection=(wchar_t *)API_HeapAlloc(hHeap, 0, (nSectionLen + 1) * sizeof(wchar_t)))
  {
    nSectionLen=MultiByteToWideChar(CP_ACP, 0, pSection, nSectionLen + 1, wpSection, nSectionLen + 1);
    lpIniSection=StackOpenIniSectionW(hIniFile, wpSection, nSectionLen - 1, bCreate);
    API_HeapFree(hHeap, 0, (LPVOID)wpSection);
  }
  return lpIniSection;
}

INISECTION* StackOpenIniSectionW(INIFILE *hIniFile, const wchar_t *wpSection, int nSectionLen, BOOL bCreate)
{
  INISECTION *lpIniSection=(INISECTION *)hIniFile->first;
  int nSectionBytes=(nSectionLen + 1) * sizeof(wchar_t);

  while (lpIniSection)
  {
    if (lpIniSection->nSectionBytes == nSectionBytes)
    {
      if (!xstrcmpiW(lpIniSection->wszSection, wpSection))
        return lpIniSection;
    }
    lpIniSection=lpIniSection->next;
  }

  if (bCreate)
  {
    if (!StackInsertIndex((stack **)&hIniFile->first, (stack **)&hIniFile->last, (stack **)&lpIniSection, -1, sizeof(INISECTION)))
    {
      lpIniSection->hIniFile=(HANDLE)hIniFile;
      lpIniSection->nSectionBytes=nSectionBytes;
      if (lpIniSection->wszSection=(wchar_t *)API_HeapAlloc(hHeap, 0, lpIniSection->nSectionBytes))
        xmemcpy(lpIniSection->wszSection, wpSection, lpIniSection->nSectionBytes);

      lpIniSection->first=0;
      lpIniSection->last=0;
    }
  }
  return lpIniSection;
}

INIKEY* StackOpenIniKeyA(INISECTION *lpIniSection, const char *pKey, int nKeyLen, BOOL bCreate)
{
  INIKEY *lpIniKey=NULL;
  wchar_t *wpKey;

  if (wpKey=(wchar_t *)API_HeapAlloc(hHeap, 0, (nKeyLen + 1) * sizeof(wchar_t)))
  {
    nKeyLen=MultiByteToWideChar(CP_ACP, 0, pKey, nKeyLen + 1, wpKey, nKeyLen + 1);
    lpIniKey=StackOpenIniKeyW(lpIniSection, wpKey, nKeyLen - 1, bCreate);
    API_HeapFree(hHeap, 0, (LPVOID)wpKey);
  }
  return lpIniKey;
}

INIKEY* StackOpenIniKeyW(INISECTION *lpIniSection, const wchar_t *wpKey, int nKeyLen, BOOL bCreate)
{
  INIKEY *lpIniKey=(INIKEY *)lpIniSection->first;
  int nKeyBytes=(nKeyLen + 1) * sizeof(wchar_t);

  while (lpIniKey)
  {
    if (lpIniKey->nKeyBytes == nKeyBytes)
    {
      if (!xstrcmpiW(lpIniKey->wszKey, wpKey))
        return lpIniKey;
    }
    lpIniKey=lpIniKey->next;
  }

  if (bCreate)
  {
    if (!StackInsertIndex((stack **)&lpIniSection->first, (stack **)&lpIniSection->last, (stack **)&lpIniKey, -1, sizeof(INIKEY)))
    {
      lpIniKey->nKeyBytes=nKeyBytes;
      if (lpIniKey->wszKey=(wchar_t *)API_HeapAlloc(hHeap, 0, lpIniKey->nKeyBytes))
        xmemcpy(lpIniKey->wszKey, wpKey, lpIniKey->nKeyBytes);

      lpIniKey->wszString=NULL;
      lpIniKey->nStringBytes=0;
    }
  }
  return lpIniKey;
}

int StackGetIniData(INIKEY *lpIniKey, int nType, unsigned char *lpData, DWORD dwDataBytes)
{
  DWORD dwStringLen;

  if (nType == INI_DWORD)
  {
    if (lpData)
    {
      DWORD dwNumber;

      dwNumber=(DWORD)xatoiW(lpIniKey->wszString, NULL);
      dwStringLen=min(sizeof(DWORD), dwDataBytes);
      xmemcpy((DWORD *)lpData, &dwNumber, dwStringLen);
      return dwStringLen;
    }
    return sizeof(DWORD);
  }
  else if (nType == INI_BINARY)
  {
    return (int)hex2binW(lpIniKey->wszString, lpData, dwDataBytes);
  }
  else if (nType == INI_STRINGANSI)
  {
    if (lpData)
    {
      dwStringLen=WideCharToMultiByte(CP_ACP, 0, lpIniKey->wszString, lpIniKey->nStringBytes / sizeof(wchar_t), (char *)lpData, dwDataBytes / sizeof(wchar_t), NULL, NULL);
      ((char *)lpData)[dwStringLen - 1]='\0';
      return dwStringLen;
    }
    return WideCharToMultiByte(CP_ACP, 0, lpIniKey->wszString, lpIniKey->nStringBytes / sizeof(wchar_t), NULL, 0, NULL, NULL);
  }
  else if (nType == INI_STRINGUNICODE)
  {
    if (lpData)
    {
      dwStringLen=min((DWORD)lpIniKey->nStringBytes, dwDataBytes);
      dwStringLen=(dwStringLen / sizeof(wchar_t)) * sizeof(wchar_t);
      xmemcpy((wchar_t *)lpData, lpIniKey->wszString, dwStringLen);
      ((wchar_t *)lpData)[dwStringLen / sizeof(wchar_t) - 1]=L'\0';
      return dwStringLen;
    }
    return lpIniKey->nStringBytes;
  }
  return 0;
}

BOOL StackSetIniData(INIKEY *lpIniKey, int nType, unsigned char *lpData, DWORD dwDataBytes)
{
  wchar_t *wszString;
  int nStringBytes;
  BOOL bResult=FALSE;

  if (nType == INI_DWORD)
  {
    wchar_t wszNumber[16];

    nStringBytes=(xuitoaW(*(DWORD *)lpData, wszNumber) + 1) * sizeof(wchar_t);
    if (wszString=(wchar_t *)API_HeapAlloc(hHeap, 0, nStringBytes))
    {
      xmemcpy(wszString, wszNumber, nStringBytes);
      bResult=TRUE;
    }
  }
  else if (nType == INI_BINARY)
  {
    nStringBytes=(dwDataBytes * 2 + 1) * sizeof(wchar_t);
    if (wszString=(wchar_t *)API_HeapAlloc(hHeap, 0, nStringBytes))
    {
      bin2hexW(lpData, dwDataBytes, wszString, nStringBytes / sizeof(wchar_t), FALSE);
      bResult=TRUE;
    }
  }
  else if (nType == INI_STRINGANSI)
  {
    dwDataBytes=max(dwDataBytes, sizeof(char));

    nStringBytes=MultiByteToWideChar(CP_ACP, 0, (char *)lpData, dwDataBytes, NULL, 0) * sizeof(wchar_t);
    if (wszString=(wchar_t *)API_HeapAlloc(hHeap, 0, nStringBytes))
    {
      MultiByteToWideChar(CP_ACP, 0, (char *)lpData, dwDataBytes, wszString, nStringBytes / sizeof(wchar_t));
      if (wszString[nStringBytes / sizeof(wchar_t) - 1] != L'\0')
        wszString[nStringBytes / sizeof(wchar_t) - 1]=L'\0';
      bResult=TRUE;
    }
  }
  else if (nType == INI_STRINGUNICODE)
  {
    dwDataBytes=max(dwDataBytes, sizeof(wchar_t));

    nStringBytes=(dwDataBytes / sizeof(wchar_t)) * sizeof(wchar_t);
    if (wszString=(wchar_t *)API_HeapAlloc(hHeap, 0, nStringBytes))
    {
      xmemcpy(wszString, (wchar_t *)lpData, nStringBytes);
      if (wszString[nStringBytes / sizeof(wchar_t) - 1] != L'\0')
        wszString[nStringBytes / sizeof(wchar_t) - 1]=L'\0';
      bResult=TRUE;
    }
  }

  if (bResult)
  {
    if (lpIniKey->wszString)
    {
      if (nStringBytes == lpIniKey->nStringBytes && !xstrcmpW(wszString, lpIniKey->wszString))
      {
        API_HeapFree(hHeap, 0, (LPVOID)wszString);
        return FALSE;
      }
      else API_HeapFree(hHeap, 0, (LPVOID)lpIniKey->wszString);
    }
    lpIniKey->wszString=wszString;
    lpIniKey->nStringBytes=nStringBytes;
  }
  return bResult;
}

BOOL StackDeleteIniKey(INISECTION *lpIniSection, INIKEY *lpIniKey)
{
  if (lpIniKey)
  {
    if (lpIniKey->wszKey) API_HeapFree(hHeap, 0, (LPVOID)lpIniKey->wszKey);
    if (lpIniKey->wszString) API_HeapFree(hHeap, 0, (LPVOID)lpIniKey->wszString);

    StackDelete((stack **)&lpIniSection->first, (stack **)&lpIniSection->last, (stack *)lpIniKey);
    ((INIFILE *)lpIniSection->hIniFile)->bModified=TRUE;
    return TRUE;
  }
  return FALSE;
}

BOOL StackDeleteIniSection(INIFILE *hIniFile, INISECTION *lpIniSection, BOOL bOnlyClear)
{
  if (lpIniSection)
  {
    INIKEY *lpIniKey=(INIKEY *)lpIniSection->first;

    while (lpIniKey)
    {
      if (lpIniKey->wszKey) API_HeapFree(hHeap, 0, (LPVOID)lpIniKey->wszKey);
      if (lpIniKey->wszString) API_HeapFree(hHeap, 0, (LPVOID)lpIniKey->wszString);

      lpIniKey=lpIniKey->next;
    }
    StackClear((stack **)&lpIniSection->first, (stack **)&lpIniSection->last);

    if (!bOnlyClear)
    {
      API_HeapFree(hHeap, 0, (LPVOID)lpIniSection->wszSection);
      StackDelete((stack **)&hIniFile->first, (stack **)&hIniFile->last, (stack *)lpIniSection);
    }
    hIniFile->bModified=TRUE;
    return TRUE;
  }
  return FALSE;
}

void StackFreeIni(INIFILE *hIniFile)
{
  INISECTION *lpIniSection=(INISECTION *)hIniFile->first;

  while (lpIniSection)
  {
    API_HeapFree(hHeap, 0, (LPVOID)lpIniSection->wszSection);
    StackDeleteIniSection(hIniFile, lpIniSection, TRUE);

    lpIniSection=lpIniSection->next;
  }
  StackClear((stack **)&hIniFile->first, (stack **)&hIniFile->last);
}

INIKEY* IniOpenKeyA(INIFILE *hIniFile, const char *pSection, const char *pKey, BOOL bCreate, INISECTION **lppIniSection)
{
  INISECTION *lpIniSection;

  if (!(lpIniSection=StackOpenIniSectionA(hIniFile, pSection, lstrlenA(pSection), bCreate)))
    return 0;
  if (lppIniSection)
    *lppIniSection=lpIniSection;
  return StackOpenIniKeyA(lpIniSection, pKey, lstrlenA(pKey), bCreate);
}

INIKEY* IniOpenKeyW(INIFILE *hIniFile, const wchar_t *wpSection, const wchar_t *wpKey, BOOL bCreate, INISECTION **lppIniSection)
{
  INISECTION *lpIniSection;

  if (!(lpIniSection=StackOpenIniSectionW(hIniFile, wpSection, (int)xstrlenW(wpSection), bCreate)))
    return 0;
  if (lppIniSection)
    *lppIniSection=lpIniSection;
  return StackOpenIniKeyW(lpIniSection, wpKey, (int)xstrlenW(wpKey), bCreate);
}

int IniGetValueA(INIFILE *hIniFile, const char *pSection, const char *pKey, int nType, unsigned char *lpData, DWORD dwDataBytes)
{
  INIKEY *lpIniKey;

  if (!(lpIniKey=IniOpenKeyA(hIniFile, pSection, pKey, FALSE, NULL)))
    return 0;
  return StackGetIniData(lpIniKey, nType, lpData, dwDataBytes);
}

int IniGetValueW(INIFILE *hIniFile, const wchar_t *wpSection, const wchar_t *wpKey, int nType, unsigned char *lpData, DWORD dwDataBytes)
{
  INIKEY *lpIniKey;

  if (!(lpIniKey=IniOpenKeyW(hIniFile, wpSection, wpKey, FALSE, NULL)))
    return 0;
  return StackGetIniData(lpIniKey, nType, lpData, dwDataBytes);
}

BOOL IniSetValueA(INIFILE *hIniFile, const char *pSection, const char *pKey, int nType, unsigned char *lpData, DWORD dwDataBytes)
{
  INIKEY *lpIniKey;

  if (!(lpIniKey=IniOpenKeyA(hIniFile, pSection, pKey, TRUE, NULL)))
    return FALSE;
  if (StackSetIniData(lpIniKey, nType, lpData, dwDataBytes))
    hIniFile->bModified=TRUE;
  return TRUE;
}

BOOL IniSetValueW(INIFILE *hIniFile, const wchar_t *wpSection, const wchar_t *wpKey, int nType, unsigned char *lpData, DWORD dwDataBytes)
{
  INIKEY *lpIniKey;

  if (!(lpIniKey=IniOpenKeyW(hIniFile, wpSection, wpKey, TRUE, NULL)))
    return FALSE;
  if (StackSetIniData(lpIniKey, nType, lpData, dwDataBytes))
    hIniFile->bModified=TRUE;
  return TRUE;
}

BOOL IniDelKeyA(INIFILE *hIniFile, const char *pSection, const char *pKey)
{
  INISECTION *lpIniSection;
  INIKEY *lpIniKey;

  if (!(lpIniKey=IniOpenKeyA(hIniFile, pSection, pKey, FALSE, &lpIniSection)))
    return FALSE;
  return StackDeleteIniKey(lpIniSection, lpIniKey);
}

BOOL IniDelKeyW(INIFILE *hIniFile, const wchar_t *wpSection, const wchar_t *wpKey)
{
  INISECTION *lpIniSection;
  INIKEY *lpIniKey;

  if (!(lpIniKey=IniOpenKeyW(hIniFile, wpSection, wpKey, FALSE, &lpIniSection)))
    return FALSE;
  return StackDeleteIniKey(lpIniSection, lpIniKey);
}


//// Read/Save options

DWORD ReadOption(OPTIONHANDLE *oh, const wchar_t *wpParam, DWORD dwType, void *lpData, DWORD dwSize)
{
  if (oh->nSaveSettings == SS_REGISTRY)
  {
    if (dwType & MOT_DWORD) dwType=REG_DWORD;
    else if (dwType & MOT_BINARY) dwType=REG_BINARY;
    else if (dwType & MOT_STRING) dwType=REG_SZ;

    if (RegQueryValueExWide((HKEY)oh->hHandle, wpParam, NULL, &dwType, (LPBYTE)lpData, &dwSize) == ERROR_SUCCESS)
      return dwSize;
    else
      return 0;
  }
  else
  {
    INIKEY *lpIniKey;

    if (dwType & MOT_DWORD) dwType=INI_DWORD;
    else if (dwType & MOT_BINARY) dwType=INI_BINARY;
    else if (dwType & MOT_STRING) dwType=INI_STRINGUNICODE;

    if (lpIniKey=StackOpenIniKeyW((INISECTION *)oh->hHandle, wpParam, (int)xstrlenW(wpParam), FALSE))
      return StackGetIniData(lpIniKey, dwType, (LPBYTE)lpData, dwSize);
    else
      return 0;
  }
}

DWORD SaveOption(OPTIONHANDLE *oh, const wchar_t *wpParam, DWORD dwType, void *lpData, DWORD dwSize)
{
  if (dwType & MOT_MAINOFFSET)
  {
    if (!oh->bForceWrite)
    {
      if ((dwType & MOT_MANUAL) && !ReadOption(oh, wpParam, dwType, NULL, 0))
      {
        //lpData is manual parameter. Save it if not saved before.
      }
      else
      {
        //lpData is MAINOPTIONS structure offset. Check if specified member in lpData has been changed.
        if (!xmemcmp(((LPBYTE)oh->mo) + (UINT_PTR)lpData, ((LPBYTE)&moInit) + (UINT_PTR)lpData, dwSize))
          return dwSize;
      }
    }
    lpData=((LPBYTE)oh->mo) + (UINT_PTR)lpData;
  }
  else if (dwType & MOT_FRAMEOFFSET)
  {
    if (!oh->bForceWrite)
    {
      //lpData is FRAMEDATA structure offset. Check if specified member in lpData has been changed.
      if (!xmemcmp(((LPBYTE)oh->fd) + (UINT_PTR)lpData, ((LPBYTE)&fdInit) + (UINT_PTR)lpData, dwSize))
        return dwSize;
    }
    lpData=((LPBYTE)oh->fd) + (UINT_PTR)lpData;
  }

  if (oh->nSaveSettings == SS_REGISTRY)
  {
    if (dwType & MOT_DWORD) dwType=REG_DWORD;
    else if (dwType & MOT_BINARY) dwType=REG_BINARY;
    else if (dwType & MOT_STRING) dwType=REG_SZ;

    return !RegSetValueExWide((HKEY)oh->hHandle, wpParam, 0, dwType, (LPBYTE)lpData, dwSize);
  }
  else
  {
    INIKEY *lpIniKey;

    if (dwType & MOT_DWORD) dwType=INI_DWORD;
    else if (dwType & MOT_BINARY) dwType=INI_BINARY;
    else if (dwType & MOT_STRING) dwType=INI_STRINGUNICODE;

    if (lpIniKey=StackOpenIniKeyW((INISECTION *)oh->hHandle, wpParam, (int)xstrlenW(wpParam), TRUE))
    {
      StackSetIniData(lpIniKey, dwType, (LPBYTE)lpData, dwSize);
      return dwSize;
    }
    else return 0;
  }
}

void ReadOptions(MAINOPTIONS *mo, FRAMEDATA *fd)
{
  OPTIONHANDLE oh;
  DWORD dwSize;

  xprintfW(wszIniFile, L"%s\\AkelPad.ini", wszExeDir);
  if (OpenIniW(&hIniFile, wszIniFile, FALSE))
    IniGetValueW(&hIniFile, L"Options", L"SaveSettings", INI_DWORD, (LPBYTE)&mo->nSaveSettings, sizeof(DWORD));
  oh.mo=mo;
  oh.fd=fd;
  oh.nSaveSettings=mo->nSaveSettings;

  if (oh.nSaveSettings == SS_REGISTRY)
  {
    wchar_t wszRegKey[MAX_PATH];

    xprintfW(wszRegKey, L"%s\\Options", APP_REGHOMEW);
    if (RegOpenKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, KEY_READ, (HKEY *)&oh.hHandle) != ERROR_SUCCESS)
      oh.hHandle=NULL;
  }
  else
  {
    if (!(oh.hHandle=(HANDLE)StackOpenIniSectionW(&hIniFile, L"Options", (int)xstrlenW(L"Options"), FALSE)))
      oh.hHandle=NULL;
  }

  if (oh.hHandle)
  {
    //Manual
    if (dwSize=ReadOption(&oh, L"CmdLineBegin", MOT_STRING, NULL, 0))
    {
      if (wpCmdLineBegin=(wchar_t *)API_HeapAlloc(hHeap, 0, dwSize))
      {
        ReadOption(&oh, L"CmdLineBegin", MOT_STRING, wpCmdLineBegin, dwSize);
        nCmdLineBeginLen=dwSize / sizeof(wchar_t) - 1;
      }
    }
    else bSaveManual=TRUE;

    if (dwSize=ReadOption(&oh, L"CmdLineEnd", MOT_STRING, NULL, 0))
    {
      if (wpCmdLineEnd=(wchar_t *)API_HeapAlloc(hHeap, 0, dwSize))
      {
        ReadOption(&oh, L"CmdLineEnd", MOT_STRING, wpCmdLineEnd, dwSize);
        nCmdLineEndLen=dwSize / sizeof(wchar_t) - 1;
      }
    }
    else bSaveManual=TRUE;

    if (!ReadOption(&oh, L"ShowModify", MOT_DWORD, &mo->dwShowModify, sizeof(DWORD)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"StatusPosType", MOT_DWORD, &mo->dwStatusPosType, sizeof(DWORD)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"StatusUserFormat", MOT_STRING, mo->wszStatusUserFormat, sizeof(mo->wszStatusUserFormat)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"WordBreak", MOT_DWORD, &mo->dwWordBreakCustom, sizeof(DWORD)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"PaintOptions", MOT_DWORD, &mo->dwPaintOptions, sizeof(DWORD)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"EditStyle", MOT_DWORD, &mo->dwEditStyle, sizeof(DWORD)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"RichEditClass", MOT_DWORD, &mo->bRichEditClass, sizeof(DWORD)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"AkelAdminResident", MOT_DWORD, &mo->bAkelAdminResident, sizeof(DWORD)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"DateLogFormat", MOT_STRING, mo->wszDateLogFormat, sizeof(mo->wszDateLogFormat)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"DateInsertFormat", MOT_STRING, mo->wszDateInsertFormat, sizeof(mo->wszDateInsertFormat)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"AkelUpdaterOptions", MOT_STRING, mo->wszAkelUpdaterOptions, sizeof(mo->wszAkelUpdaterOptions)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"UrlCommand", MOT_STRING, mo->wszUrlCommand, sizeof(mo->wszUrlCommand)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"TabNameFind", MOT_STRING, mo->wszTabNameFind, sizeof(mo->wszTabNameFind)))
      bSaveManual=TRUE;
    if (!ReadOption(&oh, L"TabNameRep", MOT_STRING, mo->wszTabNameRep, sizeof(mo->wszTabNameRep)))
      bSaveManual=TRUE;

    //Frame data
    ReadOption(&oh, L"TabStopSize", MOT_DWORD, &fd->nTabStopSize, sizeof(DWORD));
    ReadOption(&oh, L"TabStopAsSpaces", MOT_DWORD, &fd->bTabStopAsSpaces, sizeof(DWORD));
    ReadOption(&oh, L"UndoLimit", MOT_DWORD, &fd->nUndoLimit, sizeof(DWORD));
    ReadOption(&oh, L"DetailedUndo", MOT_DWORD, &fd->bDetailedUndo, sizeof(DWORD));
    ReadOption(&oh, L"WordWrap", MOT_DWORD, &fd->ei.bWordWrap, sizeof(DWORD));
    ReadOption(&oh, L"WrapType", MOT_DWORD, &fd->dwWrapType, sizeof(DWORD));
    ReadOption(&oh, L"WrapLimit", MOT_DWORD, &fd->dwWrapLimit, sizeof(DWORD));
    ReadOption(&oh, L"Marker", MOT_DWORD, &fd->dwMarker, sizeof(DWORD));
    ReadOption(&oh, L"CaretOptions", MOT_DWORD, &fd->dwCaretOptions, sizeof(DWORD));
    ReadOption(&oh, L"CaretWidth", MOT_DWORD, &fd->nCaretWidth, sizeof(DWORD));
    ReadOption(&oh, L"AltLineSkip", MOT_DWORD, &fd->dwAltLineSkip, sizeof(DWORD));
    ReadOption(&oh, L"AltLineFill", MOT_DWORD, &fd->dwAltLineFill, sizeof(DWORD));
    ReadOption(&oh, L"AltLineBorder", MOT_DWORD, &fd->bAltLineBorder, sizeof(DWORD));
    ReadOption(&oh, L"MarginsEditRect", MOT_BINARY, &fd->rcEditMargins, sizeof(RECT));
    ReadOption(&oh, L"LineGap", MOT_DWORD, &fd->dwLineGap, sizeof(DWORD));
    ReadOption(&oh, L"MouseOptions", MOT_DWORD, &fd->dwMouseOptions, sizeof(DWORD));
    ReadOption(&oh, L"ShowURL", MOT_DWORD, &fd->bShowURL, sizeof(DWORD));
    ReadOption(&oh, L"ClickURL", MOT_DWORD, &fd->nClickURL, sizeof(DWORD));
    ReadOption(&oh, L"UrlPrefixesEnable", MOT_DWORD, &fd->bUrlPrefixesEnable, sizeof(DWORD));
    ReadOption(&oh, L"UrlPrefixes", MOT_BINARY, fd->wszUrlPrefixes, sizeof(fd->wszUrlPrefixes));
    ReadOption(&oh, L"UrlDelimitersEnable", MOT_DWORD, &fd->bUrlDelimitersEnable, sizeof(DWORD));
    ReadOption(&oh, L"UrlLeftDelimiters", MOT_BINARY, fd->wszUrlLeftDelimiters, sizeof(fd->wszUrlLeftDelimiters));
    ReadOption(&oh, L"UrlRightDelimiters", MOT_BINARY, fd->wszUrlRightDelimiters, sizeof(fd->wszUrlRightDelimiters));
    ReadOption(&oh, L"WordDelimitersEnable", MOT_DWORD, &fd->bWordDelimitersEnable, sizeof(DWORD));
    ReadOption(&oh, L"WordDelimiters", MOT_BINARY, fd->wszWordDelimiters, sizeof(fd->wszWordDelimiters));
    ReadOption(&oh, L"WrapDelimitersEnable", MOT_DWORD, &fd->bWrapDelimitersEnable, sizeof(DWORD));
    ReadOption(&oh, L"WrapDelimiters", MOT_BINARY, fd->wszWrapDelimiters, sizeof(fd->wszWrapDelimiters));
    ReadOption(&oh, L"Font", MOT_BINARY, &fd->lf, offsetof(LOGFONTW, lfFaceName));
    ReadOption(&oh, L"FontFace", MOT_STRING, fd->lf.lfFaceName, sizeof(fd->lf.lfFaceName));
    ReadOption(&oh, L"Colors", MOT_BINARY, &fd->aec, sizeof(AECOLORS));
    ReadOption(&oh, L"BkImageFile", MOT_STRING, &fd->wszBkImageFile, sizeof(fd->wszBkImageFile));
    ReadOption(&oh, L"BkImageAlpha", MOT_DWORD, &fd->nBkImageAlpha, sizeof(DWORD));

    //Menu settings
    ReadOption(&oh, L"OnTop", MOT_DWORD, &mo->bOnTop, sizeof(DWORD));
    ReadOption(&oh, L"StatusBar", MOT_DWORD, &mo->bStatusBar, sizeof(DWORD));
    ReadOption(&oh, L"LanguageModule", MOT_STRING, mo->wszLangModule, sizeof(mo->wszLangModule));
    ReadOption(&oh, L"KeepSpace", MOT_DWORD, &mo->bKeepSpace, sizeof(DWORD));
    ReadOption(&oh, L"WatchFile", MOT_DWORD, &mo->bWatchFile, sizeof(DWORD));
    ReadOption(&oh, L"SaveTime", MOT_DWORD, &mo->bSaveTime, sizeof(DWORD));
    ReadOption(&oh, L"SingleOpenFile", MOT_DWORD, &mo->bSingleOpenFile, sizeof(DWORD));
    ReadOption(&oh, L"SingleOpenProgram", MOT_DWORD, &mo->dwSingleOpenProgram, sizeof(DWORD));
    ReadOption(&oh, L"MDI", MOT_DWORD, &mo->nMDI, sizeof(DWORD));
    if (mo->nMDI)
    {
      ReadOption(&oh, L"TabOptionsMDI", MOT_DWORD, &mo->dwTabOptionsMDI, sizeof(DWORD));
    }

    //Settings dialog
    if (dwSize=ReadOption(&oh, L"CodepageList", MOT_BINARY, NULL, 0))
    {
      if (lpCodepageList=(int *)API_HeapAlloc(hHeap, 0, dwSize))
      {
        ReadOption(&oh, L"CodepageList", MOT_BINARY, lpCodepageList, dwSize);
        nCodepageListLen=CodepageListLen(lpCodepageList);
      }
    }
    ReadOption(&oh, L"ExecuteCommand", MOT_STRING, mo->wszExecuteCommand, sizeof(mo->wszExecuteCommand));
    ReadOption(&oh, L"ExecuteDirectory", MOT_STRING, mo->wszExecuteDirectory, sizeof(mo->wszExecuteDirectory));
    ReadOption(&oh, L"DefaultCodepage", MOT_DWORD, &mo->nDefaultCodePage, sizeof(DWORD));
    ReadOption(&oh, L"NewFileCodepage", MOT_DWORD, &mo->nNewFileCodePage, sizeof(DWORD));
    ReadOption(&oh, L"NewFileBOM", MOT_DWORD, &mo->bNewFileBOM, sizeof(DWORD));
    ReadOption(&oh, L"NewFileNewLine", MOT_DWORD, &mo->nNewFileNewLine, sizeof(DWORD));
    ReadOption(&oh, L"CodepageRecognition", MOT_DWORD, &mo->dwLangCodepageRecognition, sizeof(DWORD));
    ReadOption(&oh, L"CodepageRecognitionBuffer", MOT_DWORD, &mo->dwCodepageRecognitionBuffer, sizeof(DWORD));
    ReadOption(&oh, L"SavePositions", MOT_DWORD, &mo->bSavePositions, sizeof(DWORD));
    ReadOption(&oh, L"SaveCodepages", MOT_DWORD, &mo->bSaveCodepages, sizeof(DWORD));
    ReadOption(&oh, L"RecentFiles", MOT_DWORD, &mo->nRecentFiles, sizeof(DWORD));
    ReadOption(&oh, L"SearchStrings", MOT_DWORD, &mo->nSearchStrings, sizeof(DWORD));
    ReadOption(&oh, L"FileTypesOpen", MOT_STRING, mo->wszFileTypesOpen, sizeof(mo->wszFileTypesOpen));
    ReadOption(&oh, L"FileTypesEdit", MOT_STRING, mo->wszFileTypesEdit, sizeof(mo->wszFileTypesEdit));
    ReadOption(&oh, L"FileTypesPrint", MOT_STRING, mo->wszFileTypesPrint, sizeof(mo->wszFileTypesPrint));
    ReadOption(&oh, L"FileTypesAssociated", MOT_DWORD, &mo->dwFileTypesAssociated, sizeof(DWORD));
    ReadOption(&oh, L"KeybLayoutOptions", MOT_DWORD, &mo->dwKeybLayoutOptions, sizeof(DWORD));
    if (mo->nMDI)
      ReadOption(&oh, L"SilentCloseEmptyMDI", MOT_DWORD, &mo->bSilentCloseEmptyMDI, sizeof(DWORD));
    ReadOption(&oh, L"DateLog", MOT_DWORD, &mo->bDateLog, sizeof(DWORD));
    ReadOption(&oh, L"SaveInReadOnlyMsg", MOT_DWORD, &mo->bSaveInReadOnlyMsg, sizeof(DWORD));
    ReadOption(&oh, L"DefaultSaveExt", MOT_STRING, mo->wszDefaultSaveExt, sizeof(mo->wszDefaultSaveExt));

    //Search dialog
    ReadOption(&oh, L"SearchOptions", MOT_DWORD, &mo->dwSearchOptions, sizeof(DWORD));

    //Open file dialog
    ReadOption(&oh, L"LastDirectory", MOT_STRING, mo->wszLastDir, sizeof(mo->wszLastDir));
    ReadOption(&oh, L"ShowPlacesBar", MOT_DWORD, &mo->bShowPlacesBar, sizeof(DWORD));

    //Print dialog
    ReadOption(&oh, L"MarginsPrint", MOT_BINARY, &mo->rcPrintMargins, sizeof(RECT));
    ReadOption(&oh, L"PrintColor", MOT_DWORD, &mo->dwPrintColor, sizeof(DWORD));
    ReadOption(&oh, L"PrintHeaderEnable", MOT_DWORD, &mo->bPrintHeaderEnable, sizeof(DWORD));
    ReadOption(&oh, L"PrintHeader", MOT_STRING, mo->wszPrintHeader, sizeof(mo->wszPrintHeader));
    ReadOption(&oh, L"PrintFooterEnable", MOT_DWORD, &mo->bPrintFooterEnable, sizeof(DWORD));
    ReadOption(&oh, L"PrintFooter", MOT_STRING, mo->wszPrintFooter, sizeof(mo->wszPrintFooter));
    ReadOption(&oh, L"PrintFontEnable", MOT_DWORD, &mo->bPrintFontEnable, sizeof(DWORD));
    ReadOption(&oh, L"PrintFont", MOT_BINARY, &mo->lfPrintFont, offsetof(LOGFONTW, lfFaceName));
    ReadOption(&oh, L"PrintFontFace", MOT_STRING, mo->lfPrintFont.lfFaceName, sizeof(mo->lfPrintFont.lfFaceName));

    //Recode dialog
    ReadOption(&oh, L"RecodeDialog", MOT_BINARY, &mo->rcRecodeCurrentDialog, sizeof(RECT));

    //Colors dialog
    ReadOption(&oh, L"ColorsDialog", MOT_BINARY, &mo->rcColorsCurrentDialog, sizeof(RECT));

    //Plugin dialog
    ReadOption(&oh, L"PluginsDialog", MOT_BINARY, &mo->rcPluginsCurrentDialog, sizeof(RECT));

    //Mdi list dialog
    if (mo->nMDI)
    {
      ReadOption(&oh, L"MdiListOptions", MOT_DWORD, &mo->dwMdiListOptions, sizeof(DWORD));
      ReadOption(&oh, L"MdiListDialog", MOT_BINARY, &mo->rcMdiListCurrentDialog, sizeof(RECT));
    }

    //Main window
    ReadOption(&oh, L"WindowStyle", MOT_DWORD, &mo->dwMainStyle, sizeof(DWORD));
    ReadOption(&oh, L"WindowPosition", MOT_BINARY, &mo->rcMainWindowRestored, sizeof(RECT));
    if (mo->nMDI)
    {
      ReadOption(&oh, L"WindowStyleMDI", MOT_DWORD, &mo->dwMdiStyle, sizeof(DWORD));
    }

    //Close handle
    if (oh.nSaveSettings == SS_REGISTRY)
    {
      RegCloseKey((HKEY)oh.hHandle);
      oh.hHandle=NULL;
    }
  }

  //Read and register plugins hotkeys
  RegisterPluginsHotkeys(mo);

  //Read theams
  ReadThemes(mo);

  //Free INI
  StackFreeIni(&hIniFile);
}

BOOL SaveOptions(MAINOPTIONS *mo, FRAMEDATA *fd, int nSaveSettings, BOOL bForceWrite)
{
  OPTIONHANDLE oh;
  BOOL bResult=FALSE;

  //Main window style
  mo->dwMainStyle=(DWORD)GetWindowLongPtrWide(hMainWnd, GWL_STYLE);
  mo->dwMainStyle=((mo->dwMainStyle & WS_MAXIMIZE) || ((mo->dwMainStyle & WS_MINIMIZE) && dwLastMainSizeType == SIZE_MAXIMIZED))?WS_MAXIMIZE:0;

  //MDI frame style
  if (nMDI == WMD_MDI)
  {
    if (lpFrameCurrent->hWndEditParent) SendMessage(hMdiClient, WM_MDIGETACTIVE, 0, (LPARAM)&bMdiMaximize);
    mo->dwMdiStyle=!bMdiMaximize?0:WS_MAXIMIZE;
  }

  if (!bForceWrite)
  {
    if (!xmemcmp(mo, &moInit, sizeof(MAINOPTIONS)))
    {
      if (!xmemcmp(&fd->lf, &fdInit.lf, offsetof(FRAMEDATA, lpEditProc) - offsetof(FRAMEDATA, lf)))
      {
        if (fd->ei.bWordWrap == fdInit.ei.bWordWrap)
        {
          if (!bCodepageListChanged && !bCmdLineChanged && !bSaveManual)
          {
            //Settings unchanged
            return TRUE;
          }
        }
      }
    }
  }
  oh.mo=mo;
  oh.fd=fd;
  oh.nSaveSettings=nSaveSettings;
  oh.bForceWrite=bForceWrite;

  if (oh.nSaveSettings == SS_REGISTRY)
  {
    wchar_t wszRegKey[MAX_PATH];

    xprintfW(wszRegKey, L"%s\\Options", APP_REGHOMEW);
    if (RegCreateKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, (HKEY *)&oh.hHandle, NULL) != ERROR_SUCCESS)
      return FALSE;
  }
  else
  {
    if (!OpenIniW(&hIniFile, wszIniFile, TRUE))
      return FALSE;
    if (!(oh.hHandle=(HANDLE)StackOpenIniSectionW(&hIniFile, L"Options", (int)xstrlenW(L"Options"), TRUE)))
      goto Error;

    if (!SaveOption(&oh, L"SaveSettings", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, nSaveSettings), sizeof(DWORD)))
      goto Error;
  }

  //Manual
  if (bForceWrite || bCmdLineChanged ||
     !ReadOption(&oh, L"CmdLineBegin", MOT_STRING, NULL, 0) ||
     !ReadOption(&oh, L"CmdLineEnd", MOT_STRING, NULL, 0))
  {
    bCmdLineChanged=FALSE;
    if (!SaveOption(&oh, L"CmdLineBegin", MOT_STRING, (void *)wpCmdLineBegin, (nCmdLineBeginLen + 1) * sizeof(wchar_t)))
      goto Error;
    if (!SaveOption(&oh, L"CmdLineEnd", MOT_STRING, (void *)wpCmdLineEnd, (nCmdLineEndLen + 1) * sizeof(wchar_t)))
      goto Error;
  }
  if (!SaveOption(&oh, L"ShowModify", MOT_DWORD|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, dwShowModify), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"StatusPosType", MOT_DWORD|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, dwStatusPosType), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"StatusUserFormat", MOT_STRING|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, wszStatusUserFormat), BytesInString(mo->wszStatusUserFormat)))
    goto Error;
  if (!SaveOption(&oh, L"WordBreak", MOT_DWORD|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, dwWordBreakCustom), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"PaintOptions", MOT_DWORD|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, dwPaintOptions), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"EditStyle", MOT_DWORD|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, dwEditStyle), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"RichEditClass", MOT_DWORD|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, bRichEditClass), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"AkelAdminResident", MOT_DWORD|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, bAkelAdminResident), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"DateLogFormat", MOT_STRING|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, wszDateLogFormat), BytesInString(mo->wszDateLogFormat)))
    goto Error;
  if (!SaveOption(&oh, L"DateInsertFormat", MOT_STRING|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, wszDateInsertFormat), BytesInString(mo->wszDateInsertFormat)))
    goto Error;
  if (!SaveOption(&oh, L"AkelUpdaterOptions", MOT_STRING|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, wszAkelUpdaterOptions), BytesInString(mo->wszAkelUpdaterOptions)))
    goto Error;
  if (!SaveOption(&oh, L"UrlCommand", MOT_STRING|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, wszUrlCommand), BytesInString(mo->wszUrlCommand)))
    goto Error;
  if (!SaveOption(&oh, L"TabNameFind", MOT_STRING|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, wszTabNameFind), BytesInString(mo->wszTabNameFind)))
    goto Error;
  if (!SaveOption(&oh, L"TabNameRep", MOT_STRING|MOT_MAINOFFSET|MOT_MANUAL, (void *)offsetof(MAINOPTIONS, wszTabNameRep), BytesInString(mo->wszTabNameRep)))
    goto Error;

  bSaveManual=FALSE;

  //Frame data
  if (!SaveOption(&oh, L"TabStopSize", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, nTabStopSize), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"TabStopAsSpaces", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, bTabStopAsSpaces), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"UndoLimit", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, nUndoLimit), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"DetailedUndo", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, bDetailedUndo), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"WordWrap", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, ei.bWordWrap), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"WrapType", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, dwWrapType), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"WrapLimit", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, dwWrapLimit), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"Marker", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, dwMarker), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"CaretOptions", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, dwCaretOptions), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"CaretWidth", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, nCaretWidth), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"AltLineSkip", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, dwAltLineSkip), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"AltLineFill", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, dwAltLineFill), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"AltLineBorder", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, bAltLineBorder), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"MarginsEditRect", MOT_BINARY|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, rcEditMargins), sizeof(RECT)))
    goto Error;
  if (!SaveOption(&oh, L"LineGap", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, dwLineGap), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"MouseOptions", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, dwMouseOptions), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"ShowURL", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, bShowURL), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"ClickURL", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, nClickURL), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"UrlPrefixesEnable", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, bUrlPrefixesEnable), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"UrlPrefixes", MOT_BINARY|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, wszUrlPrefixes), BytesInString(fd->wszUrlPrefixes)))
    goto Error;
  if (!SaveOption(&oh, L"UrlDelimitersEnable", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, bUrlDelimitersEnable), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"UrlLeftDelimiters", MOT_BINARY|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, wszUrlLeftDelimiters), BytesInString(fd->wszUrlLeftDelimiters)))
    goto Error;
  if (!SaveOption(&oh, L"UrlRightDelimiters", MOT_BINARY|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, wszUrlRightDelimiters), BytesInString(fd->wszUrlRightDelimiters)))
    goto Error;
  if (!SaveOption(&oh, L"WordDelimitersEnable", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, bWordDelimitersEnable), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"WordDelimiters", MOT_BINARY|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, wszWordDelimiters), BytesInString(fd->wszWordDelimiters)))
    goto Error;
  if (!SaveOption(&oh, L"WrapDelimitersEnable", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, bWrapDelimitersEnable), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"WrapDelimiters", MOT_BINARY|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, wszWrapDelimiters), BytesInString(fd->wszWrapDelimiters)))
    goto Error;
  if (!SaveOption(&oh, L"Font", MOT_BINARY|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, lf), offsetof(LOGFONTW, lfFaceName)))
    goto Error;
  if (!SaveOption(&oh, L"FontFace", MOT_STRING|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, lf.lfFaceName), BytesInString(fd->lf.lfFaceName)))
    goto Error;
  if (!SaveOption(&oh, L"Colors", MOT_BINARY|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, aec), sizeof(AECOLORS)))
    goto Error;
  if (!SaveOption(&oh, L"BkImageFile", MOT_STRING|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, wszBkImageFile), BytesInString(fd->wszBkImageFile)))
    goto Error;
  if (!SaveOption(&oh, L"BkImageAlpha", MOT_DWORD|MOT_FRAMEOFFSET, (void *)offsetof(FRAMEDATA, nBkImageAlpha), sizeof(DWORD)))
    goto Error;

  //Menu settings
  if (!SaveOption(&oh, L"OnTop", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bOnTop), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"StatusBar", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bStatusBar), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"LanguageModule", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszLangModule), BytesInString(mo->wszLangModule)))
    goto Error;
  if (!SaveOption(&oh, L"KeepSpace", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bKeepSpace), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"WatchFile", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bWatchFile), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"SaveTime", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bSaveTime), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"SingleOpenFile", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bSingleOpenFile), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"SingleOpenProgram", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwSingleOpenProgram), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"MDI", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, nMDI), sizeof(DWORD)))
    goto Error;
  if (nMDI)
  {
    if (!SaveOption(&oh, L"TabOptionsMDI", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwTabOptionsMDI), sizeof(DWORD)))
      goto Error;
  }

  //Settings dialog
  if (bForceWrite || bCodepageListChanged)
  {
    bCodepageListChanged=FALSE;
    if (!SaveOption(&oh, L"CodepageList", MOT_BINARY, lpCodepageList, nCodepageListLen * sizeof(int)))
      goto Error;
  }
  if (!SaveOption(&oh, L"ExecuteCommand", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszExecuteCommand), BytesInString(mo->wszExecuteCommand)))
    goto Error;
  if (!SaveOption(&oh, L"ExecuteDirectory", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszExecuteDirectory), BytesInString(mo->wszExecuteDirectory)))
    goto Error;
  if (!SaveOption(&oh, L"DefaultCodepage", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, nDefaultCodePage), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"NewFileCodepage", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, nNewFileCodePage), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"NewFileBOM", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bNewFileBOM), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"NewFileNewLine", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, nNewFileNewLine), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"CodepageRecognition", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwLangCodepageRecognition), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"CodepageRecognitionBuffer", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwCodepageRecognitionBuffer), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"SavePositions", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bSavePositions), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"SaveCodepages", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bSaveCodepages), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"RecentFiles", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, nRecentFiles), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"SearchStrings", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, nSearchStrings), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"FileTypesOpen", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszFileTypesOpen), BytesInString(mo->wszFileTypesOpen)))
    goto Error;
  if (!SaveOption(&oh, L"FileTypesEdit", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszFileTypesEdit), BytesInString(mo->wszFileTypesEdit)))
    goto Error;
  if (!SaveOption(&oh, L"FileTypesPrint", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszFileTypesPrint), BytesInString(mo->wszFileTypesPrint)))
    goto Error;
  if (!SaveOption(&oh, L"FileTypesAssociated", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwFileTypesAssociated), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"KeybLayoutOptions", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwKeybLayoutOptions), sizeof(DWORD)))
    goto Error;
  if (nMDI)
  {
    if (!SaveOption(&oh, L"SilentCloseEmptyMDI", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bSilentCloseEmptyMDI), sizeof(DWORD)))
      goto Error;
  }
  if (!SaveOption(&oh, L"DateLog", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bDateLog), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"SaveInReadOnlyMsg", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bSaveInReadOnlyMsg), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"DefaultSaveExt", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszDefaultSaveExt), BytesInString(mo->wszDefaultSaveExt)))
    goto Error;

  //Search dialog
  if (!SaveOption(&oh, L"SearchOptions", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwSearchOptions), sizeof(DWORD)))
    goto Error;

  //Open file dialog
  if (!SaveOption(&oh, L"LastDirectory", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszLastDir), BytesInString(mo->wszLastDir)))
    goto Error;
  if (!SaveOption(&oh, L"ShowPlacesBar", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bShowPlacesBar), sizeof(DWORD)))
    goto Error;

  //Print dialog
  if (!SaveOption(&oh, L"MarginsPrint", MOT_BINARY|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, rcPrintMargins), sizeof(RECT)))
    goto Error;
  if (!SaveOption(&oh, L"PrintColor", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwPrintColor), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"PrintHeaderEnable", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bPrintHeaderEnable), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"PrintHeader", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszPrintHeader), BytesInString(mo->wszPrintHeader)))
    goto Error;
  if (!SaveOption(&oh, L"PrintFooterEnable", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bPrintFooterEnable), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"PrintFooter", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, wszPrintFooter), BytesInString(mo->wszPrintFooter)))
    goto Error;
  if (!SaveOption(&oh, L"PrintFontEnable", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, bPrintFontEnable), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"PrintFont", MOT_BINARY|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, lfPrintFont), offsetof(LOGFONTW, lfFaceName)))
    goto Error;
  if (!SaveOption(&oh, L"PrintFontFace", MOT_STRING|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, lfPrintFont.lfFaceName), BytesInString(mo->lfPrintFont.lfFaceName)))
    goto Error;

  //Recode dialog
  if (!SaveOption(&oh, L"RecodeDialog", MOT_BINARY|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, rcRecodeCurrentDialog), sizeof(RECT)))
    goto Error;

  //Colors dialog
  if (!SaveOption(&oh, L"ColorsDialog", MOT_BINARY|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, rcColorsCurrentDialog), sizeof(RECT)))
    goto Error;

  //Plugin dialog
  if (!SaveOption(&oh, L"PluginsDialog", MOT_BINARY|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, rcPluginsCurrentDialog), sizeof(RECT)))
    goto Error;

  //Mdi list dialog
  if (nMDI)
  {
    if (!SaveOption(&oh, L"MdiListOptions", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwMdiListOptions), sizeof(DWORD)))
      goto Error;
    if (!SaveOption(&oh, L"MdiListDialog", MOT_BINARY|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, rcMdiListCurrentDialog), sizeof(RECT)))
      goto Error;
  }

  //Main window
  if (!SaveOption(&oh, L"WindowStyle", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwMainStyle), sizeof(DWORD)))
    goto Error;
  if (!SaveOption(&oh, L"WindowPosition", MOT_BINARY|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, rcMainWindowRestored), sizeof(RECT)))
    goto Error;
  if (nMDI)
  {
    if (!SaveOption(&oh, L"WindowStyleMDI", MOT_DWORD|MOT_MAINOFFSET, (void *)offsetof(MAINOPTIONS, dwMdiStyle), sizeof(DWORD)))
      goto Error;
  }

  if (oh.nSaveSettings == SS_REGISTRY)
    bResult=TRUE;
  else
    bResult=SaveIniW(&hIniFile, wszIniFile);

  if (bResult)
  {
    //Update initial options
    xmemcpy(&moInit, mo, sizeof(MAINOPTIONS));

    CopyFrameSettings(&fdInit, fd);
    CopyFrameSettings(&fdDefault, fd);
  }

  Error:
  if (oh.nSaveSettings == SS_REGISTRY)
    RegCloseKey((HKEY)oh.hHandle);
  else
    StackFreeIni(&hIniFile);
  return bResult;
}

BOOL SetOption(LPARAM lParam, void *lpData, int nDataSize, int nType)
{
  if (nType == INI_STRINGUNICODE)
  {
    if (xstrcmpiW((wchar_t *)lpData, (wchar_t *)lParam))
    {
      xstrcpynW((wchar_t *)lpData, (wchar_t *)lParam, nDataSize / sizeof(wchar_t));
      return TRUE;
    }
    return FALSE;
  }
  if (nType == INI_STRINGANSI)
  {
    if (xstrcmpiA((char *)lpData, (char *)lParam))
    {
      xstrcpynA((char *)lpData, (char *)lParam, nDataSize);
      return TRUE;
    }
    return FALSE;
  }

  if (xmemcmp(lpData, &lParam, nDataSize))
  {
    xmemcpy(lpData, &lParam, nDataSize);
    return TRUE;
  }
  return FALSE;
}


//// Open/Save document

int OpenDocument(HWND hWnd, const wchar_t *wpFile, DWORD dwFlags, int nCodePage, BOOL bBOM)
{
  wchar_t wszFile[MAX_PATH];
  HANDLE hFile;
  FILESTREAMDATA fsd;
  FRAMEDATA *lpFrame;
  RECENTFILE *lpRecentFile;
  HWND hWndFriend=NULL;
  int nResult=EOD_SUCCESS;
  int nDetect;
  int nFileCmp;
  int nFileLen;
  int nStreamOffset;
  int nChoice;
  DWORD dwMsgFlags;
  BOOL bFileExist=FALSE;

  if (!hWnd)
    hWnd=lpFrameCurrent->ei.hWndEdit;
  if (!hWnd)
  {
    //MDI has no windows
    DoFileNew();
    hWnd=lpFrameCurrent->ei.hWndEdit;
  }
  if (dwFlags & OD_MULTIFILE)
    dwMsgFlags=MB_YESNOCANCEL;
  else
    dwMsgFlags=MB_OKCANCEL;
  bFileExist=GetFullName(wpFile, wszFile, MAX_PATH, &nFileLen);
  nStreamOffset=GetFileStreamOffset(wszFile, nFileLen);
  if (IsEditActive(hWnd))
  {
    if (!nMDI && !(nFileCmp=xstrcmpiW(lpFrameCurrent->wszFile, wszFile)))
      dwFlags|=OD_REOPEN;
  }

  //Notification message
  if (GetWindowLongPtrWide(hWnd, GWLP_ID) == ID_EDIT)
  {
    NOPENDOCUMENT nod;
    char szFile[MAX_PATH];

    WideCharToMultiByte(CP_ACP, 0, wszFile, -1, szFile, MAX_PATH, NULL, NULL);
    nod.pFile=bOldWindows?(LPBYTE)szFile:(LPBYTE)wszFile;
    nod.szFile=szFile;
    nod.wszFile=wszFile;
    nod.nCodePage=&nCodePage;
    nod.bBOM=&bBOM;
    nod.dwFlags=&dwFlags;
    nod.bProcess=TRUE;
    SendMessage(hMainWnd, AKDN_OPENDOCUMENT_START, (WPARAM)hWnd, (LPARAM)&nod);
    if (!nod.bProcess)
    {
      nResult=EOD_STOP;
      goto End;
    }
  }
  fsd.bResult=-1;

  if (nDocumentsCount > MDI_MAXDOCUMENTS && nMDI == WMD_MDI)
  {
    API_LoadStringW(hLangLib, MSG_DOCUMENTSLIMIT, wszMsg, BUFFER_SIZE);
    API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
    nResult=EOD_DOCUMENTS_LIMIT;
    goto End;
  }
  if (!bFileExist)
  {
    //File doesn't exist
    if (dwCmdLineOptions & CLO_MSGOPENCREATEYES)
    {
    }
    else if (dwCmdLineOptions & CLO_MSGOPENCREATENO)
    {
      nResult=EOD_MSGNO;
      goto End;
    }
    else
    {
      API_LoadStringW(hLangLib, MSG_FILE_DOES_NOT_EXIST, wbuf, BUFFER_SIZE);
      xprintfW(wszMsg, wbuf, wszFile);
      nChoice=API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, dwMsgFlags|MB_ICONEXCLAMATION);
      if (nChoice == IDNO)
      {
        nResult=EOD_MSGNO;
        goto End;
      }
      else if (nChoice == IDCANCEL)
      {
        nResult=EOD_MSGCANCELCREATE;
        goto End;
      }
    }
    nCodePage=moCur.nNewFileCodePage;
    bBOM=moCur.bNewFileBOM;
  }
  else
  {
    if (IsEditActive(hWnd))
    {
      //File exists
      if (!(dwFlags & OD_REOPEN) && moCur.bSingleOpenFile)
      {
        if (!nMDI && nFileCmp)
        {
          if ((hWndFriend=FindWindowExWide(NULL, NULL, APP_SDI_CLASSW, wszFile)) &&
              (hWndFriend=GetParent(hWndFriend)))
          {
            ActivateWindow(hWndFriend);
            OpenDocumentSend(hWndFriend, NULL, wszFile, dwFlags|OD_REOPEN, nCodePage, bBOM, FALSE);
            nResult=EOD_WINDOW_EXIST;
            goto End;
          }
        }
        if (nMDI)
        {
          if (lpFrame=StackFrameGetByName(&hFramesStack, wszFile, -1))
          {
            ActivateFrameWindow(lpFrame, FWA_NOTIFY_REOPEN);
            hWnd=lpFrameCurrent->ei.hWndEdit;

            //If wpCmdLine not NULL, document just opened and does not need a reopening.
            if (SaveChanged(0) && !wpCmdLine)
            {
              OpenDocument(hWnd, wszFile, dwFlags|OD_REOPEN, nCodePage, bBOM);
            }
            nResult=EOD_WINDOW_EXIST;
            goto End;
          }
        }
      }
      if (dwFlags & OD_REOPEN)
      {
        SendMessage(hWnd, AEM_GETSCROLLPOS, 0, (LPARAM)&ptDocumentPos);
      }
    }
  }

  //Read from file
  for (;;)
  {
    if ((hFile=CreateFileWide(wszFile, GENERIC_READ, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, bFileExist?OPEN_EXISTING:OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL|FILE_FLAG_SEQUENTIAL_SCAN, NULL)) == INVALID_HANDLE_VALUE)
    {
      if (!bSetSecurity && !bOldWindows && GetLastError() == ERROR_ACCESS_DENIED && IsFile(wszFile) != ERROR_DIRECTORY)
      {
        //Allow all access to the file (UAC).
        if (AkelAdminInit(wszFile))
        {
          if ((!bFileExist && !AkelAdminSend(AAA_CREATEFILE, wszFile)) ||
              !AkelAdminSend(AAA_SECURITYGET, wszFile) ||
              !AkelAdminSend(AAA_SECURITYSETEVERYONE, wszFile))
          {
            //Reset AkelAdmin
            AkelAdminExit();
          }
          else
          {
            bSetSecurity=TRUE;
            continue;
          }
        }
      }
      API_LoadStringW(hLangLib, MSG_CANNOT_OPEN_FILE, wbuf, BUFFER_SIZE);
      xprintfW(wszMsg, wbuf, wszFile);
      API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);

      nResult=EOD_OPEN;
      break;
    }

    if (bFileExist)
    {
      //Autodetect code page
      if ((nDetect=AutodetectCodePage(wszFile, hFile, moCur.dwCodepageRecognitionBuffer, dwFlags, &nCodePage, &bBOM)) < 0)
      {
        if (!(dwFlags & OD_REOPEN))
        {
          if (nDetect == EDT_BINARY)
          {
            if (dwCmdLineOptions & CLO_MSGOPENBINARYYES)
            {
            }
            else if (dwCmdLineOptions & CLO_MSGOPENBINARYNO)
            {
              nResult=EOD_MSGNO;
              break;
            }
            else
            {
              API_LoadStringW(hLangLib, MSG_ERROR_BINARY, wbuf, BUFFER_SIZE);
              xprintfW(wszMsg, wbuf, wszFile);
              nChoice=API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, dwMsgFlags|MB_ICONEXCLAMATION|MB_DEFBUTTON2);
              if (nChoice == IDNO)
              {
                nResult=EOD_MSGNO;
                goto End;
              }
              else if (nChoice == IDCANCEL)
              {
                nResult=EOD_MSGCANCELBINARY;
                break;
              }
            }
          }
          else
          {
            nResult=nDetect;
            break;
          }
        }
      }
    }

    //Check code page
    if (!IsCodePageValid(nCodePage))
    {
      API_LoadStringW(hLangLib, MSG_CP_UNIMPLEMENTED, wbuf, BUFFER_SIZE);
      xprintfW(wszMsg, wbuf, nCodePage);
      API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
      nResult=EOD_CODEPAGE_ERROR;
      break;
    }

    //Offset BOM
    fsd.dwBytesCurrent=0;

    if (bBOM)
    {
      if (nCodePage == CP_UNICODE_UTF16LE ||
          nCodePage == CP_UNICODE_UTF16BE)
      {
        fsd.dwBytesCurrent=2;
      }
      else if (nCodePage == CP_UNICODE_UTF32LE ||
               nCodePage == CP_UNICODE_UTF32BE)
      {
        fsd.dwBytesCurrent=4;
      }
      else if (nCodePage == CP_UNICODE_UTF8)
      {
        fsd.dwBytesCurrent=3;
      }
    }
    SetFilePointer64(hFile, fsd.dwBytesCurrent, FILE_BEGIN);

    if (IsEditActive(hWnd))
    {
      //Save position of the previous file before load new document
      RecentFilesSaveFile(lpFrameCurrent);

      //Create edit window if necessary
      if (nMDI && !(dwFlags & OD_REOPEN) && (!lpFrameCurrent->hWndEditParent || lpFrameCurrent->ei.bModified || lpFrameCurrent->wszFile[0]))
      {
        DoFileNew();
        hWnd=lpFrameCurrent->ei.hWndEdit;
      }

      //Get file write time
      GetFileTime(hFile, NULL, NULL, &lpFrameCurrent->ft);
    }

    //Load text
    fsd.hWnd=hWnd;
    fsd.hFile=hFile;
    fsd.nCodePage=nCodePage;
    fsd.dwFlags=dwFlags;
    fsd.nNewLine=moCur.nNewFileNewLine;
    fsd.dwBytesMax=(UINT_PTR)-1;
    fsd.bResult=TRUE;
    lpStreamInData=&fsd;
    FileStreamIn(&fsd);
    lpStreamInData=NULL;
    break;
  }
  if (hFile != INVALID_HANDLE_VALUE)
    CloseHandle(hFile);

  //Change back file security (UAC).
  if (bSetSecurity)
  {
    if (!AkelAdminSend(AAA_SECURITYRESTORE, L"") ||
        !AkelAdminSend(AAA_SECURITYFREE, L"") ||
        !moCur.bAkelAdminResident)
    {
      //Reset AkelAdmin
      AkelAdminExit();
    }
    bSetSecurity=FALSE;
  }

  if (fsd.bResult != -1 && IsEditActive(hWnd))
  {
    if (fsd.bResult)
    {
      //Compare
      nFileCmp=xstrcmpiW(lpFrameCurrent->wszFile, wszFile);

      if (nFileCmp || lpFrameCurrent->ei.nCodePage != nCodePage)
      {
        //Read position of the new document
        if (moCur.nRecentFiles)
        {
          RecentFilesRead(&hRecentFilesStack);
          if (lpRecentFile=RecentFilesUpdate(wszFile))
            lpRecentFile->nCodePage=nCodePage;
          RecentFilesSave(&hRecentFilesStack);
          if (nFileCmp) bMenuRecentFiles=TRUE;
        }
      }

      //Update titles
      SetNewLineStatus(lpFrameCurrent, fsd.nNewLine, AENL_INPUT);
      SetModifyStatus(lpFrameCurrent, FALSE);
      SetCodePageStatus(lpFrameCurrent, nCodePage, bBOM);

      if (nFileCmp)
      {
        lpFrameCurrent->nStreamOffset=nStreamOffset;
        lpFrameCurrent->nFileLen=(int)xstrcpynW(lpFrameCurrent->wszFile, wszFile, MAX_PATH);
        WideCharToMultiByte(CP_ACP, 0, lpFrameCurrent->wszFile, lpFrameCurrent->nFileLen + 1, lpFrameCurrent->szFile, MAX_PATH, NULL, NULL);
        UpdateTitle(lpFrameCurrent);
      }

      //.LOG
      if (moCur.bDateLog && !IsReadOnly(hWnd))
      {
        AECHARINDEX ciChar;

        if (SendMessage(hWnd, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&ciChar))
        {
          if (!xstrcmpW(ciChar.lpLine->wpLine, L".LOG"))
          {
            SendMessage(hWnd, EM_SETSEL, (WPARAM)-1, (LPARAM)-1);

            GetTimeString(moCur.wszDateLogFormat, wbuf, FALSE);
            xprintfW(wbuf2, L"\r%s\r", wbuf);
            ReplaceSelW(hWnd, wbuf2, -1, AELB_ASINPUT, 0, NULL, NULL);
            goto GlobalPrint;
          }
        }
      }

      //Update selection
      if (moCur.nRecentFiles && moCur.bSavePositions)
      {
        CHARRANGE64 cr;
        int nLockScroll;

        if ((nLockScroll=(int)SendMessage(hWnd, AEM_LOCKSCROLL, (WPARAM)-1, 0)) == -1)
          SendMessage(hWnd, AEM_LOCKSCROLL, SB_BOTH, TRUE);
        cr.cpMin=hRecentFilesStack.first->cpMin;
        cr.cpMax=hRecentFilesStack.first->cpMax;
        SendMessage(hWnd, EM_EXSETSEL64, 0, (LPARAM)&cr);
        if (nLockScroll == -1)
        {
          SendMessage(hWnd, AEM_LOCKSCROLL, SB_BOTH, FALSE);

          if (!(dwFlags & OD_NOSCROLL))
          {
            if ((dwFlags & OD_REOPEN))
              SendMessage(hWnd, AEM_SETSCROLLPOS, 0, (LPARAM)&ptDocumentPos);
            else
              ScrollCaret(hWnd);
          }
        }
      }
      else UpdateStatusUser(lpFrameCurrent, CSB_CHARHEX|CSB_CHARDEC|CSB_CHARLETTER|CSB_RICHOFFSET|CSB_BYTEOFFSET|CSB_LINEALLCOUNT|CSB_RICHCOUNT|CSB_LINESELCOUNT|CSB_LINESELBEGIN|CSB_LINESELEND);

      //Print if "/p" option used in command line
      GlobalPrint:
      if (dwCmdLineOptions & CLO_GLOBALPRINT)
      {
        DoFilePrint(lpFrameCurrent, TRUE);
        dwCmdLineOptions&=~CLO_GLOBALPRINT;

        if (!nMDI)
        {
          PostMessage(hMainWnd, WM_COMMAND, IDM_FILE_EXIT, 0);
        }
        else
        {
          DestroyFrameWindow(lpFrameCurrent);
          //if (FrameNoWindows())
          //  PostMessage(hMainWnd, WM_COMMAND, IDM_FILE_EXIT, 0);
        }
      }
    }
    else nResult=EOD_STREAMIN;
  }

  End:
  if (GetWindowLongPtrWide(hWnd, GWLP_ID) == ID_EDIT)
  {
    SendMessage(hMainWnd, AKDN_OPENDOCUMENT_FINISH, (WPARAM)hWnd, nResult);
  }
  return nResult;
}

/*
//Read entire file (AEM_SETTEXTW).
void FileStreamIn(FILESTREAMDATA *lpData)
{
  unsigned char *szBuffer;
  wchar_t *wszBuffer;
  wchar_t *wpCount;
  wchar_t *wpMax;
  INT_PTR nBufferBytes;
  UINT_PTR dwFileSize;
  UINT_PTR dwBytesRead;
  UINT_PTR dwCharsConverted;

  if ((dwFileSize=GetFileSize64(lpData->hFile)) != (UINT_PTR)-1)
  {
    if (lpData->dwBytesMax == (UINT_PTR)-1)
      lpData->dwBytesMax=dwFileSize;

    if (IsCodePageUnicode(lpData->nCodePage) && lpData->nCodePage != CP_UNICODE_UTF8)
      nBufferBytes=lpData->dwBytesMax;
    else
      nBufferBytes=lpData->dwBytesMax * sizeof(wchar_t);

    if (wszBuffer=(wchar_t *)API_HeapAlloc(hHeap, 0, nBufferBytes + sizeof(wchar_t)))
    {
      if (IsCodePageUnicode(lpData->nCodePage) && lpData->nCodePage != CP_UNICODE_UTF8)
        szBuffer=(unsigned char *)wszBuffer;
      else
        szBuffer=(unsigned char *)wszBuffer + lpData->dwBytesMax;

      //Read data from file
      if (ReadFile64(lpData->hFile, szBuffer, lpData->dwBytesMax, &dwBytesRead, NULL))
      {
        //Translate data to UNICODE
        if (lpData->nCodePage == CP_UNICODE_UTF16LE)
        {
          dwCharsConverted=dwBytesRead / sizeof(wchar_t);
        }
        else if (lpData->nCodePage == CP_UNICODE_UTF16BE)
        {
          ChangeTwoBytesOrder(szBuffer, dwBytesRead, NULL);
          dwCharsConverted=dwBytesRead / sizeof(wchar_t);
        }
        else if (lpData->nCodePage == CP_UNICODE_UTF32LE)
        {
          dwCharsConverted=UTF32toUTF16((const unsigned long *)szBuffer, dwBytesRead / sizeof(unsigned long), NULL, wszBuffer, nBufferBytes / sizeof(wchar_t));
        }
        else if (lpData->nCodePage == CP_UNICODE_UTF32BE)
        {
          ChangeFourBytesOrder(szBuffer, dwBytesRead, NULL);
          dwCharsConverted=UTF32toUTF16((const unsigned long *)szBuffer, dwBytesRead / sizeof(unsigned long), NULL, wszBuffer, nBufferBytes / sizeof(wchar_t));
        }
        else
        {
          if (lpData->nCodePage == CP_UNICODE_UTF8)
            dwCharsConverted=UTF8toUTF16(szBuffer, dwBytesRead, NULL, wszBuffer, nBufferBytes / sizeof(wchar_t));
          else
            dwCharsConverted=MultiByteToWideChar64(lpData->nCodePage, 0, (char *)szBuffer, dwBytesRead, wszBuffer, nBufferBytes / sizeof(wchar_t));
        }
        wszBuffer[dwCharsConverted]=L'\0';
        wpMax=wszBuffer + dwCharsConverted;

        //Detect new line
        for (wpCount=wszBuffer; wpCount < wpMax; ++wpCount)
        {
          if (*wpCount == L'\r')
          {
            if (*(wpCount + 1) == L'\r' && *(wpCount + 2) == L'\n')
            {
              //Windows format \r\r\n
              lpData->nNewLine=NEWLINE_WIN;
              break;
            }
            else if (*(wpCount + 1) == L'\n')
            {
              //Windows format \r\n
              lpData->nNewLine=NEWLINE_WIN;
              break;
            }
            //MacOS format \r
            lpData->nNewLine=NEWLINE_MAC;
            break;
          }
          else if (*wpCount == L'\n')
          {
            //Unix format \n
            lpData->nNewLine=NEWLINE_UNIX;
            break;
          }
        }

        //Send to AkelEdit
        SendMessage(lpData->hWnd, AEM_SETNEWLINE, AENL_INPUT|AENL_OUTPUT, MAKELONG(AELB_ASIS, AELB_ASIS));
        SendMessage(lpData->hWnd, AEM_SETTEXTW, (WPARAM)dwCharsConverted, (LPARAM)wszBuffer);
      }
      else
      {
        lpData->bResult=FALSE;
        PostMessage(hMainWnd, WM_COMMAND, IDM_INTERNAL_ERRORIO_MSG, 0);
      }
      API_HeapFree(hHeap, 0, (LPVOID)wszBuffer);
    }
    else lpData->bResult=FALSE;
  }
  else lpData->bResult=FALSE;
}
*/

//Stream file reading (AEM_STREAMIN).
void FileStreamIn(FILESTREAMDATA *lpData)
{
  CPINFO cpi;

  if (lpData->dwBytesMax == (UINT_PTR)-1)
    lpData->dwBytesMax=GetFileSize64(lpData->hFile);
  SendMessage(lpData->hWnd, AEM_SETNEWLINE, AENL_INPUT|AENL_OUTPUT, MAKELONG(AELB_ASIS, AELB_ASIS));

  if (!IsCodePageUnicode(lpData->nCodePage))
  {
    GetCPInfo(lpData->nCodePage, &cpi);
    lpData->dwMaxCharSize=cpi.MaxCharSize;
  }
  else lpData->dwMaxCharSize=0;

  if (lpData->dwMaxCharSize > 2)
  {
    //MBCS. Read entire file (AEM_SETTEXTW).
    unsigned char *szBuffer;
    wchar_t *wszBuffer;
    wchar_t *wpCount;
    wchar_t *wpMax;
    INT_PTR nBufferBytes=lpData->dwBytesMax * sizeof(wchar_t);
    UINT_PTR dwBytesRead;
    UINT_PTR dwCharsConverted;

    if (wszBuffer=(wchar_t *)API_HeapAlloc(hHeap, 0, nBufferBytes + sizeof(wchar_t)))
    {
      szBuffer=(unsigned char *)wszBuffer + lpData->dwBytesMax;

      //Read data from file
      if (ReadFile64(lpData->hFile, szBuffer, lpData->dwBytesMax, &dwBytesRead, NULL))
      {
        //Translate data to UNICODE
        dwCharsConverted=MultiByteToWideChar64(lpData->nCodePage, 0, (char *)szBuffer, dwBytesRead, wszBuffer, nBufferBytes / sizeof(wchar_t));
        wszBuffer[dwCharsConverted]=L'\0';
        wpMax=wszBuffer + dwCharsConverted;

        //Detect new line
        for (wpCount=wszBuffer; wpCount < wpMax; ++wpCount)
        {
          if (*wpCount == L'\r')
          {
            if (*(wpCount + 1) == L'\r' && *(wpCount + 2) == L'\n')
            {
              //Windows format \r\r\n
              lpData->nNewLine=NEWLINE_WIN;
              break;
            }
            else if (*(wpCount + 1) == L'\n')
            {
              //Windows format \r\n
              lpData->nNewLine=NEWLINE_WIN;
              break;
            }
            //MacOS format \r
            lpData->nNewLine=NEWLINE_MAC;
            break;
          }
          else if (*wpCount == L'\n')
          {
            //Unix format \n
            lpData->nNewLine=NEWLINE_UNIX;
            break;
          }
        }

        //Send to AkelEdit
        SendMessage(lpData->hWnd, AEM_SETNEWLINE, AENL_INPUT|AENL_OUTPUT, MAKELONG(AELB_ASIS, AELB_ASIS));
        SendMessage(lpData->hWnd, AEM_SETTEXTW, (WPARAM)dwCharsConverted, (LPARAM)wszBuffer);
      }
      else
      {
        lpData->bResult=FALSE;
        PostMessage(hMainWnd, WM_COMMAND, IDM_INTERNAL_ERRORIO_MSG, 0);
      }
      API_HeapFree(hHeap, 0, (LPVOID)wszBuffer);
    }
    else lpData->bResult=FALSE;
  }
  else
  {
    //Read file stream (AEM_STREAMIN).
    AESTREAMIN aesi;

    aesi.dwCookie=(UINT_PTR)lpData;
    aesi.lpCallback=InputStreamCallback;
    aesi.nNewLine=AELB_ASIS;
    aesi.dwTextLen=lpData->dwBytesMax;
    SendMessage(lpData->hWnd, AEM_STREAMIN, 0, (LPARAM)&aesi);
    lpData->bResult=!aesi.dwError;

    //Detect new line
    if (lpData->bResult)
    {
      if (aesi.nFirstNewLine == AELB_RRN)
        lpData->nNewLine=NEWLINE_WIN;
      else if (aesi.nFirstNewLine == AELB_RN)
        lpData->nNewLine=NEWLINE_WIN;
      else if (aesi.nFirstNewLine == AELB_R)
        lpData->nNewLine=NEWLINE_MAC;
      else if (aesi.nFirstNewLine == AELB_N)
        lpData->nNewLine=NEWLINE_UNIX;
    }
  }
}

DWORD CALLBACK InputStreamCallback(UINT_PTR dwCookie, wchar_t *wszBuf, DWORD dwBufBytesSize, DWORD *dwBufBytesDone)
{
  FILESTREAMDATA *lpData=(FILESTREAMDATA *)dwCookie;
  UINT_PTR dwBytesRead;
  UINT_PTR dwCharsConverted=0;
  UINT_PTR dwBytesConverted=0;

  if (lpData->dwBytesCurrent <= lpData->dwBytesMax)
  {
    if (ReadFile64(lpData->hFile, pcTranslateBuffer, dwBufBytesSize / sizeof(wchar_t), &dwBytesRead, NULL))
    {
      lpData->dwBytesCurrent+=dwBytesRead;

      //Translate data to UNICODE
      if (lpData->nCodePage == CP_UNICODE_UTF16LE)
      {
        xmemcpy(wszBuf, pcTranslateBuffer, dwBytesRead);
        dwCharsConverted=dwBytesRead / sizeof(wchar_t);
      }
      else if (lpData->nCodePage == CP_UNICODE_UTF16BE)
      {
        ChangeTwoBytesOrder(pcTranslateBuffer, dwBytesRead, NULL);
        xmemcpy(wszBuf, pcTranslateBuffer, dwBytesRead);
        dwCharsConverted=dwBytesRead / sizeof(wchar_t);
      }
      else if (lpData->nCodePage == CP_UNICODE_UTF32LE)
      {
        dwCharsConverted=UTF32toUTF16((const unsigned long *)pcTranslateBuffer, dwBytesRead / sizeof(unsigned long), NULL, (unsigned short *)wszBuf, dwBufBytesSize / sizeof(wchar_t));
      }
      else if (lpData->nCodePage == CP_UNICODE_UTF32BE)
      {
        ChangeFourBytesOrder(pcTranslateBuffer, dwBytesRead, NULL);
        dwCharsConverted=UTF32toUTF16((const unsigned long *)pcTranslateBuffer, dwBytesRead / sizeof(unsigned long), NULL, (unsigned short *)wszBuf, dwBufBytesSize / sizeof(wchar_t));
      }
      else
      {
        if (lpData->nCodePage == CP_UNICODE_UTF8)
        {
          dwCharsConverted=UTF8toUTF16(pcTranslateBuffer, dwBytesRead, &dwBytesConverted, (unsigned short *)wszBuf, dwBufBytesSize / sizeof(wchar_t));

          if (dwBytesRead != dwBytesConverted)
          {
            //UTF-8 char was split
            lpData->dwBytesCurrent+=dwBytesConverted - dwBytesRead;
            SetFilePointer64(lpData->hFile, lpData->dwBytesCurrent, FILE_BEGIN);
          }
        }
        else
        {
          dwCharsConverted=MultiByteToWideChar(lpData->nCodePage, 0, (char *)pcTranslateBuffer, (int)dwBytesRead, wszBuf, dwBufBytesSize / sizeof(wchar_t));

          if (lpData->dwMaxCharSize == 2 && dwCharsConverted > 0 && dwBytesRead >= 2 && lpData->dwBytesCurrent < lpData->dwBytesMax)
          {
            //DBCS
            BYTE *pLastChar=pcTranslateBuffer + dwBytesRead - 1;

            //Windows 95/98/Me/2000/XP/2003. MultiByteToWideChar converts one byte of double-byte char to zero.
            if ((*pLastChar != '\0' && wszBuf[dwCharsConverted - 1] == L'\0') ||
                //Windows Vista/7/2008. MultiByteToWideChar trying to convert one byte of double-byte char.
                (IsDBCSLeadByteEx(lpData->nCodePage, *pLastChar) &&
                 //Double-byte char can consist of two lead bytes
                 (!IsDBCSLeadByteEx(lpData->nCodePage, *(pLastChar - 1)) ||
                   pLastChar - (BYTE *)CharPrevExA((WORD)lpData->nCodePage, (char *)pcTranslateBuffer, (char *)pLastChar, 0) == 2)))
            {
              //Double-byte char was split
              --lpData->dwBytesCurrent;
              SetFilePointer64(lpData->hFile, lpData->dwBytesCurrent, FILE_BEGIN);
              --dwCharsConverted;
            }
          }
        }
      }
      wszBuf[dwCharsConverted]=L'\0';
    }
    else PostMessage(hMainWnd, WM_COMMAND, IDM_INTERNAL_ERRORIO_MSG, 0);
  }
  *dwBufBytesDone=(DWORD)dwCharsConverted * sizeof(wchar_t);

  return 0;
}

UINT_PTR ReadFileContent(HANDLE hFile, UINT_PTR dwBytesMax, int nCodePage, BOOL bBOM, wchar_t **wpContent)
{
  unsigned char *szBuffer;
  wchar_t *wszBuffer=NULL;
  UINT_PTR dwFileSize;
  UINT_PTR dwBufferBytes;
  UINT_PTR dwBytesRead;
  UINT_PTR dwCharsConverted=0;

  //Offset BOM
  if (bBOM)
  {
    if (SetFilePointer64(hFile, 0, FILE_CURRENT) == 0)
    {
      if (nCodePage == CP_UNICODE_UTF16LE ||
          nCodePage == CP_UNICODE_UTF16BE)
      {
        SetFilePointer64(hFile, 2, FILE_BEGIN);
      }
      else if (nCodePage == CP_UNICODE_UTF32LE ||
               nCodePage == CP_UNICODE_UTF32BE)
      {
        SetFilePointer64(hFile, 4, FILE_BEGIN);
      }
      else if (nCodePage == CP_UNICODE_UTF8)
      {
        SetFilePointer64(hFile, 3, FILE_BEGIN);
      }
    }
  }

  if ((dwFileSize=GetFileSize64(hFile)) != (UINT_PTR)-1)
  {
    if (dwBytesMax == (UINT_PTR)-1)
      dwBytesMax=dwFileSize;

    if (IsCodePageUnicode(nCodePage) && nCodePage != CP_UNICODE_UTF8)
      dwBufferBytes=dwBytesMax;
    else
      dwBufferBytes=dwBytesMax * sizeof(wchar_t);

    if (wszBuffer=(wchar_t *)API_HeapAlloc(hHeap, 0, dwBufferBytes + sizeof(wchar_t)))
    {
      if (IsCodePageUnicode(nCodePage) && nCodePage != CP_UNICODE_UTF8)
        szBuffer=(unsigned char *)wszBuffer;
      else
        szBuffer=(unsigned char *)wszBuffer + dwBytesMax;

      //Read data from file
      if (ReadFile64(hFile, szBuffer, dwBytesMax, &dwBytesRead, NULL))
      {
        //Translate data to UNICODE
        if (nCodePage == CP_UNICODE_UTF16LE)
        {
          dwCharsConverted=dwBytesRead / sizeof(wchar_t);
        }
        else if (nCodePage == CP_UNICODE_UTF16BE)
        {
          ChangeTwoBytesOrder(szBuffer, dwBytesRead, NULL);
          dwCharsConverted=dwBytesRead / sizeof(wchar_t);
        }
        else if (nCodePage == CP_UNICODE_UTF32LE)
        {
          dwCharsConverted=UTF32toUTF16((const unsigned long *)szBuffer, dwBytesRead / sizeof(unsigned long), NULL, (unsigned short *)wszBuffer, dwBufferBytes / sizeof(wchar_t));
        }
        else if (nCodePage == CP_UNICODE_UTF32BE)
        {
          ChangeFourBytesOrder(szBuffer, dwBytesRead, NULL);
          dwCharsConverted=UTF32toUTF16((const unsigned long *)szBuffer, dwBytesRead / sizeof(unsigned long), NULL, (unsigned short *)wszBuffer, dwBufferBytes / sizeof(wchar_t));
        }
        else
        {
          if (nCodePage == CP_UNICODE_UTF8)
            dwCharsConverted=UTF8toUTF16(szBuffer, dwBytesRead, NULL, (unsigned short *)wszBuffer, dwBufferBytes / sizeof(wchar_t));
          else
            dwCharsConverted=MultiByteToWideChar64(nCodePage, 0, (char *)szBuffer, dwBytesRead, wszBuffer, dwBufferBytes / sizeof(wchar_t));
        }
        wszBuffer[dwCharsConverted]=L'\0';
      }
    }
  }
  *wpContent=wszBuffer;
  return dwCharsConverted;
}

int WriteFileContent(HANDLE hFile, const wchar_t *wpContent, INT_PTR nContentLen, int nCodePage, BOOL bBOM)
{
  const wchar_t *wpContentMax;
  const wchar_t *wpBlock;
  unsigned char *pDataToWrite;
  DWORD dwBlockLen=2048;
  UINT_PTR dwBytesToWrite=0;
  UINT_PTR dwBytesWritten;
  int nResult=ESD_SUCCESS;
  int nWrite=0;

  if (nContentLen == -1)
    nContentLen=xstrlenW(wpContent);
  wpContentMax=wpContent + nContentLen;

  if (bBOM)
  {
    if (IsCodePageUnicode(nCodePage))
    {
      if (nCodePage == CP_UNICODE_UTF16LE)
        nWrite=API_WriteFile(hFile, "\xFF\xFE", 2, &dwBytesWritten, NULL);
      else if (nCodePage == CP_UNICODE_UTF16BE)
        nWrite=API_WriteFile(hFile, "\xFE\xFF", 2, &dwBytesWritten, NULL);
      else if (nCodePage == CP_UNICODE_UTF32LE)
        nWrite=API_WriteFile(hFile, "\xFF\xFE\x00\x00", 4, &dwBytesWritten, NULL);
      else if (nCodePage == CP_UNICODE_UTF32BE)
        nWrite=API_WriteFile(hFile, "\x00\x00\xFE\xFF", 4, &dwBytesWritten, NULL);
      else if (nCodePage == CP_UNICODE_UTF8)
        nWrite=API_WriteFile(hFile, "\xEF\xBB\xBF", 3, &dwBytesWritten, NULL);
      if (!nWrite)
        return ESD_WRITE;
    }
    else bBOM=FALSE;
  }

  for (wpBlock=wpContent; wpBlock < wpContentMax; wpBlock+=dwBlockLen)
  {
    if (wpBlock + dwBlockLen > wpContentMax)
      dwBlockLen=(DWORD)(wpContentMax - wpBlock);
    pDataToWrite=pcTranslateBuffer;

    if (nCodePage == CP_UNICODE_UTF16LE)
    {
      pDataToWrite=(unsigned char *)wpBlock;
      dwBytesToWrite=dwBlockLen * sizeof(wchar_t);
    }
    else if (nCodePage == CP_UNICODE_UTF16BE)
    {
      ChangeTwoBytesOrder((unsigned char *)wpBlock, dwBytesToWrite, pcTranslateBuffer);
      dwBytesToWrite=dwBlockLen * sizeof(wchar_t);
    }
    else if (nCodePage == CP_UNICODE_UTF32LE)
    {
      dwBytesToWrite=UTF16toUTF32(wpBlock, dwBlockLen, NULL, (unsigned long *)pcTranslateBuffer, TRANSLATE_BUFFER_SIZE / sizeof(unsigned long)) * sizeof(unsigned long);
    }
    else if (nCodePage == CP_UNICODE_UTF32BE)
    {
      dwBytesToWrite=UTF16toUTF32(wpBlock, dwBlockLen, NULL, (unsigned long *)pcTranslateBuffer, TRANSLATE_BUFFER_SIZE / sizeof(unsigned long)) * sizeof(unsigned long);
      ChangeFourBytesOrder(pcTranslateBuffer, dwBytesToWrite, NULL);
    }
    else
    {
      if (nCodePage == CP_UNICODE_UTF8)
        dwBytesToWrite=UTF16toUTF8(wpBlock, dwBlockLen, NULL, pcTranslateBuffer, TRANSLATE_BUFFER_SIZE);
      else
        dwBytesToWrite=WideCharToMultiByte(nCodePage, 0, wpBlock, dwBlockLen, (char *)pcTranslateBuffer, TRANSLATE_BUFFER_SIZE, NULL, NULL);
    }
    if (!API_WriteFile(hFile, pDataToWrite, dwBytesToWrite, &dwBytesWritten, NULL))
    {
      nResult=ESD_WRITE;
      break;
    }
  }
  return nResult;
}

BOOL OpenDocumentSend(HWND hWnd, HWND hWndEditCtrl, const wchar_t *wpFile, DWORD dwFlags, int nCodePage, BOOL bBOM, BOOL bOtherProcess)
{
  OPENDOCUMENTPOSTW odp;
  COPYDATASTRUCT cds;

  odp.dwFlags=dwFlags;
  odp.nCodePage=nCodePage;
  odp.bBOM=bBOM;
  odp.hWnd=hWndEditCtrl;
  xstrcpynW(odp.szFile, wpFile, MAX_PATH);
  if (bOtherProcess)
    GetCurrentDirectoryWide(MAX_PATH, odp.szWorkDir);
  else
    odp.szWorkDir[0]=L'\0';

  cds.dwData=CD_OPENDOCUMENTW;
  cds.cbData=sizeof(OPENDOCUMENTPOSTW);
  cds.lpData=&odp;
  return (BOOL)SendMessage(hWnd, WM_COPYDATA, (WPARAM)hWnd, (LPARAM)&cds);
}

int SaveDocument(HWND hWnd, const wchar_t *wpFile, int nCodePage, BOOL bBOM, DWORD dwFlags)
{
  wchar_t wszFile[MAX_PATH];
  WIN32_FIND_DATAW wfd;
  HANDLE hFile;
  FILESTREAMDATA fsd;
  UINT_PTR dwBytesWritten;
  int nResult=ESD_SUCCESS;
  int nWrite=0;
  int nFileCmp;
  int nCodePageCmp;
  int nLostLine=0;
  int nLostCharInLine;
  int nFileLen;
  int nStreamOffset;
  BOOL bFileExist;

  if (!wpFile[0])
  {
    SetCodePageStatus(lpFrameCurrent, nCodePage, bBOM);
    return nResult;
  }
  if (!hWnd)
    hWnd=lpFrameCurrent->ei.hWndEdit;
  if (!hWnd)
    return ESD_NOWINDOW;
  GetFullName(wpFile, wszFile, MAX_PATH, &nFileLen);
  nStreamOffset=GetFileStreamOffset(wszFile, nFileLen);

  //Notification message
  if (GetWindowLongPtrWide(hWnd, GWLP_ID) == ID_EDIT)
  {
    NSAVEDOCUMENT nsd;
    char szFile[MAX_PATH];

    WideCharToMultiByte(CP_ACP, 0, wszFile, -1, szFile, MAX_PATH, NULL, NULL);
    nsd.pFile=bOldWindows?(LPBYTE)szFile:(LPBYTE)wszFile;
    nsd.szFile=szFile;
    nsd.wszFile=wszFile;
    nsd.nCodePage=&nCodePage;
    nsd.bBOM=&bBOM;
    nsd.dwFlags=dwFlags;
    nsd.bProcess=TRUE;
    SendMessage(hMainWnd, AKDN_SAVEDOCUMENT_START, (WPARAM)hWnd, (LPARAM)&nsd);
    if (!nsd.bProcess)
    {
      nResult=ESD_STOP;
      goto End;
    }
  }
  fsd.bResult=-1;

  //Check code page
  if (!IsCodePageValid(nCodePage))
  {
    if (!IsEditActive(hWnd))
      SetFocus(hWnd);
    API_LoadStringW(hLangLib, MSG_CP_UNIMPLEMENTED, wbuf, BUFFER_SIZE);
    xprintfW(wszMsg, wbuf, nCodePage);
    API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
    nResult=ESD_CODEPAGE_ERROR;
    goto End;
  }
  if (!IsCodePageUnicode(nCodePage))
  {
    if (nLostLine=(int)SendMessage(hWnd, AEM_CHECKCODEPAGE, (WPARAM)nCodePage, (LPARAM)&nLostCharInLine))
    {
      if (dwCmdLineOptions & CLO_MSGSAVELOSTSYMBOLSYES)
      {
      }
      else if (dwCmdLineOptions & CLO_MSGSAVELOSTSYMBOLSNO)
      {
        nResult=ESD_CODEPAGE_ERROR;
        goto End;
      }
      else
      {
        if (!IsEditActive(hWnd))
          SetFocus(hWnd);

        //Custom MessageBox
        {
          BUTTONMESSAGEBOX bmb[]={{IDC_MESSAGEBOX_OK,   MAKEINTRESOURCEW(STR_MESSAGEBOX_OK),     0},
                                  {IDCANCEL,            MAKEINTRESOURCEW(STR_MESSAGEBOX_CANCEL), 0},
                                  {IDC_MESSAGEBOX_GOTO, MAKEINTRESOURCEW(STR_MESSAGEBOX_GOTO),   BMB_DEFAULT},
                                  {0, 0, 0}};
          int nChoice;
          int nMessageLine;

          nMessageLine=GetLineNumber(lpFrameCurrent, nLostLine - 1) + 1;
          API_LoadStringW(hLangLib, MSG_CP_MISMATCH, wbuf, BUFFER_SIZE);
          xprintfW(wszMsg, wbuf, nMessageLine);
          nChoice=MessageBoxCustom(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_ICONEXCLAMATION, NULL, &bmb[0]);

          if (nChoice == IDC_MESSAGEBOX_GOTO ||
              nChoice == IDCANCEL)
          {
            if (nChoice == IDC_MESSAGEBOX_GOTO)
            {
              AECHARRANGE cr;

              //Lost index
              cr.ciMin.nLine=nLostLine - 1;
              cr.ciMin.nCharInLine=nLostCharInLine - 1;
              SendMessage(hWnd, AEM_INDEXUPDATE, 0, (LPARAM)&cr.ciMin);

              //Set selection
              cr.ciMax=cr.ciMin;
              SetSel(hWnd, &cr, AESELT_LOCKSCROLL, NULL);
              ScrollCaret(hWnd);
            }
            nResult=ESD_CODEPAGE_ERROR;
            goto End;
          }
        }
      }
    }
  }

  //File attributes
  if (nStreamOffset) wszFile[nStreamOffset]=L'\0';

  if (GetFileWin32Data(wszFile, &wfd))
  {
    if (moCur.bSaveInReadOnlyMsg && (wfd.dwFileAttributes & FILE_ATTRIBUTE_READONLY))
    {
      if (!IsEditActive(hWnd))
        SetFocus(hWnd);
      API_LoadStringW(hLangLib, MSG_SAVEIN_READONLY, wbuf, BUFFER_SIZE);
      xprintfW(wszMsg, wbuf, wszFile);
      if (API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_YESNO|MB_ICONEXCLAMATION) == IDNO)
      {
        nResult=ESD_READONLY;
        goto End;
      }
    }
    if ((wfd.dwFileAttributes & FILE_ATTRIBUTE_READONLY) || (wfd.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) || (wfd.dwFileAttributes & FILE_ATTRIBUTE_SYSTEM))
      SetFileAttributesWide(wszFile, wfd.dwFileAttributes & ~FILE_ATTRIBUTE_READONLY & ~FILE_ATTRIBUTE_HIDDEN & ~FILE_ATTRIBUTE_SYSTEM);
  }

  //Write to file
  for (;;)
  {
    if (nStreamOffset) wszFile[nStreamOffset]=L':';

    if (nStreamOffset)
      bFileExist=FileExistsWide(wszFile);
    else
      bFileExist=(wfd.dwFileAttributes != INVALID_FILE_ATTRIBUTES);

    if ((hFile=CreateFileWide(wszFile, GENERIC_WRITE, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, bFileExist?TRUNCATE_EXISTING:CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL)) == INVALID_HANDLE_VALUE)
    {
      if (!bSetSecurity && !bOldWindows && GetLastError() == ERROR_ACCESS_DENIED && IsFile(wszFile) != ERROR_DIRECTORY)
      {
        if (nStreamOffset) wszFile[nStreamOffset]=L'\0';

        //Allow all access to the file (UAC).
        if (AkelAdminInit(wszFile))
        {
          if ((wfd.dwFileAttributes == INVALID_FILE_ATTRIBUTES && !AkelAdminSend(AAA_CREATEFILE, wszFile)) ||
              !AkelAdminSend(AAA_SECURITYGET, wszFile) ||
              !AkelAdminSend(AAA_SECURITYSETEVERYONE, wszFile))
          {
            //Reset AkelAdmin
            AkelAdminExit();
          }
          else
          {
            //File created - refill "wfd".
            if (wfd.dwFileAttributes == INVALID_FILE_ATTRIBUTES)
              GetFileWin32Data(wszFile, &wfd);

            //Security successfully set - try to open file again.
            bSetSecurity=TRUE;
            continue;
          }
        }
        if (nStreamOffset) wszFile[nStreamOffset]=L':';
      }
      API_LoadStringW(hLangLib, MSG_CANNOT_OPEN_FILE, wbuf, BUFFER_SIZE);
      xprintfW(wszMsg, wbuf, wszFile);
      API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);

      nResult=ESD_OPEN;
      break;
    }

    if (bBOM)
    {
      if (IsCodePageUnicode(nCodePage))
      {
        if (nCodePage == CP_UNICODE_UTF16LE)
          nWrite=API_WriteFile(hFile, "\xFF\xFE", 2, &dwBytesWritten, NULL);
        else if (nCodePage == CP_UNICODE_UTF16BE)
          nWrite=API_WriteFile(hFile, "\xFE\xFF", 2, &dwBytesWritten, NULL);
        else if (nCodePage == CP_UNICODE_UTF32LE)
          nWrite=API_WriteFile(hFile, "\xFF\xFE\x00\x00", 4, &dwBytesWritten, NULL);
        else if (nCodePage == CP_UNICODE_UTF32BE)
          nWrite=API_WriteFile(hFile, "\x00\x00\xFE\xFF", 4, &dwBytesWritten, NULL);
        else if (nCodePage == CP_UNICODE_UTF8)
          nWrite=API_WriteFile(hFile, "\xEF\xBB\xBF", 3, &dwBytesWritten, NULL);

        if (!nWrite)
        {
          nResult=ESD_WRITE;
          break;
        }
      }
      else bBOM=FALSE;
    }

    fsd.hWnd=hWnd;
    fsd.hFile=hFile;
    fsd.nCodePage=nCodePage;
    fsd.dwFlags=dwFlags;
    fsd.bResult=TRUE;
    FileStreamOut(&fsd);

    if (fsd.bResult == TRUE)
    {
      //Restore last write time
      if (wfd.dwFileAttributes != INVALID_FILE_ATTRIBUTES && moCur.bSaveTime)
        SetFileTime(hFile, NULL, NULL, &wfd.ftLastWriteTime);
    }
    break;
  }
  if (hFile != INVALID_HANDLE_VALUE)
    CloseHandle(hFile);

  //Change back file attributes
  if (wfd.dwFileAttributes != INVALID_FILE_ATTRIBUTES)
  {
    if (nStreamOffset) wszFile[nStreamOffset]=L'\0';

    if ((!(wfd.dwFileAttributes & FILE_ATTRIBUTE_ARCHIVE) && fsd.bResult == TRUE) || (wfd.dwFileAttributes & FILE_ATTRIBUTE_READONLY) || (wfd.dwFileAttributes & FILE_ATTRIBUTE_HIDDEN) || (wfd.dwFileAttributes & FILE_ATTRIBUTE_SYSTEM))
      SetFileAttributesWide(wszFile, wfd.dwFileAttributes|(fsd.bResult == TRUE?FILE_ATTRIBUTE_ARCHIVE:0));

    if (nStreamOffset) wszFile[nStreamOffset]=L':';
  }

  //Change back file security (UAC).
  if (bSetSecurity)
  {
    if (!AkelAdminSend(AAA_SECURITYRESTORE, L"") ||
        !AkelAdminSend(AAA_SECURITYFREE, L"") ||
        !moCur.bAkelAdminResident)
    {
      //Reset AkelAdmin
      AkelAdminExit();
    }
    bSetSecurity=FALSE;
  }

  if (fsd.bResult != -1)
  {
    if (fsd.bResult)
    {
      //Update file info
      if (dwFlags & SD_UPDATE)
      {
        if (IsEditActive(hWnd))
        {
          //Compare
          nFileCmp=xstrcmpiW(lpFrameCurrent->wszFile, wszFile);
          nCodePageCmp=lpFrameCurrent->ei.nCodePage - nCodePage;

          if (nStreamOffset) wszFile[nStreamOffset]=L'\0';
          GetFileWriteTimeWide(wszFile, &lpFrameCurrent->ft);
          if (nStreamOffset) wszFile[nStreamOffset]=L':';

          SetModifyStatus(lpFrameCurrent, FALSE);
          SetCodePageStatus(lpFrameCurrent, nCodePage, bBOM);

          if (nFileCmp)
          {
            lpFrameCurrent->nStreamOffset=nStreamOffset;
            lpFrameCurrent->nFileLen=(int)xstrcpynW(lpFrameCurrent->wszFile, wszFile, MAX_PATH);
            WideCharToMultiByte(CP_ACP, 0, lpFrameCurrent->wszFile, lpFrameCurrent->nFileLen + 1, lpFrameCurrent->szFile, MAX_PATH, NULL, NULL);
            UpdateTitle(lpFrameCurrent);
          }
          if (nFileCmp || nCodePageCmp)
            RecentFilesSaveFile(lpFrameCurrent);

          if ((dwFlags & SD_SELECTION) || nLostLine ||
              //Is output new line format is changed?
              (HIWORD(SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETNEWLINE, 0, 0)) != AELB_ASIS &&
               SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETLINENUMBER, AEGL_LINECOUNT, 0) > 1))
          {
            OpenDocument(hWnd, lpFrameCurrent->wszFile, OD_REOPEN, lpFrameCurrent->ei.nCodePage, lpFrameCurrent->ei.bBOM);
          }
        }
        else
        {
          FILETIME ft;
          FRAMEDATA *lpFrame;

          if (lpFrame=GetFrameDataFromEditWindow(hWnd))
          {
            //Compare
            nFileCmp=xstrcmpiW(lpFrame->wszFile, wszFile);

            if (nStreamOffset) wszFile[nStreamOffset]=L'\0';
            GetFileWriteTimeWide(wszFile, &ft);
            if (nStreamOffset) wszFile[nStreamOffset]=L':';

            SetModifyStatus(lpFrame, FALSE);
            lpFrame->ei.nCodePage=nCodePage;
            lpFrame->ei.bBOM=bBOM;
            lpFrame->ft=ft;

            if (nFileCmp)
            {
              lpFrame->nStreamOffset=nStreamOffset;
              lpFrame->nFileLen=(int)xstrcpynW(lpFrame->wszFile, wszFile, MAX_PATH);
              WideCharToMultiByte(CP_ACP, 0, lpFrame->wszFile, lpFrame->nFileLen + 1, lpFrame->szFile, MAX_PATH, NULL, NULL);
              UpdateTitle(lpFrame);
            }
          }
        }
      }
    }
    else nResult=ESD_STREAMOUT;
  }

  End:
  if (GetWindowLongPtrWide(hWnd, GWLP_ID) == ID_EDIT)
  {
    SendMessage(hMainWnd, AKDN_SAVEDOCUMENT_FINISH, (WPARAM)hWnd, nResult);
  }
  return nResult;
}

void FileStreamOut(FILESTREAMDATA *lpData)
{
  AESTREAMOUT aeso;

  aeso.dwCookie=(UINT_PTR)lpData;
  aeso.lpCallback=OutputStreamCallback;
  aeso.nNewLine=AELB_ASOUTPUT;
  if (lpData->dwFlags & SD_SELECTION)
    aeso.bColumnSel=(BOOL)SendMessage(lpData->hWnd, AEM_GETCOLUMNSEL, 0, 0);
  else
    aeso.bColumnSel=FALSE;
  SendMessage(lpData->hWnd, AEM_STREAMOUT, (lpData->dwFlags & SD_SELECTION)?(AESF_SELECTION|AESF_FILLSPACES):0, (LPARAM)&aeso);
  lpData->bResult=!aeso.dwError;
}

DWORD CALLBACK OutputStreamCallback(UINT_PTR dwCookie, wchar_t *wszBuf, DWORD dwBufBytesSize, DWORD *dwBufBytesDone)
{
  FILESTREAMDATA *lpData=(FILESTREAMDATA *)dwCookie;
  unsigned char *pDataToWrite=(unsigned char *)wszBuf;
  UINT_PTR dwBytesToWrite=dwBufBytesSize;
  UINT_PTR dwBytesWritten;

  if (lpData->nCodePage == CP_UNICODE_UTF16LE)
  {
  }
  else if (lpData->nCodePage == CP_UNICODE_UTF16BE)
  {
    ChangeTwoBytesOrder((unsigned char *)wszBuf, dwBytesToWrite, NULL);
  }
  else if (lpData->nCodePage == CP_UNICODE_UTF32LE)
  {
    dwBytesToWrite=UTF16toUTF32(wszBuf, dwBufBytesSize / sizeof(wchar_t), NULL, (unsigned long *)pcTranslateBuffer, TRANSLATE_BUFFER_SIZE / sizeof(unsigned long)) * sizeof(unsigned long);
    pDataToWrite=pcTranslateBuffer;
  }
  else if (lpData->nCodePage == CP_UNICODE_UTF32BE)
  {
    dwBytesToWrite=UTF16toUTF32(wszBuf, dwBufBytesSize / sizeof(wchar_t), NULL, (unsigned long *)pcTranslateBuffer, TRANSLATE_BUFFER_SIZE / sizeof(unsigned long)) * sizeof(unsigned long);
    ChangeFourBytesOrder(pcTranslateBuffer, dwBytesToWrite, NULL);
    pDataToWrite=pcTranslateBuffer;
  }
  else
  {
    if (lpData->nCodePage == CP_UNICODE_UTF8)
      dwBytesToWrite=UTF16toUTF8(wszBuf, dwBufBytesSize / sizeof(wchar_t), NULL, pcTranslateBuffer, TRANSLATE_BUFFER_SIZE);
    else
      dwBytesToWrite=WideCharToMultiByte(lpData->nCodePage, 0, wszBuf, dwBufBytesSize / sizeof(wchar_t), (char *)pcTranslateBuffer, TRANSLATE_BUFFER_SIZE, NULL, NULL);
    pDataToWrite=pcTranslateBuffer;
  }
  if (API_WriteFile(lpData->hFile, pDataToWrite, dwBytesToWrite, &dwBytesWritten, NULL))
  {
    *dwBufBytesDone=(DWORD)dwBytesWritten;
    return 0;
  }
  return 1;
}

BOOL OpenDirectory(wchar_t *wpPath, BOOL bSubDir)
{
  wchar_t wszName[MAX_PATH];
  WIN32_FIND_DATAW wfd;
  HANDLE hSearch;
  int nOpen;
  BOOL bResult=TRUE;

  xprintfW(wszName, L"%s\\*.*", wpPath);

  if ((hSearch=FindFirstFileWide(wszName, &wfd)) != INVALID_HANDLE_VALUE)
  {
    do
    {
      if (wfd.cFileName[0] == L'.' && (wfd.cFileName[1] == L'\0' || (wfd.cFileName[1] == L'.' && wfd.cFileName[2] == L'\0'))) continue;

      xprintfW(wszName, L"%s\\%s", wpPath, wfd.cFileName);

      if (wfd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
      {
        if (bSubDir)
        {
          if (!(bResult=OpenDirectory(wszName, TRUE)))
            break;
        }
      }
      else
      {
        nOpen=OpenDocument(NULL, wszName, OD_ADT_BINARY_ERROR|OD_ADT_REG_CODEPAGE|OD_MULTIFILE, 0, FALSE);
        if (nOpen != EOD_SUCCESS && nOpen != EOD_MSGNO && nOpen != EOD_WINDOW_EXIST)
        {
          bResult=FALSE;
          break;
        }
      }
    }
    while (FindNextFileWide(hSearch, &wfd));

    FindClose(hSearch);
  }
  return bResult;
}

void DropFiles(HDROP hDrop)
{
  wchar_t wszFile[MAX_PATH];
  wchar_t wszString[MAX_PATH];
  MSG msg;
  int nDropped;
  int nOpen;
  int i;

  if (moCur.bStatusBar)
    API_LoadStringW(hLangLib, STR_COUNT, wszString, MAX_PATH);

  if (nMDI || SaveChanged(0))
  {
    nDropped=DragQueryFileWide(hDrop, 0xFFFFFFFF, NULL, 0);

    for (i=0; i < nDropped; ++i)
    {
      DragQueryFileWide(hDrop, i, wszFile, MAX_PATH);
      if (nMDI && IsFile(wszFile) == ERROR_DIRECTORY)
      {
        if (!OpenDirectory(wszFile, TRUE))
          break;
      }
      else
      {
        nOpen=OpenDocument(NULL, wszFile, OD_ADT_BINARY_ERROR|OD_ADT_REG_CODEPAGE|(i + 1 < nDropped?OD_MULTIFILE:0), 0, FALSE);
        if (nOpen != EOD_SUCCESS && nOpen != EOD_MSGNO && nOpen != EOD_WINDOW_EXIST)
          break;
      }
      if (nMDI == WMD_SDI) break;

      //Status update
      if (moCur.bStatusBar)
      {
        xprintfW(wbuf, wszString, i + 1, nDropped);
        StatusBar_SetTextWide(hStatus, SBP_MODIFY, wbuf);

        while (PeekMessageWide(&msg, hStatus, 0, 0, PM_REMOVE))
        {
          TranslateMessage(&msg);
          DispatchMessageWide(&msg);
        }
      }

      //Win7: prevent system from mark program as hanged
      PeekMessageWide(&msg, hMainWnd, 0, 0, PM_NOREMOVE);
    }

    if (moCur.bStatusBar)
      StatusBar_SetTextWide(hStatus, SBP_MODIFY, L"");
  }
  DragFinish(hDrop);
}

BOOL IsAllowWatchFile(FRAMEDATA *lpFrame)
{
  if (moCur.bWatchFile && lpFrame->wszFile[0] && (lpFrame->ft.dwLowDateTime || lpFrame->ft.dwHighDateTime))
    return TRUE;
  return FALSE;
}

BOOL CALLBACK SaveAllAsDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndCodePageCheck;
  static HWND hWndCodePageList;
  static HWND hWndBOM;
  static HWND hWndNewLineCheck;
  static HWND hWndNewLineWin;
  static HWND hWndNewLineUnix;
  static HWND hWndNewLineMac;
  static HWND hWndOK;
  static BOOL bCodePageEnable=TRUE;
  static BOOL bNewLineEnable=FALSE;
  static int nNewLine=-1;
  static int nCodePage=-1;
  static BOOL bBOM;

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndCodePageCheck=GetDlgItem(hDlg, IDC_SAVEALLAS_CODEPAGE_CHECK);
    hWndCodePageList=GetDlgItem(hDlg, IDC_SAVEALLAS_CODEPAGE_LIST);
    hWndBOM=GetDlgItem(hDlg, IDC_SAVEALLAS_BOM);
    hWndNewLineCheck=GetDlgItem(hDlg, IDC_SAVEALLAS_NEWLINE_CHECK);
    hWndNewLineWin=GetDlgItem(hDlg, IDC_SAVEALLAS_NEWLINE_WIN);
    hWndNewLineUnix=GetDlgItem(hDlg, IDC_SAVEALLAS_NEWLINE_UNIX);
    hWndNewLineMac=GetDlgItem(hDlg, IDC_SAVEALLAS_NEWLINE_MAC);
    hWndOK=GetDlgItem(hDlg, IDOK);

    if (nCodePage < 0)
      nCodePage=moCur.nNewFileCodePage;
    if (nNewLine < 0)
      nNewLine=moCur.nNewFileNewLine;

    FillComboboxCodepage(hWndCodePageList, lpCodepageList);
    SelectComboboxCodepage(hWndCodePageList, nCodePage);

    if (IsCodePageUnicode(nCodePage))
    {
      if (bBOM)
        SendMessage(hWndBOM, BM_SETCHECK, BST_CHECKED, 0);
    }
    if (bCodePageEnable)
      SendMessage(hWndCodePageCheck, BM_SETCHECK, BST_CHECKED, 0);
    if (bNewLineEnable)
      SendMessage(hWndNewLineCheck, BM_SETCHECK, BST_CHECKED, 0);
    if (nNewLine == NEWLINE_WIN)
      SendMessage(hWndNewLineWin, BM_SETCHECK, BST_CHECKED, 0);
    else if (nNewLine == NEWLINE_UNIX)
      SendMessage(hWndNewLineUnix, BM_SETCHECK, BST_CHECKED, 0);
    else if (nNewLine == NEWLINE_MAC)
      SendMessage(hWndNewLineMac, BM_SETCHECK, BST_CHECKED, 0);

    SendMessage(hDlg, WM_COMMAND, IDC_SAVEALLAS_CODEPAGE_CHECK, 0);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_SAVEALLAS_CODEPAGE_CHECK ||
        LOWORD(wParam) == IDC_SAVEALLAS_NEWLINE_CHECK)
    {
      bCodePageEnable=(BOOL)SendMessage(hWndCodePageCheck, BM_GETCHECK, 0, 0);
      bNewLineEnable=(BOOL)SendMessage(hWndNewLineCheck, BM_GETCHECK, 0, 0);

      if (!bCodePageEnable && !bNewLineEnable)
        EnableWindow(hWndOK, FALSE);
      else
        EnableWindow(hWndOK, TRUE);
      EnableWindow(hWndCodePageList, bCodePageEnable);
      EnableWindow(hWndBOM, bCodePageEnable);

      if (bCodePageEnable)
      {
        if (!IsCodePageUnicode(nCodePage))
        {
          EnableWindow(hWndBOM, FALSE);
        }
      }
      EnableWindow(hWndNewLineWin, bNewLineEnable);
      EnableWindow(hWndNewLineUnix, bNewLineEnable);
      EnableWindow(hWndNewLineMac, bNewLineEnable);
    }
    else if (LOWORD(wParam) == IDC_SAVEALLAS_CODEPAGE_LIST && HIWORD(wParam) == CBN_SELCHANGE)
    {
      nCodePage=GetDlgItemInt(hDlg, IDC_SAVEALLAS_CODEPAGE_LIST, NULL, FALSE);

      if (!IsCodePageUnicode(nCodePage))
      {
        SendMessage(hWndBOM, BM_SETCHECK, BST_UNCHECKED, 0);
        EnableWindow(hWndBOM, FALSE);
      }
      else
      {
        EnableWindow(hWndBOM, TRUE);
        SendMessage(hWndBOM, BM_SETCHECK, BST_CHECKED, 0);
      }
    }
    else if (LOWORD(wParam) == IDOK)
    {
      bBOM=(BOOL)SendMessage(hWndBOM, BM_GETCHECK, 0, 0);

      if (SendMessage(hWndNewLineWin, BM_GETCHECK, 0, 0) == BST_CHECKED)
        nNewLine=NEWLINE_WIN;
      else if (SendMessage(hWndNewLineUnix, BM_GETCHECK, 0, 0) == BST_CHECKED)
        nNewLine=NEWLINE_UNIX;
      else if (SendMessage(hWndNewLineMac, BM_GETCHECK, 0, 0) == BST_CHECKED)
        nNewLine=NEWLINE_MAC;

      EndDialog(hDlg, 1);

      //Save documents
      if (lpFrameCurrent->hWndEditParent || !nMDI)
      {
        FRAMEDATA *lpFrameInit=lpFrameCurrent;

        do
        {
          if ((bNewLineEnable && (nNewLine != lpFrameCurrent->ei.nNewLine)) ||
              (bCodePageEnable && (nCodePage != lpFrameCurrent->ei.nCodePage || bBOM != lpFrameCurrent->ei.bBOM || !lpFrameCurrent->wszFile[0] || lpFrameCurrent->ei.bModified || !FileExistsWide(lpFrameCurrent->wszFile))))
          {
            if (bNewLineEnable && nNewLine != lpFrameCurrent->ei.nNewLine)
              SetNewLineStatus(lpFrameCurrent, nNewLine, AENL_INPUT|AENL_OUTPUT);

            if (!lpFrameCurrent->wszFile[0])
            {
              if (!DoFileSaveAs(bCodePageEnable?nCodePage:lpFrameCurrent->ei.nCodePage, bCodePageEnable?bBOM:lpFrameCurrent->ei.bBOM))
                break;
            }
            else
            {
              if (SaveDocument(NULL, lpFrameCurrent->wszFile, bCodePageEnable?nCodePage:lpFrameCurrent->ei.nCodePage, bCodePageEnable?bBOM:lpFrameCurrent->ei.bBOM, SD_UPDATE) != ESD_SUCCESS)
                break;
            }
          }
          if (!nMDI) break;

          lpFrameCurrent=ActivateNextFrameWindow(lpFrameCurrent, FALSE);
        }
        while (lpFrameCurrent != lpFrameInit);
      }
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


//// Print

unsigned int CALLBACK PrintPageSetupDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndFontGroup;
  static HWND hWndFontCheck;
  static HWND hWndFontButton;
  static HWND hWndColorGroup;
  static HWND hWndColorTextCheck;
  static HWND hWndColorBkCheck;
  static HWND hWndHeadlineGroup;
  static HWND hWndHeaderCheck;
  static HWND hWndHeaderEdit;
  static HWND hWndHeaderHelp;
  static HWND hWndFooterCheck;
  static HWND hWndFooterEdit;
  static HWND hWndFooterHelp;
  static HFONT hPrintFont;
  static HFONT hGuiFont;
  static LOGFONTW lfDialog;
  BOOL bState;
  int i;

  if (uMsg == WM_INITDIALOG)
  {
    HWND hWndOK;
    HWND hWndCancel;
    HWND hWndPrinter;
    RECT rcDlg;
    RECT rcControl;
    BUTTONDRAW bd;
    int nExtend;

    xmemcpy(&lfDialog, &moCur.lfPrintFont, sizeof(LOGFONTW));
    hPrintFont=(HFONT)CreateFontIndirectWide(&lfDialog);
    hGuiFont=(HFONT)GetStockObject(DEFAULT_GUI_FONT);
    hWndOK=GetDlgItem(hDlg, IDOK);
    hWndCancel=GetDlgItem(hDlg, IDCANCEL);
    hWndPrinter=GetDlgItem(hDlg, IDC_PSD_PRINTER_BUTTON);

    //Get scale factor for ScaleX and ScaleY
    ScaleInit(NULL, hDlg);
    nExtend=ScaleY(155);

    GetWindowRect(hDlg, &rcDlg);
    rcDlg.right-=rcDlg.left;
    rcDlg.bottom-=rcDlg.top;
    if (rcDlg.top - nExtend >= 0)
      rcDlg.top-=nExtend;
    else
      rcDlg.top=0;
    SetWindowPos(hDlg, NULL, rcDlg.left, rcDlg.top, rcDlg.right, rcDlg.bottom + nExtend, SWP_NOZORDER|SWP_NOACTIVATE);

    if (hWndPrinter)
    {
      GetWindowPos(hWndPrinter, hDlg, &rcControl);
      SetWindowPos(hWndPrinter, NULL, rcControl.left, rcControl.top + nExtend, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
    }
    if (hWndOK)
    {
      GetWindowPos(hWndOK, hDlg, &rcControl);
      SetWindowPos(hWndOK, NULL, rcControl.left, rcControl.top + nExtend, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
    }
    if (hWndCancel)
    {
      GetWindowPos(hWndCancel, hDlg, &rcControl);
      SetWindowPos(hWndCancel, NULL, rcControl.left, rcControl.top + nExtend, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
    }

    //Group box
    hWndFontGroup=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|BS_GROUPBOX,
                         ScaleX(12), rcControl.top - ScaleY(6), ScaleX(96), ScaleY(70),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_FONT_GROUP,
                         hInstance,
                         NULL);
    //Font checkbox
    hWndFontCheck=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_AUTOCHECKBOX,
                         ScaleX(24), rcControl.top + ScaleY(26), ScaleX(16), ScaleY(16),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_FONT_CHECK,
                         hInstance,
                         NULL);
    //Font button
    hWndFontButton=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP,
                         ScaleX(46), rcControl.top + ScaleY(22), ScaleX(46), ScaleY(23),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_FONT_BUTTON,
                         hInstance,
                         NULL);
    //Color group box
    hWndColorGroup=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|BS_GROUPBOX,
                         ScaleX(120), rcControl.top - ScaleY(6), ScaleX(228), ScaleY(70),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_COLOR_GROUP,
                         hInstance,
                         NULL);
    //Color text checkbox
    hWndColorTextCheck=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_AUTOCHECKBOX,
                         ScaleX(136), rcControl.top + ScaleY(16), ScaleX(200), ScaleY(16),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_COLOR_TEXT_CHECK,
                         hInstance,
                         NULL);
    //Color background checkbox
    hWndColorBkCheck=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_AUTOCHECKBOX,
                         ScaleX(136), rcControl.top + ScaleY(36), ScaleX(200), ScaleY(16),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_COLOR_BACKGROUND_CHECK,
                         hInstance,
                         NULL);
    //Headline group box
    hWndHeadlineGroup=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|BS_GROUPBOX,
                         ScaleX(12), rcControl.top + ScaleY(71), ScaleX(336), ScaleY(70),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_HEADLINE_GROUP,
                         hInstance,
                         NULL);
    //Header checkbox
    hWndHeaderCheck=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_AUTOCHECKBOX|BS_RIGHTBUTTON,
                         ScaleX(18), rcControl.top + ScaleY(93), ScaleX(76), ScaleY(16),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_HEADLINE_HEADER_CHECK,
                         hInstance,
                         NULL);
    //Header edit
    hWndHeaderEdit=CreateWindowExWide(WS_EX_CLIENTEDGE,
                         L"EDIT",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP|ES_AUTOHSCROLL,
                         ScaleX(102), rcControl.top + ScaleY(91), ScaleX(216), ScaleY(18),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_HEADLINE_HEADER_EDIT,
                         hInstance,
                         NULL);
    //Header help
    hWndHeaderHelp=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_OWNERDRAW,
                         ScaleX(320), rcControl.top + ScaleY(91), ScaleX(18), ScaleY(18),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_HEADLINE_HEADER_HELP,
                         hInstance,
                         NULL);
    //Footer checkbox
    hWndFooterCheck=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_AUTOCHECKBOX|BS_RIGHTBUTTON,
                         ScaleX(18), rcControl.top + ScaleY(113), ScaleX(76), ScaleY(16),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_HEADLINE_FOOTER_CHECK,
                         hInstance,
                         NULL);
    //Footer edit
    hWndFooterEdit=CreateWindowExWide(WS_EX_CLIENTEDGE,
                         L"EDIT",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP|ES_AUTOHSCROLL,
                         ScaleX(102), rcControl.top + ScaleY(111), ScaleX(216), ScaleY(18),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_HEADLINE_FOOTER_EDIT,
                         hInstance,
                         NULL);
    //Footer help
    hWndFooterHelp=CreateWindowExWide(0,
                         L"BUTTON",
                         NULL,
                         WS_CHILD|WS_VISIBLE|WS_TABSTOP|BS_OWNERDRAW,
                         ScaleX(320), rcControl.top + ScaleY(111), ScaleX(18), ScaleY(18),
                         hDlg,
                         (HMENU)(UINT_PTR)IDC_PSD_HEADLINE_FOOTER_HELP,
                         hInstance,
                         NULL);

    SendMessage(hWndFontGroup, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
    API_LoadStringW(hLangLib, STR_FONT, wbuf, BUFFER_SIZE);
    SetWindowTextWide(hWndFontGroup, wbuf);
    if (moCur.bPrintFontEnable)
      SendMessage(hWndFontCheck, BM_SETCHECK, BST_CHECKED, 0);
    else
      EnableWindow(hWndFontButton, FALSE);
    SendMessage(hWndFontButton, WM_SETFONT, (WPARAM)hPrintFont, TRUE);
    SetWindowTextWide(hWndFontButton, L"...");

    SendMessage(hWndColorGroup, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
    API_LoadStringW(hLangLib, STR_COLOR, wbuf, BUFFER_SIZE);
    SetWindowTextWide(hWndColorGroup, wbuf);

    SendMessage(hWndColorTextCheck, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
    API_LoadStringW(hLangLib, STR_TEXT, wbuf, BUFFER_SIZE);
    SetWindowTextWide(hWndColorTextCheck, wbuf);
    if (moCur.dwPrintColor & PRNC_TEXT)
      SendMessage(hWndColorTextCheck, BM_SETCHECK, BST_CHECKED, 0);

    SendMessage(hWndColorBkCheck, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
    API_LoadStringW(hLangLib, STR_BACKGROUND, wbuf, BUFFER_SIZE);
    SetWindowTextWide(hWndColorBkCheck, wbuf);
    if (moCur.dwPrintColor & PRNC_BACKGROUND)
      SendMessage(hWndColorBkCheck, BM_SETCHECK, BST_CHECKED, 0);

    SendMessage(hWndHeadlineGroup, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
    API_LoadStringW(hLangLib, STR_HEADLINE, wbuf, BUFFER_SIZE);
    SetWindowTextWide(hWndHeadlineGroup, wbuf);

    SendMessage(hWndHeaderCheck, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
    API_LoadStringW(hLangLib, STR_TOP, wbuf, BUFFER_SIZE);
    SetWindowTextWide(hWndHeaderCheck, wbuf);
    if (moCur.bPrintHeaderEnable)
      SendMessage(hWndHeaderCheck, BM_SETCHECK, BST_CHECKED, 0);
    else
    {
      EnableWindow(hWndHeaderEdit, FALSE);
      EnableWindow(hWndHeaderHelp, FALSE);
    }
    SendMessage(hWndHeaderEdit, EM_LIMITTEXT, MAX_PATH, 0);
    SendMessage(hWndHeaderEdit, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
    SetWindowTextWide(hWndHeaderEdit, moCur.wszPrintHeader);

    //Set header menu image
    bd.dwFlags=BIF_DOWNARROW|BIF_ENABLEFOCUS;
    SetButtonDraw(hWndHeaderHelp, &bd);

    SendMessage(hWndFooterCheck, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
    API_LoadStringW(hLangLib, STR_BOTTOM, wbuf, BUFFER_SIZE);
    SetWindowTextWide(hWndFooterCheck, wbuf);
    if (moCur.bPrintFooterEnable)
      SendMessage(hWndFooterCheck, BM_SETCHECK, BST_CHECKED, 0);
    else
    {
      EnableWindow(hWndFooterEdit, FALSE);
      EnableWindow(hWndFooterHelp, FALSE);
    }
    SendMessage(hWndFooterEdit, EM_LIMITTEXT, MAX_PATH, 0);
    SendMessage(hWndFooterEdit, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
    SetWindowTextWide(hWndFooterEdit, moCur.wszPrintFooter);

    //Set footer menu image
    bd.dwFlags=BIF_DOWNARROW|BIF_ENABLEFOCUS;
    SetButtonDraw(hWndFooterHelp, &bd);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_PSD_FONT_CHECK)
    {
      bState=(BOOL)SendMessage(hWndFontCheck, BM_GETCHECK, 0, 0);
      EnableWindow(hWndFontButton, bState);
    }
    else if (LOWORD(wParam) == IDC_PSD_HEADLINE_HEADER_CHECK)
    {
      bState=(BOOL)SendMessage(hWndHeaderCheck, BM_GETCHECK, 0, 0);
      EnableWindow(hWndHeaderEdit, bState);
      EnableWindow(hWndHeaderHelp, bState);
    }
    else if (LOWORD(wParam) == IDC_PSD_HEADLINE_FOOTER_CHECK)
    {
      bState=(BOOL)SendMessage(hWndFooterCheck, BM_GETCHECK, 0, 0);
      EnableWindow(hWndFooterEdit, bState);
      EnableWindow(hWndFooterHelp, bState);
    }
    else if (LOWORD(wParam) == IDC_PSD_HEADLINE_HEADER_HELP ||
             LOWORD(wParam) == IDC_PSD_HEADLINE_FOOTER_HELP)
    {
      HWND hWndInsert=NULL;
      HWND hWndButton=NULL;
      RECT rc;

      if (LOWORD(wParam) == IDC_PSD_HEADLINE_HEADER_HELP)
      {
        hWndInsert=hWndHeaderEdit;
        hWndButton=hWndHeaderHelp;
      }
      else if (LOWORD(wParam) == IDC_PSD_HEADLINE_FOOTER_HELP)
      {
        hWndInsert=hWndFooterEdit;
        hWndButton=hWndFooterHelp;
      }

      GetWindowRect(hWndButton, &rc);
      if (i=TrackPopupMenu(hPopupHeadline, TPM_NONOTIFY|TPM_RETURNCMD|TPM_LEFTBUTTON|TPM_RIGHTBUTTON, rc.left, rc.top + (rc.bottom - rc.top), 0, hDlg, NULL))
      {
        if (i == IDM_POPUP_HEADLINE_SYMBOL)
          SendMessageA(hWndInsert, EM_REPLACESEL, TRUE, (LPARAM)"%%");
        else if (i == IDM_POPUP_HEADLINE_PAGE_NUMBER)
          SendMessageA(hWndInsert, EM_REPLACESEL, TRUE, (LPARAM)"%n[1]");
        else if (i == IDM_POPUP_HEADLINE_FILENAME)
          SendMessageA(hWndInsert, EM_REPLACESEL, TRUE, (LPARAM)"%f");
        else if (i == IDM_POPUP_HEADLINE_FILEDIR)
          SendMessageA(hWndInsert, EM_REPLACESEL, TRUE, (LPARAM)"%d");
        else if (i == IDM_POPUP_HEADLINE_ALIGN_CENTER)
          SendMessageA(hWndInsert, EM_REPLACESEL, TRUE, (LPARAM)"%c");
        else if (i == IDM_POPUP_HEADLINE_ALIGN_LEFT)
          SendMessageA(hWndInsert, EM_REPLACESEL, TRUE, (LPARAM)"%l");
        else if (i == IDM_POPUP_HEADLINE_ALIGN_RIGHT)
          SendMessageA(hWndInsert, EM_REPLACESEL, TRUE, (LPARAM)"%r");
        SetFocus(hWndInsert);
      }
    }
    else if (LOWORD(wParam) == IDC_PSD_FONT_BUTTON)
    {
      if (DoViewFont(hDlg, &lfDialog))
      {
        if (hPrintFont) DeleteObject(hPrintFont);
        hPrintFont=(HFONT)CreateFontIndirectWide(&lfDialog);
        SendMessage(hWndFontButton, WM_SETFONT, (WPARAM)hPrintFont, TRUE);
      }
    }
    else if (LOWORD(wParam) == IDOK)
    {
      if (hPrintFont) DeleteObject(hPrintFont);

      moCur.bPrintFontEnable=(BOOL)SendMessage(hWndFontCheck, BM_GETCHECK, 0, 0);
      xmemcpy(&moCur.lfPrintFont, &lfDialog, sizeof(LOGFONTW));

      moCur.dwPrintColor=0;
      if (SendMessage(hWndColorTextCheck, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwPrintColor|=PRNC_TEXT;
      if (SendMessage(hWndColorBkCheck, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwPrintColor|=PRNC_BACKGROUND;

      moCur.bPrintHeaderEnable=(BOOL)SendMessage(hWndHeaderCheck, BM_GETCHECK, 0, 0);
      GetWindowTextWide(hWndHeaderEdit, moCur.wszPrintHeader, MAX_PATH);

      moCur.bPrintFooterEnable=(BOOL)SendMessage(hWndFooterCheck, BM_GETCHECK, 0, 0);
      GetWindowTextWide(hWndFooterEdit, moCur.wszPrintFooter, MAX_PATH);
    }
    else if (LOWORD(wParam) == IDCANCEL)
    {
      if (hPrintFont) DeleteObject(hPrintFont);
    }
  }
  return FALSE;
}

BOOL GetPrinter(HWND hWndOwner, PRINTINFO *prninfo, BOOL bSilent)
{
  if (bOldWindows)
    return GetPrinterA(hWndOwner, prninfo, bSilent);
  else
    return GetPrinterW(hWndOwner, prninfo, bSilent);
}

BOOL GetPrinterA(HWND hWndOwner, PRINTINFO *prninfo, BOOL bSilent)
{
  PRINTDLGA pdA;

  xmemset(&pdA, 0, sizeof(PRINTDLGA));
  pdA.lStructSize =sizeof(PRINTDLGA);
  pdA.hwndOwner   =hWndOwner;
  pdA.Flags       =prninfo->dwPrintFlags|PD_RETURNDC;
  pdA.nMinPage    =1;
  pdA.nMaxPage    =9999;
  pdA.nFromPage   =prninfo->nFromPage;
  pdA.nToPage     =prninfo->nToPage;
  pdA.nCopies     =prninfo->nCopies;
  pdA.hDC         =prninfo->hDC;
  pdA.hDevMode    =prninfo->hDevMode;
  pdA.hDevNames   =prninfo->hDevNames;

  if (bSilent)
  {
    DEVMODEA *hDevMode;
    DEVNAMES *hDevNames;

    //If printer already had been choosen
    if (pdA.hDevMode && pdA.hDevNames)
    {
      hDevMode=(DEVMODEA *)GlobalLock(pdA.hDevMode);
      hDevNames=(DEVNAMES *)GlobalLock(pdA.hDevNames);
      pdA.hDC=CreateDCA((char *)hDevNames + hDevNames->wDriverOffset, (char *)hDevNames + hDevNames->wDeviceOffset, NULL, hDevMode);
      GlobalUnlock(pdA.hDevMode);
      GlobalUnlock(pdA.hDevNames);
    }

    //Get default printer
    if (!pdA.hDC)
    {
      pdA.Flags     =PD_RETURNDEFAULT|PD_RETURNDC;
      pdA.hDevMode  =NULL;
      pdA.hDevNames =NULL;

      if (!PrintDlgA(&pdA))
        return FALSE;
    }
  }
  else
  {
    //Display printer dialog
    if (!PrintDlgA(&pdA))
      return FALSE;
    prninfo->dwPrintFlags=pdA.Flags;
    prninfo->nFromPage=pdA.nFromPage;
    prninfo->nToPage=pdA.nToPage;
    prninfo->nCopies=pdA.nCopies;
  }
  prninfo->hDC=pdA.hDC;
  prninfo->hDevMode=pdA.hDevMode;
  prninfo->hDevNames=pdA.hDevNames;
  return TRUE;
}

BOOL GetPrinterW(HWND hWndOwner, PRINTINFO *prninfo, BOOL bSilent)
{
  PRINTDLGW pdW;

  xmemset(&pdW, 0, sizeof(PRINTDLGW));
  pdW.lStructSize =sizeof(PRINTDLGW);
  pdW.hwndOwner   =hWndOwner;
  pdW.Flags       =prninfo->dwPrintFlags|PD_RETURNDC;
  pdW.nMinPage    =1;
  pdW.nMaxPage    =9999;
  pdW.nFromPage   =prninfo->nFromPage;
  pdW.nToPage     =prninfo->nToPage;
  pdW.nCopies     =prninfo->nCopies;
  pdW.hDC         =prninfo->hDC;
  pdW.hDevMode    =prninfo->hDevMode;
  pdW.hDevNames   =prninfo->hDevNames;

  if (bSilent)
  {
    DEVMODEW *hDevMode;
    DEVNAMES *hDevNames;

    //If printer already had been choosen
    if (pdW.hDevMode && pdW.hDevNames)
    {
      hDevMode=(DEVMODEW *)GlobalLock(pdW.hDevMode);
      hDevNames=(DEVNAMES *)GlobalLock(pdW.hDevNames);
      pdW.hDC=CreateDCW((wchar_t *)hDevNames + hDevNames->wDriverOffset, (wchar_t *)hDevNames + hDevNames->wDeviceOffset, NULL, hDevMode);
      GlobalUnlock(pdW.hDevMode);
      GlobalUnlock(pdW.hDevNames);
    }

    //Get default printer
    if (!pdW.hDC)
    {
      pdW.Flags     =PD_RETURNDEFAULT|PD_RETURNDC;
      pdW.hDevMode  =NULL;
      pdW.hDevNames =NULL;

      if (!PrintDlgW(&pdW))
        return FALSE;
    }
  }
  else
  {
    //Display printer dialog
    if (!PrintDlgW(&pdW))
      return FALSE;
    prninfo->dwPrintFlags=pdW.Flags;
    prninfo->nFromPage=pdW.nFromPage;
    prninfo->nToPage=pdW.nToPage;
    prninfo->nCopies=pdW.nCopies;
  }
  prninfo->hDC=pdW.hDC;
  prninfo->hDevMode=pdW.hDevMode;
  prninfo->hDevNames=pdW.hDevNames;
  return TRUE;
}

DWORD GetMappedPrintWidth(HWND hWnd)
{
  AEPRINT prn;
  AEHPRINT hPrintDoc;
  int nAveCharWidth;
  DWORD dwWidth=0;

  if (!GetPrinter(hMainWnd, &prninfo, TRUE))
    return 0;

  if (prninfo.hDC)
  {
    xmemset(&prn, 0, sizeof(AEPRINT));
    prn.dwFlags=(prninfo.dwPageFlags & PSD_INHUNDREDTHSOFMILLIMETERS?AEPRN_INHUNDREDTHSOFMILLIMETERS:AEPRN_INTHOUSANDTHSOFINCHES);
    prn.hPrinterDC=prninfo.hDC;
    prn.rcMargins=prninfo.rtMargin;
    prn.hEditFont=(HFONT)SendMessage(hWnd, WM_GETFONT, 0, 0);

    if (hPrintDoc=(AEHPRINT)SendMessage(hWnd, AEM_STARTPRINTDOC, 0, (LPARAM)&prn))
    {
      nAveCharWidth=(int)SendMessage(hWnd, AEM_GETCHARSIZE, AECS_AVEWIDTH, 0);
      dwWidth=MulDiv(nAveCharWidth, prn.rcPageIn.right - prn.rcPageIn.left, prn.nAveCharWidth);
      SendMessage(hWnd, AEM_ENDPRINTDOC, (WPARAM)hPrintDoc, (LPARAM)&prn);
    }
    DeleteDC(prninfo.hDC);
    prninfo.hDC=NULL;
  }
  return dwWidth;
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
BOOL UpdateMappedPrintWidth(FRAMEDATA *lpFrame)
{
  if (lpFrame == lpFrameCurrent)
  {
    if (lpFrame->dwMarker == (DWORD)-1 || (lpFrame->ei.bWordWrap && lpFrame->dwWrapLimit == (DWORD)-1))
    {
      lpFrame->dwMappedPrintWidth=GetMappedPrintWidth(lpFrame->ei.hWndEdit);
      if (lpFrame->dwMarker == (DWORD)-1)
        SetMarker(lpFrame, lpFrame->dwMarker);
      if (lpFrame->ei.bWordWrap && lpFrame->dwWrapLimit == (DWORD)-1)
        SetWordWrap(lpFrame, lpFrame->dwWrapType, lpFrame->dwWrapLimit);
      return TRUE;
    }
  }
  return FALSE;
}

int PrintDocument(HWND hWnd, AEPRINT *prn, DWORD dwFlags, int nInitPage)
{
  DOCINFOW diW;
  POINT ptScreenDpi={0};
  POINT ptPrintDpi;
  PRINTPAGE *lpElement;
  AECHARRANGE crInitText;
  AECHARRANGE crInitPage;
  AEHPRINT hPrintDoc;
  HFONT hPrintFontOld;
  HDC hScreenDC;
  RECT rcHeader;
  RECT rcFooter;
  int nPageNumber=nInitPage;
  BOOL bPrintError=FALSE;
  BOOL bPrintStop=FALSE;
  WORD nCopies=1;

  //Set print settings
  prn->dwFlags=((lpFrameCurrent->dwWrapType & AEWW_SYMBOL)?AEPRN_WRAPSYMBOL:AEPRN_WRAPWORD)|
               (prninfo.dwPageFlags & PSD_INHUNDREDTHSOFMILLIMETERS?AEPRN_INHUNDREDTHSOFMILLIMETERS:AEPRN_INTHOUSANDTHSOFINCHES)|
               (dwFlags & PRND_TEST?AEPRN_TEST:0)|
               (dwFlags & PRND_ANSI?AEPRN_ANSI:0)|
               (moCur.dwPrintColor & PRNC_TEXT?AEPRN_COLOREDTEXT:0)|
               (moCur.dwPrintColor & PRNC_BACKGROUND?AEPRN_COLOREDBACKGROUND:0)|
               (moCur.dwPrintColor & PRNC_SELECTION?AEPRN_COLOREDSELECTION:0);
  if (!(dwFlags & PRND_RANGE))
  {
    if (dwFlags & PRND_SELECTION)
    {
      if (!AEC_IndexCompare(&crCurSel.ciMin, &crCurSel.ciMax))
        return 0;
      prn->crText.ciMin=crCurSel.ciMin;
      prn->crText.ciMax=crCurSel.ciMax;
    }
    else if (dwFlags & PRND_ALLTEXT)
    {
      SendMessage(hWnd, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&prn->crText.ciMin);
      SendMessage(hWnd, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&prn->crText.ciMax);
    }
  }
  xmemcpy(&crInitText, &prn->crText, sizeof(AECHARRANGE));

  if (moCur.bPrintFontEnable)
    prn->hEditFont=CreateFontIndirectWide(&moCur.lfPrintFont);
  else
    prn->hEditFont=CreateFontIndirectWide(&lpFrameCurrent->lf);

  //Document properties
  if (dwFlags & PRND_REALPRINT)
  {
    API_LoadStringW(hLangLib, STR_DOCNAME, wbuf, BUFFER_SIZE);
    diW.cbSize=sizeof(DOCINFOW);
    diW.lpszDocName=(lpFrameCurrent->wszFile[0])?lpFrameCurrent->wszFile:wbuf;
    if (prninfo.dwPrintFlags & PD_PRINTTOFILE)
      diW.lpszOutput=L"FILE:";
    else
      diW.lpszOutput=NULL;
    diW.lpszDatatype=NULL;
    diW.fwType=0;
  }

  if (!(dwFlags & PRND_REALPRINT) || StartDocWide(prn->hPrinterDC, &diW) > 0)
  {
    if (hPrintDoc=(AEHPRINT)SendMessage(hWnd, AEM_STARTPRINTDOC, 0, (LPARAM)prn))
    {
      if (dwFlags & PRND_TEST)
      {
        //Get DPI for converting coordinates
        if (hScreenDC=GetDC(hWndPreviewDlg))
        {
          ptScreenDpi.x=GetDeviceCaps(hScreenDC, LOGPIXELSX);
          ptScreenDpi.y=GetDeviceCaps(hScreenDC, LOGPIXELSY);
          ReleaseDC(hWndPreviewDlg, hScreenDC);
        }
        ptPrintDpi.x=GetDeviceCaps(prn->hPrinterDC, LOGPIXELSX);
        ptPrintDpi.y=GetDeviceCaps(prn->hPrinterDC, LOGPIXELSY);

        //Paper in screen coordinates
        rcPreviewPaper.left=2;
        rcPreviewPaper.top=2;
        rcPreviewPaper.right=rcPreviewPaper.left + MulDiv(ptScreenDpi.x, RectW(&prn->rcPageFull), ptPrintDpi.x);
        rcPreviewPaper.bottom=rcPreviewPaper.top + MulDiv(ptScreenDpi.y, RectH(&prn->rcPageFull), ptPrintDpi.y);
      }

      if (moCur.bPrintHeaderEnable)
      {
        rcHeader=prn->rcPageIn;
        rcHeader.bottom=rcHeader.top + prn->nCharHeight;
        prn->rcPageIn.top+=prn->nCharHeight;
      }
      if (moCur.bPrintFooterEnable)
      {
        rcFooter=prn->rcPageIn;
        rcFooter.top=rcFooter.bottom - prn->nCharHeight;
        prn->rcPageIn.bottom-=prn->nCharHeight;
      }

      NextCopy:
      while (!bPrintStop && !bPrintError)
      {
        ++nPageNumber;

        if (dwFlags & PRND_REALPRINT)
        {
          if (prninfo.dwPrintFlags & PD_PAGENUMS)
          {
            if (nPageNumber > prninfo.nToPage)
            {
              bPrintStop=TRUE;
              continue;
            }
            else if (nPageNumber < prninfo.nFromPage)
            {
              prn->dwFlags|=AEPRN_TEST;
              if (!SendMessage(hWnd, AEM_PRINTPAGE, (WPARAM)hPrintDoc, (LPARAM)prn))
                bPrintStop=TRUE;
              prn->dwFlags&=~AEPRN_TEST;
              continue;
            }
          }
        }

        //Print page
        if (!(dwFlags & PRND_REALPRINT) || StartPage(prn->hPrinterDC) > 0)
        {
          if (dwFlags & PRND_TEST)
          {
            if (dwFlags & PRND_SELECTION)
            {
              if (lpElement=StackPageInsert(&hPreviewSelPagesStack))
              {
                lpElement->crText.ciMin=prn->crText.ciMin;
                lpElement->crText.ciMax=prn->crText.ciMax;
              }
            }
            else
            {
              if (lpElement=StackPageInsert(&hPreviewAllPagesStack))
              {
                lpElement->crText.ciMin=prn->crText.ciMin;
                lpElement->crText.ciMax=prn->crText.ciMax;
              }
            }
          }
          if ((dwFlags & PRND_REALPRINT) && !(prninfo.dwPrintFlags & PD_COLLATE) && nCopies == 1)
            xmemcpy(&crInitPage, &prn->crText, sizeof(AECHARRANGE));

          if (moCur.bPrintHeaderEnable)
          {
            hPrintFontOld=(HFONT)SelectObject(prn->hPrinterDC, prn->hPrintFont);
            if (!PrintHeadline(prn->hPrinterDC, &rcHeader, moCur.wszPrintHeader, nPageNumber))
              bPrintStop=TRUE;
            if (hPrintFontOld) SelectObject(prn->hPrinterDC, hPrintFontOld);
          }

          if (!SendMessage(hWnd, AEM_PRINTPAGE, (WPARAM)hPrintDoc, (LPARAM)prn))
            bPrintStop=TRUE;

          if (moCur.bPrintFooterEnable)
          {
            hPrintFontOld=(HFONT)SelectObject(prn->hPrinterDC, prn->hPrintFont);
            if (!PrintHeadline(prn->hPrinterDC, &rcFooter, moCur.wszPrintFooter, nPageNumber))
              bPrintStop=TRUE;
            if (hPrintFontOld) SelectObject(prn->hPrinterDC, hPrintFontOld);
          }

          if (dwFlags & PRND_REALPRINT)
          {
            if (EndPage(prn->hPrinterDC) <= 0)
              bPrintError=TRUE;
          }
          if (dwFlags & PRND_ONEPAGE)
            break;
        }
        else bPrintError=TRUE;

        if ((dwFlags & PRND_REALPRINT) && !(prninfo.dwPrintFlags & PD_COLLATE))
        {
          if (nCopies < prninfo.nCopies)
          {
            xmemcpy(&prn->crText, &crInitPage, sizeof(AECHARRANGE));
            ++nCopies;
            --nPageNumber;
            bPrintStop=FALSE;
          }
          else nCopies=1;
        }
      }
      if ((dwFlags & PRND_REALPRINT) && (prninfo.dwPrintFlags & PD_COLLATE) && nCopies < prninfo.nCopies)
      {
        xmemcpy(&prn->crText, &crInitText, sizeof(AECHARRANGE));
        ++nCopies;
        nPageNumber=nInitPage;
        bPrintStop=FALSE;
        goto NextCopy;
      }
      SendMessage(hWnd, AEM_ENDPRINTDOC, (WPARAM)hPrintDoc, (LPARAM)prn);
    }
    if (dwFlags & PRND_REALPRINT) EndDoc(prn->hPrinterDC);
  }
  else bPrintError=TRUE;

  if (bPrintError)
  {
    API_LoadStringW(hLangLib, MSG_ERROR_PRINT, wszMsg, BUFFER_SIZE);
    API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
  }
  if (prn->hEditFont)
  {
    DeleteObject(prn->hEditFont);
    prn->hEditFont=NULL;
  }
  return nPageNumber;
}

BOOL PrintHeadline(HDC hDC, RECT *rc, wchar_t *wpHeadline, int nPageNumber)
{
  //%% == %, %n[1] == nPageNumber[nPageStart], %f == lpFrameCurrent->wszFile, %d == GetFileDir(lpFrameCurrent->wszFile), %c == DT_CENTER, %l == DT_LEFT, %r == DT_RIGHT
  wchar_t wszFileDir[MAX_PATH];
  wchar_t wszCenter[MAX_PATH];
  wchar_t wszLeft[MAX_PATH];
  wchar_t wszRight[MAX_PATH];
  wchar_t *wpCount=&wszCenter[0];
  wchar_t wszPageNumber[32];
  const wchar_t *wpFile;
  HBRUSH hbrBasicBk;
  DWORD dwAlign;
  COLORREF crTextOld=(DWORD)-1;
  int nBkModeOld=0;
  int nCenter=0;
  int nLeft=0;
  int nRight=0;
  int nCount=nCenter;
  int nAlign=DT_CENTER;
  int nPageNumberLen;
  int nPageStart;
  int nFileLen;
  int nFileDirLen;
  int a;
  BOOL bResult=TRUE;

  wpFile=GetFileName(lpFrameCurrent->wszFile, lpFrameCurrent->nFileLen);
  nFileLen=lpFrameCurrent->nFileLen - (int)(wpFile - lpFrameCurrent->wszFile);
  nFileDirLen=GetFileDir(lpFrameCurrent->wszFile, lpFrameCurrent->nFileLen, wszFileDir, MAX_PATH);

  for (a=0; wpHeadline[a] && nCount < MAX_PATH; ++a)
  {
    if (wpHeadline[a] == L'%')
    {
      if (wpHeadline[++a] == L'%')
      {
        wpCount[nCount++]=L'%';
      }
      else if (wpHeadline[a] == L'c' || wpHeadline[a] == L'C' ||
               wpHeadline[a] == L'l' || wpHeadline[a] == L'L' ||
               wpHeadline[a] == L'r' || wpHeadline[a] == L'R')
      {
        if (nAlign == DT_CENTER)
          nCenter=nCount;
        else if (nAlign == DT_LEFT)
          nLeft=nCount;
        else if (nAlign == DT_RIGHT)
          nRight=nCount;

        if (wpHeadline[a] == L'c' || wpHeadline[a] == L'C')
        {
          nAlign=DT_CENTER;
          wpCount=&wszCenter[0];
          nCount=nCenter;
        }
        else if (wpHeadline[a] == L'l' || wpHeadline[a] == L'L')
        {
          nAlign=DT_LEFT;
          wpCount=&wszLeft[0];
          nCount=nLeft;
        }
        else if (wpHeadline[a] == L'r' || wpHeadline[a] == L'R')
        {
          nAlign=DT_RIGHT;
          wpCount=&wszRight[0];
          nCount=nRight;
        }
      }
      else if (wpHeadline[a] == L'n' || wpHeadline[a] == L'N')
      {
        if (wpHeadline[++a] == L'[')
        {
          for (++a, nPageNumberLen=0; wpHeadline[a] && wpHeadline[a] != L']' && nPageNumberLen < 30; ++a, ++nPageNumberLen)
            wszPageNumber[nPageNumberLen]=wpHeadline[a];
          wszPageNumber[nPageNumberLen]=L'\0';
          nPageStart=(int)xatoiW(wszPageNumber, NULL);
          nPageNumberLen=(int)xprintfW(wszPageNumber, L"%d", nPageNumber + nPageStart - 1);

          if (nCount + nPageNumberLen <= MAX_PATH)
          {
            xstrcpyW(wpCount + nCount, wszPageNumber);
            nCount+=nPageNumberLen;
          }
          else break;
        }
        else break;
      }
      else if (wpHeadline[a] == L'f' || wpHeadline[a] == L'F')
      {
        if (nCount + nFileLen <= MAX_PATH)
        {
          xstrcpyW(wpCount + nCount, wpFile);
          nCount+=nFileLen;
        }
        else break;
      }
      else if (wpHeadline[a] == L'd' || wpHeadline[a] == L'D')
      {
        if (nCount + nFileDirLen <= MAX_PATH)
        {
          xstrcpyW(wpCount + nCount, wszFileDir);
          nCount+=nFileDirLen;
        }
        else break;
      }
      else break;
    }
    else wpCount[nCount++]=wpHeadline[a];
  }
  if (nAlign == DT_CENTER)
    nCenter=nCount;
  else if (nAlign == DT_LEFT)
    nLeft=nCount;
  else if (nAlign == DT_RIGHT)
    nRight=nCount;

  if (nCenter < MAX_PATH)
    wszCenter[nCenter]=L'\0';
  if (nLeft < MAX_PATH)
    wszLeft[nLeft]=L'\0';
  if (nRight < MAX_PATH)
    wszRight[nRight]=L'\0';

  if (moCur.dwPrintColor & PRNC_TEXT)
  {
    crTextOld=GetTextColor(hDC);
    SetTextColor(hDC, lpFrameCurrent->aec.crBasicText);
  }
  if (moCur.dwPrintColor & PRNC_BACKGROUND)
  {
    if (hbrBasicBk=CreateSolidBrush(lpFrameCurrent->aec.crBasicBk))
    {
      FillRect(hDC, rc, hbrBasicBk);
      DeleteObject(hbrBasicBk);
    }
    nBkModeOld=SetBkMode(hDC, TRANSPARENT);
  }

  dwAlign=GetTextAlign(hDC);
  if (*wszCenter && bResult)
  {
    SetTextAlign(hDC, TA_CENTER|TA_TOP|TA_NOUPDATECP);
    bResult=ExtTextOutW(hDC, rc->left + (rc->right - rc->left) / 2, rc->top, ETO_CLIPPED, rc, wszCenter, nCenter, NULL);
  }
  if (*wszLeft && bResult)
  {
    SetTextAlign(hDC, TA_LEFT|TA_TOP|TA_NOUPDATECP);
    bResult=ExtTextOutW(hDC, rc->left, rc->top, ETO_CLIPPED, rc, wszLeft, nLeft, NULL);
  }
  if (*wszRight && bResult)
  {
    SetTextAlign(hDC, TA_RIGHT|TA_TOP|TA_NOUPDATECP);
    bResult=ExtTextOutW(hDC, rc->right, rc->top, ETO_CLIPPED, rc, wszRight, nRight, NULL);
  }
  SetTextAlign(hDC, dwAlign);
  if (moCur.dwPrintColor & PRNC_TEXT)
    SetTextColor(hDC, crTextOld);
  if (moCur.dwPrintColor & PRNC_BACKGROUND)
    SetBkMode(hDC, nBkModeOld);

  return bResult;
}

BOOL CALLBACK PreviewDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndPreview;
  static HWND hWndPrevPage;
  static HWND hWndPageCount;
  static HWND hWndNextPage;
  static HWND hWndZoom;
  static HWND hWndSelection;
  int i;

  if (uMsg == WM_INITDIALOG)
  {
    DWORD dwStyle;

    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndPrevPage=GetDlgItem(hDlg, IDC_PREVIEW_PREVPAGE);
    hWndPageCount=GetDlgItem(hDlg, IDC_PREVIEW_PAGECOUNT);
    hWndNextPage=GetDlgItem(hDlg, IDC_PREVIEW_NEXTPAGE);
    hWndZoom=GetDlgItem(hDlg, IDC_PREVIEW_ZOOM);
    hWndZoomEdit=GetDlgItem(hWndZoom, IDC_COMBOBOXEDIT);
    hWndSelection=GetDlgItem(hDlg, IDC_PREVIEW_SELECTION);

    dwStyle=(DWORD)GetWindowLongPtrWide(hWndZoomEdit, GWL_STYLE);
    SetWindowLongPtrWide(hWndZoomEdit, GWL_STYLE, dwStyle|ES_NUMBER);

    //Get scale factor for ScaleX and ScaleY
    ScaleInit(NULL, hDlg);
    rcPreviewWindow.left=ScaleX(10);
    rcPreviewWindow.top=ScaleY(50);

    //Positioning dialog
    if (rcPreviewDialog.right && rcPreviewDialog.bottom)
      MoveWindow(hDlg, rcPreviewDialog.left, rcPreviewDialog.top, rcPreviewDialog.right, rcPreviewDialog.bottom, FALSE);

    //Initialize
    hWndPreviewDlg=hDlg;
    hWndPreview=NULL;
    ptPreviewScroll.x=0;
    ptPreviewScroll.y=0;
    nPreviewZoomPercent=nPreviewZoomValue;
    nPreviewPageCur=1;
    bPreviewSelection=FALSE;
    bPreviewRedrawLock=FALSE;

    if (PreviewInit(hWndSelection))
    {
      //Create preview window
      {
        WNDCLASSW wndclass;

        wndclass.style        =CS_HREDRAW|CS_VREDRAW|CS_DBLCLKS;
        wndclass.lpfnWndProc  =PreviewProc;
        wndclass.cbClsExtra   =0;
        wndclass.cbWndExtra   =DLGWINDOWEXTRA;
        wndclass.hInstance    =hInstance;
        wndclass.hIcon        =NULL;
        wndclass.hCursor      =hCursorHandOpen;
        wndclass.hbrBackground=(HBRUSH)GetStockObject(HOLLOW_BRUSH);
        wndclass.lpszMenuName =NULL;
        wndclass.lpszClassName=APP_PRINTPREVIEW_CLASSW;
        RegisterClassWide(&wndclass);

        hWndPreview=CreateWindowExWide(0,
                                       APP_PRINTPREVIEW_CLASSW,
                                       NULL,
                                       WS_CHILD|WS_VISIBLE|WS_HSCROLL|WS_VSCROLL,
                                       rcPreviewWindow.left, rcPreviewWindow.top, rcPreviewDialog.right, rcPreviewDialog.bottom,
                                       hDlg,
                                       (HMENU)(UINT_PTR)IDC_PREVIEW_BOX,
                                       hInstance,
                                       NULL);
      }

      //Set caret page
      if (!(nPreviewPageCur=StackPageFind(&hPreviewAllPagesStack, &ciCurCaret)))
        nPreviewPageCur=1;
      PostMessage(hDlg, AKDLG_PREVIEWSETPAGE, nPreviewPageCur, 0);

      //Fill zooms
      for (i=0; i <= nPreviewZoomMaxIndex; ++i)
      {
        xprintfW(wbuf, L"%d%%", lpZoom[i]);
        ComboBox_AddStringWide(hWndZoom, wbuf);
      }
      API_LoadStringW(hLangLib, STR_PAGEFIT, wbuf, BUFFER_SIZE);
      ComboBox_AddStringWide(hWndZoom, wbuf);
      API_LoadStringW(hLangLib, STR_PAGEWIDTH, wbuf, BUFFER_SIZE);
      ComboBox_AddStringWide(hWndZoom, wbuf);

      hHookKeys=SetWindowsHookEx(WH_GETMESSAGE, GetMsgProc, NULL, GetCurrentThreadId());
    }
    ShowWindow(hDlg, dwPreviewShowDialog);
  }
  else if (uMsg == WM_SIZE)
  {
    if (lParam)
    {
      if (hWndPreview)
      {
        rcPreviewWindow.right=LOWORD(lParam);
        rcPreviewWindow.bottom=HIWORD(lParam);
        MoveWindow(hWndPreview, rcPreviewWindow.left, rcPreviewWindow.top, RectW(&rcPreviewWindow), RectH(&rcPreviewWindow), TRUE);
      }
    }
    return 0;
  }
  else if (uMsg == AKDLG_PREVIEWKEYDOWN)
  {
    if (wParam == VK_HOME && (GetKeyState(VK_CONTROL) & 0x80))
    {
      PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_FIRSTPAGE, 0);
    }
    else if (wParam == VK_END && (GetKeyState(VK_CONTROL) & 0x80))
    {
      PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_LASTPAGE, 0);
    }
    else if (wParam == VK_PRIOR)
    {
      PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_PREVPAGE, 0);
    }
    else if (wParam == VK_NEXT)
    {
      PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_NEXTPAGE, 0);
    }
    else if (wParam == VK_ADD && (GetKeyState(VK_CONTROL) & 0x80))
    {
      PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_ZOOMIN, 0);
    }
    else if (wParam == VK_SUBTRACT && (GetKeyState(VK_CONTROL) & 0x80))
    {
      PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_ZOOMOUT, 0);
    }
    else if (wParam == VK_RETURN)
    {
      if ((HWND)lParam == hWndZoomEdit)
      {
        GetWindowTextWide(hWndZoomEdit, wbuf, BUFFER_SIZE);
        PostMessage(hDlg, AKDLG_PREVIEWSETZOOM, xatoiW(wbuf, NULL), FALSE);
      }
      else if ((HWND)lParam == hWndPageCount)
      {
        GetWindowTextWide(hWndPageCount, wbuf, BUFFER_SIZE);
        PostMessage(hDlg, AKDLG_PREVIEWSETPAGE, xatoiW(wbuf, NULL), 0);
      }
      else return FALSE;
    }
    else return FALSE;

    SetWindowLongPtrWide(hDlg, DWLP_MSGRESULT, 1);
    return TRUE;
  }
  else if (uMsg == AKDLG_PREVIEWMOUSEWHEEL)
  {
    if ((short)HIWORD(wParam) < 0)
    {
      if (GetKeyState(VK_CONTROL) & 0x80)
        PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_ZOOMOUT, 0);
      else
        PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_NEXTPAGE, 0);
    }
    else
    {
      if (GetKeyState(VK_CONTROL) & 0x80)
        PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_ZOOMIN, 0);
      else
        PostMessage(hDlg, WM_COMMAND, IDC_PREVIEW_PREVPAGE, 0);
    }

    SetWindowLongPtrWide(hDlg, DWLP_MSGRESULT, 1);
    return TRUE;
  }
  else if (uMsg == AKDLG_PREVIEWSETZOOM)
  {
    nPreviewZoomValue=(int)wParam;

    if (nPreviewZoomValue == PREVIEWZOOM_FIT ||
        nPreviewZoomValue == PREVIEWZOOM_WIDTH)
    {
    }
    else if (nPreviewZoomValue < lpZoom[0])
    {
      nPreviewZoomValue=lpZoom[0];
    }
    else if (nPreviewZoomValue > lpZoom[nPreviewZoomMaxIndex])
    {
      nPreviewZoomValue=lpZoom[nPreviewZoomMaxIndex];
    }

    if (nPreviewZoomValue != nPreviewZoomPercent)
    {
      nPreviewZoomPercent=nPreviewZoomValue;
      InvalidateRect(hWndPreview, NULL, TRUE);
      if (lParam) UpdateWindow(hWndPreview);
    }
    xprintfW(wbuf, L"%d%%", nPreviewZoomPercent);
    SetWindowTextWide(hWndZoomEdit, wbuf);
  }
  else if (uMsg == AKDLG_PREVIEWSETPAGE)
  {
    int nPageSum=bPreviewSelection?nPreviewSelPageSum:nPreviewAllPageSum;
    int nPageCur=(int)wParam;

    nPageCur=max(nPageCur, 1);
    nPageCur=min(nPageCur, nPageSum);

    if (nPageCur != nPreviewPageCur)
    {
      nPreviewPageCur=nPageCur;
      InvalidateRect(hWndPreview, NULL, TRUE);
    }
    if (nPreviewPageCur <= 1)
    {
      if (GetFocus() == hWndPrevPage)
        SetFocus(hWndPageCount);
      EnableWindow(hWndPrevPage, FALSE);
    }
    else EnableWindow(hWndPrevPage, TRUE);

    if (nPreviewPageCur >= nPageSum)
    {
      if (GetFocus() == hWndNextPage)
        SetFocus(hWndPageCount);
      EnableWindow(hWndNextPage, FALSE);
    }
    else EnableWindow(hWndNextPage, TRUE);

    xprintfW(wbuf, L"%d / %d", nPreviewPageCur, nPageSum);
    SetWindowTextWide(hWndPageCount, wbuf);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_PREVIEW_PRINT)
    {
      if (bPreviewSelection)
      {
        prninfo.dwPrintFlags|=PD_SELECTION;
      }
      SendMessage(hDlg, WM_COMMAND, IDCANCEL, 0);
      PostMessage(hMainWnd, WM_COMMAND, IDM_FILE_PRINT, (LPARAM)hWndPreviewEdit);
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_PAGESETUP)
    {
      //Stop message hooking
      hWndPreviewDlg=NULL;

      if (DoFilePageSetup(hDlg))
      {
        hWndPreviewDlg=hDlg;
        PreviewUninit();
        if (PreviewInit(hWndSelection))
        {
          PostMessage(hDlg, AKDLG_PREVIEWSETPAGE, nPreviewPageCur, 0);
          InvalidateRect(hWndPreview, NULL, TRUE);
        }
      }
      hWndPreviewDlg=hDlg;
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_SELECTION)
    {
      nPreviewPageCur=1;
      bPreviewSelection=(BOOL)SendMessage(hWndSelection, BM_GETCHECK, 0, 0);
      PostMessage(hDlg, AKDLG_PREVIEWSETPAGE, nPreviewPageCur, 0);
      InvalidateRect(hWndPreview, NULL, TRUE);
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_FIRSTPAGE)
    {
      PostMessage(hDlg, AKDLG_PREVIEWSETPAGE, 1, 0);
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_LASTPAGE)
    {
      PostMessage(hDlg, AKDLG_PREVIEWSETPAGE, bPreviewSelection?nPreviewSelPageSum:nPreviewAllPageSum, 0);
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_PREVPAGE)
    {
      PostMessage(hDlg, AKDLG_PREVIEWSETPAGE, nPreviewPageCur - 1, 0);
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_NEXTPAGE)
    {
      PostMessage(hDlg, AKDLG_PREVIEWSETPAGE, nPreviewPageCur + 1, 0);
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_ZOOM)
    {
      int nCurSel;

      if (HIWORD(wParam) == CBN_SELCHANGE)
      {
        if ((nCurSel=(int)SendMessage(hWndZoom, CB_GETCURSEL, 0, 0)) != CB_ERR)
        {
          PostMessage(hDlg, AKDLG_PREVIEWSETZOOM, lpZoom[nCurSel], FALSE);
        }
      }
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_ZOOMIN)
    {
      for (i=nPreviewZoomMaxIndex; i >= 0; --i)
      {
        if (lpZoom[i] <= nPreviewZoomPercent)
        {
          if (i < nPreviewZoomMaxIndex)
          {
            PostMessage(hDlg, AKDLG_PREVIEWSETZOOM, lpZoom[i + 1], FALSE);
          }
          break;
        }
      }
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_ZOOMOUT)
    {
      for (i=0; i <= nPreviewZoomMaxIndex; ++i)
      {
        if (lpZoom[i] >= nPreviewZoomPercent)
        {
          if (i > 0)
          {
            PostMessage(hDlg, AKDLG_PREVIEWSETZOOM, lpZoom[i - 1], FALSE);
          }
          break;
        }
      }
    }
    else if (LOWORD(wParam) == IDC_PREVIEW_CLOSE ||
             LOWORD(wParam) == IDCANCEL)
    {
      if (GetWindowLongPtrWide(hDlg, GWL_STYLE) & WS_MAXIMIZE)
      {
        dwPreviewShowDialog=SW_MAXIMIZE;
      }
      else
      {
        GetWindowPos(hDlg, NULL, &rcPreviewDialog);
        dwPreviewShowDialog=SW_SHOW;
      }
      hWndPreviewDlg=NULL;
      PreviewUninit();
      EndDialog(hDlg, 0);
      return TRUE;
    }
  }
  else if (uMsg == WM_DESTROY)
  {
    if (hHookKeys)
    {
      if (UnhookWindowsHookEx(hHookKeys))
        hHookKeys=NULL;
    }
  }
  return FALSE;
}

BOOL PreviewInit(HWND hWndSelection)
{
  BOOL bResult=FALSE;

  xmemset(&rcPreviewZoomed, 0, sizeof(RECT));
  nPreviewAllPageSum=0;
  nPreviewSelPageSum=0;

  //Get printer DC
  if (!prninfo.hDC)
  {
    if (!GetPrinter(hMainWnd, &prninfo, TRUE))
      return 0;
  }

  if (prninfo.hDC)
  {
    //Initialize variables
    prn.hPrinterDC=prninfo.hDC;
    prn.rcMargins=prninfo.rtMargin;

    nPreviewAllPageSum=PrintDocument(hWndPreviewEdit, &prn, PRND_TEST|PRND_ALLTEXT, 0);
    nPreviewSelPageSum=PrintDocument(hWndPreviewEdit, &prn, PRND_TEST|PRND_SELECTION, 0);
    bResult=TRUE;
  }

  if (!nPreviewSelPageSum)
    EnableWindow(hWndSelection, FALSE);
  return bResult;
}

void PreviewUninit()
{
  StackPageFree(&hPreviewAllPagesStack);
  StackPageFree(&hPreviewSelPagesStack);

  if (prninfo.hDC)
  {
    DeleteDC(prninfo.hDC);
    prninfo.hDC=NULL;
  }
}

LRESULT CALLBACK PreviewProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  LRESULT lResult;

  if (uMsg == WM_PAINT)
  {
    PAINTSTRUCT ps;

    if (BeginPaint(hWnd, &ps))
    {
      if (ps.rcPaint.right - ps.rcPaint.left > 0 &&
          ps.rcPaint.bottom - ps.rcPaint.top > 0)
      {
        PRINTPAGE *lpElement;
        HENHMETAFILE hMetaFile;
        HDC hMetaDC;
        RECT rcMeta;

        if ((bPreviewSelection == FALSE && (lpElement=StackPageGet(&hPreviewAllPagesStack, nPreviewPageCur))) ||
            (bPreviewSelection == TRUE && (lpElement=StackPageGet(&hPreviewSelPagesStack, nPreviewPageCur))))
        {
          //Create the EMF in memory
          rcMeta.left=0;
          rcMeta.top=0;
          rcMeta.right=GetDeviceCaps(prninfo.hDC, HORZSIZE) * 100;
          rcMeta.bottom=GetDeviceCaps(prninfo.hDC, VERTSIZE) * 100;
          if (bOldWindows)
            hMetaDC=CreateEnhMetaFileA(prninfo.hDC, NULL, &rcMeta, NULL);
          else
            hMetaDC=CreateEnhMetaFileW(prninfo.hDC, NULL, &rcMeta, NULL);

          //Print page on metafile device
          prn.hPrinterDC=hMetaDC;
          prn.crText.ciMin=lpElement->crText.ciMin;
          prn.crText.ciMax=lpElement->crText.ciMax;
          PrintDocument(hWndPreviewEdit, &prn, (bOldWindows?PRND_ANSI:0)|PRND_RANGE|PRND_ONEPAGE, nPreviewPageCur - 1);
          hMetaFile=CloseEnhMetaFile(hMetaDC);

          //Draw page on window
          PreviewPaint(hWnd, ps.hdc, hMetaFile);
          DeleteEnhMetaFile(hMetaFile);

          //Set zoom text
          xprintfW(wbuf, L"%d%%", nPreviewZoomPercent);
          SetWindowTextWide(hWndZoomEdit, wbuf);
        }
      }
      EndPaint(hWnd, &ps);
    }
  }

  if (lResult=PreviewMessages(hWnd, uMsg, wParam, lParam))
    return lResult;

  return DefWindowProcWide(hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK PreviewMessages(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static POINT ptMouseMove;

  if (uMsg == WM_LBUTTONDOWN ||
      uMsg == WM_LBUTTONDBLCLK)
  {
    if (!(dwMouseCapture & MSC_PREVIEWMOVE))
    {
      SetMouseCapture(hWnd, MSC_PREVIEWMOVE);
      SetCursor(hCursorHandClose);
      GetCursorPos(&ptMouseMove);
    }
    if (uMsg == WM_LBUTTONDBLCLK)
    {
      if (nPreviewZoomPercent == 100)
      {
        PostMessage(hWndPreviewDlg, AKDLG_PREVIEWSETZOOM, (WPARAM)PREVIEWZOOM_FIT, FALSE);
      }
      else
      {
        POINT ptClient;

        GetCursorPos(&ptClient);
        ScreenToClient(hWnd, &ptClient);
        ptClient.x=MulDiv(100, ptPreviewScroll.x + ptClient.x, nPreviewZoomPercent) - ptClient.x;
        ptClient.y=MulDiv(100, ptPreviewScroll.y + ptClient.y, nPreviewZoomPercent) - ptClient.y;

        bPreviewRedrawLock=TRUE;
        SendMessage(hWndPreviewDlg, AKDLG_PREVIEWSETZOOM, 100, TRUE);
        bPreviewRedrawLock=FALSE;
        PreviewHScroll(hWnd, PREVIEWSCROLL_ABSOLUTE, ptClient.x);
        PreviewVScroll(hWnd, PREVIEWSCROLL_ABSOLUTE, ptClient.y);
      }
    }
  }
  else if (uMsg == WM_MOUSEMOVE)
  {
    if (dwMouseCapture & MSC_PREVIEWMOVE)
    {
      POINT ptScreen;

      GetCursorPos(&ptScreen);
      PreviewHScroll(hWnd, PREVIEWSCROLL_OFFSET, ptMouseMove.x - ptScreen.x);
      PreviewVScroll(hWnd, PREVIEWSCROLL_OFFSET, ptMouseMove.y - ptScreen.y);
      ptMouseMove=ptScreen;
    }
  }
  else if (uMsg == WM_LBUTTONUP ||
           uMsg == WM_CAPTURECHANGED ||
           uMsg == WM_DESTROY)
  {
    if (dwMouseCapture & MSC_PREVIEWMOVE)
    {
      ReleaseMouseCapture(MSC_PREVIEWMOVE);
      SetCursor(hCursorHandOpen);
    }
  }
  else if (uMsg == WM_SIZE)
  {
    if (lParam)
    {
      PreviewScrollUpdate(hWnd);
    }
  }
  else if (uMsg == WM_HSCROLL)
  {
    PreviewHScroll(hWnd, LOWORD(wParam), 0);
    return 1;
  }
  else if (uMsg == WM_VSCROLL)
  {
    PreviewVScroll(hWnd, LOWORD(wParam), 0);
    return 1;
  }
  return 0;
}

void PreviewPaint(HWND hWnd, HDC hPaintDC, HENHMETAFILE hMetaFile)
{
  HDC hBufferDC=NULL;
  HBITMAP hBitmap=NULL;
  HBITMAP hBitmapOld=NULL;
  HBRUSH hBrush;
  RECT rcPreviewIn;
  RECT rcPreviewClient;
  RECT rcPaperFrame;
  SIZE sizePreview;

  //Zooming
  if (nPreviewZoomPercent == PREVIEWZOOM_FIT)
  {
    FitInside(RectW(&rcPreviewPaper), RectH(&rcPreviewPaper), RectW(&rcPreviewWindow) - rcPreviewPaper.left * 2, RectH(&rcPreviewWindow) - rcPreviewPaper.top * 2, &sizePreview);
    nPreviewZoomPercent=sizePreview.cx * 100 / RectW(&rcPreviewPaper);
  }
  else if (nPreviewZoomPercent == PREVIEWZOOM_WIDTH)
  {
    nPreviewZoomPercent=(RectW(&rcPreviewWindow) - rcPreviewPaper.left * 2 - GetSystemMetrics(SM_CXVSCROLL)) * 100 / RectW(&rcPreviewPaper);
  }
  rcPreviewZoomed.left=rcPreviewPaper.left;
  rcPreviewZoomed.top=rcPreviewPaper.top;
  rcPreviewZoomed.right=rcPreviewZoomed.left + MulDiv(RectW(&rcPreviewPaper), nPreviewZoomPercent, 100);
  rcPreviewZoomed.bottom=rcPreviewZoomed.top + MulDiv(RectH(&rcPreviewPaper), nPreviewZoomPercent, 100);

  //Update scroll
  PreviewScrollUpdate(hWnd);
  PreviewScrollUpdate(hWnd); //Client rect can be changed.
  rcPreviewZoomed.left-=ptPreviewScroll.x;
  rcPreviewZoomed.top-=ptPreviewScroll.y;
  rcPreviewZoomed.right-=ptPreviewScroll.x;
  rcPreviewZoomed.bottom-=ptPreviewScroll.y;

  if (!bPreviewRedrawLock)
  {
    //Create compatible DC to avoid flicking
    hBufferDC=CreateCompatibleDC(hPaintDC);
    hBitmap=CreateCompatibleBitmap(hPaintDC, RectW(&rcPreviewWindow), RectH(&rcPreviewWindow));
    hBitmapOld=(HBITMAP)SelectObject(hBufferDC, hBitmap);

    //Erase preview window contents
    GetClientRect(hWnd, &rcPreviewClient);
    FillRect(hBufferDC, &rcPreviewClient, (HBRUSH)(UINT_PTR)(COLOR_BTNFACE + 1));

    //Set margins color
    if (hBrush=CreateSolidBrush(RGB(0xE0, 0xE0, 0xE0)))
    {
      FillRect(hBufferDC, &rcPreviewZoomed, hBrush);
      DeleteObject(hBrush);
    }

    //Paper frame
    rcPaperFrame.left=rcPreviewZoomed.left - 1;
    rcPaperFrame.top=rcPreviewZoomed.top - 1;
    rcPaperFrame.right=rcPreviewZoomed.right + 1;
    rcPaperFrame.bottom=rcPreviewZoomed.bottom + 1;
    FrameRect(hBufferDC, &rcPaperFrame, (HBRUSH)GetStockObject(BLACK_BRUSH));
  }

  //Draw area
  if (moCur.bPrintHeaderEnable)
    prn.rcPageIn.top-=prn.nCharHeight;
  if (moCur.bPrintFooterEnable)
    prn.rcPageIn.bottom+=prn.nCharHeight;
  rcPreviewIn.left=rcPreviewZoomed.left + MulDiv(RectW(&rcPreviewZoomed), prn.rcPageIn.left - prn.rcPageFull.left, RectW(&prn.rcPageFull));
  rcPreviewIn.top=rcPreviewZoomed.top + MulDiv(RectH(&rcPreviewZoomed), prn.rcPageIn.top - prn.rcPageFull.top, RectH(&prn.rcPageFull));
  rcPreviewIn.right=rcPreviewIn.left + MulDiv(RectW(&rcPreviewZoomed), RectW(&prn.rcPageIn), RectW(&prn.rcPageFull));
  rcPreviewIn.bottom=rcPreviewIn.top + MulDiv(RectH(&rcPreviewZoomed), RectH(&prn.rcPageIn), RectH(&prn.rcPageFull));

  //Get character size
  nPreviewAveCharWidth=MulDiv(RectW(&rcPreviewZoomed), prn.nAveCharWidth, RectW(&prn.rcPageFull));
  nPreviewCharHeight=MulDiv(RectW(&rcPreviewZoomed), prn.nCharHeight, RectW(&prn.rcPageFull));

  if (!bPreviewRedrawLock)
  {
    //Set draw area color
    if (hBrush=CreateSolidBrush(RGB(0xFF, 0xFF, 0xFF)))
    {
      FillRect(hBufferDC, &rcPreviewIn, hBrush);
      DeleteObject(hBrush);
    }

    //Draw text
    PlayEnhMetaFile(hBufferDC, hMetaFile, &rcPreviewZoomed);

    //Transfer image
    BitBlt(hPaintDC, 0, 0, RectW(&rcPreviewWindow), RectH(&rcPreviewWindow), hBufferDC, 0, 0, SRCCOPY);

    //Clean up
    if (hBitmapOld) SelectObject(hBufferDC, hBitmapOld);
    DeleteObject(hBitmap);
    DeleteDC(hBufferDC);
  }
}

int PreviewHScroll(HWND hWnd, int nAction, int nPos)
{
  SCROLLINFO si;

  si.cbSize=sizeof(SCROLLINFO);
  si.fMask=SIF_ALL;
  GetScrollInfo(hWnd, SB_HORZ, &si);

  if (nAction == PREVIEWSCROLL_OFFSET)
  {
    si.nPos+=nPos;
  }
  else if (nAction == PREVIEWSCROLL_ABSOLUTE)
  {
    si.nPos=nPos;
  }
  else if (nAction == SB_LEFT)
  {
    si.nPos=si.nMin;
  }
  else if (nAction == SB_RIGHT)
  {
    si.nPos=si.nMax;
  }
  else if (nAction == SB_LINELEFT)
  {
    si.nPos-=nPreviewAveCharWidth;
  }
  else if (nAction == SB_LINERIGHT)
  {
    si.nPos+=nPreviewAveCharWidth;
  }
  else if (nAction == SB_PAGELEFT)
  {
    si.nPos-=si.nPage;
  }
  else if (nAction == SB_PAGERIGHT)
  {
    si.nPos+=si.nPage;
  }
  else if (nAction == SB_THUMBTRACK)
  {
    si.nPos=si.nTrackPos;
  }
  else if (nAction == SB_THUMBPOSITION)
  {
  }
  si.fMask=SIF_POS;
  SetScrollInfo(hWnd, SB_HORZ, &si, TRUE);
  si.fMask=SIF_POS;
  GetScrollInfo(hWnd, SB_HORZ, &si);

  ptPreviewScroll.x=si.nPos;
  InvalidateRect(hWnd, NULL, FALSE);
  return 0;
}

int PreviewVScroll(HWND hWnd, int nAction, int nPos)
{
  SCROLLINFO si;

  si.cbSize=sizeof(SCROLLINFO);
  si.fMask=SIF_ALL;
  GetScrollInfo(hWnd, SB_VERT, &si);

  if (nAction == PREVIEWSCROLL_OFFSET)
  {
    si.nPos+=nPos;
  }
  else if (nAction == PREVIEWSCROLL_ABSOLUTE)
  {
    si.nPos=nPos;
  }
  else if (nAction == SB_TOP)
  {
    si.nPos=si.nMin;
  }
  else if (nAction == SB_BOTTOM)
  {
    si.nPos=si.nMax;
  }
  else if (nAction == SB_LINEUP)
  {
    si.nPos-=nPreviewCharHeight;
  }
  else if (nAction == SB_LINEDOWN)
  {
    si.nPos+=nPreviewCharHeight;
  }
  else if (nAction == SB_PAGEUP)
  {
    si.nPos-=si.nPage;
  }
  else if (nAction == SB_PAGEDOWN)
  {
    si.nPos+=si.nPage;
  }
  else if (nAction == SB_THUMBTRACK)
  {
    si.nPos=si.nTrackPos;
  }
  else if (nAction == SB_THUMBPOSITION)
  {
  }
  si.fMask=SIF_POS;
  SetScrollInfo(hWnd, SB_VERT, &si, TRUE);
  si.fMask=SIF_POS;
  GetScrollInfo(hWnd, SB_VERT, &si);

  ptPreviewScroll.y=si.nPos;
  InvalidateRect(hWnd, NULL, FALSE);
  return 0;
}

void PreviewScrollUpdate(HWND hWnd)
{
  SCROLLINFO si;
  RECT rcPreviewClient;
  int nHorzMax;
  int nVertMax;

  if (rcPreviewZoomed.right && rcPreviewZoomed.bottom)
  {
    nHorzMax=RectW(&rcPreviewZoomed) + rcPreviewPaper.left * 2 - 1;
    nVertMax=RectH(&rcPreviewZoomed) + rcPreviewPaper.top * 2 - 1;
    GetClientRect(hWnd, &rcPreviewClient);

    si.cbSize=sizeof(SCROLLINFO);
    si.fMask=SIF_ALL;
    GetScrollInfo(hWnd, SB_HORZ, &si);

    if (si.nMax != nHorzMax ||
        si.nPage != (DWORD)rcPreviewClient.right)
    {
      si.fMask=SIF_RANGE|SIF_PAGE;
      si.nMin=0;
      si.nMax=nHorzMax;
      si.nPage=rcPreviewClient.right;
      SetScrollInfo(hWnd, SB_HORZ, &si, TRUE);

      si.fMask=SIF_POS;
      GetScrollInfo(hWnd, SB_HORZ, &si);

      ptPreviewScroll.x=si.nPos;
    }

    si.cbSize=sizeof(SCROLLINFO);
    si.fMask=SIF_ALL;
    GetScrollInfo(hWnd, SB_VERT, &si);

    if (si.nMax != nVertMax ||
        si.nPage != (DWORD)rcPreviewClient.bottom)
    {
      si.cbSize=sizeof(SCROLLINFO);
      si.fMask=SIF_RANGE|SIF_PAGE;
      si.nMin=0;
      si.nMax=nVertMax;
      si.nPage=rcPreviewClient.bottom;
      SetScrollInfo(hWnd, SB_VERT, &si, TRUE);

      si.fMask=SIF_POS;
      GetScrollInfo(hWnd, SB_VERT, &si);

      ptPreviewScroll.y=si.nPos;
    }
  }
}

BOOL FitInside(int nWidth, int nHeight, int nMaxWidth, int nMaxHeight, SIZE *s)
{
  s->cx=nMaxHeight * nWidth / nHeight;
  s->cy=nMaxWidth * nHeight / nWidth;

  if (s->cx <= nMaxWidth)
  {
    s->cy=s->cx * nHeight / nWidth; //portrait
    return TRUE;
  }
  else
  {
    s->cx=s->cy * nWidth / nHeight; //landscape
    return FALSE;
  }
}

int RectW(const RECT *rc)
{
  return rc->right - rc->left;
}

int RectH(const RECT *rc)
{
  return rc->bottom - rc->top;
}

PRINTPAGE* StackPageInsert(HSTACK *hStack)
{
  PRINTPAGE *lpElement;

  if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, -1, sizeof(PRINTPAGE)))
    return lpElement;
  return NULL;
}

PRINTPAGE* StackPageGet(HSTACK *hStack, int nPage)
{
  PRINTPAGE *lpElement;

  if (!StackGetElement((stack *)hStack->first, (stack *)hStack->last, (stack **)&lpElement, nPage))
    return lpElement;
  return NULL;
}

int StackPageFind(HSTACK *hStack, const AECHARINDEX *ciPos)
{
  PRINTPAGE *lpElement=(PRINTPAGE *)hStack->first;
  int nIndex=0;

  while (lpElement)
  {
    if (AEC_IndexCompare(ciPos, &lpElement->crText.ciMin) < 0)
      return nIndex;

    ++nIndex;
    lpElement=lpElement->next;
  }
  return nIndex;
}

void StackPageFree(HSTACK *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}


//// Code pages

UINT_PTR CALLBACK FileDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hDlgParent;
  static HWND hDlgList;
  static HWND hDlgComboboxLabel;
  static HWND hDlgCombobox;
  static HWND hDlgPlacesBar;
  static HWND hDlgCancel;
  static HWND hWndCodePage;
  static HWND hWndCodePageLabel;
  static HWND hWndShowPlacesBar;
  static HWND hWndFilePreview;
  static HWND hWndAutodetect;
  static RECT rcDlgParent;
  static RECT rcDlg;
  static RECT rcCodePageLabel;
  static RECT rcCodePage;
  static RECT rcAutodetect;
  static RECT rcShowPlacesBar;
  static RECT rcFilePreview;
  static wchar_t wszFile[MAX_PATH];
  static int nCodePage;
  static BOOL bBOM;

  if (uMsg == WM_INITDIALOG)
  {
    DIALOGCODEPAGE *dc=(DIALOGCODEPAGE *)ofnStruct->lCustData;

    hDlgParent=GetParent(hDlg);
    hDlgList=GetDlgItem(hDlgParent, IDC_OFN_LIST);
    hDlgComboboxLabel=GetDlgItem(hDlgParent, IDC_OFN_CODEPAGECOMBOBOX_LABEL);
    hDlgCombobox=GetDlgItem(hDlgParent, IDC_OFN_CODEPAGECOMBOBOX);
    if (moCur.bShowPlacesBar && !bOldWindows && !bWindowsNT4)
    {
      hOfnDlgCombo=GetDlgItem(hDlgParent, IDC_OFN_FILECOMBOBOX);
      hOfnDlgEdit=(HWND)SendMessage(hOfnDlgCombo, CBEM_GETEDITCONTROL, 0, 0);
      hDlgPlacesBar=GetDlgItem(hDlgParent, IDC_OFN_PLACESBAR);
    }
    else
    {
      hOfnDlgCombo=NULL;
      hOfnDlgEdit=GetDlgItem(hDlgParent, IDC_OFN_FILEEDIT);
      hDlgPlacesBar=NULL;
    }
    hDlgCancel=GetDlgItem(hDlgParent, IDCANCEL);
    hWndCodePageLabel=GetDlgItem(hDlg, IDC_OFN_CODEPAGE_LABEL);
    hWndCodePage=GetDlgItem(hDlg, IDC_OFN_CODEPAGE);
    hWndShowPlacesBar=GetDlgItem(hDlg, IDC_OFN_SHOWPLACESBAR);
    hWndFilePreview=GetDlgItem(hDlg, IDC_OFN_PREVIEW);
    hWndAutodetect=GetDlgItem(hDlg, IDC_OFN_AUTODETECT);

    //Set dialog codepage
    nCodePage=lpFrameCurrent->ei.nCodePage;
    bBOM=lpFrameCurrent->ei.bBOM;
    if (dc)
    {
      if (dc->nCodePage >= 0) nCodePage=dc->nCodePage;
      if (dc->bBOM >= 0) bBOM=dc->bBOM;
    }
    FillComboboxCodepage(hWndCodePage, lpCodepageList);
    SelectComboboxCodepage(hWndCodePage, nCodePage);

    GetClientRect(hDlgParent, &rcDlgParent);
    GetWindowPos(hDlg, hDlgParent, &rcDlg);
    GetWindowPos(hWndCodePageLabel, hDlg, &rcCodePageLabel);
    GetWindowPos(hWndCodePage, hDlg, &rcCodePage);
    GetWindowPos(hWndAutodetect, hDlg, &rcAutodetect);
    GetWindowPos(hWndShowPlacesBar, hDlg, &rcShowPlacesBar);
    GetWindowPos(hWndFilePreview, hDlg, &rcFilePreview);

    //Move controls
    {
      RECT rcControl;
      int nLeftMargin;

      if (bOldWindows || bWindowsNT4)
      {
        //Hide arrow
        ShowWindow(hWndShowPlacesBar, FALSE);
      }
      GetWindowPos(hDlgList, hDlgParent, &rcControl);
      nLeftMargin=max(rcControl.left - rcFilePreview.left + (moCur.bShowPlacesBar?2:0), 0);

      if (!bSaveDlg)
      {
        //Move dialog
        rcDlg.right=rcDlgParent.right;
        SetWindowPos(hDlg, 0, 0, 0, rcDlg.right, rcDlg.bottom, SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE);

        //Move places arrow
        if (!bOldWindows && !bWindowsNT4)
        {
          rcShowPlacesBar.left+=nLeftMargin;
          SetWindowPos(hWndShowPlacesBar, 0, rcShowPlacesBar.left, rcShowPlacesBar.top, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
        }

        //Move preview
        rcFilePreview.left+=nLeftMargin;
        rcFilePreview.right=(rcControl.left + rcControl.right) - rcFilePreview.left;
        SetWindowPos(hWndFilePreview, 0, rcFilePreview.left, rcFilePreview.top, rcFilePreview.right, rcFilePreview.bottom, SWP_NOZORDER|SWP_NOACTIVATE);
      }
      else
      {
        //Hide preview
        ShowWindow(hWndFilePreview, FALSE);

        //Move places arrow
        if (!bOldWindows && !bWindowsNT4)
        {
          rcShowPlacesBar.left+=nLeftMargin;
          rcShowPlacesBar.top-=rcDlg.bottom - rcFilePreview.top;
          SetWindowPos(hWndShowPlacesBar, 0, rcShowPlacesBar.left, rcShowPlacesBar.top, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
        }

        //Move dialog
        rcDlg.right=rcDlgParent.right;
        rcDlg.bottom=rcFilePreview.top;
        SetWindowPos(hDlg, 0, 0, 0, rcDlg.right, rcDlg.bottom, SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE);
      }

      //Move codepage label
      GetWindowPos(hDlgComboboxLabel, hDlgParent, &rcControl);
      rcCodePageLabel.left=rcControl.left;
      rcCodePageLabel.right=rcControl.right;
      SetWindowPos(hWndCodePageLabel, 0, rcCodePageLabel.left, rcCodePageLabel.top, rcCodePageLabel.right, rcCodePageLabel.bottom, SWP_NOZORDER|SWP_NOACTIVATE);

      //Move codepage combobox
      GetWindowPos(hDlgCombobox, hDlgParent, &rcControl);
      rcCodePage.left=rcControl.left;
      rcCodePage.right=rcControl.right;
      SetWindowPos(hWndCodePage, 0, rcCodePage.left, rcCodePage.top, rcCodePage.right, rcCodePage.bottom, SWP_NOZORDER|SWP_NOACTIVATE);

      //Move autodetect checkbox
      GetWindowPos(hDlgCancel, hDlgParent, &rcControl);
      rcAutodetect.left=rcControl.left;
      SetWindowPos(hWndAutodetect, 0, rcAutodetect.left, rcAutodetect.top, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);

      //Set button image
      if (!bOldWindows && !bWindowsNT4)
      {
        BUTTONDRAW bd;

        bd.dwFlags=moCur.bShowPlacesBar?BIF_RIGHTARROW:BIF_LEFTARROW;
        SetButtonDraw(hWndShowPlacesBar, &bd);
      }

      //Move places toolbar
      if (hDlgPlacesBar)
      {
        GetWindowPos(hDlgPlacesBar, hDlgParent, &rcControl);
        rcControl.bottom=(rcDlgParent.bottom - rcControl.top) + (rcShowPlacesBar.top + rcShowPlacesBar.bottom);
        SetWindowPos(hDlgPlacesBar, 0, 0, 0, rcControl.right, rcControl.bottom, SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE);
      }
    }

    if (bSaveDlg)
    {
      SetWindowTextWide(hWndAutodetect, L"&BOM");

      if (!IsCodePageUnicode(nCodePage))
      {
        SendMessage(hWndAutodetect, BM_SETCHECK, (WPARAM)FALSE, 0);
        EnableWindow(hWndAutodetect, FALSE);
      }
      else SendMessage(hWndAutodetect, BM_SETCHECK, (WPARAM)bBOM, 0);
    }
    else
    {
      API_LoadStringW(hLangLib, STR_AUTODETECT, wbuf, BUFFER_SIZE);
      SetWindowTextWide(hWndAutodetect, wbuf);
      SendMessage(hWndAutodetect, BM_SETCHECK, (WPARAM)bAutodetect, 0);
      EnableWindow(hWndCodePage, !bAutodetect);

      SendMessage(hWndFilePreview, AEM_SETCOLORS, 0, (LPARAM)&lpFrameCurrent->aec);
      SetTabStops(hWndFilePreview, lpFrameCurrent->nTabStopSize, FALSE);
      SetChosenFont(hWndFilePreview, &lpFrameCurrent->lf);

      lpOldFilePreviewProc=(WNDPROC)GetWindowLongPtrWide(hWndFilePreview, GWLP_WNDPROC);
      SetWindowLongPtrWide(hWndFilePreview, GWLP_WNDPROC, (UINT_PTR)NewFilePreviewProc);
    }

    lpOldFileParentProc=(WNDPROC)GetWindowLongPtrWide(hDlgParent, GWLP_WNDPROC);
    SetWindowLongPtrWide(hDlgParent, GWLP_WNDPROC, (UINT_PTR)NewFileParentProc);
  }
  else if (uMsg == WM_SIZE)
  {
    RECT rcTemplate;
    int x;
    int y;

    GetClientRect(hDlgParent, &rcTemplate);
    x=rcTemplate.right - rcDlgParent.right;
    y=rcTemplate.bottom - rcDlgParent.bottom;
    rcDlgParent=rcTemplate;

    //if (!bOldWindows)
    {
      //Resize codepage combobox
      rcCodePage.right+=x;
      SetWindowPos(hWndCodePage, 0, 0, 0, rcCodePage.right, rcCodePage.bottom, SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE);

      //Move autodetect checkbox
      GetWindowPos(hWndAutodetect, hDlg, &rcAutodetect);
      rcAutodetect.left+=x;
      SetWindowPos(hWndAutodetect, 0, rcAutodetect.left, rcAutodetect.top, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
    }

    //Resize preview
    if (!bSaveDlg)
    {
      rcFilePreview.right+=x;
      SetWindowPos(hWndFilePreview, 0, 0, 0, rcFilePreview.right, rcFilePreview.bottom, SWP_NOMOVE|SWP_NOZORDER|SWP_NOACTIVATE);
    }
    return 0;
  }
  else if (uMsg == WM_CONTEXTMENU)
  {
    if ((HWND)wParam == hWndFilePreview)
    {
      ShowStandardViewMenu((HWND)wParam, hPopupView, lParam != -1);
    }
    return TRUE;
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (((LPNMHDR)lParam)->code == CDN_FILEOK)
    {
      if (!bSaveDlg)
      {
        dwOfnFlags=bAutodetect?(OD_ADT_BINARY_ERROR|OD_ADT_REG_CODEPAGE):0;
        nOfnCodePage=nCodePage;
        bOfnBOM=bBOM;
      }
      else
      {
        dwOfnFlags=0;
        nOfnCodePage=nCodePage;
        bOfnBOM=(BOOL)SendMessage(hWndAutodetect, BM_GETCHECK, 0, 0);
      }
      if (!IsCodePageValid(nOfnCodePage))
      {
        API_LoadStringW(hLangLib, MSG_CP_UNIMPLEMENTED, wbuf, BUFFER_SIZE);
        xprintfW(wszMsg, wbuf, nOfnCodePage);
        API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);

        SetWindowLongPtrWide(hDlg, DWLP_MSGRESULT, 1);
        return TRUE;
      }
    }
    else if (((LPNMHDR)lParam)->code == CDN_SELCHANGE)
    {
      if (!bSaveDlg)
      {
        if (nMDI)
        {
          if (GetWindowTextLengthWide(hOfnDlgEdit) > OPENFILELIST_SIZE)
          {
            API_LoadStringW(hLangLib, MSG_LONG_FILELIST, wszMsg, BUFFER_SIZE);
            API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
          }
        }

        if (bOldWindows)
        {
          SendMessageA(((OFNOTIFY *)lParam)->hdr.hwndFrom, CDM_GETFILEPATH, BUFFER_SIZE, (WPARAM)buf);
          MultiByteToWideChar(CP_ACP, 0, buf, -1, wszFile, MAX_PATH);
        }
        else SendMessageW(((OFNOTIFY *)lParam)->hdr.hwndFrom, CDM_GETFILEPATH, MAX_PATH, (WPARAM)wszFile);

        if (FilePreview(hWndFilePreview, wszFile, PREVIEW_SIZE, bAutodetect?(OD_ADT_BINARY_ERROR|OD_ADT_REG_CODEPAGE):OD_ADT_DETECT_BOM, &nCodePage, &bBOM) < 0)
        {
          EnableWindow(hWndFilePreview, FALSE);
          if (bAutodetect) SelectComboboxCodepage(hWndCodePage, moCur.nDefaultCodePage);
          SetWindowTextWide(hWndFilePreview, L"");
          return TRUE;
        }
        EnableWindow(hWndFilePreview, TRUE);
        if (bAutodetect) SelectComboboxCodepage(hWndCodePage, nCodePage);
      }
    }
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_OFN_CODEPAGE && HIWORD(wParam) == CBN_SELCHANGE)
    {
      nCodePage=GetComboboxCodepage(hWndCodePage);

      if (bSaveDlg)
      {
        if (!IsCodePageUnicode(nCodePage))
        {
          EnableWindow(hWndAutodetect, FALSE);
          SendMessage(hWndAutodetect, BM_SETCHECK, (WPARAM)FALSE, 0);
        }
        else
        {
          EnableWindow(hWndAutodetect, TRUE);
          SendMessage(hWndAutodetect, BM_SETCHECK, (WPARAM)TRUE, 0);
        }
      }
      else
      {
        if (FilePreview(hWndFilePreview, wszFile, PREVIEW_SIZE, OD_ADT_DETECT_BOM, &nCodePage, &bBOM) < 0)
        {
          EnableWindow(hWndFilePreview, FALSE);
          SetWindowTextWide(hWndFilePreview, L"");
        }
        else EnableWindow(hWndFilePreview, TRUE);
      }
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OFN_AUTODETECT)
    {
      if (!bSaveDlg)
      {
        bAutodetect=(BOOL)SendMessage(hWndAutodetect, BM_GETCHECK, 0, 0);
        EnableWindow(hWndCodePage, !bAutodetect);
        nCodePage=GetComboboxCodepage(hWndCodePage);

        if (FilePreview(hWndFilePreview, wszFile, PREVIEW_SIZE, bAutodetect?(OD_ADT_BINARY_ERROR|OD_ADT_REG_CODEPAGE):OD_ADT_DETECT_BOM, &nCodePage, &bBOM) < 0)
        {
          EnableWindow(hWndFilePreview, FALSE);
          if (bAutodetect) SelectComboboxCodepage(hWndCodePage, moCur.nDefaultCodePage);
          SetWindowTextWide(hWndFilePreview, L"");
          return TRUE;
        }
        EnableWindow(hWndFilePreview, TRUE);
        if (bAutodetect) SelectComboboxCodepage(hWndCodePage, nCodePage);
      }
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OFN_SHOWPLACESBAR)
    {
      moCur.bShowPlacesBar=!moCur.bShowPlacesBar;
      PostMessage(hDlgParent, WM_COMMAND, IDCANCEL, 0);
    }
  }
  return FALSE;
}

LRESULT CALLBACK NewFilePreviewProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  if (uMsg == WM_KEYDOWN)
  {
    if (wParam == VK_ESCAPE)
    {
      PostMessage(GetParent(hWnd), WM_KEYDOWN, VK_ESCAPE, 0);
      return TRUE;
    }
  }

  return CallWindowProcWide(lpOldFilePreviewProc, hWnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK NewFileParentProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  if (uMsg == AKDLG_SETSTREAM)
  {
    if (nOfnStreamOffset)
    {
      wszOfnFileList[nOfnStreamOffset]=L':';
      SetWindowTextWide(hOfnDlgEdit, wszOfnFileList);
      SendMessage(hOfnDlgEdit, EM_SETSEL, 0, -1);
      nOfnStreamOffset=0;
    }
  }
  else if (uMsg == WM_COMMAND)
  {
    if (nOfnStreamOffset)
    {
      if (hOfnDlgCombo ? hOfnDlgCombo == (HWND)lParam : hOfnDlgEdit == (HWND)lParam)
      {
        if (hOfnDlgCombo ? HIWORD(wParam) == CBN_EDITCHANGE : HIWORD(wParam) == EN_CHANGE)
        {
          if (lpFrameCurrent->nStreamOffset && nOfnStreamOffset)
            PostMessage(hWnd, AKDLG_SETSTREAM, 0, 0);
        }
      }
    }
    if (LOWORD(wParam) == IDOK)
    {
      wchar_t wszFile[MAX_PATH];
      wchar_t *wpStream=NULL;
      wchar_t *wpCount;
      INT_PTR nFileLen;
      LRESULT lResult;
      BOOL bUpdateName=FALSE;

      if (nFileLen=GetWindowTextWide(hOfnDlgEdit, wszFile, MAX_PATH))
      {
        if (xstrcmpinW(L"file:", wszFile, (UINT_PTR)-1))
        {
          for (wpCount=wszFile; *wpCount; ++wpCount)
          {
            if (*wpCount == L':')
            {
              //Drive letter
              if ((wszFile + 1 == wpCount && (*(wpCount + 1) == L'\\' || *(wpCount + 1) == L'/')) || *(wpCount + 1) == L'\0')
                continue;

              //Stream
              wpStream=wpCount;
              continue;
            }
            else if (*wpCount == L'/')
            {
              *wpCount=L'\\';
              bUpdateName=TRUE;
            }

            if ((*wpCount == L'\\' && wpStream) || *wpCount == L'*' || *wpCount == L'?' || *wpCount == L'\"' || *wpCount == L'<' || *wpCount == L'>' || *wpCount == L'|')
            {
              wpStream=NULL;
              break;
            }
          }
          if (wpStream)
          {
            //In Win7 not work flag changing after dialog create: "ofnStruct->Flags|=OFN_NOVALIDATE;",
            //so we remember stream here (to append it after dialog close) and send file name
            //to dialog without stream, but with dot at the end to avoid lpstrDefExt appending.
            xstrcpynW(wszOfnStream, wpStream + 1, MAX_PATH);
            *wpStream=L'.';
            *(wpStream + 1)=L'\0';
            bUpdateName=TRUE;
          }
          if (bUpdateName)
            SetWindowTextWide(hOfnDlgEdit, wszFile);
          if (wpStream)
          {
            lResult=CallWindowProcWide(lpOldFileParentProc, hWnd, uMsg, wParam, lParam);
            *wpStream=L':';
            *(wpStream + 1)=wszOfnStream[0];
            SetWindowTextWide(hOfnDlgEdit, wszFile);
            return lResult;
          }
          if (GetKeyState(VK_CONTROL) & 0x80)
          {
            //Not work in Win7
            ofnStruct->Flags|=OFN_NODEREFERENCELINKS;
            lResult=CallWindowProcWide(lpOldFileParentProc, hWnd, uMsg, wParam, lParam);
            ofnStruct->Flags&=~OFN_NODEREFERENCELINKS;
            return lResult;
          }
        }
      }
    }
  }
  return CallWindowProcWide(lpOldFileParentProc, hWnd, uMsg, wParam, lParam);
}

void FillComboboxCodepage(HWND hWnd, int *lpCodepageList)
{
  int i;

  if (lpCodepageList)
  {
    for (i=0; lpCodepageList[i]; ++i)
    {
      if (nModelessType == MLT_RECODE &&
          (IsCodePageUnicode(lpCodepageList[i]) && lpCodepageList[i] != CP_UNICODE_UTF8))
      {
        continue;
      }
      GetCodePageName(lpCodepageList[i], wbuf, BUFFER_SIZE);
      ComboBox_AddStringWide(hWnd, wbuf);
    }
  }
}

void FillListBoxCodepage(HWND hWnd, int *lpCodepageList)
{
  int i;

  if (lpCodepageList)
  {
    for (i=0; lpCodepageList[i]; ++i)
    {
      GetCodePageName(lpCodepageList[i], wbuf, BUFFER_SIZE);
      ListBox_AddStringWide(hWnd, wbuf);
    }
  }
}

void ClearCombobox(HWND hWnd)
{
  int nCount;

  if ((nCount=(int)SendMessage(hWnd, CB_GETCOUNT, 0, 0)) > 0)
  {
    while (--nCount >= 0) SendMessage(hWnd, CB_DELETESTRING, nCount, 0);
  }
}

void ClearListBox(HWND hWnd)
{
  int nCount;

  if ((nCount=(int)SendMessage(hWnd, LB_GETCOUNT, 0, 0)) > 0)
  {
    while (--nCount >= 0) SendMessage(hWnd, LB_DELETESTRING, nCount, 0);
  }
}

int GetComboboxCodepage(HWND hWnd)
{
  int nCodePage=0;
  int nSelection;

  if ((nSelection=(int)SendMessage(hWnd, CB_GETCURSEL, 0, 0)) != CB_ERR)
  {
    ComboBox_GetLBTextWide(hWnd, nSelection, wbuf);
    nCodePage=(int)xatoiW(wbuf, NULL);
  }
  return nCodePage;
}

int GetListBoxCodepage(HWND hWnd)
{
  int nCodePage=0;
  int nSelection;

  if ((nSelection=(int)SendMessage(hWnd, LB_GETCURSEL, 0, 0)) != LB_ERR)
  {
    ListBox_GetTextWide(hWnd, nSelection, wbuf);
    nCodePage=(int)xatoiW(wbuf, NULL);
  }
  return nCodePage;
}

int SelectComboboxCodepage(HWND hWnd, int nCodepage)
{
  int nSelection;

  xprintfW(wbuf, L"%d ", nCodepage);
  nSelection=ComboBox_FindStringWide(hWnd, -1, wbuf);
  SendMessage(hWnd, CB_SETCURSEL, (WPARAM)nSelection, 0);

  return nSelection;
}

int SelectListBoxCodepage(HWND hWnd, int nCodepage)
{
  int nSelection;

  xprintfW(wbuf, L"%d ", nCodepage);
  if ((nSelection=ListBox_FindStringWide(hWnd, -1, wbuf)) != LB_ERR)
    SendMessage(hWnd, LB_SETCURSEL, (WPARAM)nSelection, 0);

  return nSelection;
}

void GetListBoxCodepageList(HWND hWnd, int **lpCodepageList)
{
  int *lpCodepageListCount;
  int nCount;
  int i;

  if ((nCount=(int)SendMessage(hWnd, LB_GETCOUNT, 0, 0)) > 0)
  {
    if (*lpCodepageList=lpCodepageListCount=(int *)API_HeapAlloc(hHeap, 0, sizeof(int) * (nCount + 1)))
    {
      for (i=0; i < nCount; ++i)
      {
        ListBox_GetTextWide(hWnd, i, wbuf);
        *lpCodepageListCount++=(int)xatoiW(wbuf, NULL);
      }
      *lpCodepageListCount=0;
      bCodepageListChanged=TRUE;
    }
  }
}

int EnumCodepageList(int **lpCodepageList)
{
  int *lpCodepageListCount;
  int i;

  lpCodepageTable=NULL;
  nCodepageTableCount=0;

  if (lpCodepageTable=(int *)API_HeapAlloc(hHeap, HEAP_ZERO_MEMORY, sizeof(int) * 65536))
  {
    if (bOldWindows)
      RegEnumSystemCodePagesA();
    else
      EnumSystemCodePagesW(EnumCodePagesProc, CP_INSTALLED);

    //CP_ACP
    //CP_OEMCP
    //CP_UNICODE_UTF16LE
    //CP_UNICODE_UTF16BE
    //CP_UNICODE_UTF8
    //CP_UNICODE_UTF7
    //CP_UNICODE_UTF32LE
    //CP_UNICODE_UTF32BE
    //0
    nCodepageTableCount+=9;

    if (*lpCodepageList=lpCodepageListCount=(int *)API_HeapAlloc(hHeap, 0, sizeof(int) * nCodepageTableCount))
    {
      *lpCodepageListCount++=nAnsiCodePage;
      if (nAnsiCodePage != nOemCodePage)
        *lpCodepageListCount++=nOemCodePage;
      *lpCodepageListCount++=CP_UNICODE_UTF16LE;
      *lpCodepageListCount++=CP_UNICODE_UTF16BE;
      *lpCodepageListCount++=CP_UNICODE_UTF8;
      *lpCodepageListCount++=CP_UNICODE_UTF7;
      *lpCodepageListCount++=CP_UNICODE_UTF32LE;
      *lpCodepageListCount++=CP_UNICODE_UTF32BE;

      for (i=0; i <= 65535; ++i)
      {
        if (lpCodepageTable[i])
          *lpCodepageListCount++=i;
      }
      *lpCodepageListCount=0;
    }
    API_HeapFree(hHeap, 0, (LPVOID)lpCodepageTable);
    lpCodepageTable=NULL;
  }
  return nCodepageTableCount;
}

void RegEnumSystemCodePagesA()
{
  HKEY hKey;
  DWORD dwSizeValue;
  DWORD dwSizeString;
  DWORD dwType;
  DWORD dwIndex=0;
  int i;

  if (RegOpenKeyExA(HKEY_LOCAL_MACHINE, "SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage", 0, KEY_READ, &hKey) == ERROR_SUCCESS)
  {
    for (;;)
    {
      dwSizeValue=BUFFER_SIZE;
      dwSizeString=BUFFER_SIZE;
      buf2[0]='\0';

      if (RegEnumValueA(hKey, dwIndex++, buf, &dwSizeValue, NULL, &dwType, (LPBYTE)buf2, &dwSizeString) != ERROR_SUCCESS)
        break;

      if (*buf2)
      {
        if ((i=(int)xatoiA(buf, NULL)) > 0 && i < 65536 &&
            i != nAnsiCodePage &&
            i != nOemCodePage &&
            i != CP_UNICODE_UTF16LE &&
            i != CP_UNICODE_UTF16BE &&
            i != CP_UNICODE_UTF8 &&
            i != CP_UNICODE_UTF7 &&
            i != CP_UNICODE_UTF32LE &&
            i != CP_UNICODE_UTF32BE)
        {
          lpCodepageTable[i]=TRUE;
          ++nCodepageTableCount;
        }
      }
    }
    RegCloseKey(hKey);
  }
}

BOOL CALLBACK EnumCodePagesProc(wchar_t *wpCodePage)
{
  int i;

  if ((i=(int)xatoiW(wpCodePage, NULL)) > 0 && i < 65536 &&
      i != nAnsiCodePage &&
      i != nOemCodePage &&
      i != CP_UNICODE_UTF16LE &&
      i != CP_UNICODE_UTF16BE &&
      i != CP_UNICODE_UTF8 &&
      i != CP_UNICODE_UTF7 &&
      i != CP_UNICODE_UTF32LE &&
      i != CP_UNICODE_UTF32BE)
  {
    lpCodepageTable[i]=TRUE;
    ++nCodepageTableCount;
  }
  return TRUE;
}

int CodepageListLen(int *lpCodepageList)
{
  int i;

  if (lpCodepageList)
  {
    for (i=0; lpCodepageList[i]; ++i);
    return ++i;
  }
  return 0;
}

int CodepageListFind(int *lpCodepageList, int nCodePage)
{
  int i;

  if (lpCodepageList)
  {
    for (i=0; lpCodepageList[i]; ++i)
    {
      if (lpCodepageList[i] == nCodePage)
        return i;
    }
  }
  return -1;
}

void CodepageListFree(int **lpCodepageList)
{
  if (lpCodepageList && *lpCodepageList)
  {
    API_HeapFree(hHeap, 0, (LPVOID)*lpCodepageList);
    *lpCodepageList=NULL;
  }
}

void GetCodePageName(int nCodePage, wchar_t *wszCodePage, int nLen)
{
  CPINFOEXW CPInfoExW;

  if (nCodePage)
  {
    if (nCodePage == CP_UNICODE_UTF16LE)
      xstrcpynW(wszCodePage, STR_UNICODE_UTF16LEW, nLen);
    else if (nCodePage == CP_UNICODE_UTF16BE)
      xstrcpynW(wszCodePage, STR_UNICODE_UTF16BEW, nLen);
    else if (nCodePage == CP_UNICODE_UTF8)
      xstrcpynW(wszCodePage, STR_UNICODE_UTF8W, nLen);
    else if (nCodePage == CP_UNICODE_UTF7)
      xstrcpynW(wszCodePage, STR_UNICODE_UTF7W, nLen);
    else if (nCodePage == CP_UNICODE_UTF32LE)
      xstrcpynW(wszCodePage, STR_UNICODE_UTF32LEW, nLen);
    else if (nCodePage == CP_UNICODE_UTF32BE)
      xstrcpynW(wszCodePage, STR_UNICODE_UTF32BEW, nLen);
    else
    {
      if (GetCPInfoExWide(nCodePage, 0, &CPInfoExW) && nCodePage == xatoiW(CPInfoExW.CodePageName, NULL))
        xstrcpynW(wszCodePage, CPInfoExW.CodePageName, nLen);
      else
        xprintfW(wszCodePage, L"%d ", nCodePage);
    }
  }
  else wszCodePage[0]=L'\0';
}

int FilePreview(HWND hWnd, wchar_t *wpFile, UINT_PTR dwPreviewBytes, DWORD dwFlags, int *nCodePage, BOOL *bBOM)
{
  HANDLE hFile;
  FILESTREAMDATA fsd;
  int nDetect;
  int nResult=EOD_SUCCESS;

  if (!(dwFlags & ADT_REG_CODEPAGE) && !(dwFlags & ADT_DETECT_CODEPAGE))
    if (!*nCodePage) return EOD_OPEN;

  if (IsFile(wpFile) != ERROR_SUCCESS)
    return EOD_OPEN;

  //Read from file
  for (;;)
  {
    if ((hFile=API_CreateFileW(wpFile, GENERIC_READ, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL|FILE_FLAG_SEQUENTIAL_SCAN, NULL)) == INVALID_HANDLE_VALUE)
      return EOD_OPEN;

    if ((nDetect=AutodetectCodePage(wpFile, hFile, moCur.dwCodepageRecognitionBuffer, dwFlags, nCodePage, bBOM)) < 0)
    {
      nResult=nDetect;
      break;
    }

    //Offset BOM
    fsd.dwBytesCurrent=0;

    if (*bBOM)
    {
      if (*nCodePage == CP_UNICODE_UTF16LE ||
          *nCodePage == CP_UNICODE_UTF16BE)
      {
        fsd.dwBytesCurrent=2;
      }
      else if (*nCodePage == CP_UNICODE_UTF32LE ||
               *nCodePage == CP_UNICODE_UTF32BE)
      {
        fsd.dwBytesCurrent=4;
      }
      else if (*nCodePage == CP_UNICODE_UTF8)
      {
        fsd.dwBytesCurrent=3;
      }
    }
    SetFilePointer64(hFile, fsd.dwBytesCurrent, FILE_BEGIN);

    fsd.hWnd=hWnd;
    fsd.hFile=hFile;
    fsd.nCodePage=*nCodePage;
    fsd.dwBytesMax=dwPreviewBytes;
    FileStreamIn(&fsd);
    break;
  }
  if (hFile != INVALID_HANDLE_VALUE)
    CloseHandle(hFile);

  return nResult;
}

int AutodetectCodePage(const wchar_t *wpFile, HANDLE hFile, UINT_PTR dwBytesToCheck, DWORD dwFlags, int *nCodePage, BOOL *bBOM)
{
  HANDLE hFileInput=hFile;
  RECENTFILE *lpRecentFile;
  UINT_PTR dwBytesRead=0;
  unsigned char *pBuffer=NULL;
  int nRegCodePage=0;
  UINT_PTR a;
  UINT_PTR b;
  int nANSIrate=5;
  int nUTF16LErate=0;
  int nUTF16BErate=0;
  int nResult=EDT_SUCCESS;

  //Remembered code page from registry
  if (dwFlags & ADT_REG_CODEPAGE)
  {
    dwFlags&=~ADT_REG_CODEPAGE;
    dwFlags&=~ADT_DETECT_CODEPAGE;
    dwFlags&=~ADT_DETECT_BOM;

    if (moCur.nRecentFiles && moCur.bSaveCodepages)
    {
      if (lpRecentFile=RecentFilesFindByName(wpFile, NULL))
        nRegCodePage=lpRecentFile->nCodePage;
    }
    if (nRegCodePage)
    {
      if (nRegCodePage == CP_UNICODE_UTF32LE || nRegCodePage == CP_UNICODE_UTF32BE)
        dwFlags&=~ADT_BINARY_ERROR;
      *nCodePage=nRegCodePage;
      dwFlags|=ADT_DETECT_BOM;
    }
    else
    {
      dwFlags|=ADT_DETECT_CODEPAGE|ADT_DETECT_BOM;
    }
  }

  //Default
  if (dwFlags & ADT_DETECT_CODEPAGE) *nCodePage=moCur.nDefaultCodePage;
  if (dwFlags & ADT_DETECT_BOM) *bBOM=FALSE;

  //Read file
  if (dwFlags & ADT_BINARY_ERROR || dwFlags & ADT_DETECT_CODEPAGE || dwFlags & ADT_DETECT_BOM)
  {
    if (!hFile)
    {
      if (dwFlags & ADT_NOMESSAGES)
        hFile=CreateFileWide(wpFile, GENERIC_READ, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL|FILE_FLAG_SEQUENTIAL_SCAN, NULL);
      else
        hFile=API_CreateFileW(wpFile, GENERIC_READ, FILE_SHARE_READ|FILE_SHARE_WRITE, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL|FILE_FLAG_SEQUENTIAL_SCAN, NULL);
    }

    if (hFile != INVALID_HANDLE_VALUE)
    {
      if (pBuffer=(unsigned char *)API_HeapAlloc(hHeap, 0, dwBytesToCheck + 1))
      {
        if (!ReadFile64(hFile, pBuffer, dwBytesToCheck, &dwBytesRead, NULL))
        {
          SendMessage(hMainWnd, WM_COMMAND, IDM_INTERNAL_ERRORIO_MSG, 0);
          API_HeapFree(hHeap, 0, (LPVOID)pBuffer);
          nResult=EDT_READ;
        }
      }
      else nResult=EDT_ALLOC;
    }
    else nResult=EDT_OPEN;

    if (hFileInput != hFile && hFile != INVALID_HANDLE_VALUE)
      CloseHandle(hFile);
    if (nResult) return nResult;
  }

  //Detect Unicode BOM
  if (dwFlags & ADT_DETECT_CODEPAGE || dwFlags & ADT_DETECT_BOM)
  {
    if (dwBytesRead >= 4)
    {
      if (pBuffer[0] == 0xFF && pBuffer[1] == 0xFE && pBuffer[2] == 0x00 && pBuffer[3] == 0x00)
      {
        if (!nRegCodePage || nRegCodePage == CP_UNICODE_UTF32LE)
        {
          if (dwFlags & ADT_DETECT_CODEPAGE) *nCodePage=CP_UNICODE_UTF32LE;
          if (dwFlags & ADT_DETECT_BOM) *bBOM=TRUE;
          goto Free;
        }
      }
      else if (pBuffer[0] == 0x00 && pBuffer[1] == 0x00 && pBuffer[2] == 0xFE && pBuffer[3] == 0xFF)
      {
        if (!nRegCodePage || nRegCodePage == CP_UNICODE_UTF32BE)
        {
          if (dwFlags & ADT_DETECT_CODEPAGE) *nCodePage=CP_UNICODE_UTF32BE;
          if (dwFlags & ADT_DETECT_BOM) *bBOM=TRUE;
          goto Free;
        }
      }
    }
    if (dwBytesRead >= 2)
    {
      if (pBuffer[0] == 0xFF && pBuffer[1] == 0xFE)
      {
        if (!nRegCodePage || nRegCodePage == CP_UNICODE_UTF16LE)
        {
          if (dwFlags & ADT_DETECT_CODEPAGE) *nCodePage=CP_UNICODE_UTF16LE;
          if (dwFlags & ADT_DETECT_BOM) *bBOM=TRUE;
          goto Free;
        }
      }
      else if (pBuffer[0] == 0xFE && pBuffer[1] == 0xFF)
      {
        if (!nRegCodePage || nRegCodePage == CP_UNICODE_UTF16BE)
        {
          if (dwFlags & ADT_DETECT_CODEPAGE) *nCodePage=CP_UNICODE_UTF16BE;
          if (dwFlags & ADT_DETECT_BOM) *bBOM=TRUE;
          goto Free;
        }
      }
    }
    if (dwBytesRead >= 3)
    {
      if (pBuffer[0] == 0xEF && pBuffer[1] == 0xBB && pBuffer[2] == 0xBF)
      {
        if (!nRegCodePage || nRegCodePage == CP_UNICODE_UTF8)
        {
          if (dwFlags & ADT_DETECT_CODEPAGE) *nCodePage=CP_UNICODE_UTF8;
          if (dwFlags & ADT_DETECT_BOM) *bBOM=TRUE;
          goto Free;
        }
      }
    }
  }

  if (dwFlags & ADT_BINARY_ERROR || dwFlags & ADT_DETECT_CODEPAGE)
  {
    if (dwBytesRead >= 2)
    {
      for (a=0, b=dwBytesRead - 1; a < b; a+=2)
      {
        //Detect binary file
        if (dwFlags & ADT_BINARY_ERROR)
        {
          if (pBuffer[a] == 0x00 && pBuffer[a + 1] == 0x00)
          {
            API_HeapFree(hHeap, 0, (LPVOID)pBuffer);
            return EDT_BINARY;
          }
        }

        //Detect UTF-16LE, UTF-16BE without BOM
        if (dwFlags & ADT_DETECT_CODEPAGE)
        {
          if (pBuffer[a + 1] == 0x00 && pBuffer[a] <= 0x7E)
          {
            ++nUTF16LErate;
            nUTF16BErate=-0xFFFF;
          }
          else if (pBuffer[a] == 0x00 && pBuffer[a + 1] <= 0x7E)
          {
            ++nUTF16BErate;
            nUTF16LErate=-0xFFFF;
          }
        }
      }

      if (dwFlags & ADT_DETECT_CODEPAGE)
      {
        if (nUTF16LErate >= nANSIrate && nUTF16LErate >= nUTF16BErate)
        {
          *nCodePage=CP_UNICODE_UTF16LE;
          dwFlags&=~ADT_DETECT_CODEPAGE;

          if (dwFlags & ADT_DETECT_BOM)
          {
            *bBOM=FALSE;
            dwFlags&=~ADT_DETECT_BOM;
          }
        }
        else if (nUTF16BErate >= nANSIrate && nUTF16BErate >= nUTF16LErate)
        {
          *nCodePage=CP_UNICODE_UTF16BE;
          dwFlags&=~ADT_DETECT_CODEPAGE;

          if (dwFlags & ADT_DETECT_BOM)
          {
            *bBOM=FALSE;
            dwFlags&=~ADT_DETECT_BOM;
          }
        }
      }
    }
  }

  //Detect non-Unicode
  if (dwFlags & ADT_DETECT_CODEPAGE || dwFlags & ADT_DETECT_BOM)
  {
    if (dwFlags & ADT_DETECT_CODEPAGE)
    {
      if (!AutodetectMultibyte(moCur.dwLangCodepageRecognition, pBuffer, dwBytesRead, DETECTCHARS_REQUIRED, nCodePage))
      {
        *nCodePage=moCur.nDefaultCodePage;
        dwFlags&=~ADT_DETECT_CODEPAGE;
      }
    }
    if (dwFlags & ADT_DETECT_BOM)
    {
      *bBOM=FALSE;
      dwFlags&=~ADT_DETECT_BOM;
    }
  }

  //Free buffer
  Free:
  if (pBuffer) API_HeapFree(hHeap, 0, (LPVOID)pBuffer);
  return EDT_SUCCESS;
}

BOOL AutodetectMultibyte(DWORD dwLangID, const unsigned char *pBuffer, UINT_PTR dwBytesToCheck, UINT_PTR dwCharsRequired, int *nCodePage)
{
  static const char lpTrailingBytesForUTF8[256]={
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, 0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
    1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1, 1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,
    2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2, 3,3,3,3,3,3,3,3,4,4,4,4,5,5,5,5
  };
  char szANSIwatermark[128];
  char szOEMwatermark[128];
  char szKOIwatermark[128];
  char szUTF8watermark[128];
  int nANSIrate=5;
  int nOEMrate=0;
  int nKOIrate=0;
  int nUTF8rate=0;
  int nCommonUTF8rate=0;
  DWORD dwCounter[0x80];
  DWORD dwMaxIndex=0;
  DWORD dwMaxCount=0;
  DWORD dwTrailing;
  UINT_PTR i;
  UINT_PTR j;
  BOOL bRated=FALSE;

  if (dwBytesToCheck == (UINT_PTR)-1)
    dwBytesToCheck=xstrlenA((char *)pBuffer);
  if (dwLangID == (DWORD)-1)
    dwLangID=moCur.dwLangCodepageRecognition;
  dwLangID=PRIMARYLANGID(dwLangID);

  //Watermarks
  szANSIwatermark[0]='\0';
  szOEMwatermark[0]='\0';
  szKOIwatermark[0]='\0';
  szUTF8watermark[0]='\0';

  if (dwLangID == LANG_RUSSIAN)
  {
    xstrcpyA(szANSIwatermark, "\xE0\xE1\xE2\xE5\xE8\xED\xEE\xEF\xF0\xF2\xC0\xC1\xC2\xC5\xC8\xCD\xCE\xCF\xD2");  //������������������?
    xstrcpyA(szKOIwatermark,  "\xC1\xC2\xD7\xC5\xC9\xCE\xCF\xD2\xD4\xE1\xE2\xF7\xE5\xE9\xEE\xEF\xF0\xF2\xF4");  //������������������?
    xstrcpyA(szOEMwatermark,  "\xAE\xA5\xA0\xA8\xAA\xAC\xAD\xE2\x8E\x45\x80\x88\x8A\x8C\x8D\x92\xB0\xB1\xB2\xB3\xBA\xDB\xCD");  //���������E������         Graphic simbols: \xB0\xB1\xB2\xB3\xBA\xDB\xCD
    xstrcpyA(szUTF8watermark, "\xD0\xD1");
  }
  else if (IsLangEasternEurope(dwLangID))
  {
    xstrcpyA(szANSIwatermark, "\xE1\xED\xEB\xE8\x9A\xFE\xE6\x9E\xE3\xF5\x9D\xFA\xF6\xF8\xF3\xEA\xBF\xFC\xF1\xF4\xCE\xC8\xF9\xF0\xAA\xCF\xC1\xE2\xDA\x8E\xB9\xBE");
    xstrcpyA(szOEMwatermark,  "\xA0\xA1\x89\x82\xA7\x86\x8B\xA3\x94\xA2\xA5\xD8\xA9\x8C\x87\x88\x96\x81\xD7\x98\x93\xAC\x80\xD0\x85\xB8\xD2\xB5\xAB\xA6\x84\x83\xB0\xB1\xB2\xDB\xCD"); //Graphic simbols: \xB0\xB1\xB2\xDB\xCD
    xstrcpyA(szUTF8watermark, "\xC3\xC5\xC4\xA1");
  }
  else if (IsLangWesternEurope(dwLangID))
  {
    xstrcpyA(szANSIwatermark, "\xE1\xF3\xED\xF0\xF1\xFA\xE5\xF8\xF6\xE6\xE3\xFE\xEB\xEA\xE0\xF9\xEF\xE8\xFC\xFD\xC9\xC1\xF4\xF5\xF2\xDC\xEE\xD6\xBF\xDA\xCC\xCA\xC4");
    xstrcpyA(szOEMwatermark,  "\xA0\x82\xA2\x84\xD0\xA4\xA3\x86\x87\x9B\x94\xC6\x91\x89\x88\x85\x97\x8B\x8A\x81\xEC\xB5\x90\x93\x95\xAD\xA8\x99\x8C\x60\xDE\xD2\xC7\xB0\xB1\xB2\xDB\xCD"); //Graphic simbols: \xB0\xB1\xB2\xDB\xCD
    xstrcpyA(szUTF8watermark, "\xC3");
  }
  else if (dwLangID == LANG_TURKISH)
  {
    xstrcpyA(szANSIwatermark, "\xFC\xFD\xFE");
    xstrcpyA(szOEMwatermark,  "\x81\x87\x8D\xB0\xB1\xB2\xB3\xBA\xDB\xCD");  //Graphic simbols: \xB0\xB1\xB2\xB3\xBA\xDB\xCD
    xstrcpyA(szUTF8watermark, "\xB0\xB1\xBC\xC3\xC4\xC5");
  }
  else if (dwLangID == LANG_CHINESE)
  {
    xstrcpyA(szANSIwatermark, "\xA1\xA2\xA3\xA4\xA5\xA6\xAB\xB2\xB4\xC0\xC1\xC2\xC3\xC4\xC5\xC6\xC7\xC8\xC9\xCA\xCB\xCC\xCD\xCE\xCF\xD0\xD1\xD2\xD3\xD4\xD5\xD6\xD7\xD8\xD9\xDA\xDB\xDC\xDD\xDE\xDF\xED");
    xstrcpyA(szUTF8watermark, "\xE4\xE5\xE6\xE7\xE8\xE9");
  }
  else if (dwLangID == LANG_JAPANESE)
  {
    xstrcpyA(szANSIwatermark, "\xC0\xC1\xC2\xC3\xC4\xC5\xC6\xC7\xC8\xC9\xCA\xCB\xCC\xCD\xCE\xCF");
    xstrcpyA(szUTF8watermark, "\xE3");
  }
  else if (dwLangID == LANG_KOREAN)
  {
    xstrcpyA(szANSIwatermark, "\xC0\xC1\xC2\xC3");
    xstrcpyA(szUTF8watermark, "\xEA\xEB\xEC\xED");
  }
  else return FALSE;

  //Zero counter
  xmemset(dwCounter, 0, 0x80 * sizeof(DWORD));

  //Count number of each character in input buffer
  for (j=0, i=0; i < dwBytesToCheck; ++i)
  {
    //Char in range 0x80 - 0xFF
    if (pBuffer[i] >= 0x80)
    {
      ++j;
      dwCounter[pBuffer[i] - 0x80]++;
    }
  }

  //Give it up if there's no representative selection
  if (j > dwCharsRequired)
  {
    //Rate top repeated characters
    j=dwCharsRequired;

    while (j)
    {
      //Get max element
      for (dwMaxCount=0, i=0; i < 0x80; ++i)
      {
        if (dwCounter[i] > dwMaxCount)
        {
          dwMaxCount=dwCounter[i];
          dwMaxIndex=(DWORD)i;
        }
      }
      if (!dwCounter[dwMaxIndex]) break;

      if (AKD_strchr(szANSIwatermark, dwMaxIndex + 0x80))
      {
        nANSIrate+=dwCounter[dwMaxIndex];
        bRated=TRUE;
      }
      if (AKD_strchr(szOEMwatermark, dwMaxIndex + 0x80))
      {
        nOEMrate+=dwCounter[dwMaxIndex];
        bRated=TRUE;
      }
      if (AKD_strchr(szUTF8watermark, dwMaxIndex + 0x80))
      {
        nUTF8rate+=dwCounter[dwMaxIndex];
        bRated=TRUE;
      }
      if (AKD_strchr(szKOIwatermark, dwMaxIndex + 0x80))
      {
        nKOIrate+=dwCounter[dwMaxIndex];
        bRated=TRUE;
      }
      dwCounter[dwMaxIndex]=0;

      if (bRated)
      {
        --j;
        bRated=FALSE;
      }
    }

    //Get common UTF8 rate
    if (nUTF8rate < nANSIrate || nUTF8rate < nOEMrate || nUTF8rate < nKOIrate)
    {
      for (i=0; i < dwBytesToCheck; ++i)
      {
        dwTrailing=lpTrailingBytesForUTF8[pBuffer[i]];

        if (i + dwTrailing < dwBytesToCheck)
        {
          if (IsCharLegalUTF8(pBuffer + i, dwTrailing + 1))
          {
            if (dwTrailing)
            {
              ++nCommonUTF8rate;
              i+=dwTrailing;
            }
          }
          else
          {
            nCommonUTF8rate=0;
            break;
          }
        }
        else break;
      }
      nUTF8rate=max(nUTF8rate, nCommonUTF8rate);
    }

    //Set code page
    if (dwLangID == LANG_RUSSIAN)
    {
      if (nANSIrate >= nOEMrate && nANSIrate >= nKOIrate && nANSIrate >= nUTF8rate)
      {
        *nCodePage=1251;
      }
      else if (nOEMrate >= nKOIrate && nOEMrate >= nUTF8rate)
      {
        *nCodePage=866;
      }
      else if (nKOIrate >= nUTF8rate)
      {
        *nCodePage=CP_KOI8_R;
      }
      else
      {
        *nCodePage=CP_UNICODE_UTF8;
      }
      return TRUE;
    }
    else if (IsLangEasternEurope(dwLangID))
    {
      if (nANSIrate >= nOEMrate && nANSIrate >= nKOIrate && nANSIrate >= nUTF8rate)
      {
        *nCodePage=1250;
      }
      else if (nOEMrate >= nKOIrate && nOEMrate >= nUTF8rate)
      {
        *nCodePage=852;
      }
      else
      {
        *nCodePage=CP_UNICODE_UTF8;
      }
      return TRUE;
    }
    else if (IsLangWesternEurope(dwLangID))
    {
      if (nANSIrate >= nOEMrate && nANSIrate >= nKOIrate && nANSIrate >= nUTF8rate)
      {
        *nCodePage=1252;
      }
      else if (nOEMrate >= nKOIrate && nOEMrate >= nUTF8rate)
      {
        *nCodePage=850;
      }
      else
      {
        *nCodePage=CP_UNICODE_UTF8;
      }
      return TRUE;
    }
    else if (dwLangID == LANG_TURKISH)
    {
      if (nANSIrate >= nOEMrate && nANSIrate >= nUTF8rate)
      {
        *nCodePage=1254;
      }
      else if (nOEMrate >= nUTF8rate)
      {
        *nCodePage=857;
      }
      else
      {
        *nCodePage=CP_UNICODE_UTF8;
      }
      return TRUE;
    }
    else if (dwLangID == LANG_CHINESE)
    {
      if (nUTF8rate > nANSIrate)
      {
        *nCodePage=CP_UNICODE_UTF8;
        return TRUE;
      }
      else
      {
        if (nAnsiCodePage == 936 ||
            nAnsiCodePage == 950)
        {
          *nCodePage=nAnsiCodePage;
          return TRUE;
        }
      }
    }
    else if (dwLangID == LANG_JAPANESE)
    {
      if (nUTF8rate > nANSIrate)
      {
        *nCodePage=CP_UNICODE_UTF8;
      }
      else
      {
        *nCodePage=932;
      }
      return TRUE;
    }
    else if (dwLangID == LANG_KOREAN)
    {
      if (nUTF8rate > nANSIrate)
      {
        *nCodePage=CP_UNICODE_UTF8;
      }
      else
      {
        *nCodePage=949;
      }
      return TRUE;
    }
  }
  return FALSE;
}

BOOL AutodetectWideChar(DWORD dwLangID, const wchar_t *wpText, INT_PTR nTextLen, INT_PTR nMinChars, int *nCodePageFrom, int *nCodePageTo)
{
  //Detect nCodePageFrom
  int lpDetectCodePage[][5]={{0,    0,               0,               0,               0},  //DETECTINDEX_NONE
                             {1251, 866,             CP_KOI8_R,       CP_UNICODE_UTF8, 0},  //DETECTINDEX_RUSSIAN
                             {1250, 852,             CP_UNICODE_UTF8, 0,               0},  //DETECTINDEX_EASTERNEUROPE
                             {1252, 850,             CP_UNICODE_UTF8, 0,               0},  //DETECTINDEX_WESTERNEUROPE
                             {1254, 857,             CP_UNICODE_UTF8, 0,               0},  //DETECTINDEX_TURKISH
                             {950,  936,             CP_UNICODE_UTF8, 0,               0},  //DETECTINDEX_CHINESE
                             {932,  CP_UNICODE_UTF8, 0,               0,               0},  //DETECTINDEX_JAPANESE
                             {949,  CP_UNICODE_UTF8, 0,               0,               0}}; //DETECTINDEX_KOREAN
  int nIndex;
  char *szText=NULL;
  INT_PTR nAnsiLen;
  int i;
  BOOL bUsedDefaultChar=TRUE;

  if (dwLangID == (DWORD)-1)
    dwLangID=moCur.dwLangCodepageRecognition;

  if (nIndex=GetDetectionIndex(dwLangID))
  {
    for (i=0; lpDetectCodePage[nIndex][i]; ++i)
    {
      nAnsiLen=WideCharToMultiByte(lpDetectCodePage[nIndex][i], WC_NO_BEST_FIT_CHARS, wpText, (int)nTextLen, NULL, 0, NULL, &bUsedDefaultChar);

      if (!bUsedDefaultChar)
      {
        if (szText=(char *)API_HeapAlloc(hHeap, 0, nAnsiLen))
        {
          WideCharToMultiByte64(lpDetectCodePage[nIndex][i], 0, wpText, nTextLen, szText, nAnsiLen, NULL, NULL);
          *nCodePageFrom=lpDetectCodePage[nIndex][i];
        }
        break;
      }
    }
  }

  if (szText)
  {
    //Detect nCodePageTo
    if (!AutodetectMultibyte(dwLangID, (unsigned char *)szText, nAnsiLen, nMinChars, nCodePageTo))
      *nCodePageTo=*nCodePageFrom;

    API_HeapFree(hHeap, 0, (LPVOID)szText);
    return TRUE;
  }
  return FALSE;
}

BOOL ConvertAnsiString(const char *pInput, INT_PTR nInputLen, int nCodePageFrom, int nCodePageTo, char **szOutput, INT_PTR *lpnOutputLen)
{
  wchar_t *wszInput;
  INT_PTR nAnsiLen;
  INT_PTR nUnicodeLen;
  BOOL bResult=FALSE;

  //Convert
  if (nCodePageFrom > 0 && nCodePageTo > 0)
  {
    //From MultiByte to Unicode
    nUnicodeLen=MultiByteToWideChar64(nCodePageFrom, 0, pInput, nInputLen, NULL, 0);

    if (wszInput=(wchar_t *)API_HeapAlloc(hHeap, 0, nUnicodeLen))
    {
      MultiByteToWideChar64(nCodePageFrom, 0, pInput, nInputLen, wszInput, nUnicodeLen);

      //From Unicode to MultiByte
      nAnsiLen=WideCharToMultiByte64(nCodePageTo, 0, wszInput, nUnicodeLen, NULL, 0, NULL, NULL);

      if (*szOutput=API_HeapAlloc(hHeap, 0, nAnsiLen + 1))
      {
        WideCharToMultiByte64(nCodePageTo, 0, wszInput, nUnicodeLen, *szOutput, nAnsiLen, NULL, NULL);
        (*szOutput)[nAnsiLen]=L'\0';
        if (lpnOutputLen) *lpnOutputLen=nAnsiLen;
        bResult=TRUE;
      }
      API_HeapFree(hHeap, 0, (LPVOID)wszInput);
    }
  }
  return bResult;
}

BOOL ConvertWideString(const wchar_t *wpInput, INT_PTR nInputLen, int nCodePageFrom, int nCodePageTo, wchar_t **wszOutput, INT_PTR *lpnOutputLen)
{
  char *szInput;
  INT_PTR nUnicodeLen;
  INT_PTR nAnsiLen;
  BOOL bResult=FALSE;

  //Convert
  if (nCodePageFrom > 0 && nCodePageTo > 0)
  {
    //From Unicode to MultiByte
    nAnsiLen=WideCharToMultiByte64(nCodePageFrom, 0, wpInput, nInputLen, NULL, 0, NULL, NULL);

    if (szInput=(char *)API_HeapAlloc(hHeap, 0, nAnsiLen))
    {
      WideCharToMultiByte64(nCodePageFrom, 0, wpInput, nInputLen, szInput, nAnsiLen, NULL, NULL);

      //From MultiByte to Unicode
      nUnicodeLen=MultiByteToWideChar64(nCodePageTo, 0, szInput, nAnsiLen, NULL, 0);

      if (*wszOutput=AllocWideStr(nUnicodeLen + 1))
      {
        MultiByteToWideChar64(nCodePageTo, 0, szInput, nAnsiLen, *wszOutput, nUnicodeLen);
        (*wszOutput)[nUnicodeLen]=L'\0';
        if (lpnOutputLen) *lpnOutputLen=nUnicodeLen;
        bResult=TRUE;
      }
      API_HeapFree(hHeap, 0, (LPVOID)szInput);
    }
  }
  return bResult;
}

int GetDetectionIndex(DWORD dwLangID)
{
  dwLangID=PRIMARYLANGID(dwLangID);

  if (dwLangID == LANG_RUSSIAN)
    return DETECTINDEX_RUSSIAN;
  else if (IsLangEasternEurope(dwLangID))
    return DETECTINDEX_EASTERNEUROPE;
  else if (IsLangWesternEurope(dwLangID))
    return DETECTINDEX_WESTERNEUROPE;
  else if (dwLangID == LANG_TURKISH)
    return DETECTINDEX_TURKISH;
  else if (dwLangID == LANG_CHINESE)
    return DETECTINDEX_CHINESE;
  else if (dwLangID == LANG_JAPANESE)
    return DETECTINDEX_JAPANESE;
  else if (dwLangID == LANG_KOREAN)
    return DETECTINDEX_KOREAN;
  else
    return DETECTINDEX_NONE;
}

BOOL IsLangEasternEurope(DWORD dwLangID)
{
  dwLangID=PRIMARYLANGID(dwLangID);

  if (dwLangID == LANG_ALBANIAN ||
      dwLangID == LANG_CROATIAN ||
      dwLangID == LANG_CZECH ||
      dwLangID == LANG_HUNGARIAN ||
      dwLangID == LANG_POLISH ||
      dwLangID == LANG_ROMANIAN ||
      dwLangID == LANG_SERBIAN ||
      dwLangID == LANG_SLOVAK ||
      dwLangID == LANG_SLOVENIAN)
  {
    return TRUE;
  }
  return FALSE;
}

BOOL IsLangWesternEurope(DWORD dwLangID)
{
  dwLangID=PRIMARYLANGID(dwLangID);

  if (dwLangID == LANG_AFRIKAANS ||
      dwLangID == LANG_BASQUE ||
      dwLangID == LANG_CATALAN ||
      dwLangID == LANG_DANISH ||
      dwLangID == LANG_DUTCH ||
      dwLangID == LANG_FAEROESE ||
      dwLangID == LANG_FINNISH ||
      dwLangID == LANG_FRENCH ||
      dwLangID == LANG_GERMAN ||
      dwLangID == LANG_ICELANDIC ||
      dwLangID == LANG_INDONESIAN ||
      dwLangID == LANG_ITALIAN ||
      dwLangID == LANG_MALAY ||
      dwLangID == LANG_NORWEGIAN ||
      dwLangID == LANG_PORTUGUESE ||
      dwLangID == LANG_SPANISH ||
      dwLangID == LANG_SWAHILI ||
      dwLangID == LANG_SWEDISH)
  {
    return TRUE;
  }
  return FALSE;
}

BOOL IsCharLegalUTF8(const unsigned char *pSource, unsigned int nTrailingBytes)
{
  const unsigned char *pSrc=pSource + nTrailingBytes;
  unsigned char nChar;

  switch (nTrailingBytes)
  {
    default: return FALSE;

    case 4: if ((nChar=(*--pSrc)) < 0x80 || nChar > 0xBF) return FALSE;
    case 3: if ((nChar=(*--pSrc)) < 0x80 || nChar > 0xBF) return FALSE;
    case 2: if ((nChar=(*--pSrc)) > 0xBF) return FALSE;

    switch (*pSource)
    {
      case 0xE0: if (nChar < 0xA0) return FALSE; break;
      case 0xED: if (nChar > 0x9F) return FALSE; break;
      case 0xF0: if (nChar < 0x90) return FALSE; break;
      case 0xF4: if (nChar > 0x8F) return FALSE; break;
      default:   if (nChar < 0x80) return FALSE;
    }
    case 1: if (*pSource >= 0x80 && *pSource < 0xC2) return FALSE;
  }
  if (*pSource > 0xF4) return FALSE;
  return TRUE;
}

void ChangeTwoBytesOrder(unsigned char *pSrc, UINT_PTR dwSrcSize, unsigned char *pDst)
{
  unsigned char *pLast=pSrc + dwSrcSize - 2;
  unsigned char *pSrcByte=pSrc;
  unsigned char *pDstByte=pDst?pDst:pSrc;
  unsigned char ch2;

  while (pSrcByte <= pLast)
  {
    ch2=*pSrcByte++;
    *pDstByte++=*pSrcByte++;
    *pDstByte++=ch2;
  }
}

void ChangeFourBytesOrder(unsigned char *pSrc, UINT_PTR dwSrcSize, unsigned char *pDst)
{
  unsigned char *pLast=pSrc + dwSrcSize - 4;
  unsigned char *pSrcByte=pSrc;
  unsigned char *pDstByte=pDst?pDst:pSrc;
  unsigned char ch2;
  unsigned char ch3;
  unsigned char ch4;

  while (pSrcByte <= pLast)
  {
    ch4=*pSrcByte++;
    ch3=*pSrcByte++;
    ch2=*pSrcByte++;
    *pDstByte++=*pSrcByte++;
    *pDstByte++=ch2;
    *pDstByte++=ch3;
    *pDstByte++=ch4;
  }
}

BOOL IsCodePageUnicode(int nCodePage)
{
  if (nCodePage == CP_UNICODE_UTF16LE ||
      nCodePage == CP_UNICODE_UTF16BE ||
      nCodePage == CP_UNICODE_UTF32LE ||
      nCodePage == CP_UNICODE_UTF32BE ||
      nCodePage == CP_UNICODE_UTF8)
  {
    return TRUE;
  }
  return FALSE;
}

BOOL IsCodePageValid(int nCodePage)
{
  char ch='A';
  wchar_t wch;

  if (IsCodePageUnicode(nCodePage))
    return TRUE;
  if (MultiByteToWideChar(nCodePage, 0, &ch, 1, &wch, 1))
    return TRUE;

  return FALSE;
}

unsigned int TranslateNewLinesToUnixW(wchar_t *wszWideString, unsigned int nWideStringLen)
{
  unsigned int a;
  unsigned int b;

  for (a=0, b=0; a < nWideStringLen && wszWideString[a]; ++a, ++b)
  {
    if (wszWideString[a] == L'\r' && wszWideString[a + 1] == L'\n') ++a;
    wszWideString[b]=wszWideString[a];
  }
  return b;
}


//// Find/Replace

void RegReadSearch()
{
  wchar_t wszRegKey[MAX_PATH];
  HKEY hKey;
  DWORD dwType;
  DWORD dwSize;

  xprintfW(wszRegKey, L"%s\\Search", APP_REGHOMEW);
  if (RegOpenKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, KEY_READ, &hKey) != ERROR_SUCCESS)
    return;

  if (RegQueryValueExWide(hKey, L"find0", NULL, &dwType, NULL, &dwSize) == ERROR_SUCCESS && dwSize > 0)
  {
    if (wszFindText=(wchar_t *)API_HeapAlloc(hHeap, 0, dwSize))
    {
      if (RegQueryValueExWide(hKey, L"find0", NULL, &dwType, (LPBYTE)wszFindText, &dwSize) == ERROR_SUCCESS)
      {
        nFindTextLen=dwSize / sizeof(wchar_t) - 1;
      }
    }
  }
  RegCloseKey(hKey);
}

BOOL CALLBACK FindAndReplaceDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndFind;
  static HWND hWndReplace;
  static HWND hWndMatchCase;
  static HWND hWndWholeWord;
  static HWND hWndEscapeSeq;
  static HWND hWndRegExp;
  static HWND hWndForward;
  static HWND hWndBackward;
  static HWND hWndBeginning;
  static HWND hWndInSelection;
  static HWND hWndAllFiles;
  static HWND hWndAllFilesGroup;
  static HWND hWndFindNextButton;
  static HWND hWndReplaceButton;
  static HWND hWndReplaceAllButton;
  static HWND hWndCancelButton;
  static BOOL bSpecialCheck=FALSE;
  static BOOL bInSelAutoCheck=FALSE;
  HWND hWndFocus=NULL;
  HWND hWndError;
  HWND hWndComboboxEdit;
  BOOL bReplace=FALSE;
  BOOL bReplaceAll=FALSE;
  BOOL bReplaceAllButtonState=FALSE;
  INT_PTR nReplaceCount;
  INT_PTR nResult;
  int i;

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndFind=GetDlgItem(hDlg, IDC_SEARCH_FIND);
    hWndFindNextButton=GetDlgItem(hDlg, IDC_SEARCH_FIND_BUTTON);
    hWndCancelButton=GetDlgItem(hDlg, IDCANCEL);
    hWndMatchCase=GetDlgItem(hDlg, IDC_SEARCH_MATCHCASE);
    hWndWholeWord=GetDlgItem(hDlg, IDC_SEARCH_WHOLEWORD);
    hWndEscapeSeq=GetDlgItem(hDlg, IDC_SEARCH_ESCAPESEQ);
    hWndRegExp=GetDlgItem(hDlg, IDC_SEARCH_REGEXP);
    hWndForward=GetDlgItem(hDlg, IDC_SEARCH_FORWARD);
    hWndBackward=GetDlgItem(hDlg, IDC_SEARCH_BACKWARD);
    hWndBeginning=GetDlgItem(hDlg, IDC_SEARCH_BEGINNING);
    hWndInSelection=GetDlgItem(hDlg, IDC_SEARCH_INSEL);
    if (nModelessType == MLT_REPLACE)
    {
      hWndReplace=GetDlgItem(hDlg, IDC_SEARCH_REPLACE);
      hWndReplaceButton=GetDlgItem(hDlg, IDC_SEARCH_REPLACE_BUTTON);
      hWndReplaceAllButton=GetDlgItem(hDlg, IDC_SEARCH_ALL_BUTTON);
    }
    hWndAllFiles=GetDlgItem(hDlg, IDC_SEARCH_ALLFILES);
    hWndAllFilesGroup=GetDlgItem(hDlg, IDC_SEARCH_ALLFILES_GROUP);

    if (nModelessType == MLT_REPLACE)
    {
      if (wLastReplaceButtonID != IDC_SEARCH_FIND_BUTTON)
        SendMessage(hDlg, DM_SETDEFID, wLastReplaceButtonID, 0);

      //Reset replace count
      lpFrameCurrent->nReplaceCount=0;
      UpdateStatusUser(lpFrameCurrent, CSB_REPLACECOUNT);
    }
    bNoSearchFinishMsg=FALSE;

    if (moCur.dwSearchOptions & FRF_CHECKINSELIFSEL)
    {
      if (crCurSel.ciMax.nLine - crCurSel.ciMin.nLine > 0)
      {
        moCur.dwSearchOptions|=FRF_SELECTION;
        bInSelAutoCheck=TRUE;
      }
    }
    if (moCur.nSearchStrings)
    {
      FillComboboxSearch(hWndFind, hWndReplace);
    }
    if (!SendMessage(hDlg, AKDLG_PUTFIND, 0, 0))
      SendMessage(hWndFind, CB_SETCURSEL, 0, 0);

    if (!nMDI)
    {
      ShowWindow(hWndAllFiles, SW_HIDE);
      ShowWindow(hWndAllFilesGroup, SW_HIDE);
    }
    if (nModelessType == MLT_REPLACE)
    {
      SendMessage(hWndReplace, CB_SETCURSEL, 0, 0);
    }

    if (moCur.dwSearchOptions & FRF_SELECTION)
      SendMessage(hWndInSelection, BM_SETCHECK, BST_CHECKED, 0);
    else if (moCur.dwSearchOptions & FRF_ALLFILES)
      SendMessage(hWndAllFiles, BM_SETCHECK, BST_CHECKED, 0);
    else if (moCur.dwSearchOptions & FRF_BEGINNING)
      SendMessage(hWndBeginning, BM_SETCHECK, BST_CHECKED, 0);
    else if (moCur.dwSearchOptions & FRF_DOWN)
      SendMessage(hWndForward, BM_SETCHECK, BST_CHECKED, 0);
    else if (moCur.dwSearchOptions & FRF_UP)
      SendMessage(hWndBackward, BM_SETCHECK, BST_CHECKED, 0);

    if (moCur.dwSearchOptions & FRF_MATCHCASE) SendMessage(hWndMatchCase, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwSearchOptions & FRF_WHOLEWORD) SendMessage(hWndWholeWord, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwSearchOptions & FRF_ESCAPESEQ) SendMessage(hWndEscapeSeq, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwSearchOptions & FRF_REGEXP) SendMessage(hWndRegExp, BM_SETCHECK, BST_CHECKED, 0);

    if (hWndComboboxEdit=GetDlgItem(hWndFind, IDC_COMBOBOXEDIT))
    {
      SendMessage(hWndComboboxEdit, EM_LIMITTEXT, PUTFIND_MAXSEL, 0);

      lpOldComboboxEdit=(WNDPROC)GetWindowLongPtrWide(hWndComboboxEdit, GWLP_WNDPROC);
      SetWindowLongPtrWide(hWndComboboxEdit, GWLP_WNDPROC, (UINT_PTR)NewComboboxEditProc);
    }
    if (nModelessType == MLT_REPLACE)
    {
      if (hWndComboboxEdit=GetDlgItem(hWndReplace, IDC_COMBOBOXEDIT))
      {
        SendMessage(hWndComboboxEdit, EM_LIMITTEXT, PUTFIND_MAXSEL, 0);

        lpOldComboboxEdit=(WNDPROC)GetWindowLongPtrWide(hWndComboboxEdit, GWLP_WNDPROC);
        SetWindowLongPtrWide(hWndComboboxEdit, GWLP_WNDPROC, (UINT_PTR)NewComboboxEditProc);
      }
    }
    if (rcFindAndReplaceDlg.right && rcFindAndReplaceDlg.bottom)
      SetWindowPos(hDlg, 0, rcFindAndReplaceDlg.left, rcFindAndReplaceDlg.top, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);

    SendMessage(hDlg, WM_COMMAND, IDC_SETREADONLY, 0);
  }
  else if (uMsg == AKDLG_PUTFIND)
  {
    CHARRANGE64 cr;
    wchar_t *wszData;
    INT_PTR nSelLen;
    BOOL bResult=FALSE;

    if (nSelLen=SendMessage(lpFrameCurrent->ei.hWndEdit, EM_EXGETSEL64, 0, (LPARAM)&cr))
    {
      if (nSelLen <= PUTFIND_MAXSEL && !(moCur.dwSearchOptions & FRF_SELECTION) && !SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETCOLUMNSEL, 0, 0))
      {
        if (ExGetRangeTextW(lpFrameCurrent->ei.hWndEdit, &crCurSel.ciMin, &crCurSel.ciMax, FALSE, &wszData, AELB_R, FALSE))
        {
          SetWindowTextWide(hWndFind, wszData);
          bResult=TRUE;
          FreeText((LPVOID)wszData);
          SendMessage(hDlg, WM_COMMAND, MAKELONG(IDC_SEARCH_FIND, CBN_EDITCHANGE), 0);
        }
      }
    }
    SetWindowLongPtrWide(hDlg, DWLP_MSGRESULT, bResult);
    return TRUE;
  }
  else if (uMsg == WM_COMMAND)
  {
    WORD wCommand=LOWORD(wParam);

    if (wCommand == IDC_SEARCH_REPLACE)
      return TRUE;
    else if (wCommand == IDC_SEARCH_MATCHCASE)
    {
      if (SendMessage(hWndMatchCase, BM_GETCHECK, 0, 0)) moCur.dwSearchOptions|=FRF_MATCHCASE;
      else moCur.dwSearchOptions&=~FRF_MATCHCASE;
      return TRUE;
    }
    else if (wCommand == IDC_SEARCH_WHOLEWORD)
    {
      if (SendMessage(hWndWholeWord, BM_GETCHECK, 0, 0)) moCur.dwSearchOptions|=FRF_WHOLEWORD;
      else moCur.dwSearchOptions&=~FRF_WHOLEWORD;
      return TRUE;
    }
    else if (wCommand == IDC_SEARCH_ESCAPESEQ)
    {
      if (SendMessage(hWndEscapeSeq, BM_GETCHECK, 0, 0)) moCur.dwSearchOptions|=FRF_ESCAPESEQ;
      else moCur.dwSearchOptions&=~FRF_ESCAPESEQ;
      if (moCur.dwSearchOptions & FRF_ESCAPESEQ)
      {
        SendMessage(hWndRegExp, BM_SETCHECK, BST_UNCHECKED, 0);
        moCur.dwSearchOptions&=~FRF_REGEXP;
      }
      return TRUE;
    }
    else if (wCommand == IDC_SEARCH_REGEXP)
    {
      if (SendMessage(hWndRegExp, BM_GETCHECK, 0, 0)) moCur.dwSearchOptions|=FRF_REGEXP;
      else moCur.dwSearchOptions&=~FRF_REGEXP;
      if (moCur.dwSearchOptions & FRF_REGEXP)
      {
        SendMessage(hWndEscapeSeq, BM_SETCHECK, BST_UNCHECKED, 0);
        moCur.dwSearchOptions&=~FRF_ESCAPESEQ;
      }
      return TRUE;
    }
    else if (wCommand == IDC_SEARCH_FORWARD)
    {
      moCur.dwSearchOptions&=~FRF_UP&~FRF_SELECTION&~FRF_ALLFILES&~FRF_BEGINNING;
      moCur.dwSearchOptions|=FRF_DOWN;
    }
    else if (wCommand == IDC_SEARCH_BACKWARD)
    {
      moCur.dwSearchOptions&=~FRF_DOWN&~FRF_SELECTION&~FRF_ALLFILES&~FRF_BEGINNING;
      moCur.dwSearchOptions|=FRF_UP;
    }
    else if (wCommand == IDC_SEARCH_BEGINNING)
    {
      moCur.dwSearchOptions&=~FRF_UP&~FRF_SELECTION&~FRF_ALLFILES;
      moCur.dwSearchOptions|=FRF_BEGINNING|FRF_DOWN;
    }
    else if (wCommand == IDC_SEARCH_INSEL)
    {
      moCur.dwSearchOptions&=~FRF_UP&~FRF_ALLFILES&~FRF_BEGINNING;
      moCur.dwSearchOptions|=FRF_SELECTION|FRF_DOWN;
    }
    else if (wCommand == IDC_SEARCH_ALLFILES)
    {
      moCur.dwSearchOptions&=~FRF_UP&~FRF_SELECTION;
      moCur.dwSearchOptions|=FRF_ALLFILES|FRF_BEGINNING|FRF_DOWN;
    }
    else if (wCommand == IDC_SETREADONLY)
    {
      if (nModelessType == MLT_REPLACE)
      {
        EnableWindow(hWndReplaceButton, !lpFrameCurrent->ei.bReadOnly);
        EnableWindow(hWndReplaceAllButton, !lpFrameCurrent->ei.bReadOnly);
      }
    }

    if (wCommand == IDC_SEARCH_FIND ||
        wCommand == IDC_SEARCH_FORWARD ||
        wCommand == IDC_SEARCH_BACKWARD ||
        wCommand == IDC_SEARCH_BEGINNING ||
        wCommand == IDC_SEARCH_INSEL ||
        wCommand == IDC_SEARCH_ALLFILES)
    {
      EnableWindow(hWndFindNextButton, TRUE);
      if (nModelessType == MLT_REPLACE)
      {
        if (!lpFrameCurrent->ei.bReadOnly)
        {
          EnableWindow(hWndReplaceButton, TRUE);
          EnableWindow(hWndReplaceAllButton, TRUE);
        }
      }

      if (wCommand != IDC_SEARCH_FIND ||
          HIWORD(wParam) == CBN_EDITCHANGE)
      {
        if (bSpecialCheck == TRUE)
        {
          bSpecialCheck=FALSE;
          if (HIWORD(wParam) == CBN_EDITCHANGE)
            moCur.dwSearchOptions|=FRF_BEGINNING;
          SendMessage(hWndBeginning, BM_SETSTATE, FALSE, 0);
          if (nMDI) SendMessage(hWndAllFiles, BM_SETSTATE, FALSE, 0);
        }
      }

      if (HIWORD(wParam) == CBN_SELCHANGE)
        i=ComboBox_GetLBTextLenWide(hWndFind, (int)SendMessage(hWndFind, CB_GETCURSEL, 0, 0));
      else
        i=GetWindowTextLengthWide(hWndFind);

      if (i == 0)
      {
        EnableWindow(hWndFindNextButton, FALSE);
        if (nModelessType == MLT_REPLACE)
        {
          EnableWindow(hWndReplaceButton, FALSE);
          EnableWindow(hWndReplaceAllButton, FALSE);
        }
      }
      return TRUE;
    }
    if (wCommand == IDC_SEARCH_FIND_BUTTON ||
        wCommand == IDC_SEARCH_REPLACE_BUTTON ||
        wCommand == IDC_SEARCH_ALL_BUTTON)
    {
      if (nModelessType == MLT_REPLACE)
      {
        if (wCommand != wLastReplaceButtonID)
        {
          SendMessage(hDlg, DM_SETDEFID, wCommand, 0);
          SetDefButtonStyle(hWndFindNextButton, (HWND)lParam);
          SetDefButtonStyle(hWndReplaceButton, (HWND)lParam);
          SetDefButtonStyle(hWndReplaceAllButton, (HWND)lParam);
          wLastReplaceButtonID=wCommand;
        }

        if (wCommand == IDC_SEARCH_REPLACE_BUTTON)
          bReplace=TRUE;
        else if (wCommand == IDC_SEARCH_ALL_BUTTON)
          bReplaceAll=TRUE;
      }
      FreeMemorySearch();

      if ((nFindTextLen=GetComboboxSearchText(hWndFind, &wszFindText, NEWLINE_MAC)) <= 0)
      {
        FreeMemorySearch();
        return TRUE;
      }
      if (nModelessType == MLT_REPLACE)
      {
        if ((nReplaceTextLen=GetComboboxSearchText(hWndReplace, &wszReplaceText, NEWLINE_MAC)) < 0)
        {
          FreeMemorySearch();
          return TRUE;
        }
      }

      hWndFocus=GetFocus();
      if (bReplaceAll)
        bReplaceAllButtonState=EnableWindow(hWndReplaceAllButton, FALSE);

      if (!(moCur.dwSearchOptions & FRF_SELECTION) && (moCur.dwSearchOptions & FRF_ALLFILES))
      {
        FRAMEDATA *lpFrameInit=lpFrameCurrent;
        INT_PTR nChanges=0;
        int nChangedFiles=0;

        if (bReplaceAll == TRUE)
        {
          if (bSpecialCheck == TRUE)
          {
            bSpecialCheck=FALSE;
            moCur.dwSearchOptions|=FRF_BEGINNING;
            SendMessage(hWndAllFiles, BM_SETSTATE, FALSE, 0);
          }

          do
          {
            TextReplaceW(lpFrameCurrent, moCur.dwSearchOptions, wszFindText, nFindTextLen, wszReplaceText, nReplaceTextLen, TRUE, &nReplaceCount);
            if (!hDlgModeless) break;

            if (nReplaceCount)
            {
              ++nChangedFiles;
              nChanges+=nReplaceCount;
            }
            lpFrameCurrent=ActivateNextFrameWindow(lpFrameCurrent, FALSE);
          }
          while (lpFrameCurrent != lpFrameInit);

          //Show result
          lpFrameCurrent->nReplaceCount=nChanges;
          UpdateStatusUser(lpFrameCurrent, CSB_REPLACECOUNT);

          if (moCur.dwSearchOptions & FRF_REPLACEALLANDCLOSE)
          {
            PostMessage(hDlg, WM_COMMAND, IDCANCEL, 0);
            return TRUE;
          }
          else if (!(moCur.dwSearchOptions & FRF_REPLACEALLNOMSG))
          {
            API_LoadStringW(hLangLib, MSG_REPLACE_COUNT_ALLFILES, wbuf, BUFFER_SIZE);
            xprintfW(wszMsg, wbuf, nChangedFiles, nChanges);
            API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONINFORMATION);
          }
        }
        else
        {
          do
          {
            if (bReplace == TRUE)
              nResult=TextReplaceW(lpFrameCurrent, moCur.dwSearchOptions & ~FRF_CYCLESEARCH, wszFindText, nFindTextLen, wszReplaceText, nReplaceTextLen, FALSE, NULL);
            else
              nResult=TextFindW(lpFrameCurrent, moCur.dwSearchOptions & ~FRF_CYCLESEARCH, wszFindText, nFindTextLen);

            if (nResult == -1)
            {
              if (lpFrameCurrent->nCompileErrorOffset &&
                  ((moCur.dwSearchOptions & FRF_REGEXP) ||
                   (moCur.dwSearchOptions & FRF_ESCAPESEQ)))
              {
                API_LoadStringW(hLangLib, MSG_ERROR_SYNTAX, wszMsg, BUFFER_SIZE);
                API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
                hWndFocus=NULL;
                if (lpFrameCurrent->bCompileErrorReplace)
                  hWndError=hWndReplace;
                else
                  hWndError=hWndFind;
                hWndFocus=NULL;
                SetFocus(hWndError);
                SendMessage(hWndError, CB_SETEDITSEL, 0, MAKELONG(lpFrameCurrent->nCompileErrorOffset - 1, 0xFFFF));

                bNoSearchFinishMsg=TRUE;
                break;
              }
              else
              {
                if (bSpecialCheck == TRUE)
                {
                  bSpecialCheck=FALSE;
                  moCur.dwSearchOptions|=FRF_BEGINNING;
                  SendMessage(hWndAllFiles, BM_SETSTATE, FALSE, 0);
                }
                lpFrameCurrent=ActivateNextFrameWindow(lpFrameCurrent, FALSE);
              }
            }
            else
            {
              if (bSpecialCheck == FALSE)
              {
                bSpecialCheck=TRUE;
                moCur.dwSearchOptions&=~FRF_BEGINNING;
                SendMessage(hWndAllFiles, BM_SETSTATE, TRUE, 0);
              }
              break;
            }
          }
          while (lpFrameCurrent != lpFrameInit);

          if (nResult == -1 && !bNoSearchFinishMsg)
          {
            API_LoadStringW(hLangLib, MSG_SEARCH_ENDED, wszMsg, BUFFER_SIZE);
            API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONINFORMATION);
          }
        }
      }
      else
      {
        if (bReplaceAll == TRUE)
        {
          if (bSpecialCheck == TRUE)
          {
            bSpecialCheck=FALSE;
            moCur.dwSearchOptions|=FRF_BEGINNING;
            SendMessage(hWndBeginning, BM_SETSTATE, FALSE, 0);
          }
        }

        if (bReplace == TRUE || bReplaceAll == TRUE)
          nResult=TextReplaceW(lpFrameCurrent, moCur.dwSearchOptions, wszFindText, nFindTextLen, wszReplaceText, nReplaceTextLen, bReplaceAll, &nReplaceCount);
        else
          nResult=TextFindW(lpFrameCurrent, moCur.dwSearchOptions, wszFindText, nFindTextLen);

        if (nResult == -1)
        {
          if (lpFrameCurrent->nCompileErrorOffset &&
              ((moCur.dwSearchOptions & FRF_REGEXP) ||
               (moCur.dwSearchOptions & FRF_ESCAPESEQ)))
          {
            API_LoadStringW(hLangLib, MSG_ERROR_SYNTAX, wszMsg, BUFFER_SIZE);
            API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
            if (lpFrameCurrent->bCompileErrorReplace)
              hWndError=hWndReplace;
            else
              hWndError=hWndFind;
            hWndFocus=NULL;
            SetFocus(hWndError);
            SendMessage(hWndError, CB_SETEDITSEL, 0, MAKELONG(lpFrameCurrent->nCompileErrorOffset - 1, 0xFFFF));
          }
          else
          {
            if (bSpecialCheck == TRUE)
            {
              bSpecialCheck=FALSE;
              moCur.dwSearchOptions|=FRF_BEGINNING;
              SendMessage(hWndBeginning, BM_SETSTATE, FALSE, 0);
            }
            if (bReplaceAll == TRUE)
            {
              //Show result
              lpFrameCurrent->nReplaceCount=nReplaceCount;
              UpdateStatusUser(lpFrameCurrent, CSB_REPLACECOUNT);

              if (moCur.dwSearchOptions & FRF_REPLACEALLANDCLOSE)
              {
                PostMessage(hDlg, WM_COMMAND, IDCANCEL, 0);
                return TRUE;
              }
              else if (!(moCur.dwSearchOptions & FRF_REPLACEALLNOMSG))
              {
                API_LoadStringW(hLangLib, MSG_REPLACE_COUNT, wbuf, BUFFER_SIZE);
                xprintfW(wszMsg, wbuf, nReplaceCount);
                API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONINFORMATION);
              }
            }
            else if (!bNoSearchFinishMsg)
            {
              API_LoadStringW(hLangLib, MSG_SEARCH_ENDED, wszMsg, BUFFER_SIZE);
              API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONINFORMATION);
            }
          }
        }
        else
        {
          if (!(moCur.dwSearchOptions & FRF_SELECTION) && (moCur.dwSearchOptions & FRF_BEGINNING))
          {
            bSpecialCheck=TRUE;
            moCur.dwSearchOptions&=~FRF_BEGINNING;
            SendMessage(hWndBeginning, BM_SETSTATE, TRUE, 0);
          }
        }
      }
      if (bReplaceAll)
        EnableWindow(hWndReplaceAllButton, !bReplaceAllButtonState);
      if (hWndFocus)
        SetFocus(hWndFocus);

      return TRUE;
    }
    else if (wCommand == IDCANCEL)
    {
      if (bSpecialCheck == TRUE)
      {
        bSpecialCheck=FALSE;
        moCur.dwSearchOptions|=FRF_BEGINNING;
      }
      if (bInSelAutoCheck)
        moCur.dwSearchOptions&=~FRF_SELECTION;
      if (moCur.nSearchStrings)
        SaveComboboxSearch(hWndFind, hWndReplace);
      GetWindowPos(hDlg, NULL, &rcFindAndReplaceDlg);
      DestroyWindow(hDlg);
      hDlgModeless=NULL;
      nModelessType=MLT_NONE;
      return TRUE;
    }
  }
  else if (uMsg == WM_KEYDOWN)
  {
    if (wParam == VK_F3)
    {
      if (GetWindowTextLengthWide(hWndFind))
      {
        FreeMemorySearch();

        if ((nFindTextLen=GetComboboxSearchText(hWndFind, &wszFindText, NEWLINE_MAC)) <= 0)
        {
          FreeMemorySearch();
          return TRUE;
        }

        if (GetKeyState(VK_SHIFT) >= 0)
          DoEditFindNextDown(lpFrameCurrent);
        else
          DoEditFindNextUp(lpFrameCurrent);

        PostMessage(hDlg, WM_COMMAND, IDCANCEL, 0);
      }
    }
  }
  return FALSE;
}

LRESULT CALLBACK NewComboboxEditProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  if (uMsg == WM_PASTE)
  {
    PasteInEditAsRichEdit(hWnd, PUTFIND_MAXSEL);
    return TRUE;
  }

  return CallWindowProcWide(lpOldComboboxEdit, hWnd, uMsg, wParam, lParam);
}

void FillComboboxSearch(HWND hWndFind, HWND hWndReplace)
{
  wchar_t wszRegKey[MAX_PATH];
  wchar_t wszRegValue[32];
  wchar_t *wszData;
  HWND hWnd;
  HKEY hKey;
  DWORD dwType;
  DWORD dwSize;
  int i;

  xprintfW(wszRegKey, L"%s\\Search", APP_REGHOMEW);
  if (RegOpenKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, KEY_READ, &hKey) != ERROR_SUCCESS)
    return;

  hWnd=hWndFind;
  xstrcpyW(wbuf, L"find%d");

  ReadW:
  for (i=0; i < moCur.nSearchStrings; ++i)
  {
    xprintfW(wszRegValue, wbuf, i);

    if (RegQueryValueExWide(hKey, wszRegValue, NULL, &dwType, NULL, &dwSize) == ERROR_SUCCESS && dwSize > 0)
    {
      if (wszData=(wchar_t *)API_HeapAlloc(hHeap, 0, dwSize))
      {
        if (RegQueryValueExWide(hKey, wszRegValue, NULL, &dwType, (LPBYTE)wszData, &dwSize) == ERROR_SUCCESS)
        {
          ComboBox_AddStringWide(hWnd, wszData);
          API_HeapFree(hHeap, 0, (LPVOID)wszData);
          continue;
        }
      }
    }
    break;
  }
  if (nModelessType == MLT_REPLACE && hWnd != hWndReplace)
  {
    hWnd=hWndReplace;
    xstrcpyW(wbuf, L"replace%d");
    goto ReadW;
  }
  RegCloseKey(hKey);
}

int GetComboboxSearchText(HWND hWnd, wchar_t **wszText, int nNewLine)
{
  int nTextLen=-1;
  int nIndex;
  int nItemLen;

  nTextLen=GetWindowTextLengthWide(hWnd) + 1;

  if (*wszText=AllocWideStr(nTextLen))
  {
    nTextLen=GetWindowTextWide(hWnd, *wszText, nTextLen);

    if (moCur.nSearchStrings)
    {
      if (**wszText)
      {
        if ((nIndex=ComboBox_FindStringExactWide(hWnd, 0, *wszText)) != CB_ERR)
          SendMessage(hWnd, CB_DELETESTRING, (WPARAM)nIndex, 0);
      }
      else
      {
        for (nIndex=0; (nItemLen=ComboBox_GetLBTextLenWide(hWnd, nIndex)) != CB_ERR && nItemLen; ++nIndex);
        if (!nItemLen) SendMessage(hWnd, CB_DELETESTRING, (WPARAM)nIndex, 0);
      }
      SendMessage(hWnd, CB_DELETESTRING, moCur.nSearchStrings - 1, 0);
      ComboBox_InsertStringWide(hWnd, 0, *wszText);
      SendMessage(hWnd, CB_SETCURSEL, 0, 0);
    }
  }
  return nTextLen;
}

void SaveComboboxSearch(HWND hWndFind, HWND hWndReplace)
{
  wchar_t wszRegKey[MAX_PATH];
  wchar_t wszRegValue[32];
  wchar_t *wszRegData;
  HWND hWnd;
  HKEY hKey;
  int nSize;
  int i;

  xprintfW(wszRegKey, L"%s\\Search", APP_REGHOMEW);
  if (RegCreateKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL) != ERROR_SUCCESS)
    return;

  hWnd=hWndFind;
  xstrcpyW(wbuf, L"find%d");

  SaveW:
  for (i=0; i < moCur.nSearchStrings; ++i)
  {
    xprintfW(wszRegValue, wbuf, i);

    if ((nSize=ComboBox_GetLBTextLenWide(hWnd, i)) != CB_ERR)
    {
      ++nSize;

      if (wszRegData=AllocWideStr(nSize + 1))
      {
        ComboBox_GetLBTextWide(hWnd, i, wszRegData);
        RegSetValueExWide(hKey, wszRegValue, 0, REG_SZ, (LPBYTE)wszRegData, nSize * sizeof(wchar_t));
        FreeWideStr(wszRegData);
        continue;
      }
    }
    break;
  }
  if (nModelessType == MLT_REPLACE && hWnd != hWndReplace)
  {
    hWnd=hWndReplace;
    xstrcpyW(wbuf, L"replace%d");
    goto SaveW;
  }
  RegCloseKey(hKey);
}

INT_PTR TextFindW(FRAMEDATA *lpFrame, DWORD dwFlags, const wchar_t *wpFindIt, int nFindItLen)
{
  AEFINDTEXTW ft;
  CHARRANGE64 cr;
  wchar_t *wszFindItEsc=(wchar_t *)wpFindIt;
  int nFindItLenEsc=nFindItLen;
  INT_PTR nResult=-1;
  BOOL bCycleCheck=TRUE;
  BOOL bFound=FALSE;

  if (!nFindItLen)
    goto End;
  if (nFindItLen == -1)
    nFindItLen=(int)xstrlenW(wpFindIt);

  if (dwFlags & FRF_ESCAPESEQ)
  {
    lpFrame->nCompileErrorOffset=0;
    lpFrame->bCompileErrorReplace=FALSE;

    if (wszFindItEsc=AllocWideStr(nFindItLen + 1))
    {
      if ((nFindItLenEsc=(int)EscapeStringToEscapeDataW(wpFindIt, nFindItLen, wszFindItEsc, NEWLINE_MAC)) < 0)
      {
        lpFrame->nCompileErrorOffset=-nFindItLenEsc;
        goto End;
      }
    }
    else goto End;
  }

  //Leave only FRF_* flags that corresponds to AEFR_*.
  ft.dwFlags=dwFlags & (AEFR_DOWN|AEFR_WHOLEWORD|AEFR_MATCHCASE|AEFR_REGEXP);
  if (dwFlags & FRF_REGEXPNONEWLINEDOT)
    ft.dwFlags|=AEFR_REGEXPNONEWLINEDOT;
  ft.pText=wszFindItEsc;
  ft.dwTextLen=nFindItLenEsc;
  ft.nNewLine=AELB_R;

  if (dwFlags & FRF_SELECTION)
  {
    if (!(dwFlags & FRF_FINDFROMREPLACE))
    {
      ft.dwFlags=dwFlags;
      ft.pText=wszFindItEsc;
      ft.dwTextLen=nFindItLenEsc;
      ft.nNewLine=AELB_R;
      if (SendMessage(lpFrame->ei.hWndEdit, AEM_ISMATCHW, (WPARAM)&crCurSel.ciMin, (LPARAM)&ft))
        ft.crSearch.ciMin=ft.crFound.ciMax;
      else
        AEC_NextCharEx(&crCurSel.ciMin, &ft.crSearch.ciMin);
    }
    else ft.crSearch.ciMin=crCurSel.ciMin;

    ft.crSearch.ciMax=crCurSel.ciMax;
  }
  else if (dwFlags & FRF_BEGINNING)
  {
    SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&ft.crSearch.ciMin);
    SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&ft.crSearch.ciMax);
    ft.dwFlags|=AEFR_REGEXPMINMATCH;
  }
  else if (dwFlags & FRF_DOWN)
  {
    ft.crSearch.ciMin=crCurSel.ciMax;
    SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&ft.crSearch.ciMax);
  }
  else if (dwFlags & FRF_UP)
  {
    SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&ft.crSearch.ciMin);
    ft.crSearch.ciMax=crCurSel.ciMin;
  }
  else goto End;

  FindIt:
  bFound=(BOOL)SendMessage(lpFrame->ei.hWndEdit, AEM_FINDTEXTW, 0, (LPARAM)&ft);
  if (dwFlags & FRF_REGEXP)
  {
    lpFrame->nCompileErrorOffset=ft.nCompileErrorOffset;
    lpFrame->bCompileErrorReplace=FALSE;
  }

  if (bCycleCheck && !bFound)
  {
    if ((dwFlags & FRF_CYCLESEARCH) && !(dwFlags & FRF_SELECTION) && !(dwFlags & FRF_BEGINNING))
    {
      AECHARINDEX ciChar;
      HWND hWndParent;
      int nAnswer=IDOK;

      if (dwFlags & FRF_DOWN)
      {
        SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&ciChar);
        if (AEC_IndexCompare(&ft.crSearch.ciMin, &ciChar))
        {
          ft.crSearch.ciMin=ciChar;
          ft.crSearch.ciMax=crCurSel.ciMax;
        }
        else nAnswer=-1;
      }
      else if (dwFlags & FRF_UP)
      {
        SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&ciChar);
        if (AEC_IndexCompare(&ft.crSearch.ciMax, &ciChar))
        {
          ft.crSearch.ciMin=crCurSel.ciMin;
          ft.crSearch.ciMax=ciChar;
        }
        else nAnswer=-1;
      }
      if (nAnswer == IDOK && (dwFlags & FRF_CYCLESEARCHPROMPT))
      {
        hWndParent=(nModelessType == MLT_FIND || nModelessType == MLT_REPLACE)?hDlgModeless:hMainWnd;

        API_LoadStringW(hLangLib, MSG_SEARCH_CYCLEPROMPT, wszMsg, BUFFER_SIZE);
        nAnswer=API_MessageBox(hWndParent, wszMsg, APP_MAIN_TITLEW, MB_OKCANCEL|MB_ICONQUESTION);
      }
      if (nAnswer == IDOK)
      {
        ft.dwFlags|=AEFR_REGEXPMINMATCH;
        bCycleCheck=FALSE;
        goto FindIt;
      }
      else if (nAnswer == IDCANCEL)
        bNoSearchFinishMsg=TRUE;
    }
  }
  lpFrame->bReachedEOF=!bFound;
  UpdateStatusUser(lpFrame, CSB_SEARCHENDED);

  if (bFound)
  {
    if (dwFlags & FRF_SELECTION)
      ft.crFound.ciMax=crCurSel.ciMax;

    SetSel(lpFrame->ei.hWndEdit, &ft.crFound, AESELT_LOCKSCROLL, NULL);
    ScrollCaret(lpFrame->ei.hWndEdit);
    SendMessage(lpFrame->ei.hWndEdit, EM_EXGETSEL64, 0, (LPARAM)&cr);
    nResult=cr.cpMin;
    goto End;
  }
  SendMessage(hMainWnd, AKDN_SEARCH_ENDED, (WPARAM)hDlgModeless, 0);

  End:
  if (wszFindItEsc != wpFindIt)
    FreeWideStr(wszFindItEsc);
  return nResult;
}

INT_PTR TextReplaceW(FRAMEDATA *lpFrame, DWORD dwFlags, const wchar_t *wpFindIt, int nFindItLen, const wchar_t *wpReplaceWith, int nReplaceWithLen, BOOL bAll, INT_PTR *nReplaceCount)
{
  AECHARRANGE crInitialSel;
  AECHARRANGE crFullText;
  AECHARRANGE crRange;
  AECHARRANGE crInsert;
  AECHARINDEX ciFirstVisibleBefore;
  CHARRANGE64 crInitialRE;
  PATREPLACE pr;
  wchar_t *wszFindItEsc=(wchar_t *)wpFindIt;
  wchar_t *wszReplaceWithEsc=(wchar_t *)wpReplaceWith;
  wchar_t *wszFullText;
  wchar_t *wszRangeText;
  wchar_t *wszResultText=NULL;
  int nFindItLenEsc;
  int nReplaceWithLenEsc;
  int nGetTextNewLine;
  int nReplaceSelNewLine;
  INT_PTR nMin=0;
  INT_PTR nMax=0;
  INT_PTR nMaxBackward=0;
  INT_PTR nFirstVisible;
  INT_PTR nFullTextLen;
  INT_PTR nRangeTextLen;
  INT_PTR nResultTextLen;
  INT_PTR nChanges=0;
  INT_PTR nResult=-1;
  int nFirstLine=0;
  int i;
  BOOL bInitialColumnSel;
  BOOL bColumnSel;
  BOOL bCaretAtStart=FALSE;

  if (IsReadOnly(lpFrame->ei.hWndEdit) || !nFindItLen)
    goto End;
  if (!AEC_IndexCompare(&crCurSel.ciMin, &ciCurCaret))
    bCaretAtStart=TRUE;
  crInitialSel.ciMin=crCurSel.ciMin;
  crInitialSel.ciMax=crCurSel.ciMax;
  if (nFindItLen == -1)
    nFindItLen=(int)xstrlenW(wpFindIt);
  nFindItLenEsc=nFindItLen;

  if (nReplaceWithLen == -1)
    nReplaceWithLen=(int)xstrlenW(wpReplaceWith);
  nReplaceWithLenEsc=nReplaceWithLen;

  if (dwFlags & FRF_ESCAPESEQ)
  {
    lpFrame->nCompileErrorOffset=0;
    lpFrame->bCompileErrorReplace=FALSE;

    if (wszFindItEsc=AllocWideStr(nFindItLen + 1))
    {
      if ((nFindItLenEsc=(int)EscapeStringToEscapeDataW(wpFindIt, nFindItLen, wszFindItEsc, NEWLINE_MAC)) < 0)
      {
        lpFrame->nCompileErrorOffset=-nFindItLenEsc;
        goto End;
      }
    }
    else goto End;

    if (wszReplaceWithEsc=AllocWideStr(nReplaceWithLen + 1))
    {
      if ((nReplaceWithLenEsc=(int)EscapeStringToEscapeDataW(wpReplaceWith, nReplaceWithLen, wszReplaceWithEsc, NEWLINE_MAC)) < 0)
      {
        lpFrame->nCompileErrorOffset=-nReplaceWithLenEsc;
        lpFrame->bCompileErrorReplace=TRUE;
        goto End;
      }
    }
    else goto End;
  }
  if (dwFlags & FRF_REGEXP)
  {
    pr.wpPat=wszFindItEsc;
    pr.wpMaxPat=wszFindItEsc + nFindItLenEsc;
    pr.wpRep=wszReplaceWithEsc;
    pr.wpMaxRep=wszReplaceWithEsc + nReplaceWithLenEsc;
    pr.dwOptions=(dwFlags & FRF_MATCHCASE?RESE_MATCHCASE:0) |
                 (dwFlags & FRF_REGEXPNONEWLINEDOT?RESE_NONEWLINEDOT:0) |
                 RESE_MULTILINE;
    pr.wpDelim=lpFrame->wszWordDelimiters;
    pr.wpMaxDelim=lpFrame->wszWordDelimiters + xstrlenW(lpFrame->wszWordDelimiters);
    pr.wpNewLine=GetNewLineString(lpFrame->ei.nNewLine);
    lpFrame->nCompileErrorOffset=0;
    lpFrame->bCompileErrorReplace=FALSE;
  }

  if (bAll)
  {
    bInitialColumnSel=(BOOL)SendMessage(lpFrame->ei.hWndEdit, AEM_GETCOLUMNSEL, 0, 0);

    if (dwFlags & FRF_SELECTION)
    {
      crRange.ciMin=crCurSel.ciMin;
      crRange.ciMax=crCurSel.ciMax;
      bColumnSel=bInitialColumnSel;

      if (dwFlags & FRF_WHOLEWORD)
      {
        if (SendMessage(lpFrame->ei.hWndEdit, AEM_ISDELIMITER, AEDLM_WORD|AEDLM_PREVCHAR, (LPARAM)&crRange.ciMin))
          dwFlags|=FRF_WHOLEWORDGOODSTART;
        if (SendMessage(lpFrame->ei.hWndEdit, AEM_ISDELIMITER, AEDLM_WORD, (LPARAM)&crRange.ciMax))
          dwFlags|=FRF_WHOLEWORDGOODEND;
      }
    }
    else if (dwFlags & FRF_BEGINNING)
    {
      SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&crRange.ciMin);
      SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&crRange.ciMax);
      bColumnSel=FALSE;

      if (dwFlags & FRF_WHOLEWORD)
      {
        dwFlags|=FRF_WHOLEWORDGOODSTART|FRF_WHOLEWORDGOODEND;
      }
    }
    else if (dwFlags & FRF_DOWN)
    {
      crRange.ciMin=crCurSel.ciMin;
      SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&crRange.ciMax);
      bColumnSel=FALSE;

      if (dwFlags & FRF_WHOLEWORD)
      {
        if (SendMessage(lpFrame->ei.hWndEdit, AEM_ISDELIMITER, AEDLM_WORD|AEDLM_PREVCHAR, (LPARAM)&crRange.ciMin))
          dwFlags|=FRF_WHOLEWORDGOODSTART|FRF_WHOLEWORDGOODEND;
        else
          dwFlags|=FRF_WHOLEWORDGOODEND;
      }
    }
    else if (dwFlags & FRF_UP)
    {
      SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&crRange.ciMin);
      crRange.ciMax=crCurSel.ciMax;
      bColumnSel=FALSE;

      if (dwFlags & FRF_WHOLEWORD)
      {
        if (SendMessage(lpFrame->ei.hWndEdit, AEM_ISDELIMITER, AEDLM_WORD, (LPARAM)&crRange.ciMax))
          dwFlags|=FRF_WHOLEWORDGOODSTART|FRF_WHOLEWORDGOODEND;
        else
          dwFlags|=FRF_WHOLEWORDGOODSTART;
      }
    }
    else goto End;

    if (!(dwFlags & FRF_REGEXP))
    {
      //Find new line in wszFindItEsc and wszReplaceWithEsc.
      nGetTextNewLine=HIWORD(SendMessage(lpFrame->ei.hWndEdit, AEM_GETNEWLINE, 0, 0));

      if (nGetTextNewLine == AELB_ASIS)
      {
        for (i=0; i < nFindItLenEsc; ++i)
        {
          if (wszFindItEsc[i] == L'\r')
          {
            nGetTextNewLine=AELB_R;
            break;
          }
        }
        if (nGetTextNewLine == AELB_ASIS)
        {
          for (i=0; i < nReplaceWithLenEsc; ++i)
          {
            if (wszReplaceWithEsc[i] == L'\r')
            {
              nGetTextNewLine=AELB_R;
              break;
            }
          }
        }
      }
      else nGetTextNewLine=AELB_R;
    }
    else nGetTextNewLine=AELB_ASIS;

    //Check for possible backward offset
    crFullText.ciMin=crRange.ciMin;
    crFullText.ciMax=crRange.ciMax;

    if (dwFlags & FRF_REGEXP)
    {
      STACKREGROUP sreg;

      sreg.first=0;
      sreg.last=0;
      sreg.dwOptions=pr.dwOptions & (RESE_MATCHCASE|RESE_MULTILINE|RESE_WHOLEWORD|RESE_NONEWLINEDOT);
      sreg.wpDelim=pr.wpDelim;
      sreg.wpMaxDelim=pr.wpMaxDelim;
      if (lpFrame->nCompileErrorOffset=PatCompile(&sreg, pr.wpPat, pr.wpMaxPat))
        goto End;
      PatFree(&sreg);

      if (sreg.nMaxBackward)
      {
        AECHARINDEX ciCount=crFullText.ciMin;
        INT_PTR nCount=sreg.nMaxBackward;

        //Find backward offset
        while (nCount > 0 && AEC_PrevChar(&ciCount))
          --nCount;
        if (!nCount)
        {
          crFullText.ciMin=ciCount;
          nMaxBackward=sreg.nMaxBackward;
        }
      }
    }

    if (nFullTextLen=ExGetRangeTextW(lpFrame->ei.hWndEdit, &crFullText.ciMin, &crFullText.ciMax, bColumnSel, &wszFullText, nGetTextNewLine, TRUE))
    {
      wszRangeText=wszFullText + nMaxBackward;
      nRangeTextLen=nFullTextLen - nMaxBackward;

      //Calculate result string length
      if (dwFlags & FRF_REGEXP)
      {
        pr.wpStr=wszRangeText;
        pr.wpMaxStr=wszRangeText + nRangeTextLen;
        pr.wpText=wszFullText;
        pr.wpMaxText=pr.wpMaxStr;
        pr.dwOptions|=RESE_GLOBAL;
        if (dwFlags & FRF_WHOLEWORD)
          pr.dwOptions|=RESE_WHOLEWORD;
        pr.wszResult=NULL;
        nResultTextLen=PatReplace(&pr);

        if (pr.nReplaceCount)
          wszResultText=AllocWideStr(nResultTextLen);
        else
          lpFrame->nCompileErrorOffset=pr.nErrorOffset;
      }
      else
      {
        if (nFindItLenEsc < nReplaceWithLenEsc)
        {
          if (StrReplaceW(wszRangeText, nRangeTextLen, wszFindItEsc, nFindItLenEsc, wszReplaceWithEsc, nReplaceWithLenEsc, dwFlags, NULL, &nResultTextLen, NULL, NULL, NULL))
            wszResultText=AllocWideStr(nResultTextLen + 1);
        }
        else
        {
          //We don't need allocate new buffer since output string is less than input
          wszResultText=wszRangeText;
        }
      }

      if (wszResultText)
      {
        if (dwFlags & FRF_REGEXP)
        {
          //Remember scroll
          nFirstLine=SaveLineScroll(lpFrame->ei.hWndEdit);
        }
        else
        {
          //Remember selection
          if (nGetTextNewLine == AELB_ASIS)
          {
            crInitialRE.cpMin=-IndexSubtract(lpFrame->ei.hWndEdit, NULL, &crCurSel.ciMin, AELB_ASIS, FALSE);
            crInitialRE.cpMax=crInitialRE.cpMin + IndexSubtract(lpFrame->ei.hWndEdit, &crCurSel.ciMax, &crCurSel.ciMin, AELB_ASIS, FALSE);
          }
          else SendMessage(lpFrame->ei.hWndEdit, EM_EXGETSEL64, 0, (LPARAM)&crInitialRE);

          if (dwFlags & FRF_SELECTION)
          {
            nMin=0;
            nMax=crInitialRE.cpMax - crInitialRE.cpMin;
          }
          else if ((dwFlags & FRF_BEGINNING) || (dwFlags & FRF_UP))
          {
            nMin=crInitialRE.cpMin;
            nMax=crInitialRE.cpMax;
          }
          else if (dwFlags & FRF_DOWN)
          {
            nMin=0;
            nMax=crInitialRE.cpMax - crInitialRE.cpMin;
          }
          if (nMin == nMax)
            nMax=-MAXINT_PTR;

          //Remember scroll
          SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTVISIBLELINE, (LPARAM)&ciFirstVisibleBefore);

          if (AEC_IndexCompare(&ciFirstVisibleBefore, &crRange.ciMin) >= 0)
          {
            if (nGetTextNewLine == AELB_ASIS)
              nFirstVisible=-IndexSubtract(lpFrame->ei.hWndEdit, NULL, &ciFirstVisibleBefore, AELB_ASIS, FALSE);
            else
              nFirstVisible=AkelIndexToRichOffset(lpFrame->ei.hWndEdit, &ciFirstVisibleBefore);
          }
          else nFirstVisible=-MAXINT_PTR;
        }

        //Replace operation
        if (dwFlags & FRF_REGEXP)
        {
          pr.wszResult=wszResultText;
          nResultTextLen=PatReplace(&pr);
          nChanges=pr.nReplaceCount;
        }
        else nChanges=StrReplaceW(wszRangeText, nRangeTextLen, wszFindItEsc, nFindItLenEsc, wszReplaceWithEsc, nReplaceWithLenEsc, dwFlags, wszResultText, &nResultTextLen, &nMin, (nMax == -MAXINT_PTR)?NULL:&nMax, (nFirstVisible == -MAXINT_PTR)?NULL:&nFirstVisible);

        if (nChanges)
        {
          if (wszResultText != wszRangeText)
          {
            //Data for ReplaceSelW now in wszResultText
            FreeText(wszFullText);
            wszFullText=NULL;
            wszRangeText=NULL;
          }

          //Stop redraw
          SendMessage(lpFrame->ei.hWndEdit, WM_SETREDRAW, FALSE, 0);

          if (!(dwFlags & FRF_SELECTION))
            SetSel(lpFrame->ei.hWndEdit, &crRange, AESELT_LOCKSCROLL, NULL);

          if (!(dwFlags & FRF_REGEXP))
          {
            nReplaceSelNewLine=nGetTextNewLine;
            if (nReplaceSelNewLine == AELB_R)
            {
              if (lpFrame->ei.nNewLine == NEWLINE_WIN)
                nReplaceSelNewLine=AELB_RN;
              else if (lpFrame->ei.nNewLine == NEWLINE_UNIX)
                nReplaceSelNewLine=AELB_N;
              else if (lpFrame->ei.nNewLine == NEWLINE_MAC)
                nReplaceSelNewLine=AELB_R;
            }
          }
          else nReplaceSelNewLine=AELB_ASIS;

          ReplaceSelW(lpFrame->ei.hWndEdit, wszResultText, nResultTextLen, nReplaceSelNewLine, (bColumnSel?AEREPT_COLUMNON:0)|AEREPT_LOCKSCROLL, &crInsert.ciMin, &crInsert.ciMax);

          //Restore selection
          if (dwFlags & FRF_SELECTION)
          {
            SetSel(lpFrame->ei.hWndEdit, &crInsert, (bInitialColumnSel?AESELT_COLUMNON:0)|AESELT_LOCKSCROLL, bCaretAtStart?&crInsert.ciMin:&crInsert.ciMax);
          }
          else
          {
            if (dwFlags & FRF_REGEXP)
            {
              SendMessage(lpFrame->ei.hWndEdit, AEM_INDEXUPDATE, 0, (LPARAM)&crInitialSel.ciMin);
              SendMessage(lpFrame->ei.hWndEdit, AEM_INDEXUPDATE, 0, (LPARAM)&crInitialSel.ciMax);
            }
            else
            {
              if (nMax == -MAXINT_PTR)
                nMax=nMin;

              if ((dwFlags & FRF_BEGINNING) || (dwFlags & FRF_UP))
              {
                if (nGetTextNewLine == AELB_ASIS)
                {
                  SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&crInitialSel.ciMin);
                  IndexOffset(lpFrame->ei.hWndEdit, &crInitialSel.ciMin, nMin, AELB_ASIS);
                }
                else RichOffsetToAkelIndex(lpFrame->ei.hWndEdit, nMin, &crInitialSel.ciMin);

                crInitialSel.ciMax=crInitialSel.ciMin;
                IndexOffset(lpFrame->ei.hWndEdit, &crInitialSel.ciMax, nMax - nMin, nGetTextNewLine);
              }
              else if (dwFlags & FRF_DOWN)
              {
                SendMessage(lpFrame->ei.hWndEdit, AEM_INDEXUPDATE, 0, (LPARAM)&crInitialSel.ciMin);
                crInitialSel.ciMax=crInitialSel.ciMin;
                IndexOffset(lpFrame->ei.hWndEdit, &crInitialSel.ciMax, nMax - nMin, nGetTextNewLine);
              }
            }
            SetSel(lpFrame->ei.hWndEdit, &crInitialSel, (bInitialColumnSel?AESELT_COLUMNON:0)|AESELT_LOCKSCROLL, bCaretAtStart?&crInitialSel.ciMin:&crInitialSel.ciMax);
          }

          //Start redraw
          SendMessage(lpFrame->ei.hWndEdit, WM_SETREDRAW, TRUE, 0);
          InvalidateRect(lpFrame->ei.hWndEdit, NULL, TRUE);

          //Restore scroll
          if (dwFlags & FRF_REGEXP)
          {
            RestoreLineScroll(lpFrame->ei.hWndEdit, nFirstLine);
          }
          else
          {
            if (nFirstVisible != -MAXINT_PTR)
            {
              if (nGetTextNewLine == AELB_ASIS)
              {
                SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&ciFirstVisibleBefore);
                IndexOffset(lpFrame->ei.hWndEdit, &ciFirstVisibleBefore, nFirstVisible, AELB_ASIS);
              }
              else RichOffsetToAkelIndex(lpFrame->ei.hWndEdit, nFirstVisible, &ciFirstVisibleBefore);

              if ((int)SendMessage(lpFrame->ei.hWndEdit, AEM_GETLINENUMBER, AEGL_FIRSTVISIBLELINE, 0) != ciFirstVisibleBefore.nLine)
              {
                POINT64 ptScrollPos;

                ptScrollPos.x=-1;
                ptScrollPos.y=SendMessage(lpFrame->ei.hWndEdit, AEM_VPOSFROMLINE, AECT_GLOBAL, ciFirstVisibleBefore.nLine);
                SendMessage(lpFrame->ei.hWndEdit, AEM_SETSCROLLPOS, 0, (LPARAM)&ptScrollPos);
              }
            }
          }
        }

        if (wszResultText != wszRangeText)
          FreeWideStr(wszResultText);
      }
      if (wszFullText) FreeText(wszFullText);
    }
  }
  else
  {
    if (dwFlags & FRF_REGEXP)
    {
      if (!(dwFlags & FRF_WHOLEWORD) ||
          (SendMessage(lpFrame->ei.hWndEdit, AEM_ISDELIMITER, AEDLM_WORD|AEDLM_PREVCHAR, (LPARAM)&crCurSel.ciMin) &&
           SendMessage(lpFrame->ei.hWndEdit, AEM_ISDELIMITER, AEDLM_WORD, (LPARAM)&crCurSel.ciMax)))
      {
        pr.wpStr=NULL;
        pr.wpMaxStr=NULL;
        pr.ciStr=crCurSel.ciMin;
        pr.ciMaxStr=crCurSel.ciMax;
        pr.dwOptions|=RESE_ISMATCH;
        pr.wszResult=NULL;
        nResultTextLen=PatReplace(&pr);

        if (pr.nReplaceCount && !AEC_IndexCompare(&pr.ciLeftStr, &pr.ciStr) &&
            (!AEC_IndexCompare(&pr.ciRightStr, &pr.ciMaxStr) || (dwFlags & FRF_SELECTION)))
        {
          if (pr.wszResult=AllocWideStr(nResultTextLen))
          {
            nResultTextLen=PatReplace(&pr);

            if (dwFlags & FRF_SELECTION)
            {
              SendMessage(lpFrame->ei.hWndEdit, WM_SETREDRAW, FALSE, 0);
              crRange.ciMin=pr.ciStr;
              crRange.ciMax=pr.ciRightStr;
              SetSel(lpFrame->ei.hWndEdit, &crRange, AESELT_LOCKSCROLL, NULL);
              ReplaceSelW(lpFrame->ei.hWndEdit, pr.wszResult, nResultTextLen, AELB_ASIS, 0, &crInsert.ciMin, &crInsert.ciMax);

              //Restore selection
              SendMessage(lpFrame->ei.hWndEdit, AEM_INDEXUPDATE, 0, (LPARAM)&crInitialSel.ciMax);
              crRange.ciMin=crInsert.ciMax;
              crRange.ciMax=crInitialSel.ciMax;
              SetSel(lpFrame->ei.hWndEdit, &crRange, 0, &crRange.ciMin);
              SendMessage(lpFrame->ei.hWndEdit, WM_SETREDRAW, TRUE, 0);
              InvalidateRect(lpFrame->ei.hWndEdit, NULL, TRUE);
            }
            else
            {
              ReplaceSelW(lpFrame->ei.hWndEdit, pr.wszResult, nResultTextLen, AELB_ASIS, 0, NULL, NULL);
            }
            nChanges=1;
            FreeWideStr(pr.wszResult);
          }
        }
      }
    }
    else
    {
      AEFINDTEXTW ft;

      ft.dwFlags=dwFlags;
      ft.pText=wszFindItEsc;
      ft.dwTextLen=nFindItLenEsc;
      ft.nNewLine=AELB_R;

      if (SendMessage(lpFrame->ei.hWndEdit, AEM_ISMATCHW, (WPARAM)&crCurSel.ciMin, (LPARAM)&ft))
      {
        if (dwFlags & FRF_SELECTION)
        {
          SendMessage(lpFrame->ei.hWndEdit, WM_SETREDRAW, FALSE, 0);
          SetSel(lpFrame->ei.hWndEdit, &ft.crFound, AESELT_LOCKSCROLL, NULL);
          ReplaceSelW(lpFrame->ei.hWndEdit, wszReplaceWithEsc, nReplaceWithLenEsc, AELB_ASINPUT, 0, &crInsert.ciMin, &crInsert.ciMax);

          //Restore selection
          SendMessage(lpFrame->ei.hWndEdit, AEM_INDEXUPDATE, 0, (LPARAM)&crInitialSel.ciMax);
          crRange.ciMin=crInsert.ciMax;
          crRange.ciMax=crInitialSel.ciMax;
          SetSel(lpFrame->ei.hWndEdit, &crRange, 0, &crRange.ciMin);
          SendMessage(lpFrame->ei.hWndEdit, WM_SETREDRAW, TRUE, 0);
          InvalidateRect(lpFrame->ei.hWndEdit, NULL, TRUE);
          nChanges=1;
        }
        else
        {
          if (!AEC_IndexCompare(&crCurSel.ciMax, &ft.crFound.ciMax))
          {
            ReplaceSelW(lpFrame->ei.hWndEdit, wszReplaceWithEsc, nReplaceWithLenEsc, AELB_ASINPUT, 0, NULL, NULL);
            nChanges=1;
          }
        }
      }
    }
    nResult=TextFindW(lpFrame, dwFlags|FRF_FINDFROMREPLACE, wpFindIt, nFindItLen);
  }

  End:
  if (wszFindItEsc != wpFindIt)
    FreeWideStr(wszFindItEsc);
  if (wszReplaceWithEsc != wpReplaceWith)
    FreeWideStr(wszReplaceWithEsc);
  if (nReplaceCount) *nReplaceCount=nChanges;
  return nResult;
}

INT_PTR StrReplaceW(const wchar_t *wpText, INT_PTR nTextLen, const wchar_t *wpIt, int nItLen, const wchar_t *wpWith, int nWithLen, DWORD dwFlags, wchar_t *wszResult, INT_PTR *nResultLen, INT_PTR *nMin, INT_PTR *nMax, INT_PTR *nFirstVis)
{
  const wchar_t *wpTextMax;
  const wchar_t *wpTextCount;
  const wchar_t *wpMatchCount;
  const wchar_t *wpItMax;
  const wchar_t *wpItCount;
  const wchar_t *wpWithMax;
  const wchar_t *wpWithCount;
  const wchar_t *wpMin=NULL;
  const wchar_t *wpMax=NULL;
  const wchar_t *wpFirstVis=NULL;
  wchar_t *wpResultCount;
  INT_PTR nChanges=0;
  int nDiff;

  if (nTextLen == -1)
    nTextLen=xstrlenW(wpText) + 1;
  if (nItLen == -1)
    nItLen=(int)xstrlenW(wpIt);
  if (nWithLen == -1)
    nWithLen=(int)xstrlenW(wpWith);
  wpTextMax=wpText + nTextLen;
  wpItMax=wpIt + nItLen;
  wpWithMax=wpWith + nWithLen;
  wpResultCount=wszResult;

  nDiff=nItLen - nWithLen;
  if (nMin) wpMin=wpText + *nMin;
  if (nMax) wpMax=wpText + *nMax;
  if (nFirstVis) wpFirstVis=wpText + *nFirstVis;

  for (wpTextCount=wpText; wpTextCount < wpTextMax; ++wpTextCount)
  {
    if (dwFlags & FRF_WHOLEWORD)
    {
      if (wpTextCount == wpText)
      {
        if (dwFlags & FRF_WHOLEWORDGOODSTART)
          goto Find;
      }
      else if (AKD_wcschr(lpFrameCurrent->wszWordDelimiters, *(wpTextCount - 1)))
        goto Find;
      goto Next;
    }

    Find:
    wpMatchCount=wpTextCount;
    wpItCount=wpIt;

    while (*wpMatchCount == *wpItCount ||
           (!(dwFlags & FRF_MATCHCASE) && WideCharLower(*wpMatchCount) == WideCharLower(*wpItCount)))
    {
      if (++wpItCount >= wpItMax)
      {
        if (dwFlags & FRF_WHOLEWORD)
        {
          if (wpMatchCount + 1 >= wpTextMax)
          {
            if (dwFlags & FRF_WHOLEWORDGOODEND)
              goto Replace;
          }
          else if (AKD_wcschr(lpFrameCurrent->wszWordDelimiters, *(wpMatchCount + 1)))
            goto Replace;
          goto Next;
        }

        Replace:
        if (wszResult)
        {
          if (nMin)
          {
            if (wpMin > wpMatchCount) *nMin-=nDiff;
            else if (wpMin > wpTextCount) *nMin-=(wpMin - wpTextCount);
          }
          if (nMax)
          {
            if (wpMax > wpMatchCount) *nMax-=nDiff;
            else if (wpMax > wpTextCount) *nMax=*nMax - (wpMax - wpTextCount) + nWithLen;
          }
          if (nFirstVis)
          {
            if (wpFirstVis > wpMatchCount) *nFirstVis-=nDiff;
            else if (wpFirstVis > wpTextCount) *nFirstVis-=(wpFirstVis - wpTextCount);
          }

          for (wpWithCount=wpWith; wpWithCount < wpWithMax; ++wpWithCount)
            *wpResultCount++=*wpWithCount;
        }
        else wpResultCount+=nWithLen;

        wpTextCount=wpMatchCount + 1;
        wpItCount=wpIt;
        ++nChanges;
        if (wpTextCount >= wpTextMax) goto End;
      }
      if (++wpMatchCount >= wpTextMax) break;
    }

    Next:
    if (wszResult) *wpResultCount=*wpTextCount;
    ++wpResultCount;
  }

  End:
  if (nResultLen) *nResultLen=wpResultCount - wszResult;
  return nChanges;
}

INT_PTR EscapeStringToEscapeDataW(const wchar_t *wpInput, INT_PTR nInputLen, wchar_t *wszOutput, int nNewLine)
{
  const wchar_t *a=wpInput;
  const wchar_t *wpInputMax;
  wchar_t *b=wszOutput;
  wchar_t whex[5];
  int nDec;

  if (nInputLen == -1)
    nInputLen=xstrlenW(wpInput);
  wpInputMax=wpInput + nInputLen;

  for (whex[4]=L'\0'; a < wpInputMax; ++a, ++b)
  {
    if (*a == L'\\')
    {
      if (*++a == L'\\') *b=L'\\';
      else if (*a == L'n')
      {
        if (nNewLine == NEWLINE_MAC) *b=L'\r';
        else if (nNewLine == NEWLINE_UNIX) *b=L'\n';
        else if (nNewLine == NEWLINE_WIN)
        {
          *b=L'\r';
          *++b=L'\n';
        }
      }
      else if (*a == L't') *b=L'\t';
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
          *b=(wchar_t)nDec;
          while (*++a == L' ');
        }
        while (*a && *a != L']' && ++b);

        if (!*a) goto Error;
      }
      else goto Error;
    }
    else *b=*a;
  }
  *b=L'\0';
  return b - wszOutput;

  Error:
  *wszOutput=L'\0';
  return (wpInput - a) - 1;
}

void EscapeDataToEscapeStringW(const wchar_t *wpInput, wchar_t *wszOutput)
{
  const wchar_t *a=wpInput;
  wchar_t *b=wszOutput;

  for (; *a; ++b, ++a)
  {
    if (*a == L'\t') *b=L'\\', *++b=L't';
    else if (*a == L'\r') *b=L'\\', *++b=L'n';
    else if (*a == L'\n') *b=L'\\', *++b=L'n';
    else if (*a == L'\\') *b=L'\\', *++b=L'\\';
    else *b=*a;
  }
  *b=L'\0';
}

BOOL SetDefButtonStyle(HWND hWnd, HWND hWndNewDef)
{
  DWORD dwStyle=(DWORD)GetWindowLongPtrWide(hWnd, GWL_STYLE);

  if (hWnd == hWndNewDef)
  {
    if (dwStyle & BS_DEFPUSHBUTTON)
      return FALSE;
    SendMessage(hWnd, BM_SETSTYLE, (dwStyle & ~BS_PUSHBUTTON)|BS_DEFPUSHBUTTON, TRUE);
  }
  else
  {
    if (dwStyle & BS_PUSHBUTTON)
      return FALSE;
    SendMessage(hWnd, BM_SETSTYLE, (dwStyle & ~BS_DEFPUSHBUTTON)|BS_PUSHBUTTON, TRUE);
  }
  return TRUE;
}


//// Paste operation

BOOL GetSel(HWND hWnd, AECHARRANGE *crSel, BOOL *bColumnSel, AECHARINDEX *ciCaret)
{
  AESELECTION aes;
  BOOL bSelExist;

  bSelExist=(BOOL)SendMessage(hWnd, AEM_GETSEL, (WPARAM)ciCaret, (LPARAM)&aes);
  if (crSel)
  {
    crSel->ciMin=aes.crSel.ciMin;
    crSel->ciMax=aes.crSel.ciMax;
  }
  if (bColumnSel) *bColumnSel=(aes.dwFlags & AESELT_COLUMNON);
  return bSelExist;
}

void SetSel(HWND hWnd, AECHARRANGE *crSel, DWORD dwFlags, AECHARINDEX *ciCaret)
{
  AESELECTION aes;

  aes.crSel.ciMin=crSel->ciMin;
  aes.crSel.ciMax=crSel->ciMax;
  aes.dwFlags=dwFlags;
  aes.dwType=0;
  SendMessage(hWnd, AEM_SETSEL, (WPARAM)ciCaret, (LPARAM)&aes);
}

void SetSelRE(HWND hWnd, INT_PTR nSelStart, INT_PTR nSelEnd)
{
  int nLockScroll;

  if ((nLockScroll=(int)SendMessage(hWnd, AEM_LOCKSCROLL, (WPARAM)-1, 0)) == -1)
    SendMessage(hWnd, AEM_LOCKSCROLL, SB_BOTH, TRUE);
  SendMessage(hWnd, EM_SETSEL, nSelStart, nSelEnd);
  if (nLockScroll == -1)
    SendMessage(hWnd, AEM_LOCKSCROLL, SB_BOTH, FALSE);
  ScrollCaret(hWnd);
}

void ReplaceSelA(HWND hWnd, const char *pData, INT_PTR nDataLen, DWORD dwFlags, int nNewLine, AECHARINDEX *ciInsertStart, AECHARINDEX *ciInsertEnd)
{
  AEREPLACESELA rs;

  rs.pText=pData;
  rs.dwTextLen=(UINT_PTR)nDataLen;
  rs.nNewLine=nNewLine;
  rs.dwFlags=dwFlags;
  rs.ciInsertStart=ciInsertStart;
  rs.ciInsertEnd=ciInsertEnd;
  rs.nCodePage=CP_ACP;
  SendMessage(hWnd, AEM_REPLACESELA, 0, (LPARAM)&rs);
}

void ReplaceSelW(HWND hWnd, const wchar_t *wpData, INT_PTR nDataLen, int nNewLine, DWORD dwFlags, AECHARINDEX *ciInsertStart, AECHARINDEX *ciInsertEnd)
{
  AEREPLACESELW rs;

  rs.pText=wpData;
  rs.dwTextLen=(UINT_PTR)nDataLen;
  rs.nNewLine=nNewLine;
  rs.dwFlags=dwFlags;
  rs.ciInsertStart=ciInsertStart;
  rs.ciInsertEnd=ciInsertEnd;
  SendMessage(hWnd, AEM_REPLACESELW, 0, (LPARAM)&rs);
}

INT_PTR IndexSubtract(HWND hWnd, AECHARINDEX *ciChar1, AECHARINDEX *ciChar2, int nNewLine, BOOL bColumnSel)
{
  AEINDEXSUBTRACT aeis;

  aeis.ciChar1=ciChar1;
  aeis.ciChar2=ciChar2;
  aeis.nNewLine=nNewLine;
  aeis.bColumnSel=bColumnSel;
  return SendMessage(hWnd, AEM_INDEXSUBTRACT, 0, (LPARAM)&aeis);
}

INT_PTR IndexOffset(HWND hWnd, AECHARINDEX *ciChar, INT_PTR nOffset, int nNewLine)
{
  AEINDEXOFFSET aeio;

  aeio.ciCharIn=ciChar;
  aeio.ciCharOut=ciChar;
  aeio.nOffset=nOffset;
  aeio.nNewLine=nNewLine;
  return SendMessage(hWnd, AEM_INDEXOFFSET, 0, (LPARAM)&aeio);
}

INT_PTR AkelIndexToRichOffset(HWND hWnd, AECHARINDEX *ciChar)
{
  return SendMessage(hWnd, AEM_INDEXTORICHOFFSET, 0, (LPARAM)ciChar);
}

void RichOffsetToAkelIndex(HWND hWnd, INT_PTR nOffset, AECHARINDEX *ciChar)
{
  SendMessage(hWnd, AEM_RICHOFFSETTOINDEX, nOffset, (LPARAM)ciChar);
}

int GetLineNumber(FRAMEDATA *lpFrame, int nWrappedLine)
{
  if (!(moCur.dwStatusPosType & SPT_LINEWRAP) && lpFrame->ei.bWordWrap)
    return (int)SendMessage(lpFrame->ei.hWndEdit, AEM_GETUNWRAPLINE, (WPARAM)nWrappedLine, 0);
  else
    return nWrappedLine;
}

int GetIndexColumn(FRAMEDATA *lpFrame, const AECHARINDEX *ciChar)
{
  if (moCur.dwStatusPosType & SPT_COLUMN)
    return (int)SendMessage(lpFrame->ei.hWndEdit, AEM_INDEXTOCOLUMN, MAKELONG(lpFrame->nTabStopSize, !(moCur.dwStatusPosType & SPT_LINEWRAP)), (LPARAM)ciChar);
  else
    return (int)SendMessage(lpFrame->ei.hWndEdit, AEM_INDEXTOCOLUMN, MAKELONG(1, !(moCur.dwStatusPosType & SPT_LINEWRAP)), (LPARAM)ciChar);
}

INT_PTR GetTextLength(HWND hWnd)
{
  return SendMessage(hWnd, AEM_GETRICHOFFSET, AEGI_LASTCHAR, 0);
}

INT_PTR GetRangeTextA(HWND hWnd, INT_PTR nMin, INT_PTR nMax, char **pText)
{
  TEXTRANGE64A txtrngA;
  INT_PTR nLen;

  if (nMax == -1)
  {
    nMax=GetTextLength(hWnd);
  }
  if (nMin < nMax)
  {
    nLen=(nMax - nMin);

    if (*pText=(char *)API_HeapAlloc(hHeap, 0, nLen + 1))
    {
      txtrngA.chrg.cpMin=nMin;
      txtrngA.chrg.cpMax=nMax;
      txtrngA.lpstrText=*pText;
      if (!(nLen=SendMessage(hWnd, EM_GETTEXTRANGE64A, 0, (LPARAM)&txtrngA)))
      {
        API_HeapFree(hHeap, 0, (LPVOID)*pText);
        *pText=NULL;
      }
      return nLen;
    }
  }
  *pText=NULL;
  return 0;
}

INT_PTR GetRangeTextW(HWND hWnd, INT_PTR nMin, INT_PTR nMax, wchar_t **wpText)
{
  TEXTRANGE64W txtrngW;
  INT_PTR nLen;

  if (nMax == -1)
  {
    nMax=GetTextLength(hWnd);
  }
  if (nMin < nMax)
  {
    nLen=(nMax - nMin);

    if (*wpText=(wchar_t *)API_HeapAlloc(hHeap, 0, (nLen + 1) * sizeof(wchar_t)))
    {
      txtrngW.chrg.cpMin=nMin;
      txtrngW.chrg.cpMax=nMax;
      txtrngW.lpstrText=*wpText;
      if (!(nLen=SendMessage(hWnd, EM_GETTEXTRANGE64W, 0, (LPARAM)&txtrngW)))
      {
        API_HeapFree(hHeap, 0, (LPVOID)*wpText);
        *wpText=NULL;
      }
      return nLen;
    }
  }
  *wpText=NULL;
  return 0;
}

INT_PTR ExGetRangeTextA(HWND hWnd, int nCodePage, const char *lpDefaultChar, BOOL *lpUsedDefChar, AECHARINDEX *ciMin, AECHARINDEX *ciMax, BOOL bColumnSel, char **pText, int nNewLine, BOOL bFillSpaces)
{
  AETEXTRANGEA tr;
  INT_PTR nLen=0;

  if (AEC_IndexCompare(ciMin, ciMax))
  {
    tr.cr.ciMin=*ciMin;
    tr.cr.ciMax=*ciMax;
    tr.bColumnSel=bColumnSel;
    tr.pBuffer=NULL;
    tr.dwBufferMax=(UINT_PTR)-1;
    tr.nNewLine=nNewLine;
    tr.nCodePage=nCodePage;
    tr.lpDefaultChar=lpDefaultChar;
    tr.lpUsedDefChar=lpUsedDefChar;
    tr.bFillSpaces=bFillSpaces;

    if (nLen=SendMessage(hWnd, AEM_GETTEXTRANGEA, 0, (LPARAM)&tr))
    {
      if (pText)
      {
        if (tr.pBuffer=(char *)API_HeapAlloc(hHeap, 0, nLen))
        {
          if (!(nLen=SendMessage(hWnd, AEM_GETTEXTRANGEA, 0, (LPARAM)&tr)))
          {
            API_HeapFree(hHeap, 0, (LPVOID)tr.pBuffer);
            tr.pBuffer=NULL;
          }
        }
      }
    }
  }
  else tr.pBuffer=NULL;

  if (pText) *pText=tr.pBuffer;
  return nLen;
}

INT_PTR ExGetRangeTextW(HWND hWnd, AECHARINDEX *ciMin, AECHARINDEX *ciMax, BOOL bColumnSel, wchar_t **wpText, int nNewLine, BOOL bFillSpaces)
{
  AETEXTRANGEW tr;
  INT_PTR nLen=0;

  if (AEC_IndexCompare(ciMin, ciMax))
  {
    tr.cr.ciMin=*ciMin;
    tr.cr.ciMax=*ciMax;
    tr.bColumnSel=bColumnSel;
    tr.pBuffer=NULL;
    tr.dwBufferMax=(UINT_PTR)-1;
    tr.nNewLine=nNewLine;
    tr.bFillSpaces=bFillSpaces;

    if (nLen=SendMessage(hWnd, AEM_GETTEXTRANGEW, 0, (LPARAM)&tr))
    {
      if (wpText)
      {
        if (tr.pBuffer=(wchar_t *)API_HeapAlloc(hHeap, 0, nLen * sizeof(wchar_t)))
        {
          if (!(nLen=SendMessage(hWnd, AEM_GETTEXTRANGEW, 0, (LPARAM)&tr)))
          {
            API_HeapFree(hHeap, 0, (LPVOID)tr.pBuffer);
            tr.pBuffer=NULL;
          }
        }
      }
    }
  }
  else tr.pBuffer=NULL;

  if (wpText) *wpText=tr.pBuffer;
  return nLen;
}

BOOL FreeText(LPVOID pText)
{
  if (pText)
    return API_HeapFree(hHeap, 0, (LPVOID)pText);
  return FALSE;
}

BOOL PasteInEditAsRichEdit(HWND hWnd, int nMaxLenght)
{
  HGLOBAL hData;
  LPVOID pData;
  BOOL bResult=FALSE;

  if (OpenClipboard(hWnd))
  {
    if (!bOldWindows && (hData=GetClipboardData(CF_UNICODETEXT)))
    {
      if (pData=GlobalLock(hData))
      {
        //Convert \r\r\n->\r, \r\n->\r, \n->\r
        wchar_t *wpSource=(wchar_t *)pData;
        wchar_t *wpSourceMax;
        wchar_t *wpTarget;
        wchar_t *wpSourceCount;
        wchar_t *wpTargetCount;
        INT_PTR nTargetLen;

        nTargetLen=xstrlenW(wpSource);
        if (nMaxLenght > 0)
        {
          if (nTargetLen > nMaxLenght)
            nTargetLen=nMaxLenght;
          wpSourceMax=wpSource + nTargetLen;
        }
        else wpSourceMax=(wchar_t *)MAXUINT_PTR;

        if (wpTarget=AllocWideStr(nTargetLen + 1))
        {
          for (wpTargetCount=wpTarget, wpSourceCount=wpSource; wpSourceCount < wpSourceMax; ++wpSourceCount, ++wpTargetCount)
          {
            if (*wpSourceCount == L'\r')
            {
              *wpTargetCount=L'\r';

              if (*(wpSourceCount + 1) == L'\r' &&
                  *(wpSourceCount + 2) == L'\n')
              {
                wpSourceCount+=2;
              }
              else if (*(wpSourceCount + 1) == L'\n')
              {
                wpSourceCount+=1;
              }
            }
            else if (*wpSourceCount == L'\n')
            {
              *wpTargetCount=L'\r';
            }
            else *wpTargetCount=*wpSourceCount;
          }
          *wpTargetCount=L'\0';

          SendMessageW(hWnd, EM_REPLACESEL, TRUE, (LPARAM)wpTarget);
          bResult=TRUE;
          FreeWideStr(wpTarget);
        }
        GlobalUnlock(hData);
      }
    }
    else if (hData=GetClipboardData(CF_TEXT))
    {
      if (pData=GlobalLock(hData))
      {
        //Convert \r\r\n->\r, \r\n->\r, \n->\r
        char *pSource=(char *)pData;
        char *pSourceMax;
        char *pTarget;
        char *pSourceCount;
        char *pTargetCount;
        int nTargetLen;

        nTargetLen=lstrlenA(pSource);
        if (nMaxLenght > 0)
        {
          if (nTargetLen > nMaxLenght)
            nTargetLen=nMaxLenght;
          pSourceMax=pSource + nTargetLen;
        }
        else pSourceMax=(char *)MAXUINT_PTR;

        if (pTarget=(char *)API_HeapAlloc(hHeap, 0, nTargetLen + 1))
        {
          for (pTargetCount=pTarget, pSourceCount=pSource; pSourceCount < pSourceMax; ++pSourceCount, ++pTargetCount)
          {
            if (*pSourceCount == L'\r')
            {
              *pTargetCount=L'\r';

              if (*(pSourceCount + 1) == L'\r' &&
                  *(pSourceCount + 2) == L'\n')
              {
                pSourceCount+=2;
              }
              else if (*(pSourceCount + 1) == L'\n')
              {
                pSourceCount+=1;
              }
            }
            else if (*pSourceCount == L'\n')
            {
              *pTargetCount=L'\r';
            }
            else *pTargetCount=*pSourceCount;
          }
          *pTargetCount=L'\0';

          SendMessageA(hWnd, EM_REPLACESEL, TRUE, (LPARAM)pTarget);
          bResult=TRUE;
          API_HeapFree(hHeap, 0, (LPVOID)pTarget);
        }
        GlobalUnlock(hData);
      }
    }
    CloseClipboard();
  }
  return bResult;
}

int PasteCase(HWND hWnd, BOOL bAnsi)
{
  HGLOBAL hData;
  LPVOID pData;
  wchar_t *wszData=NULL;
  INT_PTR nDataLen=0;
  int nCase=0;

  if (OpenClipboard(NULL))
  {
    if (!bAnsi && (hData=GetClipboardData(CF_UNICODETEXT)))
    {
      if (pData=GlobalLock(hData))
      {
        nDataLen=xstrlenW((wchar_t *)pData) + 1;
        if (wszData=AllocWideStr(nDataLen))
          xmemcpy(wszData, (wchar_t *)pData, nDataLen * sizeof(wchar_t));
        GlobalUnlock(hData);
      }
    }
    else if (hData=GetClipboardData(CF_TEXT))
    {
      if (pData=GlobalLock(hData))
      {
        nDataLen=MultiByteToWideChar(CP_ACP, 0, (char *)pData, -1, NULL, 0);
        if (wszData=AllocWideStr(nDataLen))
          MultiByteToWideChar(CP_ACP, 0, (char *)pData, -1, wszData, (int)nDataLen);
        GlobalUnlock(hData);
      }
    }
    CloseClipboard();
  }

  //Paste
  if (wszData)
  {
    if (--nDataLen > 0)
    {
      if (nCase=DetectSelCase(hWnd))
        ConvertCase(wszData, nDataLen, nCase);
    }
    ReplaceSelW(hWnd, wszData, nDataLen, AELB_ASINPUT, AEREPT_COLUMNASIS, NULL, NULL);
    FreeWideStr(wszData);
  }
  return nCase;
}

void ShowStandardViewMenu(HWND hWnd, HMENU hMenu, BOOL bMouse)
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
  EnableMenuItem(hMenu, IDM_EDIT_COPY, (cr.cpMin < cr.cpMax)?MF_ENABLED:MF_GRAYED);

  nCmd=TrackPopupMenu(hMenu, TPM_RETURNCMD|TPM_NONOTIFY|TPM_LEFTBUTTON|TPM_RIGHTBUTTON, pt.x, pt.y, 0, hWnd, NULL);

  if (nCmd == IDM_EDIT_COPY)
    SendMessage(hWnd, WM_COPY, 0, 0);
  else if (nCmd == IDM_EDIT_SELECTALL)
    SendMessage(hWnd, EM_SETSEL, 0, -1);
}


//// Go to line

BOOL CALLBACK GoToDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndNumber;
  static HWND hWndLine;
  static HWND hWndOffset;
  int nLine;
  int nColumn;
  INT_PTR nOffset;
  int nNumberLen;

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndNumber=GetDlgItem(hDlg, IDC_GOTO_NUMBER);
    hWndLine=GetDlgItem(hDlg, IDC_GOTO_LINE);
    hWndOffset=GetDlgItem(hDlg, IDC_GOTO_OFFSET);

    if (rcGotoDlg.right && rcGotoDlg.bottom)
      SetWindowPos(hDlg, 0, rcGotoDlg.left, rcGotoDlg.top, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);

    if (dwGotoType & GT_LINE)
    {
      SendMessage(hWndLine, BM_SETCHECK, BST_CHECKED, 0);
      PostMessage(hDlg, WM_COMMAND, IDC_GOTO_LINE, 0);
    }
    else if (dwGotoType & GT_OFFSETBYTE)
    {
      SendMessage(hWndOffset, BM_SETCHECK, BST_CHECKED, 0);
      PostMessage(hDlg, WM_COMMAND, IDC_GOTO_OFFSET, 0);
    }
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_GOTO_LINE ||
        LOWORD(wParam) == IDC_GOTO_OFFSET)
    {
      if (LOWORD(wParam) == IDC_GOTO_LINE)
        dwGotoType=GT_LINE;
      else if (LOWORD(wParam) == IDC_GOTO_OFFSET)
        dwGotoType=GT_OFFSETBYTE;

      if (dwGotoType & GT_LINE)
      {
        if (!SendMessage(hWndNumber, EM_GETMODIFY, 0, 0))
        {
          //Caret line
          nLine=GetLineNumber(lpFrameCurrent, ciCurCaret.nLine);
          nNumberLen=(int)xprintfW(wbuf, L"%d", nLine + 1);

          //Caret column
          nColumn=GetIndexColumn(lpFrameCurrent, &ciCurCaret);

          xprintfW(wbuf, L"%d:%d", nLine + 1, nColumn + 1);
          SetWindowTextWide(hWndNumber, wbuf);
          SendMessage(hWndNumber, EM_SETSEL, 0, nNumberLen);
        }
      }
      else if (dwGotoType & GT_OFFSETBYTE)
      {
        if (!SendMessage(hWndNumber, EM_GETMODIFY, 0, 0))
        {
          //Caret offset
          nOffset=-IndexSubtract(lpFrameCurrent->ei.hWndEdit, NULL, &ciCurCaret, AELB_ASIS, FALSE);
          SetDlgItemInt(hDlg, IDC_GOTO_NUMBER, (UINT)nOffset, FALSE);
          SendMessage(hWndNumber, EM_SETSEL, 0, -1);
        }
      }
    }
    else if (LOWORD(wParam) == IDOK)
    {
      if (GetWindowTextWide(hWndNumber, wbuf, BUFFER_SIZE))
      {
        if (!GoTo(dwGotoType, wbuf))
        {
          API_LoadStringW(hLangLib, MSG_WRONG_STRING, wszMsg, BUFFER_SIZE);
          API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
        }
      }
    }
    if (LOWORD(wParam) == IDOK ||
        LOWORD(wParam) == IDCANCEL)
    {
      GetWindowPos(hDlg, NULL, &rcGotoDlg);
      DestroyWindow(hDlg);
      hDlgModeless=NULL;
      nModelessType=MLT_NONE;
      return TRUE;
    }
  }
  return FALSE;
}

BOOL GoTo(DWORD dwGotoType, const wchar_t *wpString)
{
  if (wpString && ((dwGotoType & GT_LINE) || (dwGotoType & GT_OFFSETBYTE) || (dwGotoType & GT_OFFSETCHAR)))
  {
    AECHARRANGE cr;
    wchar_t wszFirst[MAX_PATH];
    int nLineCount=0;
    int nFirst=1;
    int nSecond=1;
    int i;

    //Only numeral
    for (i=0; *wpString; ++wpString)
    {
      if (*wpString >= L'0' && *wpString <= L'9')
        wszFirst[i++]=*wpString;
      else if (*wpString == L'-' && i == 0)
        wszFirst[i++]=*wpString;
      else if (*wpString == L':' && i > 0)
      {
        nSecond=(int)xatoiW(wpString + 1, NULL);
        nSecond=max(1, nSecond);
        break;
      }
      else if (*wpString != L' ' && *wpString != L',' && *wpString != L'.' && i > 0)
        break;
    }
    wszFirst[i]=L'\0';

    nFirst=(int)xatoiW(wszFirst, NULL);

    if (dwGotoType & GT_LINE)
    {
      //Line
      if (!(moCur.dwStatusPosType & SPT_LINEWRAP) && lpFrameCurrent->ei.bWordWrap)
        nLineCount=(int)SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETLINENUMBER, AEGL_LINEUNWRAPCOUNT, 0);
      else
        nLineCount=(int)SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETLINENUMBER, AEGL_LINECOUNT, 0);

      if (!nFirst)
      {
        return FALSE;
      }
      else if (nFirst < 0)
      {
        nFirst=nLineCount + nFirst + 1;
        if (nFirst <= 0) nFirst=1;
      }
      nFirst=min(nFirst, nLineCount);

      if (!(moCur.dwStatusPosType & SPT_LINEWRAP) && lpFrameCurrent->ei.bWordWrap)
        SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETWRAPLINE, nFirst - 1, (LPARAM)&cr.ciMin);
      else
        SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETLINEINDEX, nFirst - 1, (LPARAM)&cr.ciMin);

      //Column
      cr.ciMin.nCharInLine=nSecond - 1;
      if (moCur.dwStatusPosType & SPT_COLUMN)
        SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_COLUMNTOINDEX, MAKELONG(lpFrameCurrent->nTabStopSize, !(moCur.dwStatusPosType & SPT_LINEWRAP)), (LPARAM)&cr.ciMin);
      else
        SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_COLUMNTOINDEX, MAKELONG(1, !(moCur.dwStatusPosType & SPT_LINEWRAP)), (LPARAM)&cr.ciMin);
    }
    else if (dwGotoType & GT_OFFSETBYTE)
    {
      if (nFirst >= 0)
      {
        SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&cr.ciMin);
        IndexOffset(lpFrameCurrent->ei.hWndEdit, &cr.ciMin, nFirst, AELB_ASIS);
      }
      else
      {
        SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&cr.ciMin);
        IndexOffset(lpFrameCurrent->ei.hWndEdit, &cr.ciMin, nFirst + 1, AELB_ASIS);
      }
    }
    else if (dwGotoType & GT_OFFSETCHAR)
    {
      SetSelRE(lpFrameCurrent->ei.hWndEdit, nFirst, nFirst);
      return TRUE;
    }

    //Set selection
    cr.ciMax=cr.ciMin;
    SetSel(lpFrameCurrent->ei.hWndEdit, &cr, AESELT_LOCKSCROLL, NULL);
    ScrollCaret(lpFrameCurrent->ei.hWndEdit);
    return TRUE;
  }
  return FALSE;
}


//// Recent files

RECENTFILE* RecentFilesInsert(RECENTFILESTACK *hStack, int nIndex)
{
  RECENTFILE *lpRecentFile=NULL;

  if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpRecentFile, nIndex, sizeof(RECENTFILE)))
    ++hStack->nElements;

  return lpRecentFile;
}

void RecentFilesDelete(RECENTFILESTACK *hStack, RECENTFILE *lpRecentFile)
{
  StackRecentFileParamFree(lpRecentFile);
  StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)lpRecentFile);
  --hStack->nElements;
}

void RecentFilesZero(RECENTFILESTACK *hStack)
{
  RECENTFILE *lpRecentFile;

  for (lpRecentFile=hStack->first; lpRecentFile; lpRecentFile=lpRecentFile->next)
  {
    StackRecentFileParamFree(lpRecentFile);
  }
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
  hStack->nElements=0;
  hStack->dwSaveTime=0;
}

RECENTFILE* RecentFilesFindByName(const wchar_t *wpFile, int *lpIndex)
{
  RECENTFILE *lpRecentFile;
  int nFileLen=(int)xstrlenW(wpFile);
  int nIndex=0;

  if (*wpFile)
  {
    for (lpRecentFile=hRecentFilesStack.first; lpRecentFile; lpRecentFile=lpRecentFile->next)
    {
      if (lpRecentFile->nFileLen == nFileLen && !xstrcmpiW(lpRecentFile->wszFile, wpFile))
      {
        if (lpIndex) *lpIndex=nIndex;
        return lpRecentFile;
      }
      ++nIndex;
    }
  }
  if (lpIndex) *lpIndex=-1;
  return NULL;
}

RECENTFILE* RecentFilesFindByIndex(int nIndex)
{
  RECENTFILE *lpRecentFile;
  int i=0;

  for (lpRecentFile=hRecentFilesStack.first; lpRecentFile; lpRecentFile=lpRecentFile->next)
  {
    if (i++ == nIndex)
      return lpRecentFile;
  }
  return NULL;
}

RECENTFILE* RecentFilesUpdate(const wchar_t *wpFile)
{
  RECENTFILE *lpRecentFile;

  if (lpRecentFile=RecentFilesFindByName(wpFile, NULL))
  {
    StackMoveIndex((stack **)&hRecentFilesStack.first, (stack **)&hRecentFilesStack.last, (stack *)lpRecentFile, 1);

    //Letters case could be different
    xstrcpynW(lpRecentFile->wszFile, wpFile, MAX_PATH);
  }
  else
  {
    if (lpRecentFile=RecentFilesInsert(&hRecentFilesStack, 1))
    {
      lpRecentFile->nFileLen=(int)xstrcpynW(lpRecentFile->wszFile, wpFile, MAX_PATH);
      if (hRecentFilesStack.nElements > moCur.nRecentFiles)
        RecentFilesDelete(&hRecentFilesStack, hRecentFilesStack.last);
    }
  }
  return lpRecentFile;
}

void RecentFilesRefresh(RECENTFILESTACK *hStack)
{
  RecentFilesZero(hStack);
  RecentFilesRead(hStack);
  bMenuRecentFiles=TRUE;

  //Clean save
  {
    wchar_t wszRegKey[MAX_PATH];

    xprintfW(wszRegKey, L"%s\\Recent", APP_REGHOMEW);
    RegClearKeyWide(HKEY_CURRENT_USER, wszRegKey);
    RecentFilesSave(hStack);
  }
}

int RecentFilesDeleteOld(RECENTFILESTACK *hStack)
{
  RECENTFILE *lpRecentFile;
  RECENTFILE *lpNextRecentFile;
  int nDead=0;

  //Delete files from recent files array if they doesn't exist
  for (lpRecentFile=hStack->first; lpRecentFile; lpRecentFile=lpNextRecentFile)
  {
    lpNextRecentFile=lpRecentFile->next;

    if (!FileExistsWide(lpRecentFile->wszFile))
    {
      RecentFilesDelete(hStack, lpRecentFile);
      ++nDead;
    }
  }
  return nDead;
}

int RecentFilesRead(RECENTFILESTACK *hStack)
{
  wchar_t wszRegKey[MAX_PATH];
  wchar_t wszRegValue[32];
  wchar_t *wszRegData;
  const wchar_t *wpCount;
  HKEY hKey;
  DWORD dwType;
  DWORD dwSize;
  DWORD dwDataMax;
  DWORD dwSaveTime;
  int i=0;

  //Params
  RECENTFILE *lpRecentFile;
  RECENTFILEPARAM *lpRecentFileParam;
  const wchar_t *wpParamName;
  const wchar_t *wpParamValue;
  int nParamNameLen;
  int nParamValueLen;

  if (!moCur.nRecentFiles) return 0;

  //Read recent files array
  xprintfW(wszRegKey, L"%s\\Recent", APP_REGHOMEW);
  if (RegOpenKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, KEY_READ, &hKey) != ERROR_SUCCESS)
    return 0;

  //Tick count
  dwSize=sizeof(DWORD);
  RegQueryValueExWide(hKey, L"SaveTime", NULL, &dwType, (LPBYTE)&dwSaveTime, &dwSize);

  if (hStack->dwSaveTime != dwSaveTime)
  {
    RecentFilesZero(hStack);

    dwDataMax=BUFFER_SIZE;
    if (wszRegData=AllocWideStr(dwDataMax))
    {
      for (i=0; i < moCur.nRecentFiles; ++i)
      {
        xprintfW(wszRegValue, L"file%d", i);
        if (RegQueryValueExWide(hKey, wszRegValue, NULL, &dwType, NULL, &dwSize) != ERROR_SUCCESS)
          break;

        if (dwSize / sizeof(wchar_t) + 1 > dwDataMax)
        {
          FreeWideStr(wszRegData);
          dwDataMax=dwSize / sizeof(wchar_t) + 1;
          if (!(wszRegData=AllocWideStr(dwDataMax)))
            break;
        }
        RegQueryValueExWide(hKey, wszRegValue, NULL, &dwType, (LPBYTE)wszRegData, &dwSize);
        if (!*wszRegData || dwType != REG_MULTI_SZ)
          break;
        wszRegData[dwSize / sizeof(wchar_t)]=L'\0';

        if (lpRecentFile=RecentFilesInsert(hStack, -1))
        {
          //File
          wpCount=wszRegData;
          lpRecentFile->nFileLen=(int)xstrcpynW(lpRecentFile->wszFile, wpCount, MAX_PATH);
          wpCount+=lpRecentFile->nFileLen + 1;

          //Codepage
          if (*wpCount)
          {
            lpRecentFile->nCodePage=(int)xatoiW(wpCount, NULL);
            wpCount+=xstrlenW(wpCount) + 1;
          }

          //Position
          if (*wpCount)
          {
            lpRecentFile->cpMin=lpRecentFile->cpMax=xatoiW(wpCount, &wpCount);
            if (*wpCount == L'-')
              lpRecentFile->cpMax=xatoiW(++wpCount, NULL);
            wpCount+=xstrlenW(wpCount) + 1;
          }

          //Parameters
          if (*wpCount)
          {
            for (wpParamName=wpCount; *wpCount; ++wpCount)
            {
              if (*wpCount == L'=')
              {
                nParamNameLen=(int)(wpCount - wpParamName);
                wpParamValue=wpCount + 1;
                while (*++wpCount);
                nParamValueLen=(int)(wpCount - wpParamValue);

                if (lpRecentFileParam=StackRecentFileParamAdd(lpRecentFile))
                {
                  if (lpRecentFileParam->pParamName=AllocWideStr(nParamNameLen + 1))
                    xstrcpynW(lpRecentFileParam->pParamName, wpParamName, nParamNameLen + 1);
                  if (lpRecentFileParam->pParamValue=AllocWideStr(nParamValueLen + 1))
                    xstrcpynW(lpRecentFileParam->pParamValue, wpParamValue, nParamValueLen + 1);
                }
                wpParamName=wpCount + 1;
              }
            }
          }
        }
      }
      FreeWideStr(wszRegData);
    }
  }
  RegCloseKey(hKey);
  return i;
}

void RecentFilesSave(RECENTFILESTACK *hStack)
{
  wchar_t wszRegKey[MAX_PATH];
  wchar_t wszRegValue[32];
  wchar_t *wszRegData;
  wchar_t wchNull=L'\0';
  HKEY hKey;
  int nDataLen;
  DWORD dwSaveTime;
  int i=0;

  //Params
  RECENTFILE *lpRecentFile;
  RECENTFILEPARAM *lpRecentFileParam;
  DWORD dwParamsLen;
  wchar_t *wszRecentFileParams=NULL;

  //Save recent files array
  xprintfW(wszRegKey, L"%s\\Recent", APP_REGHOMEW);
  if (RegCreateKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL) != ERROR_SUCCESS)
    return;

  for (lpRecentFile=hStack->first; lpRecentFile; lpRecentFile=lpRecentFile->next)
  {
    xprintfW(wszRegValue, L"file%d", i++);

    //Alloc params
    for (dwParamsLen=0, lpRecentFileParam=(RECENTFILEPARAM *)lpRecentFile->lpParamsStack.first; lpRecentFileParam; lpRecentFileParam=lpRecentFileParam->next)
    {
      dwParamsLen+=(DWORD)xprintfW(NULL, L"%s=%s%c", lpRecentFileParam->pParamName, lpRecentFileParam->pParamValue, wchNull) - 1;
    }
    if (dwParamsLen)
    {
      if (wszRecentFileParams=AllocWideStr(dwParamsLen + 1))
      {
        for (dwParamsLen=0, lpRecentFileParam=(RECENTFILEPARAM *)lpRecentFile->lpParamsStack.first; lpRecentFileParam; lpRecentFileParam=lpRecentFileParam->next)
        {
          dwParamsLen+=(DWORD)xprintfW(wszRecentFileParams + dwParamsLen, L"%s=%s%c", lpRecentFileParam->pParamName, lpRecentFileParam->pParamValue, wchNull);
        }
      }
    }

    //Set value
    for (wszRegData=NULL;;)
    {
      nDataLen=(int)xprintfW(wszRegData, L"%s%c%d%c%Id-%Id%c", lpRecentFile->wszFile, wchNull, lpRecentFile->nCodePage, wchNull, lpRecentFile->cpMin, lpRecentFile->cpMax, wchNull);
      if (!wszRegData)
        wszRegData=AllocWideStr(nDataLen + dwParamsLen);
      else
        break;
    }
    if (wszRegData)
    {
      xmemcpy(wszRegData + nDataLen, wszRecentFileParams, dwParamsLen * sizeof(wchar_t));
      wszRegData[nDataLen + dwParamsLen]=L'\0';

      if (RegSetValueExWide(hKey, wszRegValue, 0, REG_MULTI_SZ, (LPBYTE)wszRegData, (nDataLen + dwParamsLen + 1) * sizeof(wchar_t)) != ERROR_SUCCESS)
      {
        //Too long string
      }
      FreeWideStr(wszRegData);
    }

    //Free params
    if (wszRecentFileParams)
    {
      FreeWideStr(wszRecentFileParams);
      wszRecentFileParams=NULL;
    }
  }

  for (;;)
  {
    xprintfW(wszRegValue, L"file%d", i++);
    if (RegDeleteValueWide(hKey, wszRegValue) != ERROR_SUCCESS)
      break;
  }

  //Tick count
  dwSaveTime=(DWORD)GetTickCount();
  RegSetValueExWide(hKey, L"SaveTime", 0, REG_DWORD, (LPBYTE)&dwSaveTime, sizeof(DWORD));
  hStack->dwSaveTime=dwSaveTime;

  RegCloseKey(hKey);
}

void RecentFilesSaveFile(FRAMEDATA *lpFrame)
{
  RECENTFILE *lpRecentFile;
  CHARRANGE64 cr;
  BOOL bSwitch=FALSE;

  if (moCur.nRecentFiles && lpFrame->wszFile[0])
  {
    if (!nMainOnFinish || !nMDI || xstrcmpiW(fdDefault.wszFile, lpFrame->wszFile))
    {
      RecentFilesRead(&hRecentFilesStack);

      //Get selection
      SendMessage(lpFrame->ei.hWndEdit, EM_EXGETSEL64, 0, (LPARAM)&cr);
      if (lpFrame == lpFrameCurrent)
      {
        if (AEC_IndexCompare(&ciCurCaret, &crCurSel.ciMax) < 0)
          bSwitch=TRUE;
      }
      else
      {
        AESELECTION aes;
        AECHARINDEX ciCaret;

        SendMessage(lpFrame->ei.hWndEdit, AEM_GETSEL, (WPARAM)&ciCaret, (LPARAM)&aes);
        if (AEC_IndexCompare(&ciCaret, &aes.crSel.ciMax) < 0)
          bSwitch=TRUE;
      }

      if (lpRecentFile=RecentFilesUpdate(lpFrame->wszFile))
      {
        lpRecentFile->nCodePage=lpFrame->ei.nCodePage;
        lpRecentFile->cpMin=bSwitch?cr.cpMax:cr.cpMin;
        lpRecentFile->cpMax=bSwitch?cr.cpMin:cr.cpMax;
      }
      RecentFilesSave(&hRecentFilesStack);
      bMenuRecentFiles=TRUE;
    }
  }
}

void RecentFilesMenu()
{
  RECENTFILE *lpRecentFile;
  int i;

  for (i=1; DeleteMenu(hMainMenu, IDM_RECENT_FILES + i, MF_BYCOMMAND); ++i);

  if (moCur.nRecentFiles && hRecentFilesStack.first)
  {
    API_LoadStringW(hLangLib, STR_RECENT_FILES_DELETE_DEAD, wbuf, BUFFER_SIZE);
    ModifyMenuWide(hMainMenu, IDM_RECENT_FILES, MF_BYCOMMAND|MF_STRING, IDM_RECENT_FILES, wbuf);

    for (i=0, lpRecentFile=hRecentFilesStack.first; lpRecentFile; lpRecentFile=lpRecentFile->next)
    {
      TrimPath(wbuf2, lpRecentFile->wszFile, RECENTFILES_RECORD_LENGTH);
      FixAmpW(wbuf2, wbuf, BUFFER_SIZE);
      InsertMenuWide(hMainMenu, IDM_RECENT_FILES, MF_BYCOMMAND|MF_STRING, IDM_RECENT_FILES + (++i), wbuf);
    }
    InsertMenuWide(hMainMenu, IDM_RECENT_FILES, MF_BYCOMMAND|MF_SEPARATOR, IDM_RECENT_FILES + (++i), NULL);
  }
  else
  {
    API_LoadStringW(hLangLib, STR_RECENT_FILES_EMPTY, wbuf, BUFFER_SIZE);
    ModifyMenuWide(hMainMenu, IDM_RECENT_FILES, MF_BYCOMMAND|MF_STRING|MF_GRAYED, IDM_RECENT_FILES, wbuf);
  }
}

RECENTFILEPARAM* StackRecentFileParamAdd(RECENTFILE *lpRecentFile)
{
  RECENTFILEPARAM *lpRecentFileParam=NULL;

  if (!StackInsertIndex((stack **)&lpRecentFile->lpParamsStack.first, (stack **)&lpRecentFile->lpParamsStack.last, (stack **)&lpRecentFileParam, -1, sizeof(RECENTFILEPARAM)))
  {
    lpRecentFileParam->file=lpRecentFile;
  }
  return lpRecentFileParam;
}

RECENTFILEPARAM* StackRecentFileParamGetByName(RECENTFILE *lpRecentFile, const wchar_t *wpParamName)
{
  RECENTFILEPARAM *lpRecentFileParam;

  for (lpRecentFileParam=(RECENTFILEPARAM *)lpRecentFile->lpParamsStack.first; lpRecentFileParam; lpRecentFileParam=lpRecentFileParam->next)
  {
    if (!xstrcmpiW(lpRecentFileParam->pParamName, wpParamName))
      return lpRecentFileParam;
  }
  return NULL;
}

void StackRecentFileParamDelete(RECENTFILEPARAM *lpRecentFileParam)
{
  if (lpRecentFileParam->pParamName) FreeWideStr(lpRecentFileParam->pParamName);
  if (lpRecentFileParam->pParamValue) FreeWideStr(lpRecentFileParam->pParamValue);
  StackDelete((stack **)&lpRecentFileParam->file->lpParamsStack.first, (stack **)&lpRecentFileParam->file->lpParamsStack.last, (stack *)lpRecentFileParam);
}

void StackRecentFileParamFree(RECENTFILE *lpRecentFile)
{
  RECENTFILEPARAM *lpRecentFileParam;

  for (lpRecentFileParam=(RECENTFILEPARAM *)lpRecentFile->lpParamsStack.first; lpRecentFileParam; lpRecentFileParam=lpRecentFileParam->next)
  {
    if (lpRecentFileParam->pParamName) FreeWideStr(lpRecentFileParam->pParamName);
    if (lpRecentFileParam->pParamValue) FreeWideStr(lpRecentFileParam->pParamValue);
  }
  StackClear((stack **)&lpRecentFile->lpParamsStack.first, (stack **)&lpRecentFile->lpParamsStack.last);
}

int TrimPath(wchar_t *wszResult, const wchar_t *wpPath, int nMax)
{
  int nFileLen;
  int nLen=(int)xstrlenW(wpPath);
  const wchar_t *wpPathMax=wpPath + nLen;
  const wchar_t *wpFile=wpPathMax - 1;

  if (nLen <= nMax)
    return (int)xstrcpyW(wszResult, wpPath);

  while (wpFile >= wpPath && *wpFile != L'\\') --wpFile;

  if (wpFile >= wpPath)
  {
    nFileLen=(int)(wpPathMax - wpFile);

    if (nFileLen + 3 < nMax)
      return (int)xprintfW(wszResult, L"%.%ds...%s", nMax - nFileLen - 4, wpPath, wpFile);
  }
  return (int)xprintfW(wszResult, L"...%s", wpPathMax - nMax + 4);
}

int FixAmpW(const wchar_t *wpInput, wchar_t *wszOutput, int nOutputMax)
{
  int a;
  int b;

  //Replace "&" with "&&"
  nOutputMax-=2;

  if (nOutputMax >= 0)
  {
    for (a=0, b=0; wpInput[a] && b < nOutputMax; ++a, ++b)
    {
      wszOutput[b]=wpInput[a];
      if (wpInput[a] == L'&')
        wszOutput[++b]=L'&';
    }
    wszOutput[b]=L'\0';
    return b;
  }
  return -1;
}


//// Language menu

void LanguageMenu()
{
  WIN32_FIND_DATAW wfd;
  HANDLE hFind;
  int nCommand=0;
  int i;
  nLangModuleCount=0;

  for (i=1; DeleteMenu(hMainMenu, IDM_LANGUAGE + i, MF_BYCOMMAND); ++i);

  xprintfW(wbuf, L"%s\\AkelFiles\\Langs\\*.dll", wszExeDir);

  if ((hFind=FindFirstFileWide(wbuf, &wfd)) != INVALID_HANDLE_VALUE)
  {
    i=1;

    do
    {
      if (!xstrcmpiW(moCur.wszLangModule, wfd.cFileName))
        nCommand=IDM_LANGUAGE + i;
      GetBaseName(wfd.cFileName, wbuf, BUFFER_SIZE);
      InsertMenuWide(hMainMenu, IDM_LANGUAGE, MF_BYCOMMAND|MF_STRING, IDM_LANGUAGE + i, wbuf);
      ++i;
    }
    while (FindNextFileWide(hFind, &wfd));

    nLangModuleCount=i - 1;
    InsertMenuWide(hMainMenu, IDM_LANGUAGE, MF_BYCOMMAND|MF_SEPARATOR, IDM_LANGUAGE + i, NULL);
    FindClose(hFind);
  }

  API_LoadStringW(hLangLib, STR_LANGUAGE_NONE, wbuf, BUFFER_SIZE);
  if (!*moCur.wszLangModule)
    nCommand=IDM_LANGUAGE;
  ModifyMenuWide(hMainMenu, IDM_LANGUAGE, MF_BYCOMMAND|MF_STRING, IDM_LANGUAGE, wbuf);

  CheckMenuRadioItem(hMainMenu, nCommand, nCommand, nCommand, MF_BYCOMMAND);
}


//// Code page popup menu

void FillMenuPopupCodepage()
{
  int i;

  if (lpCodepageList)
  {
    while (DeleteMenu(hPopupOpenCodepage, 1, MF_BYPOSITION));
    while (DeleteMenu(hPopupSaveCodepage, 1, MF_BYPOSITION));

    for (i=0; lpCodepageList[i]; ++i)
    {
      GetCodePageName(lpCodepageList[i], wbuf, BUFFER_SIZE);

      if (!i)
      {
        ModifyMenuWide(hPopupOpenCodepage, i, MF_BYPOSITION|MF_STRING, IDM_POPUP_OPENAS, wbuf);
        ModifyMenuWide(hPopupSaveCodepage, i, MF_BYPOSITION|MF_STRING, IDM_POPUP_SAVEAS, wbuf);
      }
      else
      {
        InsertMenuWide(hPopupOpenCodepage, i + 1, MF_BYPOSITION|MF_STRING, IDM_POPUP_OPENAS + i, wbuf);
        InsertMenuWide(hPopupSaveCodepage, i + 1, MF_BYPOSITION|MF_STRING, IDM_POPUP_SAVEAS + i, wbuf);
      }
    }
  }
}

void ShowMenuPopupCodepage(POINT *ptScreen, BOOL bMouse)
{
  NCONTEXTMENU ncm;

  ncm.hWnd=hStatus;
  ncm.uType=NCM_STATUS;
  ncm.pt=*ptScreen;
  ncm.bMouse=bMouse;
  ncm.bProcess=TRUE;
  SendMessage(hMainWnd, AKDN_CONTEXTMENU, 0, (LPARAM)&ncm);

  if (ncm.bProcess)
  {
    TrackPopupMenu(hPopupCodepage, TPM_LEFTBUTTON|TPM_RIGHTBUTTON, ptScreen->x, ptScreen->y, 0, hMainWnd, NULL);
  }
}


//// Recode

void RecodeTextW(FRAMEDATA *lpFrame, HWND hWndPreview, DWORD dwFlags, int *nCodePageFrom, int *nCodePageTo)
{
  AECHARRANGE crRange;
  AECHARINDEX ciInitialCaret=ciCurCaret;
  wchar_t *wszSelText;
  wchar_t *wszText;
  int nFirstLine=0;
  INT_PTR nUnicodeLen;
  BOOL bCaretAtStart=FALSE;
  BOOL bSelection;

  if (!hWndPreview && !(dwFlags & RCS_DETECTONLY))
  {
    if (IsReadOnly(lpFrame->ei.hWndEdit))
      return;
    if (!AEC_IndexCompare(&crCurSel.ciMin, &ciCurCaret))
      bCaretAtStart=TRUE;

    //Save scroll
    nFirstLine=SaveLineScroll(lpFrame->ei.hWndEdit);
  }

  if (!AEC_IndexCompare(&crCurSel.ciMin, &crCurSel.ciMax))
  {
    SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_FIRSTCHAR, (LPARAM)&crRange.ciMin);
    SendMessage(lpFrame->ei.hWndEdit, AEM_GETINDEX, AEGI_LASTCHAR, (LPARAM)&crRange.ciMax);
    bSelection=FALSE;
  }
  else
  {
    crRange.ciMin=crCurSel.ciMin;
    crRange.ciMax=crCurSel.ciMax;
    bSelection=TRUE;
  }

  if (nUnicodeLen=ExGetRangeTextW(lpFrame->ei.hWndEdit, &crRange.ciMin, &crRange.ciMax, -1, &wszSelText, AELB_ASIS, TRUE))
  {
    //Detect
    if (*nCodePageFrom == -1 && *nCodePageTo == -1)
    {
      AutodetectWideChar(moCur.dwLangCodepageRecognition, wszSelText, nUnicodeLen, DETECTCHARS_REQUIRED, nCodePageFrom, nCodePageTo);
    }

    //Convert
    if (!(dwFlags & RCS_DETECTONLY))
    {
      if (ConvertWideString(wszSelText, nUnicodeLen, *nCodePageFrom, *nCodePageTo, &wszText, &nUnicodeLen))
      {
        FreeText(wszSelText);
        wszSelText=NULL;

        if (!hWndPreview)
        {
          if (!bSelection)
            SetSel(lpFrame->ei.hWndEdit, &crRange, AESELT_LOCKSCROLL, &crRange.ciMax);

          ReplaceSelW(lpFrame->ei.hWndEdit, wszText, nUnicodeLen, AELB_ASINPUT, AEREPT_COLUMNASIS|AEREPT_LOCKSCROLL, &crRange.ciMin, &crRange.ciMax);

          //Update selection
          if (!bSelection)
          {
            SendMessage(lpFrame->ei.hWndEdit, AEM_INDEXUPDATE, 0, (LPARAM)&ciInitialCaret);
            crRange.ciMin=ciInitialCaret;
            crRange.ciMax=ciInitialCaret;
          }
          SetSel(lpFrame->ei.hWndEdit, &crRange, AESELT_COLUMNASIS|AESELT_LOCKSCROLL, bCaretAtStart?&crRange.ciMin:&crRange.ciMax);
        }
        else SendMessage(hWndPreview, AEM_SETTEXTW, (WPARAM)nUnicodeLen, (LPARAM)wszText);

        FreeWideStr(wszText);
      }
    }
    if (wszSelText) FreeText(wszSelText);
  }

  if (!hWndPreview && !(dwFlags & RCS_DETECTONLY))
  {
    //Restore scroll
    RestoreLineScroll(lpFrame->ei.hWndEdit, nFirstLine);
  }
}

BOOL CALLBACK RecodeDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndCodePageFromList;
  static HWND hWndCodePageToList;
  static HWND hWndCodePageAutodetect;
  static HWND hWndCodePagePreview;
  static HWND hWndOK;
  static HWND hWndCancel;
  static int nCodePageFrom=-1;
  static int nCodePageTo=-1;
  static BOOL bAutodetect=TRUE;
  static DIALOGRESIZE drs[]={{&hWndCodePageFromList,       DRS_SIZE|DRS_X, 0},
                             {&hWndCodePageToList,         DRS_SIZE|DRS_X, 0},
                             {&hWndCodePageAutodetect,     DRS_MOVE|DRS_X, 0},
                             {&hWndCodePagePreview,        DRS_SIZE|DRS_X, 0},
                             {&hWndCodePagePreview,        DRS_SIZE|DRS_Y, 0},
                             {&hWndOK,                     DRS_MOVE|DRS_X, 0},
                             {&hWndOK,                     DRS_MOVE|DRS_Y, 0},
                             {&hWndCancel,                 DRS_MOVE|DRS_X, 0},
                             {&hWndCancel,                 DRS_MOVE|DRS_Y, 0},
                             {0, 0, 0}};

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndCodePageFromList=GetDlgItem(hDlg, IDC_RECODE_CODEPAGEFROM_LIST);
    hWndCodePageToList=GetDlgItem(hDlg, IDC_RECODE_CODEPAGETO_LIST);
    hWndCodePageAutodetect=GetDlgItem(hDlg, IDC_RECODE_AUTODETECT);
    hWndCodePagePreview=GetDlgItem(hDlg, IDC_RECODE_PREVIEW);
    hWndOK=GetDlgItem(hDlg, IDOK);
    hWndCancel=GetDlgItem(hDlg, IDCANCEL);

    EnsureWindowInMonitor(&moCur.rcRecodeCurrentDialog);

    API_LoadStringW(hLangLib, STR_AUTODETECT, wbuf, BUFFER_SIZE);
    SetWindowTextWide(hWndCodePageAutodetect, wbuf);

    FillComboboxCodepage(hWndCodePageFromList, lpCodepageList);
    if (nCodePageFrom > 0) SelectComboboxCodepage(hWndCodePageFromList, nCodePageFrom);
    FillComboboxCodepage(hWndCodePageToList, lpCodepageList);
    if (nCodePageTo > 0) SelectComboboxCodepage(hWndCodePageToList, nCodePageTo);

    if (nCodePageFrom <= 0 && nCodePageTo <= 0 && bAutodetect == FALSE)
    {
      nCodePageFrom=-1;
      nCodePageTo=-1;
      bAutodetect=TRUE;
    }
    if (bAutodetect)
    {
      SendMessage(hWndCodePageAutodetect, BM_SETCHECK, BST_CHECKED, 0);
      EnableWindow(hWndCodePageFromList, FALSE);
      EnableWindow(hWndCodePageToList, FALSE);
    }

    SendMessage(hWndCodePagePreview, AEM_SETCOLORS, 0, (LPARAM)&lpFrameCurrent->aec);
    SetTabStops(hWndCodePagePreview, lpFrameCurrent->nTabStopSize, FALSE);
    SetChosenFont(hWndCodePagePreview, &lpFrameCurrent->lf);

    SendMessage(hDlg, WM_COMMAND, IDC_SETREADONLY, 0);
    PostMessage(hDlg, AKDLG_RECODEUPDATE, 0, 0);
  }
  else if (uMsg == WM_CONTEXTMENU)
  {
    if ((HWND)wParam == hWndCodePagePreview)
    {
      ShowStandardViewMenu((HWND)wParam, hPopupView, lParam != -1);
    }
    return TRUE;
  }
  else if (uMsg == AKDLG_RECODEUPDATE)
  {
    if (bAutodetect)
    {
      nCodePageFrom=-1;
      nCodePageTo=-1;
    }
    else
    {
      nCodePageFrom=GetComboboxCodepage(hWndCodePageFromList);
      nCodePageTo=GetComboboxCodepage(hWndCodePageToList);
    }

    if (nCodePageFrom != 0 && nCodePageTo != 0)
    {
      RecodeTextW(lpFrameCurrent, hWndCodePagePreview, 0, &nCodePageFrom, &nCodePageTo);

      if (nCodePageFrom <= 0 || SelectComboboxCodepage(hWndCodePageFromList, nCodePageFrom) == CB_ERR)
        SendMessage(hWndCodePageFromList, CB_SETCURSEL, (WPARAM)-1, 0);
      if (nCodePageTo <= 0 || SelectComboboxCodepage(hWndCodePageToList, nCodePageTo) == CB_ERR)
        SendMessage(hWndCodePageToList, CB_SETCURSEL, (WPARAM)-1, 0);

      if (nCodePageFrom <= 0 && nCodePageTo <= 0)
      {
        bAutodetect=FALSE;
        SendMessage(hWndCodePageAutodetect, BM_SETCHECK, BST_UNCHECKED, 0);
        EnableWindow(hWndCodePageFromList, TRUE);
        EnableWindow(hWndCodePageToList, TRUE);
      }
    }
    if (!lpFrameCurrent->ei.bReadOnly && nCodePageFrom > 0 && nCodePageTo > 0 && nCodePageFrom != nCodePageTo)
      EnableWindow(hWndOK, TRUE);
    else
      EnableWindow(hWndOK, FALSE);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_RECODE_CODEPAGEFROM_LIST ||
        LOWORD(wParam) == IDC_RECODE_CODEPAGETO_LIST)
    {
      if (HIWORD(wParam) == CBN_SELCHANGE)
        PostMessage(hDlg, AKDLG_RECODEUPDATE, 0, 0);
    }
    else if (LOWORD(wParam) == IDC_RECODE_AUTODETECT)
    {
      bAutodetect=(BOOL)SendMessage(hWndCodePageAutodetect, BM_GETCHECK, 0, 0);
      EnableWindow(hWndCodePageFromList, !bAutodetect);
      EnableWindow(hWndCodePageToList, !bAutodetect);
      PostMessage(hDlg, AKDLG_RECODEUPDATE, 0, 0);
    }
    else if (LOWORD(wParam) == IDC_SETREADONLY)
    {
      EnableWindow(hWndOK, !lpFrameCurrent->ei.bReadOnly);
    }
    else if (LOWORD(wParam) == IDOK ||
             LOWORD(wParam) == IDCANCEL)
    {
      if (LOWORD(wParam) == IDOK)
      {
        nCodePageFrom=GetComboboxCodepage(hWndCodePageFromList);
        nCodePageTo=GetComboboxCodepage(hWndCodePageToList);
      }
      DestroyWindow(hDlg);
      hDlgModeless=NULL;
      nModelessType=MLT_NONE;

      if (LOWORD(wParam) == IDOK)
      {
        RecodeTextW(lpFrameCurrent, NULL, 0, &nCodePageFrom, &nCodePageTo);
      }
      return TRUE;
    }
  }
  DialogResizeMessages(&drs[0], &rcRecodeMinMaxDialog, &moCur.rcRecodeCurrentDialog, DRM_PAINTSIZEGRIP, hDlg, uMsg, wParam, lParam);

  return FALSE;
}


//// Colors

BOOL CALLBACK ColorsDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndThemeName;
  static HWND hWndThemeSave;
  static HWND hWndThemeDelete;
  static HWND hWndList;
  static HWND hWndBkImageFileLabel;
  static HWND hWndBkImageFileEdit;
  static HWND hWndBkImageFileBrowse;
  static HWND hWndBkImageAlphaLabel;
  static HWND hWndBkImageAlphaEdit;
  static HWND hWndBkImageAlphaSpin;
  static HWND hWndOK;
  static HWND hWndCancel;
  static AECOLORS aecColorsDlg;
  static wchar_t wszBkImageFileDlg[MAX_PATH];
  static int nBkImageAlphaDlg;
  static BOOL bColorsChanged;
  static DIALOGRESIZE drs[]={{&hWndThemeName,         DRS_SIZE|DRS_X, 0},
                             {&hWndThemeSave,         DRS_MOVE|DRS_X, 0},
                             {&hWndThemeDelete,       DRS_MOVE|DRS_X, 0},
                             {&hWndList,              DRS_SIZE|DRS_X, 0},
                             {&hWndList,              DRS_SIZE|DRS_Y, 0},
                             {&hWndBkImageFileLabel,  DRS_SIZE|DRS_X, 0},
                             {&hWndBkImageFileLabel,  DRS_MOVE|DRS_Y, 0},
                             {&hWndBkImageFileEdit,   DRS_SIZE|DRS_X, 0},
                             {&hWndBkImageFileEdit,   DRS_MOVE|DRS_Y, 0},
                             {&hWndBkImageFileBrowse, DRS_MOVE|DRS_X, 0},
                             {&hWndBkImageFileBrowse, DRS_MOVE|DRS_Y, 0},
                             {&hWndBkImageAlphaLabel, DRS_MOVE|DRS_X, 0},
                             {&hWndBkImageAlphaLabel, DRS_MOVE|DRS_Y, 0},
                             {&hWndBkImageAlphaEdit,  DRS_MOVE|DRS_X, 0},
                             {&hWndBkImageAlphaEdit,  DRS_MOVE|DRS_Y, 0},
                             {&hWndBkImageAlphaSpin,  DRS_MOVE|DRS_X, 0},
                             {&hWndBkImageAlphaSpin,  DRS_MOVE|DRS_Y, 0},
                             {&hWndOK,                DRS_MOVE|DRS_X, 0},
                             {&hWndOK,                DRS_MOVE|DRS_Y, 0},
                             {&hWndCancel,            DRS_MOVE|DRS_X, 0},
                             {&hWndCancel,            DRS_MOVE|DRS_Y, 0},
                             {0, 0, 0}};

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndThemeName=GetDlgItem(hDlg, IDC_COLORS_THEME_NAME);
    hWndThemeSave=GetDlgItem(hDlg, IDC_COLORS_THEME_SAVE);
    hWndThemeDelete=GetDlgItem(hDlg, IDC_COLORS_THEME_DELETE);
    hWndList=GetDlgItem(hDlg, IDC_COLORS_LIST);
    hWndBkImageFileLabel=GetDlgItem(hDlg, IDC_COLORS_BKIMAGEFILE_LABEL);
    hWndBkImageFileEdit=GetDlgItem(hDlg, IDC_COLORS_BKIMAGEFILE_EDIT);
    hWndBkImageFileBrowse=GetDlgItem(hDlg, IDC_COLORS_BKIMAGEFILE_BROWSE);
    hWndBkImageAlphaLabel=GetDlgItem(hDlg, IDC_COLORS_BKIMAGEALPHA_LABEL);
    hWndBkImageAlphaEdit=GetDlgItem(hDlg, IDC_COLORS_BKIMAGEALPHA_EDIT);
    hWndBkImageAlphaSpin=GetDlgItem(hDlg, IDC_COLORS_BKIMAGEALPHA_SPIN);
    hWndOK=GetDlgItem(hDlg, IDOK);
    hWndCancel=GetDlgItem(hDlg, IDCANCEL);

    SendMessage(hWndList, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_GRIDLINES, LVS_EX_GRIDLINES);
    SendMessage(hWndBkImageAlphaSpin, UDM_SETRANGE, 0, MAKELONG(255, 0));
    SendMessage(hWndBkImageAlphaSpin, UDM_SETBUDDY, (WPARAM)hWndBkImageAlphaEdit, 0);

    //Columns
    {
      LVCOLUMNW lvcW;

      lvcW.mask=LVCF_TEXT|LVCF_WIDTH|LVCF_SUBITEM;
      lvcW.pszText=wbuf;

      API_LoadStringW(hLangLib, STR_ELEMENT, wbuf, BUFFER_SIZE);
      lvcW.cx=140;
      lvcW.iSubItem=LVSI_COLOR_ELEMENT;
      ListView_InsertColumnWide(hWndList, LVSI_COLOR_ELEMENT, &lvcW);

      API_LoadStringW(hLangLib, STR_TEXT, wbuf, BUFFER_SIZE);
      lvcW.cx=70;
      lvcW.iSubItem=LVSI_COLOR_TEXT;
      ListView_InsertColumnWide(hWndList, LVSI_COLOR_TEXT, &lvcW);

      API_LoadStringW(hLangLib, STR_BACKGROUND, wbuf, BUFFER_SIZE);
      lvcW.cx=70;
      lvcW.iSubItem=LVSI_COLOR_BACKGROUND;
      ListView_InsertColumnWide(hWndList, LVSI_COLOR_BACKGROUND, &lvcW);

      API_LoadStringW(hLangLib, STR_SAMPLE, wbuf, BUFFER_SIZE);
      lvcW.cx=70;
      lvcW.iSubItem=LVSI_COLOR_SAMPLE;
      ListView_InsertColumnWide(hWndList, LVSI_COLOR_SAMPLE, &lvcW);
    }

    //Rows
    {
      LVITEMW lviW;
      int i;

      for (i=LVI_COLOR_BASIC; i <= LVI_COLOR_VISITURL; ++i)
      {
        API_LoadStringW(hLangLib, STR_BASIC + i, wbuf, BUFFER_SIZE);
        lviW.mask=LVIF_TEXT;
        lviW.pszText=wbuf;
        lviW.iItem=i;
        lviW.iSubItem=LVSI_COLOR_ELEMENT;
        ListView_InsertItemWide(hWndList, &lviW);
      }

      //Set "Sample" text
      API_LoadStringW(hLangLib, STR_SAMPLE, wbuf, BUFFER_SIZE);

      for (i=LVI_COLOR_BASIC; i <= LVI_COLOR_ALTLINE; ++i)
      {
        lviW.mask=LVIF_TEXT;
        lviW.pszText=wbuf;
        lviW.iItem=i;
        lviW.iSubItem=LVSI_COLOR_SAMPLE;
        ListView_SetItemWide(hWndList, &lviW);
      }
    }

    //Combobox
    {
      COLORTHEME *ctElement;
      int nSelection=-1;

      FillComboboxThemes(hWndThemeName);

      xmemcpy(&aecColorsDlg, &lpFrameCurrent->aec, sizeof(AECOLORS));
      xstrcpynW(wszBkImageFileDlg, lpFrameCurrent->wszBkImageFile, MAX_PATH);
      nBkImageAlphaDlg=lpFrameCurrent->nBkImageAlpha;

      if (!aecColorsDlg.dwFlags)
      {
        aecColorsDlg.dwFlags=AECLR_DEFAULT|AECLR_ALL;
        SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETCOLORS, 0, (LPARAM)&aecColorsDlg);
      }
      if (ctElement=StackThemeGetByData(&hThemesStack, &aecColorsDlg, wszBkImageFileDlg, nBkImageAlphaDlg))
        nSelection=ComboBox_FindStringWide(hWndThemeName, -1, ctElement->wszName);
      SendMessage(hWndThemeName, CB_SETCURSEL, (WPARAM)nSelection, 0);

      if (nSelection <= 0)
      {
        EnableWindow(hWndThemeSave, FALSE);
        EnableWindow(hWndThemeDelete, FALSE);
      }
      else if (nSelection > 0)
      {
        EnableWindow(hWndThemeSave, FALSE);
        EnableWindow(hWndThemeDelete, TRUE);
      }
    }

    //Background image
    SetWindowTextWide(hWndBkImageFileEdit, wszBkImageFileDlg);
    SetDlgItemInt(hDlg, IDC_COLORS_BKIMAGEALPHA_EDIT, nBkImageAlphaDlg, FALSE);

    //Colors inheriting is locked. Don't let user press OK, but allow to save as new theme.
    if ((lpFrameCurrent->dwLockInherit & LI_COLORS) || (lpFrameCurrent->dwLockInherit & LI_BKIMAGE))
      EnableWindow(hWndOK, FALSE);
    if (lpFrameCurrent->dwLockInherit & LI_BKIMAGE)
    {
      EnableWindow(hWndBkImageFileEdit, FALSE);
      EnableWindow(hWndBkImageFileBrowse, FALSE);
      EnableWindow(hWndBkImageAlphaEdit, FALSE);
      EnableWindow(hWndBkImageAlphaSpin, FALSE);
    }
    else
    {
      EnableWindow(hWndBkImageAlphaEdit, wszBkImageFileDlg[0]);
      EnableWindow(hWndBkImageAlphaSpin, wszBkImageFileDlg[0]);
    }
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (wParam == IDC_COLORS_LIST)
    {
      if (((NMHDR *)lParam)->code == (UINT)NM_CUSTOMDRAW)
      {
        LPNMLVCUSTOMDRAW lplvcd=(LPNMLVCUSTOMDRAW)lParam;
        LRESULT lResult;

        if (lplvcd->nmcd.dwDrawStage == CDDS_PREPAINT)
        {
          lResult=CDRF_NOTIFYITEMDRAW;
        }
        else if (lplvcd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT)
        {
          lResult=CDRF_NOTIFYSUBITEMDRAW;
        }
        else if (lplvcd->nmcd.dwDrawStage == (CDDS_SUBITEM|CDDS_ITEMPREPAINT))
        {
          if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_BASIC)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crBasicText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crBasicBk;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrText=aecColorsDlg.crBasicText;
              lplvcd->clrTextBk=aecColorsDlg.crBasicBk;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_SELECTION)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crSelText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crSelBk;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrText=aecColorsDlg.crSelText;
              lplvcd->clrTextBk=aecColorsDlg.crSelBk;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_ACTIVELINE)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crActiveLineText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crActiveLineBk;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrText=aecColorsDlg.crActiveLineText;
              lplvcd->clrTextBk=aecColorsDlg.crActiveLineBk;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_ACTIVELINEBORDER)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crActiveLineBorder;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crActiveLineBorder;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrText=aecColorsDlg.crActiveLineBorder;
              lplvcd->clrTextBk=aecColorsDlg.crActiveLineBorder;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_ALTLINE)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crAltLineText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crAltLineBk;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrText=aecColorsDlg.crAltLineText;
              lplvcd->clrTextBk=aecColorsDlg.crAltLineBk;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_ALTLINEBORDER)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crAltLineBorder;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crAltLineBorder;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrTextBk=aecColorsDlg.crAltLineBorder;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_ACTIVECOLUMN)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crActiveColumn;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crActiveColumn;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrTextBk=aecColorsDlg.crActiveColumn;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_COLUMNMARKER)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crColumnMarker;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crColumnMarker;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrTextBk=aecColorsDlg.crColumnMarker;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_CARET)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crCaret;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crCaret;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrTextBk=aecColorsDlg.crCaret;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_URL)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crUrlText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crUrlText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrTextBk=aecColorsDlg.crUrlText;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_ACTIVEURL)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crUrlCursorText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crUrlCursorText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrTextBk=aecColorsDlg.crUrlCursorText;
            }
          }
          else if (lplvcd->nmcd.dwItemSpec == LVI_COLOR_VISITURL)
          {
            if (lplvcd->iSubItem == LVSI_COLOR_TEXT)
            {
              lplvcd->clrTextBk=aecColorsDlg.crUrlVisitText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_BACKGROUND)
            {
              lplvcd->clrTextBk=aecColorsDlg.crUrlVisitText;
            }
            else if (lplvcd->iSubItem == LVSI_COLOR_SAMPLE)
            {
              lplvcd->clrTextBk=aecColorsDlg.crUrlVisitText;
            }
          }
          lResult=CDRF_DODEFAULT;
        }
        else lResult=CDRF_DODEFAULT;

        SetWindowLongPtrWide(hDlg, DWLP_MSGRESULT, lResult);
        return TRUE;
      }
      else if (((NMHDR *)lParam)->code == (UINT)NM_CLICK)
      {
        LVHITTESTINFO lvhti;
        BOOL bResult=FALSE;

        GetCursorPos(&lvhti.pt);
        ScreenToClient(hWndList, &lvhti.pt);
        SendMessage(hWndList, LVM_SUBITEMHITTEST, 0, (LPARAM)&lvhti);

        if (lvhti.iItem >= 0)
        {
          if (lvhti.iSubItem == LVSI_COLOR_TEXT)
          {
            if (lvhti.iItem == LVI_COLOR_BASIC)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crBasicText);
            else if (lvhti.iItem == LVI_COLOR_SELECTION)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crSelText);
            else if (lvhti.iItem == LVI_COLOR_ACTIVELINE)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crActiveLineText);
            else if (lvhti.iItem == LVI_COLOR_ACTIVELINEBORDER)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crActiveLineBorder);
            else if (lvhti.iItem == LVI_COLOR_ALTLINE)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crAltLineText);
            else if (lvhti.iItem == LVI_COLOR_ALTLINEBORDER)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crAltLineBorder);
            else if (lvhti.iItem == LVI_COLOR_ACTIVECOLUMN)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crActiveColumn);
            else if (lvhti.iItem == LVI_COLOR_COLUMNMARKER)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crColumnMarker);
            else if (lvhti.iItem == LVI_COLOR_CARET)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crCaret);
            else if (lvhti.iItem == LVI_COLOR_URL)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crUrlText);
            else if (lvhti.iItem == LVI_COLOR_ACTIVEURL)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crUrlCursorText);
            else if (lvhti.iItem == LVI_COLOR_VISITURL)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crUrlVisitText);
          }
          else if (lvhti.iSubItem == LVSI_COLOR_BACKGROUND)
          {
            if (lvhti.iItem == LVI_COLOR_BASIC)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crBasicBk);
            else if (lvhti.iItem == LVI_COLOR_SELECTION)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crSelBk);
            else if (lvhti.iItem == LVI_COLOR_ACTIVELINE)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crActiveLineBk);
            else if (lvhti.iItem == LVI_COLOR_ACTIVELINEBORDER)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crActiveLineBorder);
            else if (lvhti.iItem == LVI_COLOR_ALTLINE)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crAltLineBk);
            else if (lvhti.iItem == LVI_COLOR_ALTLINEBORDER)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crAltLineBorder);
            else if (lvhti.iItem == LVI_COLOR_ACTIVECOLUMN)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crActiveColumn);
            else if (lvhti.iItem == LVI_COLOR_COLUMNMARKER)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crColumnMarker);
            else if (lvhti.iItem == LVI_COLOR_CARET)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crCaret);
            else if (lvhti.iItem == LVI_COLOR_URL)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crUrlText);
            else if (lvhti.iItem == LVI_COLOR_ACTIVEURL)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crUrlCursorText);
            else if (lvhti.iItem == LVI_COLOR_VISITURL)
              bResult=SelectColorDialog(hDlg, &aecColorsDlg.crUrlVisitText);
          }
          if (bResult)
          {
            COLORTHEME *ctElement;
            int nSelection=-1;

            if (ctElement=StackThemeGetByData(&hThemesStack, &aecColorsDlg, wszBkImageFileDlg, nBkImageAlphaDlg))
              nSelection=ComboBox_FindStringWide(hWndThemeName, -1, ctElement->wszName);
            SendMessage(hWndThemeName, CB_SETCURSEL, (WPARAM)nSelection, 0);
            if (nSelection > 0)
              EnableWindow(hWndThemeDelete, TRUE);
            else
              EnableWindow(hWndThemeDelete, FALSE);

            bColorsChanged=TRUE;
            InvalidateRect(hWndList, NULL, TRUE);
          }
        }
      }
    }
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_COLORS_THEME_NAME)
    {
      if (HIWORD(wParam) == CBN_EDITCHANGE)
      {
        COLORTHEME *ctElement=NULL;
        int nSelection=-1;
        int nNameLen;

        if (nNameLen=GetWindowTextWide(hWndThemeName, wbuf, BUFFER_SIZE))
        {
          if (ctElement=StackThemeGetByName(&hThemesStack, wbuf))
          {
            nSelection=ComboBox_FindStringWide(hWndThemeName, -1, ctElement->wszName);
          }
          ctElement=StackThemeGetByData(&hThemesStack, &aecColorsDlg, wszBkImageFileDlg, nBkImageAlphaDlg);
        }

        if (!nSelection)
        {
          EnableWindow(hWndThemeSave, FALSE);
          EnableWindow(hWndThemeDelete, FALSE);
        }
        else
        {
          EnableWindow(hWndThemeSave, (nNameLen && (nSelection < 0 || !ctElement))?TRUE:FALSE);
          EnableWindow(hWndThemeDelete, TRUE);
        }
      }
      else if (HIWORD(wParam) == CBN_SELCHANGE)
      {
        COLORTHEME *ctElement;
        int nSelection;

        if ((nSelection=(int)SendMessage(hWndThemeName, CB_GETCURSEL, 0, 0)) != CB_ERR)
        {
          if (ComboBox_GetLBTextWide(hWndThemeName, nSelection, wbuf))
          {
            if (ctElement=StackThemeGetByName(&hThemesStack, wbuf))
            {
              xmemcpy(&aecColorsDlg, &ctElement->aec, sizeof(AECOLORS));
              xstrcpynW(wszBkImageFileDlg, ctElement->wszBkImageFile, MAX_PATH);
              SetWindowTextWide(hWndBkImageFileEdit, wszBkImageFileDlg);
              nBkImageAlphaDlg=ctElement->nBkImageAlpha;
              SetDlgItemInt(hDlg, IDC_COLORS_BKIMAGEALPHA_EDIT, nBkImageAlphaDlg, FALSE);
              bColorsChanged=TRUE;
              InvalidateRect(hWndList, NULL, TRUE);
            }
          }
        }

        if (!nSelection)
        {
          EnableWindow(hWndThemeSave, FALSE);
          EnableWindow(hWndThemeDelete, FALSE);
        }
        else
        {
          EnableWindow(hWndThemeSave, FALSE);
          EnableWindow(hWndThemeDelete, TRUE);
        }
      }
    }
    else if (LOWORD(wParam) == IDC_COLORS_THEME_SAVE)
    {
      COLORTHEME *ctElement;
      wchar_t wszThemeName[MAX_PATH];
      BOOL bProcess=TRUE;

      if (GetWindowTextWide(hWndThemeName, wszThemeName, MAX_PATH))
      {
        if (ctElement=StackThemeGetByName(&hThemesStack, wszThemeName))
        {
          API_LoadStringW(hLangLib, MSG_OVERWRITEPROMPT, wbuf, BUFFER_SIZE);
          xprintfW(wszMsg, wbuf, ctElement->wszName);
          if (API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_YESNO|MB_ICONQUESTION) == IDYES)
          {
            xstrcpynW(ctElement->wszName, wszThemeName, MAX_PATH);
            xmemcpy(&ctElement->aec, &aecColorsDlg, sizeof(AECOLORS));
            xstrcpynW(ctElement->wszBkImageFile, wszBkImageFileDlg, MAX_PATH);
            ctElement->nBkImageAlpha=nBkImageAlphaDlg;
          }
          else bProcess=FALSE;
        }
        else
        {
          StackThemeAdd(&hThemesStack, wszThemeName, &aecColorsDlg, wszBkImageFileDlg, nBkImageAlphaDlg, -1);
          ComboBox_AddStringWide(hWndThemeName, wszThemeName);
        }

        if (bProcess)
        {
          EnableWindow(hWndThemeSave, FALSE);
          EnableWindow(hWndThemeDelete, TRUE);
        }
      }
    }
    else if (LOWORD(wParam) == IDC_COLORS_THEME_DELETE)
    {
      if (GetWindowTextWide(hWndThemeName, wbuf, BUFFER_SIZE))
      {
        COLORTHEME *ctElement;
        int nSelection;

        if (ctElement=StackThemeGetByName(&hThemesStack, wbuf))
        {
          nSelection=ComboBox_FindStringWide(hWndThemeName, -1, ctElement->wszName);
          SendMessageW(hWndThemeName, CB_DELETESTRING, nSelection, 0);
          SendMessage(hWndThemeName, CB_SETCURSEL, (WPARAM)-1, 0);
          StackThemeDelete(&hThemesStack, ctElement);

          SetFocus(hWndThemeName);
          EnableWindow(hWndThemeSave, FALSE);
          EnableWindow(hWndThemeDelete, FALSE);
        }
      }
    }
    else if (LOWORD(wParam) == IDC_COLORS_BKIMAGEFILE_EDIT)
    {
      if (HIWORD(wParam) == EN_CHANGE)
      {
        if (!(lpFrameCurrent->dwLockInherit & LI_BKIMAGE))
        {
          COLORTHEME *ctElement;
          int nSelection=-1;

          GetWindowTextWide(hWndBkImageFileEdit, wszBkImageFileDlg, MAX_PATH);
          if (ctElement=StackThemeGetByData(&hThemesStack, &aecColorsDlg, wszBkImageFileDlg, nBkImageAlphaDlg))
            nSelection=ComboBox_FindStringWide(hWndThemeName, -1, ctElement->wszName);
          SendMessage(hWndThemeName, CB_SETCURSEL, (WPARAM)nSelection, 0);
          if (nSelection > 0)
            EnableWindow(hWndThemeDelete, TRUE);
          else
            EnableWindow(hWndThemeDelete, FALSE);
          EnableWindow(hWndBkImageAlphaEdit, wszBkImageFileDlg[0]);
          EnableWindow(hWndBkImageAlphaSpin, wszBkImageFileDlg[0]);

          bColorsChanged=TRUE;
        }
      }
    }
    else if (LOWORD(wParam) == IDC_COLORS_BKIMAGEFILE_BROWSE)
    {
      OPENFILENAMEW efnW;

      xmemset(&efnW, 0, sizeof(OPENFILENAMEW));
      efnW.lStructSize  =sizeof(OPENFILENAMEW);
      efnW.hwndOwner    =hDlg;
      efnW.lpstrFile    =wszBkImageFileDlg;
      efnW.nMaxFile     =MAX_PATH;
      efnW.lpstrFilter  =L"*.bmp;*.jpg;*.jpeg;*.gif\0*.bmp;*.jpg;*.jpeg;*.gif\0*.*\0*.*\0\0";;
      efnW.nFilterIndex =1;
      efnW.Flags        =OFN_HIDEREADONLY|OFN_PATHMUSTEXIST;

      if (GetOpenFileNameWide(&efnW))
      {
        SetWindowTextWide(hWndBkImageFileEdit, wszBkImageFileDlg);
        PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_COLORS_BKIMAGEFILE_EDIT, EN_CHANGE), 0);
      }
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_COLORS_BKIMAGEALPHA_EDIT)
    {
      if (HIWORD(wParam) == EN_CHANGE)
      {
        nBkImageAlphaDlg=GetDlgItemInt(hDlg, IDC_COLORS_BKIMAGEALPHA_EDIT, NULL, FALSE);
        if (nBkImageAlphaDlg > 255)
        {
          nBkImageAlphaDlg=255;
          SetDlgItemInt(hDlg, IDC_COLORS_BKIMAGEALPHA_EDIT, nBkImageAlphaDlg, FALSE);
        }
        PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_COLORS_BKIMAGEFILE_EDIT, EN_CHANGE), 0);
      }
    }
    else if (LOWORD(wParam) == IDOK)
    {
      if (bColorsChanged)
      {
        //Colors
        xmemcpy(&lpFrameCurrent->aec, &aecColorsDlg, sizeof(AECOLORS));
        lpFrameCurrent->aec.dwFlags=AECLR_ALL;
        SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_SETCOLORS, 0, (LPARAM)&lpFrameCurrent->aec);

        //Background image
        SetBkImage(lpFrameCurrent, wszBkImageFileDlg, nBkImageAlphaDlg);
      }

      SaveThemes(moCur.nSaveSettings);
      EndDialog(hDlg, 1);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDCANCEL)
    {
      EndDialog(hDlg, 0);
      return TRUE;
    }
  }
  DialogResizeMessages(&drs[0], &rcColorsMinMaxDialog, &moCur.rcColorsCurrentDialog, DRM_PAINTSIZEGRIP, hDlg, uMsg, wParam, lParam);

  return FALSE;
}

void FillComboboxThemes(HWND hWnd)
{
  COLORTHEME *ctElement=(COLORTHEME *)hThemesStack.first;

  while (ctElement)
  {
    ComboBox_AddStringWide(hWnd, ctElement->wszName);

    ctElement=ctElement->next;
  }
}

COLORTHEME* StackThemeAdd(HSTACK *hStack, const wchar_t *wpName, AECOLORS *aec, const wchar_t *wpFile, int nBkImageAlpha, int nIndex)
{
  COLORTHEME *ctElement;

  if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&ctElement, nIndex, sizeof(COLORTHEME)))
  {
    xstrcpynW(ctElement->wszName, wpName, MAX_PATH);
    ctElement->nNameLen=(int)xstrlenW(ctElement->wszName);
    if (aec) xmemcpy(&ctElement->aec, aec, sizeof(AECOLORS));
    xstrcpynW(ctElement->wszBkImageFile, wpFile, MAX_PATH);
    ctElement->nBkImageAlpha=nBkImageAlpha;
  }
  return ctElement;
}

COLORTHEME* StackThemeGetByName(HSTACK *hStack, const wchar_t *wpName)
{
  COLORTHEME *ctElement=(COLORTHEME *)hStack->first;
  int nNameLen=(int)xstrlenW(wpName);

  while (ctElement)
  {
    if (ctElement->nNameLen == nNameLen)
      if (!xstrcmpiW(ctElement->wszName, wpName))
        return ctElement;

    ctElement=ctElement->next;
  }
  return NULL;
}

COLORTHEME* StackThemeGetByData(HSTACK *hStack, AECOLORS *aec, const wchar_t *wpBkImageFile, int nBkImageAlpha)
{
  COLORTHEME *ctElement;

  for (ctElement=(COLORTHEME *)hStack->first; ctElement; ctElement=ctElement->next)
  {
    if (!xmemcmp(&ctElement->aec.crCaret, &aec->crCaret, sizeof(AECOLORS) - sizeof(DWORD)) &&
        !xstrcmpiW(ctElement->wszBkImageFile, wpBkImageFile) &&
        ctElement->nBkImageAlpha == nBkImageAlpha)
      return ctElement;
  }
  return NULL;
}

void StackThemeDelete(HSTACK *hStack, COLORTHEME *ctElement)
{
  StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)ctElement);
}

void StackThemeFree(HSTACK *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}

void ReadThemes(MAINOPTIONS *mo)
{
  COLORTHEME ct;
  HKEY hKey;
  DWORD dwSizeData=sizeof(COLORTHEME) - offsetof(COLORTHEME, aec);
  DWORD dwType;
  DWORD dwSizeValue;
  DWORD dwSizeString;
  DWORD dwIndex=0;

  if (mo->nSaveSettings == SS_REGISTRY)
  {
    wchar_t wszRegKey[MAX_PATH];

    xprintfW(wszRegKey, L"%s\\Themes", APP_REGHOMEW);
    if (RegOpenKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
      for (;;)
      {
        xmemset(&ct.aec, 0, dwSizeData);
        ct.nBkImageAlpha=EDIT_BKIMAGEALPHA;

        dwSizeValue=MAX_PATH;
        dwSizeString=dwSizeData;
        if (RegEnumValueWide(hKey, dwIndex, ct.wszName, &dwSizeValue, NULL, &dwType, (LPBYTE)&ct.aec, &dwSizeString) != ERROR_SUCCESS)
          break;

        StackThemeAdd(&hThemesStack, ct.wszName, &ct.aec, ct.wszBkImageFile, ct.nBkImageAlpha, -1);
        ++dwIndex;
      }
      RegCloseKey(hKey);
    }
  }
  else
  {
    INISECTION *lpIniSection;
    INIKEY *lpIniKey;

    if (lpIniSection=StackOpenIniSectionW(&hIniFile, L"Themes", (int)xstrlenW(L"Themes"), FALSE))
    {
      for (lpIniKey=lpIniSection->first; lpIniKey; lpIniKey=lpIniKey->next)
      {
        xmemset(&ct.aec, 0, dwSizeData);
        ct.nBkImageAlpha=EDIT_BKIMAGEALPHA;

        StackGetIniData(lpIniKey, INI_BINARY, (LPBYTE)&ct.aec, dwSizeData);
        StackThemeAdd(&hThemesStack, lpIniKey->wszKey, &ct.aec, ct.wszBkImageFile, ct.nBkImageAlpha, -1);
        ++dwIndex;
      }
    }
  }
}

BOOL SaveThemes(int nSaveSettings)
{
  COLORTHEME *ctElement=(COLORTHEME *)hThemesStack.first;
  HKEY hKey;
  INISECTION *lpIniSection;
  DWORD dwSizeData=sizeof(COLORTHEME) - offsetof(COLORTHEME, aec);
  BOOL bResult=FALSE;

  if (nSaveSettings == SS_REGISTRY)
  {
    wchar_t wszRegKey[MAX_PATH];

    xprintfW(wszRegKey, L"%s\\Themes", APP_REGHOMEW);
    //Clean old
    RegClearKeyWide(HKEY_CURRENT_USER, wszRegKey);
    if (RegCreateKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL) != ERROR_SUCCESS)
      return FALSE;
  }
  else
  {
    if (!OpenIniW(&hIniFile, wszIniFile, TRUE))
      return FALSE;

    //Clean old
    if (lpIniSection=StackOpenIniSectionW(&hIniFile, L"Themes", (int)xstrlenW(L"Themes"), FALSE))
      StackDeleteIniSection(&hIniFile, lpIniSection, TRUE);
  }

  //Skip "Standard" theme
  if (ctElement) ctElement=ctElement->next;

  while (ctElement)
  {
    if (nSaveSettings == SS_REGISTRY)
    {
      if (RegSetValueExWide(hKey, ctElement->wszName, 0, REG_BINARY, (LPBYTE)&ctElement->aec, dwSizeData) != ERROR_SUCCESS)
        break;
    }
    else
    {
      if (!IniSetValueW(&hIniFile, L"Themes", ctElement->wszName, INI_BINARY, (LPBYTE)&ctElement->aec, dwSizeData))
        break;
    }
    ctElement=ctElement->next;
  }

  if (nSaveSettings == SS_REGISTRY)
  {
    RegCloseKey(hKey);
    bResult=TRUE;
  }
  else
  {
    bResult=SaveIniW(&hIniFile, wszIniFile);
    StackFreeIni(&hIniFile);
  }
  return bResult;
}


//// Background image

BKIMAGEITEM* StackBkImageInsert(HSTACK *hStack, const wchar_t *wpFile)
{
  BKIMAGEITEM *lpElement=NULL;
  wchar_t wszFileExp[MAX_PATH];
  const wchar_t *wpExt;

  if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, -1, sizeof(BKIMAGEITEM)))
  {
    xstrcpynW(lpElement->wszBkImageFile, wpFile, MAX_PATH);
    TranslateFileString(wpFile, wszFileExp, MAX_PATH);

    if ((wpExt=GetFileExt(wpFile, -1)) && !xstrcmpiW(L"bmp", wpExt))
    {
      if (bOldWindows)
      {
        WideCharToMultiByte(CP_ACP, 0, wszFileExp, -1, buf, MAX_PATH, NULL, NULL);
        lpElement->hBkImageBitmap=(HBITMAP)LoadImageA(NULL, buf, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE|LR_LOADFROMFILE);
      }
      else lpElement->hBkImageBitmap=(HBITMAP)LoadImageW(NULL, wszFileExp, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE|LR_LOADFROMFILE);
    }
    else lpElement->hBkImageBitmap=LoadPictureFile(wszFileExp);
  }
  return lpElement;
}

BKIMAGEITEM* StackBkImageGet(HSTACK *hStack, const wchar_t *wpFile)
{
  BKIMAGEITEM *lpBkImage;

  for (lpBkImage=(BKIMAGEITEM *)hStack->first; lpBkImage; lpBkImage=lpBkImage->next)
  {
    if (!xstrcmpiW(lpBkImage->wszBkImageFile, wpFile))
      return lpBkImage;
  }
  return NULL;
}

void StackBkImageFree(HSTACK *hStack)
{
  BKIMAGEITEM *lpBkImage;

  for (lpBkImage=(BKIMAGEITEM *)hStack->first; lpBkImage; lpBkImage=lpBkImage->next)
  {
    if (lpBkImage->hBkImageBitmap) DeleteObject(lpBkImage->hBkImageBitmap);
  }
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}

BOOL SetBkImage(FRAMEDATA *lpFrame, const wchar_t *wpFile, int nBkImageAlpha)
{
  BKIMAGEITEM *bkfi;

  xstrcpynW(lpFrame->wszBkImageFile, wpFile, MAX_PATH);
  lpFrame->nBkImageAlpha=nBkImageAlpha;

  if (*lpFrame->wszBkImageFile)
  {
    if (!(bkfi=StackBkImageGet(&hBkImagesStack, lpFrame->wszBkImageFile)))
      bkfi=StackBkImageInsert(&hBkImagesStack, lpFrame->wszBkImageFile);
    lpFrame->hBkImageBitmap=bkfi->hBkImageBitmap;
  }
  else lpFrame->hBkImageBitmap=NULL;

  return (BOOL)SendMessage(lpFrame->ei.hWndEdit, AEM_SETBACKGROUNDIMAGE, (WPARAM)lpFrame->hBkImageBitmap, (LPARAM)lpFrame->nBkImageAlpha);
}

HBITMAP LoadPictureFile(wchar_t *wpFile)
{
  HBITMAP hBitmapResult=NULL;
  HBITMAP hBitmapOle=NULL;
  IPicture *lpPic=NULL;

  if (OleLoadPicturePath(wpFile, NULL, 0, 0, &IID_IPicture, (void **)&lpPic) == S_OK)
  {
    if (lpPic->lpVtbl->get_Handle(lpPic, (UINT*)&hBitmapOle) == S_OK)
    {
      hBitmapResult=(HBITMAP)CopyImage(hBitmapOle, IMAGE_BITMAP, 0, 0, 0);
    }
    lpPic->lpVtbl->Release(lpPic);
  }
  return hBitmapResult;
}


//// Plugins stack

void RegisterPluginsHotkeys(MAINOPTIONS *mo)
{
  if (mo->nSaveSettings == SS_REGISTRY)
  {
    wchar_t wszRegKey[MAX_PATH];
    HKEY hKey;
    DWORD dwType;
    DWORD dwSizeValue;
    DWORD dwSizeString;
    DWORD dwIndex=0;
    DWORD dwHotkey=0;

    xprintfW(wszRegKey, L"%s\\Plugs", APP_REGHOMEW);
    if (RegOpenKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
      for (;;)
      {
        dwSizeValue=BUFFER_SIZE;
        dwSizeString=sizeof(DWORD);
        if (RegEnumValueWide(hKey, dwIndex++, wbuf, &dwSizeValue, NULL, &dwType, (LPBYTE)&dwHotkey, &dwSizeString) != ERROR_SUCCESS)
          break;

        if (dwHotkey)
        {
          StackPluginAdd(&hPluginsStack, wbuf, (int)xstrlenW(wbuf), LOWORD(dwHotkey), HIWORD(dwHotkey), NULL, NULL);
        }
      }
      RegCloseKey(hKey);
    }
  }
  else
  {
    INISECTION *lpIniSection;
    INIKEY *lpIniKey;
    DWORD dwHotkey=0;

    if (lpIniSection=StackOpenIniSectionW(&hIniFile, L"Plugs", (int)xstrlenW(L"Plugs"), FALSE))
    {
      lpIniKey=(INIKEY *)lpIniSection->first;

      while (lpIniKey)
      {
        if (dwHotkey=(DWORD)xatoiW(lpIniKey->wszString, NULL))
        {
          StackPluginAdd(&hPluginsStack, lpIniKey->wszKey, lpIniKey->nKeyBytes / sizeof(wchar_t) - 1, LOWORD(dwHotkey), HIWORD(dwHotkey), NULL, NULL);
        }
        lpIniKey=lpIniKey->next;
      }
    }
  }
}

PLUGINFUNCTION* StackPluginFind(HSTACK *hStack, const wchar_t *wpPluginFunction, int nPluginFunctionLen)
{
  PLUGINFUNCTION *pfElement=(PLUGINFUNCTION *)hStack->first;

  if (!wpPluginFunction) return NULL;
  if (nPluginFunctionLen == -1)
    nPluginFunctionLen=(int)xstrlenW(wpPluginFunction);

  while (pfElement)
  {
    if (pfElement->nFunctionLen == nPluginFunctionLen)
    {
      if (!xstrcmpiW(pfElement->wszFunction, wpPluginFunction))
        break;
    }
    pfElement=pfElement->next;
  }
  return pfElement;
}

PLUGINFUNCTION* StackHotkeyFind(HSTACK *hStack, WORD wHotkey)
{
  PLUGINFUNCTION *pfElement=(PLUGINFUNCTION *)hStack->first;

  while (pfElement)
  {
    if (pfElement->wHotkey == wHotkey)
        break;

    pfElement=pfElement->next;
  }
  return pfElement;
}

PLUGINFUNCTION* StackPluginAdd(HSTACK *hStack, const wchar_t *wpPluginFunction, int nPluginFunctionLen, WORD wHotkey, BOOL bAutoLoad, PLUGINPROC PluginProc, void *lpParameter)
{
  PLUGINFUNCTION *pfElement;

  if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&pfElement, -1, sizeof(PLUGINFUNCTION)))
  {
    pfElement->wHotkey=wHotkey;
    pfElement->bAutoLoad=bAutoLoad;
    pfElement->bRunning=FALSE;
    pfElement->PluginProc=PluginProc;
    pfElement->lpParameter=lpParameter;

    if (nPluginFunctionLen == -1)
      pfElement->nFunctionLen=(int)xstrlenW(wpPluginFunction);
    else
      pfElement->nFunctionLen=nPluginFunctionLen;
    pfElement->pFunction=bOldWindows?(LPBYTE)pfElement->szFunction:(LPBYTE)pfElement->wszFunction;
    xstrcpynW(pfElement->wszFunction, wpPluginFunction, MAX_PATH);
    WideCharToMultiByte(CP_ACP, 0, pfElement->wszFunction, pfElement->nFunctionLen + 1, pfElement->szFunction, MAX_PATH, NULL, NULL);
  }
  return pfElement;
}

void StackPluginDelete(HSTACK *hStack, PLUGINFUNCTION *pfElement)
{
  if (!pfElement->nRefCount)
    StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)pfElement);
}

BOOL StackPluginSave(HSTACK *hStack, int nSaveSettings)
{
  PLUGINFUNCTION *pfElement=(PLUGINFUNCTION *)hStack->first;
  HKEY hKey;
  INISECTION *lpIniSection;
  DWORD dwHotkey;
  BOOL bResult=FALSE;

  if (nSaveSettings == SS_REGISTRY)
  {
    wchar_t wszRegKey[MAX_PATH];

    xprintfW(wszRegKey, L"%s\\Plugs", APP_REGHOMEW);
    //Clean old
    RegClearKeyWide(HKEY_CURRENT_USER, wszRegKey);
    if (RegCreateKeyExWide(HKEY_CURRENT_USER, wszRegKey, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL) != ERROR_SUCCESS)
      return FALSE;
  }
  else
  {
    if (!OpenIniW(&hIniFile, wszIniFile, TRUE))
      return FALSE;

    //Clean old
    if (lpIniSection=StackOpenIniSectionW(&hIniFile, L"Plugs", (int)xstrlenW(L"Plugs"), FALSE))
      StackDeleteIniSection(&hIniFile, lpIniSection, TRUE);
  }

  while (pfElement)
  {
    if (IsMainFunctionW(pfElement->wszFunction))
    {
      if (pfElement->wHotkey || pfElement->bAutoLoad)
      {
        dwHotkey=MAKELONG(pfElement->wHotkey, pfElement->bAutoLoad);

        if (nSaveSettings == SS_REGISTRY)
        {
          if (RegSetValueExWide(hKey, pfElement->wszFunction, 0, REG_DWORD, (LPBYTE)&dwHotkey, sizeof(DWORD)) != ERROR_SUCCESS)
            break;
        }
        else
        {
          if (!IniSetValueW(&hIniFile, L"Plugs", pfElement->wszFunction, INI_DWORD, (LPBYTE)&dwHotkey, sizeof(DWORD)))
            break;
        }
      }
    }
    pfElement=pfElement->next;
  }

  if (nSaveSettings == SS_REGISTRY)
  {
    RegCloseKey(hKey);
    bResult=TRUE;
  }
  else
  {
    bResult=SaveIniW(&hIniFile, wszIniFile);
    StackFreeIni(&hIniFile);
  }
  return bResult;
}

void StackPluginCleanUp(HSTACK *hStack, BOOL bDeleteNonExistentDLL)
{
  PLUGINFUNCTION *pfElement=(PLUGINFUNCTION *)hStack->first;
  PLUGINFUNCTION *pfNextElement;
  wchar_t wszDLL[MAX_PATH];
  wchar_t wszPlugin[MAX_PATH];

  while (pfElement)
  {
    pfNextElement=pfElement->next;

    if (!pfElement->wHotkey && !pfElement->bAutoLoad && !pfElement->bRunning)
    {
      StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)pfElement);
      pfElement=NULL;
    }
    if (bDeleteNonExistentDLL && pfElement)
    {
      if (ParsePluginNameW(pfElement->wszFunction, wszPlugin, NULL))
      {
        xprintfW(wszDLL, L"%s\\AkelFiles\\Plugs\\%s.dll", wszExeDir, wszPlugin);

        if (!FileExistsWide(wszDLL))
        {
          StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)pfElement);
          pfElement=NULL;
        }
      }
    }
    pfElement=pfNextElement;
  }
}

void StackPluginFree(HSTACK *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}

void CallPluginsOnStart(HSTACK *hStack)
{
  PLUGINFUNCTION *pfElement=(PLUGINFUNCTION *)hStack->first;
  PLUGINFUNCTION *pfNextElement;
  PLUGINCALLSENDW pcs;

  while (pfElement)
  {
    pfNextElement=pfElement->next;

    if (pfElement->bAutoLoad)
    {
      pcs.pFunction=pfElement->wszFunction;
      pcs.lParam=0;
      pcs.dwSupport=0;
      CallPluginSend(&pfElement, &pcs, DLLCF_ONPROGRAMLOAD);
    }
    pfElement=pfNextElement;
  }
}

int CallPluginSend(PLUGINFUNCTION **ppfElement, PLUGINCALLSENDW *pcs, DWORD dwFlags)
{
  PLUGINFUNCTION *pfElement=NULL;
  int nResult=UD_FAILED;

  if (pcs)
  {
    if (ppfElement)
      pfElement=*ppfElement;

    if (!pfElement)
    {
      if (!(pfElement=StackPluginFind(&hPluginsStack, pcs->pFunction, -1)))
        if (!(pfElement=StackPluginAdd(&hPluginsStack, pcs->pFunction, (int)xstrlenW(pcs->pFunction), 0, FALSE, NULL, NULL)))
          return UD_FAILED;
    }
    ++pfElement->nRefCount;

    if (pfElement->PluginProc)
    {
      if ((pfElement->PluginProc)(pfElement->lpParameter, pcs->lParam, pcs->dwSupport))
        nResult=UD_NONUNLOAD_UNCHANGE;
      else
        nResult=UD_NONUNLOAD_UNCHANGE|UD_HOTKEY_DODEFAULT;
      --pfElement->nRefCount;
    }
    else
    {
      nResult=CallPlugin(pfElement, pcs, dwFlags);
      --pfElement->nRefCount;

      if (nResult != UD_FAILED)
      {
        if (!(nResult & UD_NONUNLOAD_UNCHANGE))
        {
          if (nResult & UD_NONUNLOAD_ACTIVE)
          {
            if ((dwFlags & DLLCF_SWITCHAUTOLOAD) && !(pcs->dwSupport & PDS_NOAUTOLOAD) && !(pcs->dwSupport & PDS_GETSUPPORT))
            {
              if (!pfElement->bAutoLoad)
              {
                pfElement->bAutoLoad=TRUE;
                if (dwFlags & DLLCF_SAVENOW)
                  PostMessage(hMainWnd, AKD_DLLSAVE, DLLSF_NOW, 0);
                else if (dwFlags & DLLCF_SAVEONEXIT)
                  PostMessage(hMainWnd, AKD_DLLSAVE, DLLSF_ONEXIT, 0);
              }
            }
            pfElement->bRunning=TRUE;
          }
          else
          {
            if (dwFlags & DLLCF_SWITCHAUTOLOAD)
            {
              if (pfElement->bAutoLoad)
              {
                pfElement->bAutoLoad=FALSE;
                if (dwFlags & DLLCF_SAVENOW)
                  PostMessage(hMainWnd, AKD_DLLSAVE, DLLSF_NOW, 0);
                else if (dwFlags & DLLCF_SAVEONEXIT)
                  PostMessage(hMainWnd, AKD_DLLSAVE, DLLSF_ONEXIT, 0);
              }
            }
            if (pfElement->wHotkey || pfElement->bAutoLoad)
            {
              pfElement->bRunning=FALSE;
            }
            else
            {
              StackPluginDelete(&hPluginsStack, pfElement);
              pfElement=NULL;
            }
          }
        }
      }
      else
      {
        StackPluginDelete(&hPluginsStack, pfElement);
        pfElement=NULL;
      }
    }

    if (ppfElement)
      *ppfElement=pfElement;
  }
  return nResult;
}

int CallPlugin(PLUGINFUNCTION *lpPluginFunction, PLUGINCALLSENDW *pcs, DWORD dwFlags)
{
  wchar_t wszPlugin[MAX_PATH];
  wchar_t wszFunction[MAX_PATH];
  wchar_t wszDLL[MAX_PATH];
  wchar_t wszPluginWord[MAX_PATH];
  char szFunction[MAX_PATH];
  char szFullName[MAX_PATH];
  HMODULE hModule;
  PLUGINVERSION pv;
  PLUGINDATA pd;
  BOOL bInMemory=TRUE;
  BOOL bCalled=FALSE;
  int nWordLen;
  void (*PluginIDPtr)(PLUGINVERSION *);
  void (*PluginFunctionPtr)(PLUGINDATA *);

  if (pcs->pFunction)
  {
    if (ParsePluginNameW(pcs->pFunction, wszPlugin, wszFunction))
    {
      WideCharToMultiByte(CP_ACP, 0, wszFunction, -1, szFunction, MAX_PATH, NULL, NULL);
      WideCharToMultiByte(CP_ACP, 0, pcs->pFunction, -1, szFullName, MAX_PATH, NULL, NULL);
      xprintfW(wszDLL, L"%s\\AkelFiles\\Plugs\\%s.dll", wszExeDir, wszPlugin);
      nWordLen=API_LoadStringW(hLangLib, STR_PLUGIN, wbuf, BUFFER_SIZE);
      wbuf[0]=WideCharLower(wbuf[0]);
      xprintfW(wszPluginWord, L"%s %s", wszPlugin, wbuf);

      if (!(hModule=GetModuleHandleWide(wszDLL)))
      {
        if (hModule=LoadLibraryWide(wszDLL))
        {
          StackHandleIncrease(&hHandlesStack, hModule);
          bInMemory=FALSE;
        }
      }

      if (hModule)
      {
        if (PluginIDPtr=(void (*)(PLUGINVERSION *))GetProcAddress(hModule, "DllAkelPadID"))
        {
          pv.cb=sizeof(PLUGINVERSION);
          pv.hMainWnd=hMainWnd;
          pv.dwAkelDllVersion=MAKE_IDENTIFIER(-1, -1, -1, -1);
          pv.dwExeMinVersion3x=MAKE_IDENTIFIER(-1, -1, -1, -1);
          pv.dwExeMinVersion4x=MAKE_IDENTIFIER(-1, -1, -1, -1);
          pv.pPluginName=NULL;
          (*PluginIDPtr)(&pv);

          if (LOBYTE(pv.dwAkelDllVersion) == LOBYTE(AKELDLL) &&
              HIBYTE(pv.dwAkelDllVersion) == HIBYTE(AKELDLL))
          {
            if (pv.dwExeMinVersion4x != MAKE_IDENTIFIER(-1, -1, -1, -1))
            {
              if (VersionCompare(pv.dwExeMinVersion4x, dwExeVersion) <= 0)
              {
                if (PluginFunctionPtr=(void (*)(PLUGINDATA *))GetProcAddress(hModule, szFunction))
                {
                  pd.cb=sizeof(PLUGINDATA);
                  pd.dwSupport=pcs->dwSupport;
                  if (!(pd.dwSupport & PDS_STRANSI) && !(pd.dwSupport & PDS_STRWIDE))
                    pd.dwSupport|=bOldWindows?PDS_STRANSI:PDS_STRWIDE;
                  pd.pFunction=bOldWindows?(LPBYTE)szFullName:(LPBYTE)pcs->pFunction;
                  pd.szFunction=szFullName;
                  pd.wszFunction=pcs->pFunction;
                  pd.hInstanceDLL=hModule;
                  pd.lpPluginFunction=lpPluginFunction;
                  pd.nUnload=UD_UNLOAD;
                  pd.bInMemory=bInMemory;
                  pd.bOnStart=(dwFlags & DLLCF_ONPROGRAMLOAD);
                  pd.lParam=pcs->lParam;
                  pd.pAkelDir=bOldWindows?(LPBYTE)szExeDir:(LPBYTE)wszExeDir;
                  pd.szAkelDir=szExeDir;
                  pd.wszAkelDir=wszExeDir;
                  pd.hInstanceEXE=hInstance;
                  pd.hPluginsStack=&hPluginsStack;
                  pd.nSaveSettings=moCur.nSaveSettings;
                  pd.hMainWnd=hMainWnd;
                  pd.lpFrameData=lpFrameCurrent;
                  pd.hWndEdit=lpFrameCurrent->ei.hWndEdit;
                  pd.hDocEdit=lpFrameCurrent->ei.hDocEdit;
                  pd.hStatus=hStatus;
                  pd.hMdiClient=hMdiClient;
                  pd.hTab=hTab;
                  pd.hMainMenu=hMainMenu;
                  pd.hMenuRecentFiles=hMenuRecentFiles;
                  pd.hMenuLanguage=hMenuLanguage;
                  pd.hPopupMenu=hPopupMenu;
                  pd.hMainIcon=hMainIcon;
                  pd.hGlobalAccel=hGlobalAccel;
                  pd.hMainAccel=hMainAccel;
                  pd.bOldWindows=bOldWindows;
                  pd.bOldRichEdit=bOldRichEdit;
                  pd.bOldComctl32=bOldComctl32;
                  pd.bAkelEdit=bAkelEdit;
                  pd.nMDI=nMDI;
                  pd.pLangModule=bOldWindows?(LPBYTE)moCur.szLangModule:(LPBYTE)moCur.wszLangModule;
                  pd.szLangModule=moCur.szLangModule;
                  pd.wszLangModule=moCur.wszLangModule;
                  pd.wLangSystem=(WORD)dwLangSystem;
                  pd.wLangModule=(WORD)dwLangModule;

                  (*PluginFunctionPtr)(&pd);
                  SendMessage(hMainWnd, AKDN_DLLCALL, 0, (LPARAM)&pd);

                  pcs->dwSupport=pd.dwSupport;
                  if ((pcs->dwSupport & PDS_GETSUPPORT) && bInMemory)
                    return UD_NONUNLOAD_UNCHANGE;
                  if ((pd.nUnload & UD_NONUNLOAD_ACTIVE) ||
                      (pd.nUnload & UD_NONUNLOAD_NONACTIVE) ||
                      (pd.nUnload & UD_NONUNLOAD_UNCHANGE))
                    return pd.nUnload;
                  bCalled=TRUE;
                }
                else
                {
                  API_LoadStringW(hLangLib, MSG_FUNCTION_NOT_FOUND, wbuf, BUFFER_SIZE);
                  xprintfW(wszMsg, wbuf, wszPlugin, wszFunction, wszDLL);
                  API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
                  if (bInMemory) return UD_FAILED;
                }
              }
              else
              {
                API_LoadStringW(hLangLib, MSG_UPDATE_PROGRAM, wbuf, BUFFER_SIZE);
                xprintfW(wszMsg, wbuf, LOBYTE(pv.dwExeMinVersion4x), HIBYTE(pv.dwExeMinVersion4x), LOBYTE(HIWORD(pv.dwExeMinVersion4x)), HIBYTE(HIWORD(pv.dwExeMinVersion4x)),
                                       LOBYTE(dwExeVersion), HIBYTE(dwExeVersion), LOBYTE(HIWORD(dwExeVersion)), HIBYTE(HIWORD(dwExeVersion)));
                API_MessageBox(hMainWnd, wszMsg, wszPluginWord, MB_OK|MB_ICONEXCLAMATION);
              }
            }
            else
            {
              API_LoadStringW(hLangLib, MSG_PROGRAM_IS_NOT_SUPPORTED, wbuf, BUFFER_SIZE);
              xprintfW(wszMsg, wbuf, wszDLL);
              API_MessageBox(hMainWnd, wszMsg, wszPluginWord, MB_OK|MB_ICONEXCLAMATION);
            }
          }
          else
          {
            wchar_t wszStr[MAX_PATH];

            if (VersionCompare(pv.dwAkelDllVersion, AKELDLL) < 0)
              xstrcpynW(wszStr, wszPluginWord, MAX_PATH);
            else
              xstrcpynW(wszStr, L"AkelPad", MAX_PATH);
            API_LoadStringW(hLangLib, MSG_UPDATE_PLUGIN, wbuf, BUFFER_SIZE);
            xprintfW(wszMsg, wbuf, wszStr,
                                   LOBYTE(AKELDLL), HIBYTE(AKELDLL), LOBYTE(HIWORD(AKELDLL)), HIBYTE(HIWORD(AKELDLL)),
                                   wszPlugin,
                                   LOBYTE(pv.dwAkelDllVersion), HIBYTE(pv.dwAkelDllVersion), LOBYTE(HIWORD(pv.dwAkelDllVersion)), HIBYTE(HIWORD(pv.dwAkelDllVersion)));
            API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
          }
        }
        else
        {
          API_LoadStringW(hLangLib, MSG_PLUGIN_IS_NOT_SUPPORTED, wbuf, BUFFER_SIZE);
          xprintfW(wszMsg, wbuf, wszDLL);
          API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
        }

        if (FreeLibrary(hModule))
        {
          StackHandleDecrease(&hHandlesStack, hModule);
          if (bCalled) return pd.nUnload; //UD_UNLOAD
        }
      }
      else
      {
        API_LoadStringW(hLangLib, MSG_CANNOT_OPEN_FILE, wbuf, BUFFER_SIZE);
        xprintfW(wszMsg, wbuf, wszDLL);
        API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
      }
    }
  }
  return UD_FAILED;
}

DWORD TranslateMessageAll(DWORD dwType, LPMSG lpMsg)
{
  int nHotkeyStatus=0;

  if ((dwType & TMSG_GLOBAL) && TranslateMessageGlobal(lpMsg))
    return TMSG_GLOBAL;

  if ((dwType & TMSG_DIALOG) && TranslateMessageDialog(&hDocksStack, lpMsg))
    return TMSG_DIALOG;

  if ((dwType & TMSG_PLUGIN) && TranslateMessagePlugin(lpMsg))
    return TMSG_PLUGIN;

  if ((dwType & TMSG_HOTKEY) && (nHotkeyStatus=TranslateMessageHotkey(&hPluginsStack, lpMsg)) > 0)
    return TMSG_HOTKEY;

  if (lpFrameCurrent->ei.hWndEdit)
  {
    if (lpMsg->message == WM_KEYDOWN && lpMsg->wParam == VK_ESCAPE)
    {
      //Escape cancel column marker movement
      if (SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_GETMOUSESTATE, AEMS_CAPTURE, 0) & AEMSC_MOUSEMARKER)
        dwType&=~TMSG_ACCELERATOR;
    }
  }
  if ((dwType & TMSG_ACCELERATOR) && nHotkeyStatus == 0 && TranslateAcceleratorWide(hMainWnd, hMainAccel, lpMsg))
    return TMSG_ACCELERATOR;

  if (dwType & TMSG_DEFAULT)
  {
    TranslateMessage(lpMsg);
    DispatchMessageWide(lpMsg);
    return TMSG_DEFAULT;
  }

  return 0;
}

BOOL TranslateMessageGlobal(LPMSG lpMsg)
{
  if (TranslateAcceleratorWide(hMainWnd, hGlobalAccel, lpMsg))
    return TRUE;

  if (lpMsg->message == WM_KEYDOWN)
  {
    if (lpMsg->wParam == VK_ESCAPE)
    {
      if (dwMouseCapture)
      {
        ReleaseCapture();
        return TRUE;
      }
    }
  }
  return FALSE;
}

BOOL TranslateMessagePlugin(LPMSG lpMsg)
{
  if (lpMsg->message == AKD_POSTMESSAGE)
  {
    POSTMESSAGE *pm=(POSTMESSAGE *)lpMsg->lParam;

    if (pm)
    {
      SendMessageWide(pm->hWnd, pm->uMsg, pm->wParam, pm->lParam);
      GlobalFree((HGLOBAL)pm);
    }
    return TRUE;
  }
  else if (lpMsg->message == AKD_CALLPROC)
  {
    CALLPROC lpProcedure=(CALLPROC)lpMsg->wParam;
    void *lpParameter=(void *)lpMsg->lParam;

    (lpProcedure)(lpParameter);
    return TRUE;
  }
  else if (lpMsg->message == AKD_DLLCALL ||
           lpMsg->message == AKD_DLLCALLA ||
           lpMsg->message == AKD_DLLCALLW)
  {
    PLUGINCALLPOSTA *lpCallPostA=(PLUGINCALLPOSTA *)lpMsg->lParam;
    PLUGINCALLPOSTW *lpCallPostW=(PLUGINCALLPOSTW *)lpMsg->lParam;
    PLUGINCALLSENDW pcsW;
    wchar_t wszPluginFunction[MAX_PATH];

    if (lpMsg->lParam)
    {
      if (lpMsg->message == AKD_DLLCALLA || (bOldWindows && lpMsg->message == AKD_DLLCALL))
        xprintfW(wszPluginFunction, L"%S", (char *)lpCallPostA->szFunction);
      else
        xprintfW(wszPluginFunction, L"%s", (wchar_t *)lpCallPostW->szFunction);
      pcsW.pFunction=wszPluginFunction;
      pcsW.lParam=lpCallPostW->lParam;
      pcsW.dwSupport=lpCallPostW->dwSupport;
      CallPluginSend(NULL, &pcsW, (DWORD)lpMsg->wParam);
      GlobalFree((HGLOBAL)lpMsg->lParam);
    }
    return TRUE;
  }
  else if (lpMsg->message == AKD_DLLUNLOAD)
  {
    HMODULE hInstanceDLL=(HMODULE)lpMsg->wParam;
    HANDLE hThread=(HANDLE)lpMsg->lParam;
    char szPluginName[MAX_PATH];
    wchar_t wszPluginName[MAX_PATH];

    if (hInstanceDLL)
    {
      if (hThread)
      {
        WaitForSingleObject(hThread, INFINITE);
        CloseHandle(hThread);
      }

      if (GetModuleFileNameWide(hInstanceDLL, wbuf, BUFFER_SIZE))
      {
        GetBaseName(wbuf, wszPluginName, MAX_PATH);
        xprintfW(wszPluginName, L"%s::", wszPluginName);
        WideCharToMultiByte(CP_ACP, 0, wszPluginName, -1, szPluginName, MAX_PATH, NULL, NULL);

        if (FreeLibrary(hInstanceDLL))
        {
          PLUGINFUNCTION *pfElement=(PLUGINFUNCTION *)hPluginsStack.first;
          PLUGINFUNCTION *pfNextElement;
          UNISTRING us;

          StackHandleDecrease(&hHandlesStack, hInstanceDLL);

          //Clean-up plugins stack
          while (pfElement)
          {
            pfNextElement=pfElement->next;

            if (!xstrcmpinW(wszPluginName, pfElement->wszFunction, (UINT_PTR)-1))
            {
              if (pfElement->wHotkey || pfElement->bAutoLoad)
                pfElement->bRunning=FALSE;
              else
                StackPluginDelete(&hPluginsStack, pfElement);
            }
            pfElement=pfNextElement;
          }
          us.pString=bOldWindows?(LPBYTE)szPluginName:(LPBYTE)wszPluginName;
          us.szString=szPluginName;
          us.wszString=wszPluginName;
          SendMessage(hMainWnd, AKDN_DLLUNLOAD, 0, (WPARAM)&us);
        }
      }
    }
    return TRUE;
  }
  return FALSE;
}

int TranslateMessageHotkey(HSTACK *hStack, LPMSG lpMsg)
{
  PLUGINFUNCTION *pfElement=(PLUGINFUNCTION *)hStack->first;
  BYTE nMod=0;
  WORD wHotkey;
  int nResult;

  if (lpMsg->message == WM_SYSKEYDOWN ||  //Alt+..., Shift+Alt+...
      lpMsg->message == WM_KEYDOWN)       //Other combinations
  {
    if ((lpMsg->lParam >> 24) & 1) nMod|=HOTKEYF_EXT;
    if (GetKeyState(VK_CONTROL) & 0x80) nMod|=HOTKEYF_CONTROL;
    if (GetKeyState(VK_MENU) & 0x80) nMod|=HOTKEYF_ALT;
    if (GetKeyState(VK_SHIFT) & 0x80) nMod|=HOTKEYF_SHIFT;
    wHotkey=MAKEWORD(lpMsg->wParam, nMod);

    while (pfElement)
    {
      if (pfElement->wHotkey == wHotkey)
      {
        PLUGINCALLSENDW pcs;

        pcs.pFunction=pfElement->wszFunction;
        pcs.lParam=0;
        pcs.dwSupport=0;
        if (CallPluginSend(&pfElement, &pcs, 0) & UD_HOTKEY_DODEFAULT)
          break;
        return TRUE;
      }
      pfElement=pfElement->next;
    }
    nResult=FALSE;
    SendMessage(hMainWnd, AKDN_HOTKEY, (WPARAM)wHotkey, (LPARAM)&nResult);
    return nResult;
  }
  return FALSE;
}

BOOL ParsePluginNameW(const wchar_t *wpFullName, wchar_t *wszPlugin, wchar_t *wszFunction)
{
  wchar_t *wpFunction;

  if ((wpFunction=AKD_wcschr(wpFullName, L':')) && *(wpFunction + 1) == L':')
  {
    if (wszPlugin) xstrcpynW(wszPlugin, wpFullName, min(wpFunction - wpFullName + 1, MAX_PATH));
    if (wszFunction) xstrcpynW(wszFunction, wpFunction + 2, MAX_PATH);
    return TRUE;
  }
  return FALSE;
}

BOOL IsMainFunctionW(const wchar_t *wpFunction)
{
  int nCount=0;
  int i;

  for (i=0; wpFunction[i]; ++i)
  {
    if (wpFunction[i] == L':' && wpFunction[i + 1] == L':')
    {
      if (++nCount == 2) return FALSE;
    }
  }
  if (nCount == 1) return TRUE;

  return FALSE;
}

int CheckHotkey(WORD wHotkey, wchar_t *wszHotkeyOwner)
{
  PLUGINFUNCTION *pfElement;
  BOOL bExist=FALSE;
  int nOwner=HKO_NONE;

  if (wHotkey)
  {
    if (pfElement=StackHotkeyFind(&hPluginsStack, wHotkey))
    {
      xstrcpynW(wszHotkeyOwner, pfElement->wszFunction, MAX_PATH);
      nOwner=HKO_PLUGINFUNCTION;
    }
    else if ((pfElement=StackPluginFind(&hPluginsStack, L"Hotkeys::Main", -1)) && pfElement->bRunning)
    {
      //Hotkeys external call defines
      #define DLLA_HOTKEYS_CHECKHOTKEY  1

      typedef struct {
        UINT_PTR dwStructSize;
        INT_PTR nAction;
        UINT_PTR dwHotkey;
        BOOL *lpbExist;
        wchar_t *wszName;
        INT_PTR nNameMax;
      } DLLEXTHOTKEYS;

      //Hotkeys external call
      PLUGINCALLSENDW pcs;
      DLLEXTHOTKEYS deh;
      wchar_t wszName[MAX_PATH];

      wszName[0]=L'\0';
      deh.dwStructSize=sizeof(DLLEXTHOTKEYS);
      deh.nAction=DLLA_HOTKEYS_CHECKHOTKEY;
      deh.dwHotkey=wHotkey;
      deh.lpbExist=&bExist;
      deh.wszName=wszName;
      deh.nNameMax=MAX_PATH - 32;

      pcs.pFunction=L"Hotkeys::Main";
      pcs.lParam=(LPARAM)&deh;
      pcs.dwSupport=PDS_STRWIDE;
      CallPluginSend(NULL, &pcs, 0);

      if (bExist)
      {
        xprintfW(wszHotkeyOwner, L"Hotkeys::Main::%s", wszName);
        nOwner=HKO_HOTKEYSPLUGIN;
      }
    }
  }
  return nOwner;
}


//// Plugins core

BOOL CALLBACK PluginsDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndList;
  static HWND hWndListInfo;
  static HWND hWndHotkeyLabel;
  static HWND hWndHotkey;
  static HWND hWndAssign;
  static HWND hWndCall;
  static HWND hWndOK;
  static HWND hWndCancel;
  static int nSelItem;
  static BOOL bListChanged;
  static DIALOGRESIZE drs[]={{&hWndList,        DRS_SIZE|DRS_X, 0},
                             {&hWndList,        DRS_SIZE|DRS_Y, 0},
                             {&hWndListInfo,    DRS_MOVE|DRS_Y, 0},
                             {&hWndHotkeyLabel, DRS_MOVE|DRS_X, 0},
                             {&hWndHotkey,      DRS_MOVE|DRS_X, 0},
                             {&hWndAssign,      DRS_MOVE|DRS_X, 0},
                             {&hWndCall,        DRS_MOVE|DRS_Y, 0},
                             {&hWndOK,          DRS_MOVE|DRS_X, 0},
                             {&hWndOK,          DRS_MOVE|DRS_Y, 0},
                             {&hWndCancel,      DRS_MOVE|DRS_X, 0},
                             {&hWndCancel,      DRS_MOVE|DRS_Y, 0},
                             {0, 0, 0}};

  if (uMsg == WM_INITDIALOG)
  {
    LVCOLUMNW lvcW;
    wchar_t wszPlugin[MAX_PATH];
    wchar_t wszFunction[MAX_PATH];

    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndList=GetDlgItem(hDlg, IDC_PLUGINS_LIST);
    hWndListInfo=GetDlgItem(hDlg, IDC_PLUGINS_LIST_INFO);
    hWndHotkeyLabel=GetDlgItem(hDlg, IDC_PLUGINS_HOTKEY_LABEL);
    hWndHotkey=GetDlgItem(hDlg, IDC_PLUGINS_HOTKEY);
    hWndAssign=GetDlgItem(hDlg, IDC_PLUGINS_ASSIGN);
    hWndCall=GetDlgItem(hDlg, IDC_PLUGINS_CALL);
    hWndOK=GetDlgItem(hDlg, IDOK);
    hWndCancel=GetDlgItem(hDlg, IDCANCEL);

    EnableWindow(hWndHotkey, FALSE);
    EnableWindow(hWndAssign, FALSE);
    EnableWindow(hWndCall, FALSE);
    SendMessage(hWndList, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT|LVS_EX_GRIDLINES|LVS_EX_CHECKBOXES, LVS_EX_FULLROWSELECT|LVS_EX_GRIDLINES|LVS_EX_CHECKBOXES);
    SendMessage(hMainWnd, AKD_SETHOTKEYINPUT, (WPARAM)hWndHotkey, 0);

    //Columns
    lvcW.mask=LVCF_TEXT|LVCF_WIDTH|LVCF_SUBITEM;
    lvcW.pszText=wbuf;

    API_LoadStringW(hLangLib, STR_PLUGIN, wszPlugin, MAX_PATH);
    API_LoadStringW(hLangLib, STR_FUNCTION, wszFunction, MAX_PATH);
    xprintfW(wbuf, L"%s::%s", wszPlugin, wszFunction);
    lvcW.cx=209;
    lvcW.iSubItem=LVSI_FUNCTION_NAME;
    ListView_InsertColumnWide(hWndList, LVSI_FUNCTION_NAME, &lvcW);

    API_LoadStringW(hLangLib, STR_HOTKEY, wbuf, BUFFER_SIZE);
    lvcW.cx=105;
    lvcW.iSubItem=LVSI_FUNCTION_HOTKEY;
    ListView_InsertColumnWide(hWndList, LVSI_FUNCTION_HOTKEY, &lvcW);

    API_LoadStringW(hLangLib, STR_STATUS, wbuf, BUFFER_SIZE);
    lvcW.cx=63;
    lvcW.iSubItem=LVSI_FUNCTION_STATUS;
    ListView_InsertColumnWide(hWndList, LVSI_FUNCTION_STATUS, &lvcW);

    nSelItem=-1;
    FillPluginList(hWndList);
    bListChanged=FALSE;
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_PLUGINS_ASSIGN)
    {
      PLUGINLISTITEM *pliElement;
      LVITEMW lviW;
      wchar_t wszHotkeyOwner[MAX_PATH];
      int nOwner=HKO_NONE;
      WORD wHotkey;

      if (wHotkey=(WORD)SendMessage(hWndHotkey, HKM_GETHOTKEY, 0, 0))
        nOwner=CheckHotkey(wHotkey, wszHotkeyOwner);

      if (nOwner == HKO_NONE)
      {
        GetHotkeyString(wHotkey, wbuf);
        lviW.mask=LVIF_TEXT;
        lviW.pszText=wbuf;
        lviW.iItem=nSelItem;
        lviW.iSubItem=LVSI_FUNCTION_HOTKEY;
        ListView_SetItemWide(hWndList, &lviW);

        if (pliElement=GetPluginListItem(&hPluginListStack, nSelItem + 1))
        {
          SendMessage(hWndHotkey, HKM_SETHOTKEY, wHotkey, 0);
          pliElement->pf->wHotkey=wHotkey;
          bListChanged=TRUE;
        }
      }
      else
      {
        API_LoadStringW(hLangLib, MSG_HOTKEY_EXISTS, wbuf, BUFFER_SIZE);
        xprintfW(wszMsg, wbuf, wszHotkeyOwner);
        API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
        SetFocus(hWndHotkey);
      }
    }
    else if (LOWORD(wParam) == IDC_PLUGINS_CALL)
    {
      PLUGINLISTITEM *pliElement;
      PLUGINCALLPOSTW *pcp=NULL;

      if (nSelItem >= 0)
      {
        if (pliElement=GetPluginListItem(&hPluginListStack, nSelItem + 1))
        {
          if (pcp=(PLUGINCALLPOSTW *)GlobalAlloc(GPTR, sizeof(PLUGINCALLPOSTW)))
          {
            xstrcpynW(pcp->szFunction, pliElement->pf->wszFunction, MAX_PATH);
            xstrcpynW(wszLastFunction, pliElement->pf->wszFunction, MAX_PATH);
            pcp->lParam=0;
          }
        }

        FreePluginList(&hPluginListStack);
        StackPluginCleanUp(&hPluginsStack, bListChanged?TRUE:FALSE);
        if (bListChanged) StackPluginSave(&hPluginsStack, moCur.nSaveSettings);

        EndDialog(hDlg, 2);
        if (pcp) PostMessage(hMainWnd, AKD_DLLCALLW, 0, (LPARAM)pcp);
      }
    }
    else if (LOWORD(wParam) == IDOK)
    {
      FreePluginList(&hPluginListStack);
      StackPluginCleanUp(&hPluginsStack, TRUE);
      StackPluginSave(&hPluginsStack, moCur.nSaveSettings);

      EndDialog(hDlg, 1);
    }
    else if (LOWORD(wParam) == IDCANCEL)
    {
      PLUGINLISTITEM *pliElement=(PLUGINLISTITEM *)hPluginListStack.first;

      while (pliElement)
      {
        pliElement->pf->wHotkey=pliElement->wInitialHotkey;
        pliElement->pf->bAutoLoad=pliElement->bInitialOnStart;

        pliElement=pliElement->next;
      }
      FreePluginList(&hPluginListStack);
      StackPluginCleanUp(&hPluginsStack, FALSE);

      EndDialog(hDlg, 0);
    }
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (wParam == IDC_PLUGINS_LIST)
    {
      if (((NMLISTVIEW *)lParam)->hdr.code == LVN_ITEMCHANGING)
      {
        PLUGINLISTITEM *pliElement;
        BOOL bNewState;
        BOOL bOldState;

        if (((NMLISTVIEW *)lParam)->uNewState & LVIS_STATEIMAGEMASK)
        {
          bNewState=((((NMLISTVIEW *)lParam)->uNewState & LVIS_STATEIMAGEMASK) >> 12) - 1;
          bOldState=((((NMLISTVIEW *)lParam)->uOldState & LVIS_STATEIMAGEMASK) >> 12) - 1;

          if (bNewState >=0 && bOldState >=0 && bNewState != bOldState)
          {
            if (pliElement=GetPluginListItem(&hPluginListStack, ((NMLISTVIEW *)lParam)->iItem + 1))
            {
              if (pliElement->pf)
              {
                pliElement->pf->bAutoLoad=bNewState;
                bListChanged=TRUE;

                if (bNewState)
                {
                  if (pliElement->nCallResult == UD_FAILED || pliElement->nAutoLoad == -1)
                  {
                    //Check plugin autoload support
                    PLUGINCALLSENDW pcs;

                    pcs.pFunction=pliElement->pf->wszFunction;
                    pcs.lParam=0;
                    pcs.dwSupport=PDS_GETSUPPORT;
                    if ((pliElement->nCallResult=CallPluginSend(NULL, &pcs, 0)) != UD_FAILED)
                    {
                      if (pcs.dwSupport & PDS_NOAUTOLOAD)
                        pliElement->nAutoLoad=0;
                      else
                        pliElement->nAutoLoad=1;
                    }
                  }
                  if (pliElement->nCallResult == UD_FAILED || pliElement->nAutoLoad == 0)
                  {
                    if (pliElement->nAutoLoad == 0)
                    {
                      API_LoadStringW(hLangLib, MSG_AUTOLOAD_IS_NOT_SUPPORTED, wbuf, BUFFER_SIZE);
                      xprintfW(wszMsg, wbuf, pliElement->pf->wszFunction);
                      API_MessageBox(hDlg, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
                    }
                    pliElement->pf->bAutoLoad=FALSE;

                    SetWindowLongPtrWide(hDlg, DWLP_MSGRESULT, 1);
                    return TRUE;
                  }
                }
              }
            }
          }
        }
        if (((NMLISTVIEW *)lParam)->uNewState & LVIS_SELECTED)
        {
          nSelItem=((NMLISTVIEW *)lParam)->iItem;

          if (pliElement=GetPluginListItem(&hPluginListStack, nSelItem + 1))
          {
            SendMessage(hWndHotkey, HKM_SETHOTKEY, pliElement->pf->wHotkey, 0);
            EnableWindow(hWndHotkey, TRUE);
            EnableWindow(hWndAssign, TRUE);
            EnableWindow(hWndCall, TRUE);
          }
        }
        else if (((NMLISTVIEW *)lParam)->uOldState & LVIS_SELECTED)
        {
          nSelItem=-1;

          SendMessage(hWndHotkey, HKM_SETHOTKEY, 0, 0);
          EnableWindow(hWndHotkey, FALSE);
          EnableWindow(hWndAssign, FALSE);
          EnableWindow(hWndCall, FALSE);
        }
      }
      else if (((NMHDR *)lParam)->code == (UINT)NM_DBLCLK)
      {
        LVHITTESTINFO lvhti;

        GetCursorPos(&lvhti.pt);
        ScreenToClient(hWndList, &lvhti.pt);
        SendMessage(hWndList, LVM_HITTEST, 0, (LPARAM)&lvhti);

        if (lvhti.flags & LVHT_ONITEMLABEL)
          PostMessage(hDlg, WM_COMMAND, IDC_PLUGINS_CALL, 0);
      }
    }
  }
  DialogResizeMessages(&drs[0], &rcPluginsMinMaxDialog, &moCur.rcPluginsCurrentDialog, DRM_PAINTSIZEGRIP, hDlg, uMsg, wParam, lParam);

  return FALSE;
}

LRESULT CALLBACK GetMsgProc(int code, WPARAM wParam, LPARAM lParam)
{
  if (code >= 0)
  {
    MSG *msg=(MSG *)lParam;

    if (hWndPreviewDlg)
    {
      if (msg->message == WM_KEYDOWN)
      {
        if (SendMessage(hWndPreviewDlg, AKDLG_PREVIEWKEYDOWN, msg->wParam, (LPARAM)msg->hwnd))
          msg->message=WM_NULL;
      }
      else if (msg->message == WM_MOUSEWHEEL)
      {
        if (SendMessage(hWndPreviewDlg, AKDLG_PREVIEWMOUSEWHEEL, msg->wParam, (LPARAM)msg->hwnd))
          msg->message=WM_NULL;
      }
    }
  }
  return CallNextHookEx(hHookKeys, code, wParam, lParam);
}

void FillPluginList(HWND hWnd)
{
  WIN32_FIND_DATAW wfd;
  PLUGINLISTDATA pld={0};
  LVITEMW lvi;
  HANDLE hFind;
  HMODULE hInstance;
  wchar_t wszBaseName[MAX_PATH];
  const wchar_t *wpPluginExt=L"dll";

  nLastFunctionIndex=-1;
  pld.hWndList=hWnd;
  xprintfW(wbuf, L"%s\\AkelFiles\\Plugs\\*.%s", wszExeDir, wpPluginExt);

  if ((hFind=FindFirstFileWide(wbuf, &wfd)) != INVALID_HANDLE_VALUE)
  {
    do
    {
      //Avoid FindFirstFile/FindNextFile bug: "*.dll_ANYSYMBOLS" is also matched
      if (xstrcmpiW(wpPluginExt, GetFileExt(wfd.cFileName, -1)))
        continue;

      xprintfW(wbuf, L"%s\\AkelFiles\\Plugs\\%s", wszExeDir, wfd.cFileName);

      if (hInstance=LoadLibraryWide(wbuf))
      {
        if (GetProcAddress(hInstance, "DllAkelPadID"))
        {
          GetBaseName(wfd.cFileName, wszBaseName, MAX_PATH);
          pld.pBaseName=(unsigned char *)wszBaseName;
          GetExportNames(hInstance, FillPluginListProc, (LPARAM)&pld);
        }
        FreeLibrary(hInstance);
      }
    }
    while (FindNextFileWide(hFind, &wfd));

    FindClose(hFind);
  }

  if (nLastFunctionIndex >= 0)
  {
    lvi.stateMask=LVIS_SELECTED|LVIS_FOCUSED;
    lvi.state=LVIS_SELECTED|LVIS_FOCUSED;
    SendMessage(hWnd, LVM_SETITEMSTATE, (WPARAM)nLastFunctionIndex, (LPARAM)&lvi);
    SendMessage(hWnd, LVM_ENSUREVISIBLE, (WPARAM)nLastFunctionIndex, TRUE);
  }
}

BOOL CALLBACK FillPluginListProc(char *pExportName, LPARAM lParam)
{
  wchar_t wszExportName[MAX_PATH];
  wchar_t wszPluginFunction[MAX_PATH];
  PLUGINLISTDATA *pld=(PLUGINLISTDATA *)lParam;
  PLUGINLISTITEM *pliElement;
  PLUGINFUNCTION *pfElement;
  LVITEMW lviW;
  int nIndex;
  int nPluginFunctionLen;

  if (xmemcmp(pExportName, "Dll", 3))
  {
    MultiByteToWideChar(CP_ACP, 0, pExportName, -1, wszExportName, MAX_PATH);
    nPluginFunctionLen=(int)xprintfW(wszPluginFunction, L"%s::%s", (wchar_t *)pld->pBaseName, wszExportName);
    lviW.mask=LVIF_TEXT;
    lviW.pszText=wszPluginFunction;
    lviW.iItem=0;
    lviW.iSubItem=LVSI_FUNCTION_NAME;
    nIndex=ListView_InsertItemWide(pld->hWndList, &lviW);

    if (nLastFunctionIndex < 0)
    {
      if (!xstrcmpiW(wszPluginFunction, wszLastFunction))
        nLastFunctionIndex=nIndex;
    }
    else if (nLastFunctionIndex >= nIndex)
      ++nLastFunctionIndex;

    if (!StackInsertIndex((stack **)&hPluginListStack.first, (stack **)&hPluginListStack.last, (stack **)&pliElement, nIndex + 1, sizeof(PLUGINLISTITEM)))
    {
      if (pfElement=StackPluginFind(&hPluginsStack, wszPluginFunction, nPluginFunctionLen))
      {
        if (pfElement->bAutoLoad)
        {
          lviW.mask=LVIF_STATE;
          lviW.iItem=nIndex;
          lviW.iSubItem=LVSI_FUNCTION_NAME;
          lviW.state=((pfElement->bAutoLoad + 1) << 12);
          lviW.stateMask=LVIS_STATEIMAGEMASK;
          ListView_SetItemWide(pld->hWndList, &lviW);
        }
        if (pfElement->wHotkey)
        {
          GetHotkeyString(pfElement->wHotkey, wbuf);
          lviW.mask=LVIF_TEXT;
          lviW.pszText=wbuf;
          lviW.iItem=nIndex;
          lviW.iSubItem=LVSI_FUNCTION_HOTKEY;
          ListView_SetItemWide(pld->hWndList, &lviW);
        }
        if (pfElement->bRunning)
        {
          API_LoadStringW(hLangLib, STR_RUNNING, wbuf, BUFFER_SIZE);
          lviW.mask=LVIF_TEXT;
          lviW.pszText=wbuf;
          lviW.iItem=nIndex;
          lviW.iSubItem=LVSI_FUNCTION_STATUS;
          ListView_SetItemWide(pld->hWndList, &lviW);
        }
      }
      else pfElement=StackPluginAdd(&hPluginsStack, wszPluginFunction, nPluginFunctionLen, 0, FALSE, NULL, NULL);

      pliElement->pf=pfElement;
      pliElement->wInitialHotkey=pfElement->wHotkey;
      pliElement->bInitialOnStart=pfElement->bAutoLoad;
      pliElement->nAutoLoad=-1;
      pliElement->nCallResult=UD_FAILED;
    }
  }
  return TRUE;
}

PLUGINLISTITEM* GetPluginListItem(HSTACK *hStack, int nIndex)
{
  PLUGINLISTITEM *pliElement;

  if (!StackGetElement((stack *)hStack->first, (stack *)hStack->last, (stack **)&pliElement, nIndex))
    return pliElement;
  return NULL;
}

void FreePluginList(HSTACK *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}

BOOL GetExportNames(HMODULE hInstance, EXPORTNAMESPROC lpExportNamesProc, LPARAM lParam)
{
  PIMAGE_DOS_HEADER pDOSHead;
  PIMAGE_NT_HEADERS pPEHeader;
  PIMAGE_EXPORT_DIRECTORY pImportDesc;
  unsigned long *pchName;
  char *pName;
  DWORD i;

  pDOSHead=(PIMAGE_DOS_HEADER)hInstance;
  if (pDOSHead->e_magic != IMAGE_DOS_SIGNATURE)
    return FALSE;

  pPEHeader=(PIMAGE_NT_HEADERS)((UINT_PTR)hInstance + pDOSHead->e_lfanew);
  if (pPEHeader->Signature != IMAGE_NT_SIGNATURE)
    return FALSE;

  pImportDesc=(PIMAGE_EXPORT_DIRECTORY)((UINT_PTR)hInstance + pPEHeader->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_EXPORT].VirtualAddress);
  if (!pImportDesc->AddressOfNames)
    return FALSE;

  pchName=(unsigned long *)((UINT_PTR)hInstance + pImportDesc->AddressOfNames);

  for (i=0; i < pImportDesc->NumberOfNames; ++i)
  {
    pName=(char *)((UINT_PTR)hInstance + pchName[i]);
    if ((*lpExportNamesProc)(pName, lParam) == FALSE) return TRUE;
  }
  return TRUE;
}

int GetHotkeyString(WORD wHotkey, wchar_t *wszString)
{
  wchar_t wszKeyText[100];
  LONG lScan;
  int nResult=0;

  wszString[0]=L'\0';
  wszKeyText[0]=L'\0';

  if (wHotkey)
  {
    if (bOldWindows)
      lScan=MapVirtualKeyA(LOBYTE(wHotkey), 0) << 16;
    else
      lScan=MapVirtualKeyW(LOBYTE(wHotkey), 0) << 16;
    if (HIBYTE(wHotkey) & HOTKEYF_EXT)
      lScan|=0x01000000L;
    GetKeyNameTextWide(lScan, wszKeyText, 100);

    nResult=(int)xprintfW(wszString, L"%s%s%s%s",
                                     (HIBYTE(wHotkey) & HOTKEYF_CONTROL)?L"Ctrl + ":L"",
                                     (HIBYTE(wHotkey) & HOTKEYF_SHIFT)?L"Shift + ":L"",
                                     (HIBYTE(wHotkey) & HOTKEYF_ALT)?L"Alt + ":L"",
                                     wszKeyText);
  }
  return nResult;
}


//// Options

LRESULT CALLBACK CBTPropertySheetProc(int iCode, WPARAM wParam, LPARAM lParam)
{
  if (iCode == HCBT_ACTIVATE)
  {
    if (hHookPropertySheet)
    {
      if (UnhookWindowsHookEx(hHookPropertySheet))
        hHookPropertySheet=NULL;
    }

    //Center PropertySheet
    {
      RECT rcMain;
      RECT rcSheet;

      GetWindowRect(hMainWnd, &rcMain);
      GetWindowRect((HWND)wParam, &rcSheet);

      rcSheet.left=rcMain.left + (RectW(&rcMain) / 2) - (RectW(&rcSheet) / 2);
      rcSheet.top=rcMain.top + (RectH(&rcMain) / 2) - (RectH(&rcSheet) / 2);
      if (rcSheet.left < 0) rcSheet.left=0;
      if (rcSheet.top < 0) rcSheet.top=0;

      SetWindowPos((HWND)wParam, NULL, rcSheet.left, rcSheet.top, 0, 0, SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
    }
  }
  return CallNextHookEx(hHookPropertySheet, iCode, wParam, lParam);
}

int CALLBACK PropSheetProc(HWND hDlg, UINT uMsg, LPARAM lParam)
{
  //Remove "?"
  if (uMsg == PSCB_PRECREATE)
  {
    ((DLGTEMPLATE *)lParam)->style&=~DS_CONTEXTHELP;
  }
  else if (uMsg == PSCB_INITIALIZED)
  {
    hPropertyTab=(HWND)SendMessage(hDlg, PSM_GETTABCONTROL, 0, 0);
  }
  return TRUE;
}

BOOL CALLBACK OptionsGeneralDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndCommand;
  static HWND hWndDirectory;
  static HWND hWndAutodetectCP;
  static HWND hWndAutodetectCPBuffer;
  static HWND hWndDefaultCP;
  static HWND hWndNewFileCP;
  static HWND hWndNewFileBOM;
  static HWND hWndNewFileNewLine;
  int i;

  if (uMsg == WM_INITDIALOG)
  {
    hWndCommand=GetDlgItem(hDlg, IDC_OPTIONS_EXECCOM);
    hWndDirectory=GetDlgItem(hDlg, IDC_OPTIONS_EXECDIR);
    hWndAutodetectCP=GetDlgItem(hDlg, IDC_OPTIONS_CODEPAGE_RECOGNITION);
    hWndAutodetectCPBuffer=GetDlgItem(hDlg, IDC_OPTIONS_CODEPAGE_RECOGNITION_BUFFER);
    hWndDefaultCP=GetDlgItem(hDlg, IDC_OPTIONS_DEFAULT_CODEPAGE);
    hWndNewFileCP=GetDlgItem(hDlg, IDC_OPTIONS_NEWFILE_CODEPAGE);
    hWndNewFileBOM=GetDlgItem(hDlg, IDC_OPTIONS_NEWFILE_BOM);
    hWndNewFileNewLine=GetDlgItem(hDlg, IDC_OPTIONS_NEWFILE_NEWLINE);

    SendMessage(hWndCommand, EM_LIMITTEXT, (WPARAM)BUFFER_SIZE, 0);
    SendMessage(hWndDirectory, EM_LIMITTEXT, (WPARAM)MAX_PATH, 0);

    SetWindowTextWide(hWndCommand, moCur.wszExecuteCommand);
    SetWindowTextWide(hWndDirectory, moCur.wszExecuteDirectory);

    API_LoadStringW(hLangLib, STR_NONE, wbuf, BUFFER_SIZE);
    ComboBox_AddStringWide(hWndAutodetectCP, wbuf);
    for (i=STR_AUTODETECT_RUSSIAN; i <= STR_AUTODETECT_KOREAN; ++i)
    {
      API_LoadStringW(hLangLib, i, wbuf, BUFFER_SIZE);
      ComboBox_AddStringWide(hWndAutodetectCP, wbuf);
    }

    i=GetDetectionIndex(moCur.dwLangCodepageRecognition);
    SendMessage(hWndAutodetectCP, CB_SETCURSEL, (WPARAM)i, 0);

    SetDlgItemInt(hDlg, IDC_OPTIONS_CODEPAGE_RECOGNITION_BUFFER, moCur.dwCodepageRecognitionBuffer, FALSE);

    FillComboboxCodepage(hWndDefaultCP, lpCodepageList);
    SelectComboboxCodepage(hWndDefaultCP, moCur.nDefaultCodePage);

    FillComboboxCodepage(hWndNewFileCP, lpCodepageList);
    SelectComboboxCodepage(hWndNewFileCP, moCur.nNewFileCodePage);

    if (moCur.bNewFileBOM)
      SendMessage(hWndNewFileBOM, BM_SETCHECK, BST_CHECKED, 0);
    else if (!IsCodePageUnicode(moCur.nNewFileCodePage))
      EnableWindow(hWndNewFileBOM, FALSE);

    ComboBox_AddStringWide(hWndNewFileNewLine, STR_NEWLINE_WIN);
    ComboBox_AddStringWide(hWndNewFileNewLine, STR_NEWLINE_UNIX);
    ComboBox_AddStringWide(hWndNewFileNewLine, STR_NEWLINE_MAC);
    SendMessage(hWndNewFileNewLine, CB_SETCURSEL, (WPARAM)(moCur.nNewFileNewLine - 1), 0);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_OPTIONS_EXECCOM_BROWSE)
    {
      wchar_t wszExeFilter[MAX_PATH];
      BOOL bResult;

      //GetOpenFileName dialog file filter
      API_LoadStringW(hLangLib, STR_EXECUTABLE_FILTER, wszExeFilter, MAX_PATH);
      for (i=0; wszExeFilter[i]; ++i)
        if (wszExeFilter[i] == L'|') wszExeFilter[i]=L'\0';
      wszExeFilter[++i]=L'\0';

      //File browse
      {
        OPENFILENAMEW efnW;

        wbuf[0]=L'\0';
        xmemset(&efnW, 0, sizeof(OPENFILENAMEW));
        efnW.lStructSize      =sizeof(OPENFILENAMEW);
        efnW.hwndOwner        =hDlg;
        efnW.lpstrFile        =wbuf;
        efnW.nMaxFile         =MAX_PATH;
        efnW.lpstrFilter      =wszExeFilter;
        efnW.nFilterIndex     =1;
        efnW.Flags            =OFN_HIDEREADONLY|OFN_PATHMUSTEXIST;

        bResult=GetOpenFileNameWide(&efnW);
      }

      if (bResult)
      {
        xprintfW(wbuf2, L"\"%s\"", wbuf);
        SetWindowTextWide(hWndCommand, wbuf2);
      }
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_EXECDIR_BROWSE)
    {
      BROWSEINFOW biW;
      LPITEMIDLIST pIdList;
      LPMALLOC pMalloc;

      GetWindowTextWide(hWndDirectory, wbuf, BUFFER_SIZE);
      biW.hwndOwner=hDlg;
      biW.pidlRoot=NULL;
      biW.pszDisplayName=wbuf;
      biW.lpszTitle=NULL;
      biW.ulFlags=BIF_RETURNONLYFSDIRS|BIF_NEWDIALOGSTYLE|BIF_EDITBOX;
      biW.lpfn=BrowseCallbackProc;
      biW.lParam=(LPARAM)wbuf;
      biW.iImage=0;

      if (pIdList=SHBrowseForFolderWide(&biW))
      {
        SHGetPathFromIDListWide(pIdList, wbuf);

        if (SHGetMalloc(&pMalloc))
        {
          pMalloc->lpVtbl->Free(pMalloc, pIdList);
          pMalloc->lpVtbl->Release(pMalloc);
        }
        SetWindowTextWide(hWndDirectory, wbuf);
      }
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_CODEPAGE_FILTER)
    {
      if (API_DialogBox(hLangLib, MAKEINTRESOURCEW(IDD_OPTIONS_GENERAL_FILTER), hDlg, (DLGPROC)OptionsGeneralFilterDlgProc))
      {
        ClearCombobox(hWndDefaultCP);
        FillComboboxCodepage(hWndDefaultCP, lpCodepageList);
        if (SelectComboboxCodepage(hWndDefaultCP, moCur.nDefaultCodePage) == CB_ERR)
          SendMessage(hWndDefaultCP, CB_SETCURSEL, 0, 0);

        ClearCombobox(hWndNewFileCP);
        FillComboboxCodepage(hWndNewFileCP, lpCodepageList);
        if (SelectComboboxCodepage(hWndNewFileCP, moCur.nNewFileCodePage) == CB_ERR)
          SendMessage(hWndNewFileCP, CB_SETCURSEL, 0, 0);
      }
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_NEWFILE_CODEPAGE)
    {
      if (HIWORD(wParam) == CBN_SELCHANGE)
      {
        if (IsCodePageUnicode(GetComboboxCodepage(hWndNewFileCP)))
        {
          EnableWindow(hWndNewFileBOM, TRUE);
          SendMessage(hWndNewFileBOM, BM_SETCHECK, BST_CHECKED, 0);
        }
        else
        {
          EnableWindow(hWndNewFileBOM, FALSE);
          SendMessage(hWndNewFileBOM, BM_SETCHECK, BST_UNCHECKED, 0);
        }
      }
    }
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (((NMHDR *)lParam)->code == (UINT)PSN_SETACTIVE)
    {
      nPropertyStartPage=(int)SendMessage(hPropertyTab, TCM_GETCURSEL, 0, 0);
    }
    else if (((NMHDR *)lParam)->code == (UINT)PSN_APPLY)
    {
      //Execute
      GetDlgItemTextWide(hDlg, IDC_OPTIONS_EXECCOM, moCur.wszExecuteCommand, BUFFER_SIZE);
      GetDlgItemTextWide(hDlg, IDC_OPTIONS_EXECDIR, moCur.wszExecuteDirectory, MAX_PATH);

      //Autodetect codepage
      i=(int)SendMessage(hWndAutodetectCP, CB_GETCURSEL, 0, 0);

      if (i == DETECTINDEX_NONE)
        moCur.dwLangCodepageRecognition=0;
      else if (i == DETECTINDEX_RUSSIAN)
        moCur.dwLangCodepageRecognition=LANGID_RUSSIAN;
      else if (i == DETECTINDEX_EASTERNEUROPE)
      {
        if (!IsLangEasternEurope(moCur.dwLangCodepageRecognition))
          moCur.dwLangCodepageRecognition=LANGID_POLISH;
      }
      else if (i == DETECTINDEX_WESTERNEUROPE)
      {
        if (!IsLangWesternEurope(moCur.dwLangCodepageRecognition))
          moCur.dwLangCodepageRecognition=LANGID_GERMAN;
      }
      else if (i == DETECTINDEX_TURKISH)
        moCur.dwLangCodepageRecognition=LANGID_TURKISH;
      else if (i == DETECTINDEX_CHINESE)
        moCur.dwLangCodepageRecognition=LANGID_CHINESE;
      else if (i == DETECTINDEX_JAPANESE)
        moCur.dwLangCodepageRecognition=LANGID_JAPANESE;
      else if (i == DETECTINDEX_KOREAN)
        moCur.dwLangCodepageRecognition=LANGID_KOREAN;

      //Autodetect codepage buffer
      moCur.dwCodepageRecognitionBuffer=GetDlgItemInt(hDlg, IDC_OPTIONS_CODEPAGE_RECOGNITION_BUFFER, NULL, FALSE);

      //Default codepage
      moCur.nDefaultCodePage=GetComboboxCodepage(hWndDefaultCP);

      //New file
      moCur.nNewFileCodePage=GetComboboxCodepage(hWndNewFileCP);

      moCur.bNewFileBOM=(BOOL)SendMessage(hWndNewFileBOM, BM_GETCHECK, 0, 0);

      moCur.nNewFileNewLine=(int)SendMessage(hWndNewFileNewLine, CB_GETCURSEL, 0, 0) + 1;
    }
  }
  return FALSE;
}

int CALLBACK BrowseCallbackProc(HWND hWnd, UINT uMsg, LPARAM lParam, LPARAM lpData)
{
  if (uMsg == BFFM_INITIALIZED)
  {
    if (bOldWindows)
    {
      char szData[MAX_PATH];

      WideCharToMultiByte(CP_ACP, 0, (wchar_t *)lpData, -1, szData, MAX_PATH, NULL, NULL);
      SendMessage(hWnd, BFFM_SETSELECTIONA, TRUE, (LPARAM)szData);
    }
    else SendMessage(hWnd, BFFM_SETSELECTIONW, TRUE, lpData);
  }
  return 0;
}

BOOL CALLBACK OptionsGeneralFilterDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndCustomList;
  static HWND hWndSystemList;
  static HWND hWndOK;
  static int *lpFullCodepageList;
  int nSelection;
  int nCount;

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndCustomList=GetDlgItem(hDlg, IDC_OPTIONS_CODEPAGES_FILTER_CUSTOM);
    hWndSystemList=GetDlgItem(hDlg, IDC_OPTIONS_CODEPAGES_FILTER_SYSTEM);
    hWndOK=GetDlgItem(hDlg, IDOK);

    FillListBoxCodepage(hWndCustomList, lpCodepageList);
    EnumCodepageList(&lpFullCodepageList);
    FillListBoxCodepage(hWndSystemList, lpFullCodepageList);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_OPTIONS_CODEPAGES_FILTER_UP)
    {
      if ((nSelection=(int)SendMessage(hWndCustomList, LB_GETCURSEL, 0, 0)) > 0)
      {
        ListBox_GetTextWide(hWndCustomList, nSelection, wbuf);
        SendMessage(hWndCustomList, LB_DELETESTRING, nSelection, 0);
        ListBox_InsertStringWide(hWndCustomList, nSelection - 1, wbuf);
        SendMessage(hWndCustomList, LB_SETCURSEL, nSelection - 1, 0);
      }
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_CODEPAGES_FILTER_DOWN)
    {
      if ((nSelection=(int)SendMessage(hWndCustomList, LB_GETCURSEL, 0, 0)) != LB_ERR &&
          (nCount=(int)SendMessage(hWndCustomList, LB_GETCOUNT, 0, 0)) != LB_ERR &&
           nSelection != nCount - 1)
      {
        ListBox_GetTextWide(hWndCustomList, nSelection, wbuf);
        SendMessage(hWndCustomList, LB_DELETESTRING, nSelection, 0);
        ListBox_InsertStringWide(hWndCustomList, nSelection + 1, wbuf);
        SendMessage(hWndCustomList, LB_SETCURSEL, nSelection + 1, 0);
      }
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_CODEPAGES_FILTER_ADD)
    {
      if ((nSelection=(int)SendMessage(hWndSystemList, LB_GETCURSEL, 0, 0)) != LB_ERR)
      {
        ListBox_GetTextWide(hWndSystemList, nSelection, wbuf);

        if (ListBox_FindStringWide(hWndCustomList, -1, wbuf) == LB_ERR)
        {
          nSelection=(int)SendMessage(hWndCustomList, LB_GETCURSEL, 0, 0);
          ListBox_InsertStringWide(hWndCustomList, nSelection, wbuf);
          EnableWindow(hWndOK, TRUE);
        }
      }
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_CODEPAGES_FILTER_DELETE)
    {
      if ((nSelection=(int)SendMessage(hWndCustomList, LB_GETCURSEL, 0, 0)) != LB_ERR)
      {
        SendMessage(hWndCustomList, LB_DELETESTRING, nSelection, 0);

        if ((nCount=(int)SendMessage(hWndCustomList, LB_GETCOUNT, 0, 0)) > 0)
        {
          if (nSelection < nCount)
            SendMessage(hWndCustomList, LB_SETCURSEL, nSelection, 0);
          else
            SendMessage(hWndCustomList, LB_SETCURSEL, nCount - 1, 0);
        }
        else EnableWindow(hWndOK, FALSE);
      }
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_CODEPAGES_FILTER_ADDALL)
    {
      ClearListBox(hWndCustomList);
      FillListBoxCodepage(hWndCustomList, lpFullCodepageList);
      EnableWindow(hWndOK, TRUE);
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_CODEPAGES_FILTER_DELETEALL)
    {
      ClearListBox(hWndCustomList);
      EnableWindow(hWndOK, FALSE);
    }
    else if (LOWORD(wParam) == IDOK)
    {
      bMenuPopupCodepage=TRUE;
      CodepageListFree(&lpCodepageList);
      CodepageListFree(&lpFullCodepageList);
      GetListBoxCodepageList(hWndCustomList, &lpCodepageList);
      nCodepageListLen=CodepageListLen(lpCodepageList);
      EndDialog(hDlg, 1);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDCANCEL)
    {
      CodepageListFree(&lpFullCodepageList);
      EndDialog(hDlg, 0);
      return TRUE;
    }
  }
  return FALSE;
}

BOOL CALLBACK OptionsRegistryDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndSaveRegistry;
  static HWND hWndSaveINI;
  static HWND hWndSavePositions;
  static HWND hWndSaveCodepages;
  static HWND hWndRecentFiles;
  static HWND hWndRecentFilesSpin;
  static HWND hWndRecentFilesClear;
  static HWND hWndSearchStrings;
  static HWND hWndSearchStringsSpin;
  static HWND hWndSearchStringsClear;
  static HWND hWndFileTypesOpen;
  static HWND hWndAssociateOpen;
  static HWND hWndFileTypesEdit;
  static HWND hWndAssociateEdit;
  static HWND hWndFileTypesPrint;
  static HWND hWndAssociatePrint;
  BOOL bState;

  if (uMsg == WM_INITDIALOG)
  {
    hWndSaveRegistry=GetDlgItem(hDlg, IDC_OPTIONS_SAVESETTINGS_REGISTRY);
    hWndSaveINI=GetDlgItem(hDlg, IDC_OPTIONS_SAVESETTINGS_INI);
    hWndSavePositions=GetDlgItem(hDlg, IDC_OPTIONS_SAVEPOSITIONS);
    hWndSaveCodepages=GetDlgItem(hDlg, IDC_OPTIONS_SAVECODEPAGES);
    hWndRecentFiles=GetDlgItem(hDlg, IDC_OPTIONS_RECENTFILES_AMOUNT);
    hWndRecentFilesSpin=GetDlgItem(hDlg, IDC_OPTIONS_RECENTFILES_AMOUNT_SPIN);
    hWndRecentFilesClear=GetDlgItem(hDlg, IDC_OPTIONS_RECENTFILES_CLEAR);
    hWndSearchStrings=GetDlgItem(hDlg, IDC_OPTIONS_SEARCHSTRINGS_AMOUNT);
    hWndSearchStringsSpin=GetDlgItem(hDlg, IDC_OPTIONS_SEARCHSTRINGS_AMOUNT_SPIN);
    hWndSearchStringsClear=GetDlgItem(hDlg, IDC_OPTIONS_SEARCHSTRINGS_CLEAR);
    hWndFileTypesOpen=GetDlgItem(hDlg, IDC_OPTIONS_FILETYPES_OPEN);
    hWndAssociateOpen=GetDlgItem(hDlg, IDC_OPTIONS_ASSOCIATE_OPEN);
    hWndFileTypesEdit=GetDlgItem(hDlg, IDC_OPTIONS_FILETYPES_EDIT);
    hWndAssociateEdit=GetDlgItem(hDlg, IDC_OPTIONS_ASSOCIATE_EDIT);
    hWndFileTypesPrint=GetDlgItem(hDlg, IDC_OPTIONS_FILETYPES_PRINT);
    hWndAssociatePrint=GetDlgItem(hDlg, IDC_OPTIONS_ASSOCIATE_PRINT);

    SendMessage(hWndRecentFilesSpin, UDM_SETBUDDY, (WPARAM)hWndRecentFiles, 0);
    SendMessage(hWndRecentFilesSpin, UDM_SETRANGE, 0, MAKELONG(999, 0));
    SendMessage(hWndSearchStringsSpin, UDM_SETBUDDY, (WPARAM)hWndSearchStrings, 0);
    SendMessage(hWndSearchStringsSpin, UDM_SETRANGE, 0, MAKELONG(999, 0));

    if (moCur.dwFileTypesAssociated & FTA_OPEN) SendMessage(hWndAssociateOpen, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwFileTypesAssociated & FTA_EDIT) SendMessage(hWndAssociateEdit, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwFileTypesAssociated & FTA_PRINT) SendMessage(hWndAssociatePrint, BM_SETCHECK, BST_CHECKED, 0);
    if (!(moCur.dwFileTypesAssociated & FTA_OPEN)) EnableWindow(hWndFileTypesOpen, FALSE);
    if (!(moCur.dwFileTypesAssociated & FTA_EDIT)) EnableWindow(hWndFileTypesEdit, FALSE);
    if (!(moCur.dwFileTypesAssociated & FTA_PRINT)) EnableWindow(hWndFileTypesPrint, FALSE);
    SendMessage(hWndFileTypesOpen, EM_LIMITTEXT, (WPARAM)MAX_PATH, 0);
    SendMessage(hWndFileTypesEdit, EM_LIMITTEXT, (WPARAM)MAX_PATH, 0);
    SendMessage(hWndFileTypesPrint, EM_LIMITTEXT, (WPARAM)MAX_PATH, 0);

    SetDlgItemInt(hDlg, IDC_OPTIONS_RECENTFILES_AMOUNT, moCur.nRecentFiles, FALSE);
    if (!moCur.nRecentFiles || !hRecentFilesStack.first) EnableWindow(hWndRecentFilesClear, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_SEARCHSTRINGS_AMOUNT, moCur.nSearchStrings, FALSE);
    if (!moCur.nSearchStrings) EnableWindow(hWndSearchStringsClear, FALSE);

    SetWindowTextWide(hWndFileTypesOpen, moCur.wszFileTypesOpen);
    SetWindowTextWide(hWndFileTypesEdit, moCur.wszFileTypesEdit);
    SetWindowTextWide(hWndFileTypesPrint, moCur.wszFileTypesPrint);

    if (moCur.nSaveSettings == SS_REGISTRY)
      SendMessage(hWndSaveRegistry, BM_SETCHECK, BST_CHECKED, 0);
    else
      SendMessage(hWndSaveINI, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.bSavePositions)
      SendMessage(hWndSavePositions, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.bSaveCodepages)
      SendMessage(hWndSaveCodepages, BM_SETCHECK, BST_CHECKED, 0);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_OPTIONS_RECENTFILES_CLEAR)
    {
      RecentFilesZero(&hRecentFilesStack);
      RecentFilesSave(&hRecentFilesStack);
      bMenuRecentFiles=TRUE;
      SetFocus(hWndRecentFiles);
      EnableWindow((HWND)lParam, FALSE);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_SEARCHSTRINGS_CLEAR)
    {
      wchar_t wszRegKey[MAX_PATH];

      xprintfW(wszRegKey, L"%s\\Search", APP_REGHOMEW);
      RegClearKeyWide(HKEY_CURRENT_USER, wszRegKey);
      SetFocus(hWndSearchStrings);
      EnableWindow((HWND)lParam, FALSE);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_ASSOCIATE_OPEN)
    {
      bState=(BOOL)SendMessage(hWndAssociateOpen, BM_GETCHECK, 0, 0);
      EnableWindow(hWndFileTypesOpen, bState);
      if (!bState) SetWindowTextWide(hWndFileTypesOpen, moCur.wszFileTypesOpen);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_ASSOCIATE_EDIT)
    {
      bState=(BOOL)SendMessage(hWndAssociateEdit, BM_GETCHECK, 0, 0);
      EnableWindow(hWndFileTypesEdit, bState);
      if (!bState) SetWindowTextWide(hWndFileTypesEdit, moCur.wszFileTypesEdit);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_ASSOCIATE_PRINT)
    {
      bState=(BOOL)SendMessage(hWndAssociatePrint, BM_GETCHECK, 0, 0);
      EnableWindow(hWndFileTypesPrint, bState);
      if (!bState) SetWindowTextWide(hWndFileTypesPrint, moCur.wszFileTypesPrint);
      return TRUE;
    }
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (((NMHDR *)lParam)->code == (UINT)PSN_SETACTIVE)
    {
      nPropertyStartPage=(int)SendMessage(hPropertyTab, TCM_GETCURSEL, 0, 0);
    }
    else if (((NMHDR *)lParam)->code == (UINT)PSN_APPLY)
    {
      wchar_t wszWindowText[MAX_PATH];
      BOOL bShellRefresh=FALSE;
      int a;
      int b;

      //Save settings
      if (SendMessage(hWndSaveRegistry, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a=SS_REGISTRY;
      else
        a=SS_INI;
      if (moCur.nSaveSettings != a)
      {
        moCur.nSaveSettings=a;
        bOptionsSave=TRUE;
      }

      //Recent files
      a=GetDlgItemInt(hDlg, IDC_OPTIONS_RECENTFILES_AMOUNT, NULL, FALSE);
      if (moCur.nRecentFiles != a)
      {
        moCur.nRecentFiles=a;
        RecentFilesRefresh(&hRecentFilesStack);
      }
      a=(int)SendMessage(hWndSavePositions, BM_GETCHECK, 0, 0);
      b=(int)SendMessage(hWndSaveCodepages, BM_GETCHECK, 0, 0);
      if (a != moCur.bSavePositions || b != moCur.bSaveCodepages)
      {
        moCur.bSavePositions=a;
        moCur.bSaveCodepages=b;
      }

      //Search history
      moCur.nSearchStrings=GetDlgItemInt(hDlg, IDC_OPTIONS_SEARCHSTRINGS_AMOUNT, NULL, FALSE);

      //Associations
      bState=(BOOL)SendMessage(hWndAssociateOpen, BM_GETCHECK, 0, 0);
      if (bState && !(moCur.dwFileTypesAssociated & FTA_OPEN))
      {
        GetWindowTextWide(hWndFileTypesOpen, moCur.wszFileTypesOpen, MAX_PATH);
        AssociateFileTypesW(hInstance, moCur.wszFileTypesOpen, FTA_OPEN|FTA_ASSOCIATE);
        moCur.dwFileTypesAssociated|=FTA_OPEN;
        bShellRefresh=TRUE;
      }
      else if (bState && (moCur.dwFileTypesAssociated & FTA_OPEN))
      {
        GetWindowTextWide(hWndFileTypesOpen, wszWindowText, MAX_PATH);
        if (xstrcmpiW(wszWindowText, moCur.wszFileTypesOpen))
        {
          AssociateFileTypesW(hInstance, moCur.wszFileTypesOpen, FTA_OPEN|FTA_DEASSOCIATE);
          AssociateFileTypesW(hInstance, wszWindowText, FTA_OPEN|FTA_ASSOCIATE);
          xstrcpynW(moCur.wszFileTypesOpen, wszWindowText, MAX_PATH);
        }
        bShellRefresh=TRUE;
      }
      else if (!bState && (moCur.dwFileTypesAssociated & FTA_OPEN))
      {
        AssociateFileTypesW(hInstance, moCur.wszFileTypesOpen, FTA_OPEN|FTA_DEASSOCIATE);
        moCur.dwFileTypesAssociated&=~FTA_OPEN;
        bShellRefresh=TRUE;
      }

      bState=(BOOL)SendMessage(hWndAssociateEdit, BM_GETCHECK, 0, 0);
      if (bState && !(moCur.dwFileTypesAssociated & FTA_EDIT))
      {
        GetWindowTextWide(hWndFileTypesEdit, moCur.wszFileTypesEdit, MAX_PATH);
        AssociateFileTypesW(hInstance, moCur.wszFileTypesEdit, FTA_EDIT|FTA_ASSOCIATE);
        moCur.dwFileTypesAssociated|=FTA_EDIT;
        bShellRefresh=TRUE;
      }
      else if (bState && (moCur.dwFileTypesAssociated & FTA_EDIT))
      {
        GetWindowTextWide(hWndFileTypesEdit, wszWindowText, MAX_PATH);
        if (xstrcmpiW(wszWindowText, moCur.wszFileTypesEdit))
        {
          AssociateFileTypesW(hInstance, moCur.wszFileTypesEdit, FTA_EDIT|FTA_DEASSOCIATE);
          AssociateFileTypesW(hInstance, wszWindowText, FTA_EDIT|FTA_ASSOCIATE);
          xstrcpynW(moCur.wszFileTypesEdit, wszWindowText, MAX_PATH);
        }
        bShellRefresh=TRUE;
      }
      else if (!bState && (moCur.dwFileTypesAssociated & FTA_EDIT))
      {
        AssociateFileTypesW(hInstance, moCur.wszFileTypesEdit, FTA_EDIT|FTA_DEASSOCIATE);
        moCur.dwFileTypesAssociated&=~FTA_EDIT;
        bShellRefresh=TRUE;
      }

      bState=(BOOL)SendMessage(hWndAssociatePrint, BM_GETCHECK, 0, 0);
      if (bState && !(moCur.dwFileTypesAssociated & FTA_PRINT))
      {
        GetWindowTextWide(hWndFileTypesPrint, moCur.wszFileTypesPrint, MAX_PATH);
        AssociateFileTypesW(hInstance, moCur.wszFileTypesPrint, FTA_PRINT|FTA_ASSOCIATE);
        moCur.dwFileTypesAssociated|=FTA_PRINT;
        bShellRefresh=TRUE;
      }
      else if (bState && (moCur.dwFileTypesAssociated & FTA_PRINT))
      {
        GetWindowTextWide(hWndFileTypesPrint, wszWindowText, MAX_PATH);
        if (xstrcmpiW(wszWindowText, moCur.wszFileTypesPrint))
        {
          AssociateFileTypesW(hInstance, moCur.wszFileTypesPrint, FTA_PRINT|FTA_DEASSOCIATE);
          AssociateFileTypesW(hInstance, wszWindowText, FTA_PRINT|FTA_ASSOCIATE);
          xstrcpynW(moCur.wszFileTypesPrint, wszWindowText, MAX_PATH);
        }
        bShellRefresh=TRUE;
      }
      else if (!bState && (moCur.dwFileTypesAssociated & FTA_PRINT))
      {
        AssociateFileTypesW(hInstance, moCur.wszFileTypesPrint, FTA_PRINT|FTA_DEASSOCIATE);
        moCur.dwFileTypesAssociated&=~FTA_PRINT;
        bShellRefresh=TRUE;
      }
      if (bShellRefresh) SHChangeNotify(SHCNE_ASSOCCHANGED, 0, 0, 0);
    }
  }
  return FALSE;
}

BOOL CALLBACK OptionsEditor1DlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndTabSize;
  static HWND hWndTabSizeSpin;
  static HWND hWndTabSizeSpaces;
  static HWND hWndUndoLimit;
  static HWND hWndUndoLimitSpin;
  static HWND hWndDetailedUndo;
  static HWND hWndWrapByWords;
  static HWND hWndWrapBySymbols;
  static HWND hWndWrapLimit;
  static HWND hWndWrapLimitSpin;
  static HWND hWndMarker;
  static HWND hWndMarkerSpin;
  static HWND hWndCaretOutEdge;
  static HWND hWndCaretVertLine;
  static HWND hWndCaretActiveLine;
  static HWND hWndCaretActiveLineBorder;
  static HWND hWndCaretWidth;
  static HWND hWndCaretWidthSpin;
  static HWND hWndAltLineSkip;
  static HWND hWndAltLineSkipSpin;
  static HWND hWndAltLineFill;
  static HWND hWndAltLineFillSpin;
  static HWND hWndAltLineBorder;
  static HWND hWndMarginLeft;
  static HWND hWndMarginLeftSpin;
  static HWND hWndMarginRight;
  static HWND hWndMarginRightSpin;
  static HWND hWndMarginTop;
  static HWND hWndMarginTopSpin;
  static HWND hWndMarginBottom;
  static HWND hWndMarginBottomSpin;
  static HWND hWndLineGap;
  static HWND hWndLineGapSpin;
  BOOL bState;

  if (uMsg == WM_INITDIALOG)
  {
    hWndTabSize=GetDlgItem(hDlg, IDC_OPTIONS_TABSIZE);
    hWndTabSizeSpin=GetDlgItem(hDlg, IDC_OPTIONS_TABSIZE_SPIN);
    hWndTabSizeSpaces=GetDlgItem(hDlg, IDC_OPTIONS_TABSIZE_SPACES);
    hWndUndoLimit=GetDlgItem(hDlg, IDC_OPTIONS_UNDO_LIMIT);
    hWndUndoLimitSpin=GetDlgItem(hDlg, IDC_OPTIONS_UNDO_LIMIT_SPIN);
    hWndDetailedUndo=GetDlgItem(hDlg, IDC_OPTIONS_UNDO_DETAILED);
    hWndWrapByWords=GetDlgItem(hDlg, IDC_OPTIONS_WRAP_BY_WORDS);
    hWndWrapBySymbols=GetDlgItem(hDlg, IDC_OPTIONS_WRAP_BY_SYMBOLS);
    hWndWrapLimit=GetDlgItem(hDlg, IDC_OPTIONS_WRAP_LIMIT);
    hWndWrapLimitSpin=GetDlgItem(hDlg, IDC_OPTIONS_WRAP_LIMIT_SPIN);
    hWndMarker=GetDlgItem(hDlg, IDC_OPTIONS_MARKER);
    hWndMarkerSpin=GetDlgItem(hDlg, IDC_OPTIONS_MARKER_SPIN);
    hWndCaretOutEdge=GetDlgItem(hDlg, IDC_OPTIONS_CARETOUTEDGE);
    hWndCaretVertLine=GetDlgItem(hDlg, IDC_OPTIONS_CARETVERTLINE);
    hWndCaretActiveLine=GetDlgItem(hDlg, IDC_OPTIONS_CARETACTIVELINE);
    hWndCaretActiveLineBorder=GetDlgItem(hDlg, IDC_OPTIONS_CARETACTIVELINEBORDER);
    hWndCaretWidth=GetDlgItem(hDlg, IDC_OPTIONS_CARETWIDTH);
    hWndCaretWidthSpin=GetDlgItem(hDlg, IDC_OPTIONS_CARETWIDTH_SPIN);
    hWndAltLineSkip=GetDlgItem(hDlg, IDC_OPTIONS_ALTLINESKIP);
    hWndAltLineSkipSpin=GetDlgItem(hDlg, IDC_OPTIONS_ALTLINESKIP_SPIN);
    hWndAltLineFill=GetDlgItem(hDlg, IDC_OPTIONS_ALTLINEFILL);
    hWndAltLineFillSpin=GetDlgItem(hDlg, IDC_OPTIONS_ALTLINEFILL_SPIN);
    hWndAltLineBorder=GetDlgItem(hDlg, IDC_OPTIONS_ALTLINEBORDER);
    hWndMarginLeft=GetDlgItem(hDlg, IDC_OPTIONS_EDITMARGIN_LEFT);
    hWndMarginLeftSpin=GetDlgItem(hDlg, IDC_OPTIONS_EDITMARGIN_LEFT_SPIN);
    hWndMarginRight=GetDlgItem(hDlg, IDC_OPTIONS_EDITMARGIN_RIGHT);
    hWndMarginRightSpin=GetDlgItem(hDlg, IDC_OPTIONS_EDITMARGIN_RIGHT_SPIN);
    hWndMarginTop=GetDlgItem(hDlg, IDC_OPTIONS_EDITMARGIN_TOP);
    hWndMarginTopSpin=GetDlgItem(hDlg, IDC_OPTIONS_EDITMARGIN_TOP_SPIN);
    hWndMarginBottom=GetDlgItem(hDlg, IDC_OPTIONS_EDITMARGIN_BOTTOM);
    hWndMarginBottomSpin=GetDlgItem(hDlg, IDC_OPTIONS_EDITMARGIN_BOTTOM_SPIN);
    hWndLineGap=GetDlgItem(hDlg, IDC_OPTIONS_LINEGAP);
    hWndLineGapSpin=GetDlgItem(hDlg, IDC_OPTIONS_LINEGAP_SPIN);

    SendMessage(hWndTabSizeSpin, UDM_SETBUDDY, (WPARAM)hWndTabSize, 0);
    SendMessage(hWndTabSizeSpin, UDM_SETRANGE, 0, MAKELONG(100, 1));
    SendMessage(hWndUndoLimitSpin, UDM_SETBUDDY, (WPARAM)hWndUndoLimit, 0);
    SendMessage(hWndUndoLimitSpin, UDM_SETRANGE, 0, MAKELONG(9999, 0));
    SendMessage(hWndWrapLimitSpin, UDM_SETBUDDY, (WPARAM)hWndWrapLimit, 0);
    SendMessage(hWndWrapLimitSpin, UDM_SETRANGE32, (WPARAM)-1, 9999);
    SendMessage(hWndMarkerSpin, UDM_SETBUDDY, (WPARAM)hWndMarker, 0);
    SendMessage(hWndMarkerSpin, UDM_SETRANGE32, (WPARAM)-1, 9999);
    SendMessage(hWndCaretWidthSpin, UDM_SETBUDDY, (WPARAM)hWndCaretWidth, 0);
    SendMessage(hWndCaretWidthSpin, UDM_SETRANGE, 0, MAKELONG(100, 1));
    SendMessage(hWndAltLineSkipSpin, UDM_SETBUDDY, (WPARAM)hWndAltLineSkip, 0);
    SendMessage(hWndAltLineSkipSpin, UDM_SETRANGE, 0, MAKELONG(999, 0));
    SendMessage(hWndAltLineFillSpin, UDM_SETBUDDY, (WPARAM)hWndAltLineFill, 0);
    SendMessage(hWndAltLineFillSpin, UDM_SETRANGE, 0, MAKELONG(999, 0));
    SendMessage(hWndMarginLeftSpin, UDM_SETBUDDY, (WPARAM)hWndMarginLeft, 0);
    SendMessage(hWndMarginLeftSpin, UDM_SETRANGE, 0, MAKELONG(999, 0));
    SendMessage(hWndMarginRightSpin, UDM_SETBUDDY, (WPARAM)hWndMarginRight, 0);
    SendMessage(hWndMarginRightSpin, UDM_SETRANGE, 0, MAKELONG(999, 0));
    SendMessage(hWndMarginTopSpin, UDM_SETBUDDY, (WPARAM)hWndMarginTop, 0);
    SendMessage(hWndMarginTopSpin, UDM_SETRANGE, 0, MAKELONG(999, 0));
    SendMessage(hWndMarginBottomSpin, UDM_SETBUDDY, (WPARAM)hWndMarginBottom, 0);
    SendMessage(hWndMarginBottomSpin, UDM_SETRANGE, 0, MAKELONG(999, 0));
    SendMessage(hWndLineGapSpin, UDM_SETBUDDY, (WPARAM)hWndLineGap, 0);
    SendMessage(hWndLineGapSpin, UDM_SETRANGE, 0, MAKELONG(100, 0));

    SetDlgItemInt(hDlg, IDC_OPTIONS_TABSIZE, lpFrameCurrent->nTabStopSize, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_UNDO_LIMIT, lpFrameCurrent->nUndoLimit, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_WRAP_LIMIT, lpFrameCurrent->dwWrapLimit, TRUE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_MARKER, lpFrameCurrent->dwMarker, TRUE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_CARETWIDTH, lpFrameCurrent->nCaretWidth, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_ALTLINESKIP, lpFrameCurrent->dwAltLineSkip, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_ALTLINEFILL, lpFrameCurrent->dwAltLineFill, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_EDITMARGIN_LEFT, lpFrameCurrent->rcEditMargins.left, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_EDITMARGIN_RIGHT, lpFrameCurrent->rcEditMargins.right, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_EDITMARGIN_TOP, lpFrameCurrent->rcEditMargins.top, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_EDITMARGIN_BOTTOM, lpFrameCurrent->rcEditMargins.bottom, FALSE);
    SetDlgItemInt(hDlg, IDC_OPTIONS_LINEGAP, lpFrameCurrent->dwLineGap, FALSE);

    if (lpFrameCurrent->bTabStopAsSpaces)
      SendMessage(hWndTabSizeSpaces, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->bDetailedUndo)
      SendMessage(hWndDetailedUndo, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->dwWrapType & AEWW_WORD)
      SendMessage(hWndWrapByWords, BM_SETCHECK, BST_CHECKED, 0);
    else if (lpFrameCurrent->dwWrapType & AEWW_SYMBOL)
      SendMessage(hWndWrapBySymbols, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->dwCaretOptions & CO_CARETOUTEDGE)
      SendMessage(hWndCaretOutEdge, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->dwCaretOptions & CO_CARETVERTLINE)
      SendMessage(hWndCaretVertLine, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->dwCaretOptions & CO_CARETACTIVELINE)
      SendMessage(hWndCaretActiveLine, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->dwCaretOptions & CO_CARETACTIVELINEBORDER)
      SendMessage(hWndCaretActiveLineBorder, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->bAltLineBorder)
      SendMessage(hWndAltLineBorder, BM_SETCHECK, BST_CHECKED, 0);

    PostMessage(hDlg, WM_COMMAND, IDC_OPTIONS_CARETACTIVELINE, 0);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_OPTIONS_CARETACTIVELINE)
    {
      bState=(BOOL)SendMessage(hWndCaretActiveLine, BM_GETCHECK, 0, 0);
      EnableWindow(hWndCaretActiveLineBorder, bState);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_ALTLINESKIP ||
             LOWORD(wParam) == IDC_OPTIONS_ALTLINEFILL)
    {
      if (GetDlgItemInt(hDlg, IDC_OPTIONS_ALTLINESKIP, NULL, FALSE) &&
          GetDlgItemInt(hDlg, IDC_OPTIONS_ALTLINEFILL, NULL, FALSE))
      {
        bState=TRUE;
      }
      else bState=FALSE;

      EnableWindow(hWndAltLineBorder, bState);
      return TRUE;
    }
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (((NMHDR *)lParam)->code == (UINT)PSN_SETACTIVE)
    {
      nPropertyStartPage=(int)SendMessage(hPropertyTab, TCM_GETCURSEL, 0, 0);
    }
    else if (((NMHDR *)lParam)->code == (UINT)PSN_APPLY)
    {
      RECT rcEditMargins;
      int a;
      int b;

      //Tab stops
      a=GetDlgItemInt(hDlg, IDC_OPTIONS_TABSIZE, NULL, FALSE);
      SetFrameInfo(lpFrameCurrent, FIS_TABSTOPSIZE, a);
      a=(int)SendMessage(hWndTabSizeSpaces, BM_GETCHECK, 0, 0);
      SetFrameInfo(lpFrameCurrent, FIS_TABSTOPASSPACES, a);

      //Undo
      a=GetDlgItemInt(hDlg, IDC_OPTIONS_UNDO_LIMIT, NULL, FALSE);
      SetFrameInfo(lpFrameCurrent, FIS_UNDOLIMIT, a);
      a=(BOOL)SendMessage(hWndDetailedUndo, BM_GETCHECK, 0, 0);
      SetFrameInfo(lpFrameCurrent, FIS_DETAILEDUNDO, a);

      //Wrap limit
      if (SendMessage(hWndWrapByWords, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a=AEWW_WORD;
      else
        a=AEWW_SYMBOL;
      b=GetDlgItemInt(hDlg, IDC_OPTIONS_WRAP_LIMIT, NULL, TRUE);
      SetFrameInfo(lpFrameCurrent, FIS_WRAP, MAKELONG(a, b));

      //Marker
      a=GetDlgItemInt(hDlg, IDC_OPTIONS_MARKER, NULL, TRUE);
      SetFrameInfo(lpFrameCurrent, FIS_MARKER, a);

      //Caret options
      a=0;
      if (SendMessage(hWndCaretOutEdge, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a|=CO_CARETOUTEDGE;
      if (SendMessage(hWndCaretVertLine, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a|=CO_CARETVERTLINE;
      if (SendMessage(hWndCaretActiveLine, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a|=CO_CARETACTIVELINE;
      if (SendMessage(hWndCaretActiveLineBorder, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a|=CO_CARETACTIVELINEBORDER;
      SetFrameInfo(lpFrameCurrent, FIS_CARETOPTIONS, a);

      //Caret width
      a=GetDlgItemInt(hDlg, IDC_OPTIONS_CARETWIDTH, NULL, FALSE);
      SetFrameInfo(lpFrameCurrent, FIS_CARETWIDTH, a);

      //Alternating lines
      a=GetDlgItemInt(hDlg, IDC_OPTIONS_ALTLINESKIP, NULL, FALSE);
      b=GetDlgItemInt(hDlg, IDC_OPTIONS_ALTLINEFILL, NULL, FALSE);
      SetFrameInfo(lpFrameCurrent, FIS_ALTLINES, MAKELONG(a, b));

      //Alternating line border
      a=(BOOL)SendMessage(hWndAltLineBorder, BM_GETCHECK, 0, 0);
      SetFrameInfo(lpFrameCurrent, FIS_ALTLINEBORDER, a);

      //Margins
      rcEditMargins.left=GetDlgItemInt(hDlg, IDC_OPTIONS_EDITMARGIN_LEFT, NULL, FALSE);
      rcEditMargins.right=GetDlgItemInt(hDlg, IDC_OPTIONS_EDITMARGIN_RIGHT, NULL, FALSE);
      rcEditMargins.top=GetDlgItemInt(hDlg, IDC_OPTIONS_EDITMARGIN_TOP, NULL, FALSE);
      rcEditMargins.bottom=GetDlgItemInt(hDlg, IDC_OPTIONS_EDITMARGIN_BOTTOM, NULL, FALSE);
      SetFrameInfo(lpFrameCurrent, FIS_RECTMARGINS, (UINT_PTR)&rcEditMargins);

      //Line gap
      a=GetDlgItemInt(hDlg, IDC_OPTIONS_LINEGAP, NULL, FALSE);
      SetFrameInfo(lpFrameCurrent, FIS_LINEGAP, a);
    }
  }
  return FALSE;
}

BOOL CALLBACK OptionsEditor2DlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndMarginSelection;
  static HWND hWndMouseRichEdit;
  static HWND hWndMouseDragging;
  static HWND hWndRClickMoveCaret;
  static HWND hWndShowURL;
  static HWND hWndSingleClickURL;
  static HWND hWndDoubleClickURL;
  static HWND hWndUrlPrefixesEnable;
  static HWND hWndUrlPrefixes;
  static HWND hWndUrlDelimitersEnable;
  static HWND hWndUrlLeftDelimiters;
  static HWND hWndUrlRightDelimiters;
  static HWND hWndWordDelimitersEnable;
  static HWND hWndWordDelimiters;
  static HWND hWndWrapDelimitersEnable;
  static HWND hWndWrapDelimiters;
  BOOL bState;

  if (uMsg == WM_INITDIALOG)
  {
    hWndMarginSelection=GetDlgItem(hDlg, IDC_OPTIONS_MOUSELEFTMARGIN);
    hWndMouseRichEdit=GetDlgItem(hDlg, IDC_OPTIONS_MOUSERICHEDIT);
    hWndMouseDragging=GetDlgItem(hDlg, IDC_OPTIONS_MOUSEDRAGGING);
    hWndRClickMoveCaret=GetDlgItem(hDlg, IDC_OPTIONS_RCLICKMOVECARET);
    hWndShowURL=GetDlgItem(hDlg, IDC_OPTIONS_URL_SHOW);
    hWndSingleClickURL=GetDlgItem(hDlg, IDC_OPTIONS_URL_SINGLECLICK);
    hWndDoubleClickURL=GetDlgItem(hDlg, IDC_OPTIONS_URL_DOUBLECLICK);
    hWndUrlPrefixesEnable=GetDlgItem(hDlg, IDC_OPTIONS_URL_PREFIXES_ENABLE);
    hWndUrlPrefixes=GetDlgItem(hDlg, IDC_OPTIONS_URL_PREFIXES);
    hWndUrlDelimitersEnable=GetDlgItem(hDlg, IDC_OPTIONS_URL_DELIMITERS_ENABLE);
    hWndUrlLeftDelimiters=GetDlgItem(hDlg, IDC_OPTIONS_URL_LEFTDELIMITERS);
    hWndUrlRightDelimiters=GetDlgItem(hDlg, IDC_OPTIONS_URL_RIGHTDELIMITERS);
    hWndWordDelimitersEnable=GetDlgItem(hDlg, IDC_OPTIONS_WORD_DELIMITERS_ENABLE);
    hWndWordDelimiters=GetDlgItem(hDlg, IDC_OPTIONS_WORD_DELIMITERS);
    hWndWrapDelimitersEnable=GetDlgItem(hDlg, IDC_OPTIONS_WRAP_DELIMITERS_ENABLE);
    hWndWrapDelimiters=GetDlgItem(hDlg, IDC_OPTIONS_WRAP_DELIMITERS);

    if (lpFrameCurrent->dwMouseOptions & MO_LEFTMARGINSELECTION)
      SendMessage(hWndMarginSelection, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->dwMouseOptions & MO_RICHEDITMOUSE)
      SendMessage(hWndMouseRichEdit, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->dwMouseOptions & MO_MOUSEDRAGGING)
      SendMessage(hWndMouseDragging, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->dwMouseOptions & MO_RCLICKMOVECARET)
      SendMessage(hWndRClickMoveCaret, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->bShowURL)
      SendMessage(hWndShowURL, BM_SETCHECK, BST_CHECKED, 0);
    if (lpFrameCurrent->nClickURL == 1)
      SendMessage(hWndSingleClickURL, BM_SETCHECK, BST_CHECKED, 0);
    else
      SendMessage(hWndDoubleClickURL, BM_SETCHECK, BST_CHECKED, 0);

    if (lpFrameCurrent->bUrlPrefixesEnable)
      SendMessage(hWndUrlPrefixesEnable, BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(hWndUrlPrefixes, lpFrameCurrent->bUrlPrefixesEnable);
    SetWindowTextWide(hWndUrlPrefixes, lpFrameCurrent->wszUrlPrefixes);
    SendMessage(hWndUrlPrefixes, EM_LIMITTEXT, (WPARAM)URL_PREFIXES_SIZE, 0);

    if (lpFrameCurrent->bUrlDelimitersEnable)
      SendMessage(hWndUrlDelimitersEnable, BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(hWndUrlLeftDelimiters, lpFrameCurrent->bUrlDelimitersEnable);
    EscapeDataToEscapeStringW(lpFrameCurrent->wszUrlLeftDelimiters, wbuf);
    SetWindowTextWide(hWndUrlLeftDelimiters, wbuf);
    SendMessage(hWndUrlLeftDelimiters, EM_LIMITTEXT, (WPARAM)URL_DELIMITERS_SIZE, 0);

    EnableWindow(hWndUrlRightDelimiters, lpFrameCurrent->bUrlDelimitersEnable);
    EscapeDataToEscapeStringW(lpFrameCurrent->wszUrlRightDelimiters, wbuf);
    SetWindowTextWide(hWndUrlRightDelimiters, wbuf);
    SendMessage(hWndUrlRightDelimiters, EM_LIMITTEXT, (WPARAM)URL_DELIMITERS_SIZE, 0);

    if (lpFrameCurrent->bWordDelimitersEnable)
      SendMessage(hWndWordDelimitersEnable, BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(hWndWordDelimiters, lpFrameCurrent->bWordDelimitersEnable);
    EscapeDataToEscapeStringW(lpFrameCurrent->wszWordDelimiters, wbuf);
    SetWindowTextWide(hWndWordDelimiters, wbuf);
    SendMessage(hWndWordDelimiters, EM_LIMITTEXT, (WPARAM)WORD_DELIMITERS_SIZE, 0);

    if (lpFrameCurrent->bWrapDelimitersEnable)
      SendMessage(hWndWrapDelimitersEnable, BM_SETCHECK, BST_CHECKED, 0);
    EnableWindow(hWndWrapDelimiters, lpFrameCurrent->bWrapDelimitersEnable);
    EscapeDataToEscapeStringW(lpFrameCurrent->wszWrapDelimiters, wbuf);
    SetWindowTextWide(hWndWrapDelimiters, wbuf);
    SendMessage(hWndWrapDelimiters, EM_LIMITTEXT, (WPARAM)WRAP_DELIMITERS_SIZE, 0);

    PostMessage(hDlg, WM_COMMAND, IDC_OPTIONS_URL_SHOW, 0);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_OPTIONS_URL_SHOW)
    {
      bState=(BOOL)SendMessage(hWndShowURL, BM_GETCHECK, 0, 0);
      EnableWindow(hWndSingleClickURL, bState);
      EnableWindow(hWndDoubleClickURL, bState);
      EnableWindow(hWndUrlPrefixesEnable, bState);
      EnableWindow(hWndUrlDelimitersEnable, bState);

      if (!bState)
      {
        EnableWindow(hWndUrlPrefixes, FALSE);
        EnableWindow(hWndUrlLeftDelimiters, FALSE);
        EnableWindow(hWndUrlRightDelimiters, FALSE);
      }
      else
      {
        bState=(BOOL)SendMessage(hWndUrlPrefixesEnable, BM_GETCHECK, 0, 0);
        EnableWindow(hWndUrlPrefixes, bState);
        bState=(BOOL)SendMessage(hWndUrlDelimitersEnable, BM_GETCHECK, 0, 0);
        EnableWindow(hWndUrlLeftDelimiters, bState);
        EnableWindow(hWndUrlRightDelimiters, bState);
      }
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_URL_PREFIXES_ENABLE)
    {
      bState=(BOOL)SendMessage(hWndUrlPrefixesEnable, BM_GETCHECK, 0, 0);
      EnableWindow(hWndUrlPrefixes, bState);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_URL_DELIMITERS_ENABLE)
    {
      bState=(BOOL)SendMessage(hWndUrlDelimitersEnable, BM_GETCHECK, 0, 0);
      EnableWindow(hWndUrlLeftDelimiters, bState);
      EnableWindow(hWndUrlRightDelimiters, bState);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_WORD_DELIMITERS_ENABLE)
    {
      bState=(BOOL)SendMessage(hWndWordDelimitersEnable, BM_GETCHECK, 0, 0);
      EnableWindow(hWndWordDelimiters, bState);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_WORD_DELIMITERS_RESET)
    {
      EscapeDataToEscapeStringW(STR_WORD_DELIMITERSW, wbuf);
      SetWindowTextWide(hWndWordDelimiters, wbuf);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_WRAP_DELIMITERS_ENABLE)
    {
      bState=(BOOL)SendMessage(hWndWrapDelimitersEnable, BM_GETCHECK, 0, 0);
      EnableWindow(hWndWrapDelimiters, bState);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_OPTIONS_WRAP_DELIMITERS_RESET)
    {
      EscapeDataToEscapeStringW(STR_WRAP_DELIMITERSW, wbuf);
      SetWindowTextWide(hWndWrapDelimiters, wbuf);
      return TRUE;
    }
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (((NMHDR *)lParam)->code == (UINT)PSN_SETACTIVE)
    {
      nPropertyStartPage=(int)SendMessage(hPropertyTab, TCM_GETCURSEL, 0, 0);
    }
    else if (((NMHDR *)lParam)->code == (UINT)PSN_APPLY)
    {
      int a;

      //Mouse options
      a=lpFrameCurrent->dwMouseOptions & ~MO_LEFTMARGINSELECTION & ~MO_RICHEDITMOUSE & ~MO_MOUSEDRAGGING & ~MO_RCLICKMOVECARET;
      if (SendMessage(hWndMarginSelection, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a|=MO_LEFTMARGINSELECTION;
      if (SendMessage(hWndMouseRichEdit, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a|=MO_RICHEDITMOUSE;
      if (SendMessage(hWndMouseDragging, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a|=MO_MOUSEDRAGGING;
      if (SendMessage(hWndRClickMoveCaret, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a|=MO_RCLICKMOVECARET;
      SetFrameInfo(lpFrameCurrent, FIS_MOUSEOPTIONS, a);

      //HyperLinks
      a=(int)SendMessage(hWndShowURL, BM_GETCHECK, 0, 0);
      SetFrameInfo(lpFrameCurrent, FIS_SHOWURL, a);

      if (SendMessage(hWndSingleClickURL, BM_GETCHECK, 0, 0) == BST_CHECKED)
        a=1;
      else
        a=2;
      SetFrameInfo(lpFrameCurrent, FIS_CLICKURL, a);

      //Url prefixes
      GetWindowTextWide(hWndUrlPrefixes, wbuf, URL_PREFIXES_SIZE);
      SetFrameInfo(lpFrameCurrent, FIS_URLPREFIXES, (UINT_PTR)wbuf);

      a=(BOOL)SendMessage(hWndUrlPrefixesEnable, BM_GETCHECK, 0, 0);
      SetFrameInfo(lpFrameCurrent, FIS_URLPREFIXESENABLE, a);

      //Url delimiters
      GetWindowTextWide(hWndUrlLeftDelimiters, wbuf, BUFFER_SIZE);
      EscapeStringToEscapeDataW(wbuf, -1, wbuf2, NEWLINE_UNIX);
      SetFrameInfo(lpFrameCurrent, FIS_URLLEFTDELIMITERS, (UINT_PTR)wbuf2);

      GetWindowTextWide(hWndUrlRightDelimiters, wbuf, BUFFER_SIZE);
      EscapeStringToEscapeDataW(wbuf, -1, wbuf2, NEWLINE_UNIX);
      SetFrameInfo(lpFrameCurrent, FIS_URLRIGHTDELIMITERS, (UINT_PTR)wbuf2);

      a=(BOOL)SendMessage(hWndUrlDelimitersEnable, BM_GETCHECK, 0, 0);
      SetFrameInfo(lpFrameCurrent, FIS_URLDELIMITERSENABLE, a);

      //Word delimiters
      GetWindowTextWide(hWndWordDelimiters, wbuf, BUFFER_SIZE);
      EscapeStringToEscapeDataW(wbuf, -1, wbuf2, NEWLINE_UNIX);
      SetFrameInfo(lpFrameCurrent, FIS_WORDDELIMITERS, (UINT_PTR)wbuf2);

      a=(BOOL)SendMessage(hWndWordDelimitersEnable, BM_GETCHECK, 0, 0);
      SetFrameInfo(lpFrameCurrent, FIS_WORDDELIMITERSENABLE, a);

      //Wrap delimiters
      GetWindowTextWide(hWndWrapDelimiters, wbuf, BUFFER_SIZE);
      EscapeStringToEscapeDataW(wbuf, -1, wbuf2, NEWLINE_UNIX);
      SetFrameInfo(lpFrameCurrent, FIS_WRAPDELIMITERS, (UINT_PTR)wbuf2);

      a=(BOOL)SendMessage(hWndWrapDelimitersEnable, BM_GETCHECK, 0, 0);
      SetFrameInfo(lpFrameCurrent, FIS_WRAPDELIMITERSENABLE, a);
    }
  }
  return FALSE;
}

BOOL CALLBACK OptionsAdvancedDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndDefaultSaveExt;
  static HWND hWndSwitchKeybLayout;
  static HWND hWndRememberKeybLayout;
  static HWND hWndSilentCloseEmpty;
  static HWND hWndDateLog;
  static HWND hWndSaveInReadOnlyMsg;
  static HWND hWndReplaceAllAndClose;
  static HWND hWndReplaceAllNoMsg;
  static HWND hWndInSelIfSel;
  static HWND hWndCycleSearch;
  static HWND hWndCycleSearchPrompt;
  BOOL bState;

  if (uMsg == WM_INITDIALOG)
  {
    hWndDefaultSaveExt=GetDlgItem(hDlg, IDC_OPTIONS_DEFAULT_SAVE_EXT);
    hWndSwitchKeybLayout=GetDlgItem(hDlg, IDC_OPTIONS_SWITCH_KEYBLAYOUT);
    hWndRememberKeybLayout=GetDlgItem(hDlg, IDC_OPTIONS_REMEMBER_KEYBLAYOUT);
    hWndSilentCloseEmpty=GetDlgItem(hDlg, IDC_OPTIONS_SILENTCLOSEEMPTY);
    hWndDateLog=GetDlgItem(hDlg, IDC_OPTIONS_LOGDATE);
    hWndSaveInReadOnlyMsg=GetDlgItem(hDlg, IDC_OPTIONS_SAVEIN_READONLY_MSG);
    hWndReplaceAllAndClose=GetDlgItem(hDlg, IDC_OPTIONS_REPLACEALL_CLOSE);
    hWndReplaceAllNoMsg=GetDlgItem(hDlg, IDC_OPTIONS_REPLACEALL_NOMSG);
    hWndInSelIfSel=GetDlgItem(hDlg, IDC_OPTIONS_INSELIFSEL);
    hWndCycleSearch=GetDlgItem(hDlg, IDC_OPTIONS_CYCLESEARCH);
    hWndCycleSearchPrompt=GetDlgItem(hDlg, IDC_OPTIONS_CYCLESEARCHPROMPT);

    if (moCur.dwKeybLayoutOptions & KLO_SWITCHLAYOUT)
      SendMessage(hWndSwitchKeybLayout, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwKeybLayoutOptions & KLO_REMEMBERLAYOUT)
      SendMessage(hWndRememberKeybLayout, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.bSilentCloseEmptyMDI)
      SendMessage(hWndSilentCloseEmpty, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.bDateLog)
      SendMessage(hWndDateLog, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.bSaveInReadOnlyMsg)
      SendMessage(hWndSaveInReadOnlyMsg, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwSearchOptions & FRF_REPLACEALLANDCLOSE)
      SendMessage(hWndReplaceAllAndClose, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwSearchOptions & FRF_REPLACEALLNOMSG)
      SendMessage(hWndReplaceAllNoMsg, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwSearchOptions & FRF_CHECKINSELIFSEL)
      SendMessage(hWndInSelIfSel, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwSearchOptions & FRF_CYCLESEARCH)
      SendMessage(hWndCycleSearch, BM_SETCHECK, BST_CHECKED, 0);
    if (moCur.dwSearchOptions & FRF_CYCLESEARCHPROMPT)
      SendMessage(hWndCycleSearchPrompt, BM_SETCHECK, BST_CHECKED, 0);

    SetWindowTextWide(hWndDefaultSaveExt, moCur.wszDefaultSaveExt);

    SendMessage(hDlg, WM_COMMAND, IDC_OPTIONS_CYCLESEARCH, 0);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_OPTIONS_CYCLESEARCH)
    {
      bState=(BOOL)SendMessage(hWndCycleSearch, BM_GETCHECK, 0, 0);
      EnableWindow(hWndCycleSearchPrompt, bState);
    }
  }
  else if (uMsg == WM_NOTIFY)
  {
    if (((NMHDR *)lParam)->code == (UINT)PSN_SETACTIVE)
    {
      nPropertyStartPage=(int)SendMessage(hPropertyTab, TCM_GETCURSEL, 0, 0);
    }
    else if (((NMHDR *)lParam)->code == (UINT)PSN_APPLY)
    {
      //Default save extention
      GetWindowTextWide(hWndDefaultSaveExt, moCur.wszDefaultSaveExt, MAX_PATH);

      //Auto switch keyboard layout
      if (SendMessage(hWndSwitchKeybLayout, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwKeybLayoutOptions|=KLO_SWITCHLAYOUT;
      else
        moCur.dwKeybLayoutOptions&=~KLO_SWITCHLAYOUT;

      //Remember keyboard layout for each tab (MDI)
      if (SendMessage(hWndRememberKeybLayout, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwKeybLayoutOptions|=KLO_REMEMBERLAYOUT;
      else
        moCur.dwKeybLayoutOptions&=~KLO_REMEMBERLAYOUT;

      //Silently close unsaved empty tab (MDI)
      moCur.bSilentCloseEmptyMDI=(BOOL)SendMessage(hWndSilentCloseEmpty, BM_GETCHECK, 0, 0);

      //.LOG feature
      moCur.bDateLog=(BOOL)SendMessage(hWndDateLog, BM_GETCHECK, 0, 0);

      //Save in read only file message
      moCur.bSaveInReadOnlyMsg=(BOOL)SendMessage(hWndSaveInReadOnlyMsg, BM_GETCHECK, 0, 0);

      //"Replace all" closes dialog
      if (SendMessage(hWndReplaceAllAndClose, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwSearchOptions|=FRF_REPLACEALLANDCLOSE;
      else
        moCur.dwSearchOptions&=~FRF_REPLACEALLANDCLOSE;

      //"Replace all" without message
      if (SendMessage(hWndReplaceAllNoMsg, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwSearchOptions|=FRF_REPLACEALLNOMSG;
      else
        moCur.dwSearchOptions&=~FRF_REPLACEALLNOMSG;

      //Check "In selection" if selection not empty
      if (SendMessage(hWndInSelIfSel, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwSearchOptions|=FRF_CHECKINSELIFSEL;
      else
        moCur.dwSearchOptions&=~FRF_CHECKINSELIFSEL;

      //Cycle search
      if (SendMessage(hWndCycleSearch, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwSearchOptions|=FRF_CYCLESEARCH;
      else
        moCur.dwSearchOptions&=~FRF_CYCLESEARCH;

      //Cycle search prompt
      if (SendMessage(hWndCycleSearchPrompt, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwSearchOptions|=FRF_CYCLESEARCHPROMPT;
      else
        moCur.dwSearchOptions&=~FRF_CYCLESEARCHPROMPT;
    }
  }
  return FALSE;
}


//// MDI list

BOOL CALLBACK MdiListDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static HWND hWndList;
  static HWND hWndStats;
  static HWND hWndSearch;
  static HWND hWndOK;
  static HWND hWndArrangeGroup;
  static HWND hWndHorz;
  static HWND hWndVert;
  static HWND hWndTabsGroup;
  static HWND hWndUp;
  static HWND hWndDown;
  static HWND hWndSort;
  static HWND hWndFilesGroup;
  static HWND hWndSave;
  static HWND hWndClose;
  static HWND hWndNames;
  static HWND hWndCancel;
  static HMENU hMenuList;
  static int nModifyFilter;
  static BOOL bListChanged;
  static DIALOGRESIZE drs[]={{&hWndList,         DRS_SIZE|DRS_X, 0},
                             {&hWndList,         DRS_SIZE|DRS_Y, 0},
                             {&hWndStats,        DRS_MOVE|DRS_X, 0},
                             {&hWndStats,        DRS_MOVE|DRS_Y, 0},
                             {&hWndSearch,       DRS_SIZE|DRS_X, 0},
                             {&hWndSearch,       DRS_MOVE|DRS_Y, 0},
                             {&hWndOK,           DRS_MOVE|DRS_X, 0},
                             {&hWndArrangeGroup, DRS_MOVE|DRS_X, 0},
                             {&hWndHorz,         DRS_MOVE|DRS_X, 0},
                             {&hWndVert,         DRS_MOVE|DRS_X, 0},
                             {&hWndTabsGroup,    DRS_MOVE|DRS_X, 0},
                             {&hWndUp,           DRS_MOVE|DRS_X, 0},
                             {&hWndDown,         DRS_MOVE|DRS_X, 0},
                             {&hWndSort,         DRS_MOVE|DRS_X, 0},
                             {&hWndFilesGroup,   DRS_MOVE|DRS_X, 0},
                             {&hWndSave,         DRS_MOVE|DRS_X, 0},
                             {&hWndClose,        DRS_MOVE|DRS_X, 0},
                             {&hWndNames,        DRS_MOVE|DRS_X, 0},
                             {&hWndCancel,       DRS_MOVE|DRS_X, 0},
                             {&hWndCancel,       DRS_MOVE|DRS_Y, 0},
                             {0, 0, 0}};
  static int lpMenuItems[]={IDC_MDILIST_HORZ,
                            IDC_MDILIST_VERT,
                            -1,
                            IDC_MDILIST_UP,
                            IDC_MDILIST_DOWN,
                            IDC_MDILIST_SORT,
                            -1,
                            IDC_MDILIST_SAVE,
                            IDC_MDILIST_CLOSE,
                            -1,
                            IDC_MDILIST_ALL,
                            IDC_MDILIST_ONLYMODIFIED,
                            IDC_MDILIST_ONLYUNMODIFIED,
                            -1,
                            IDC_MDILIST_INVERTSELECTION,
                            0};
  int nItem;

  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    hWndList=GetDlgItem(hDlg, IDC_MDILIST_LIST);
    hWndStats=GetDlgItem(hDlg, IDC_MDILIST_STATS);
    hWndSearch=GetDlgItem(hDlg, IDC_MDILIST_SEARCH);
    hWndOK=GetDlgItem(hDlg, IDOK);
    hWndArrangeGroup=GetDlgItem(hDlg, IDC_MDILIST_ARRANGE_GROUP);
    hWndHorz=GetDlgItem(hDlg, IDC_MDILIST_HORZ);
    hWndVert=GetDlgItem(hDlg, IDC_MDILIST_VERT);
    hWndTabsGroup=GetDlgItem(hDlg, IDC_MDILIST_TABS_GROUP);
    hWndUp=GetDlgItem(hDlg, IDC_MDILIST_UP);
    hWndDown=GetDlgItem(hDlg, IDC_MDILIST_DOWN);
    hWndSort=GetDlgItem(hDlg, IDC_MDILIST_SORT);
    hWndFilesGroup=GetDlgItem(hDlg, IDC_MDILIST_FILES_GROUP);
    hWndSave=GetDlgItem(hDlg, IDC_MDILIST_SAVE);
    hWndClose=GetDlgItem(hDlg, IDC_MDILIST_CLOSE);
    hWndNames=GetDlgItem(hDlg, IDC_MDILIST_ONLYNAMES);
    hWndCancel=GetDlgItem(hDlg, IDCANCEL);

    SendMessage(hWndSearch, EM_LIMITTEXT, MAX_PATH, 0);
    if (nMDI == WMD_PMDI)
    {
      EnableWindow(hWndArrangeGroup, FALSE);
      EnableWindow(hWndHorz, FALSE);
      EnableWindow(hWndVert, FALSE);
    }
    if (moCur.dwMdiListOptions & MLO_ONLYNAMES)
      SendMessage(hWndNames, BM_SETCHECK, BST_CHECKED, 0);

    nModifyFilter=MLMF_ALL;
    bListChanged=FALSE;
    FillMdiListListBox(hWndList, FALSE, nModifyFilter, (moCur.dwMdiListOptions & MLO_ONLYNAMES));
    if ((nItem=(int)SendMessage(hTab, TCM_GETCURSEL, 0, 0)) != -1)
      SendMessage(hWndList, LB_SETSEL, TRUE, nItem);

    if (hMenuList=CreatePopupMenu())
    {
      for (nItem=0; lpMenuItems[nItem]; ++nItem)
      {
        if (lpMenuItems[nItem] == -1)
        {
          AppendMenuWide(hMenuList, MF_SEPARATOR, (UINT)-1, NULL);
        }
        else
        {
          if (lpMenuItems[nItem] == IDC_MDILIST_ALL)
            API_LoadStringW(hLangLib, STR_ALLFILES, wbuf, BUFFER_SIZE);
          else if (lpMenuItems[nItem] == IDC_MDILIST_ONLYMODIFIED)
            API_LoadStringW(hLangLib, STR_ONLYMODIFIED, wbuf, BUFFER_SIZE);
          else if (lpMenuItems[nItem] == IDC_MDILIST_ONLYUNMODIFIED)
            API_LoadStringW(hLangLib, STR_ONLYUNMODIFIED, wbuf, BUFFER_SIZE);
          else if (lpMenuItems[nItem] == IDC_MDILIST_INVERTSELECTION)
            API_LoadStringW(hLangLib, STR_INVERTSELECTION, wbuf, BUFFER_SIZE);
          else
            GetDlgItemTextWide(hDlg, lpMenuItems[nItem], wbuf, BUFFER_SIZE);
          AppendMenuWide(hMenuList, MF_STRING, lpMenuItems[nItem], wbuf);
        }
      }
    }

    PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MDILIST_LIST, LBN_SELCHANGE), 0);
  }
  else if (uMsg == WM_CONTEXTMENU)
  {
    if ((HWND)wParam == hWndList)
    {
      POINT ptScreen;
      HWND hWndControl;
      BOOL bEnable;

      if (lParam == -1)
      {
        ptScreen.x=0;
        ptScreen.y=0;
        ClientToScreen(hWndList, &ptScreen);
      }
      else GetCursorPos(&ptScreen);

      for (nItem=0; lpMenuItems[nItem]; ++nItem)
      {
        if (lpMenuItems[nItem] != -1)
        {
          if (hWndControl=GetDlgItem(hDlg, lpMenuItems[nItem]))
          {
            bEnable=IsWindowEnabled(hWndControl);
            EnableMenuItem(hMenuList, lpMenuItems[nItem], bEnable?MF_ENABLED:MF_GRAYED);
          }
        }
      }
      CheckMenuRadioItem(hMenuList, IDC_MDILIST_ALL, IDC_MDILIST_ONLYUNMODIFIED, IDC_MDILIST_ALL + nModifyFilter, MF_BYCOMMAND);
      TrackPopupMenu(hMenuList, TPM_LEFTBUTTON|TPM_RIGHTBUTTON, ptScreen.x, ptScreen.y, 0, hDlg, NULL);
    }
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDC_MDILIST_SEARCH)
    {
      if (HIWORD(wParam) == EN_CHANGE)
      {
        wchar_t wszSearch[MAX_PATH];

        SendMessage(hWndList, LB_SETSEL, FALSE, -1);

        if (GetWindowTextWide(hWndSearch, wszSearch, MAX_PATH))
        {
          for (nItem=0;; ++nItem)
          {
            if (ListBox_GetTextWide(hWndList, nItem, wbuf) == LB_ERR)
              break;
            if (xstrstrW(wbuf, -1, wszSearch, -1, FALSE, NULL, NULL))
              SendMessage(hWndList, LB_SETSEL, TRUE, nItem);
          }
        }
        PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MDILIST_LIST, LBN_SELCHANGE), 0);
      }
    }
    else if (LOWORD(wParam) == IDC_MDILIST_HORZ)
    {
      ArrangeListBoxSelItems(hWndList, SB_HORZ);
    }
    else if (LOWORD(wParam) == IDC_MDILIST_VERT)
    {
      ArrangeListBoxSelItems(hWndList, SB_VERT);
    }
    else if (LOWORD(wParam) == IDC_MDILIST_UP)
    {
      if (ShiftListBoxSelItems(hWndList, FALSE))
      {
        bListChanged=TRUE;
        PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MDILIST_LIST, LBN_SELCHANGE), 0);
      }
    }
    else if (LOWORD(wParam) == IDC_MDILIST_DOWN)
    {
      if (ShiftListBoxSelItems(hWndList, TRUE))
      {
        bListChanged=TRUE;
        PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MDILIST_LIST, LBN_SELCHANGE), 0);
      }
    }
    else if (LOWORD(wParam) == IDC_MDILIST_SORT)
    {
      FillMdiListListBox(hWndList, TRUE, nModifyFilter, (moCur.dwMdiListOptions & MLO_ONLYNAMES));
      bListChanged=TRUE;
      PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MDILIST_LIST, LBN_SELCHANGE), 0);
    }
    else if (LOWORD(wParam) == IDC_MDILIST_SAVE)
    {
      SaveListBoxSelItems(hWndList);
    }
    else if (LOWORD(wParam) == IDC_MDILIST_CLOSE)
    {
      if (CloseListBoxSelItems(hWndList))
        SetFocus(hWndList);
      PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MDILIST_LIST, LBN_SELCHANGE), 0);
    }
    else if (LOWORD(wParam) == IDC_MDILIST_ONLYNAMES)
    {
      FRAMEDATA *lpFrame;
      int nListBoxItem=0;
      BOOL bSelected;

      if (SendMessage(hWndNames, BM_GETCHECK, 0, 0) == BST_CHECKED)
        moCur.dwMdiListOptions|=MLO_ONLYNAMES;
      else
        moCur.dwMdiListOptions&=~MLO_ONLYNAMES;

      while ((INT_PTR)(lpFrame=(FRAMEDATA *)SendMessage(hWndList, LB_GETITEMDATA, nListBoxItem, 0)) != LB_ERR)
      {
        xprintfW(wbuf, L"%s%s", (moCur.dwMdiListOptions & MLO_ONLYNAMES)?GetFileName(lpFrame->wszFile, lpFrame->nFileLen):lpFrame->wszFile, lpFrame->ei.bModified?L" *":L"");
        bSelected=(BOOL)SendMessage(hWndList, LB_GETSEL, nListBoxItem, 0);
        SendMessage(hWndList, LB_DELETESTRING, nListBoxItem, 0);
        ListBox_InsertStringWide(hWndList, nListBoxItem, wbuf);
        SendMessage(hWndList, LB_SETITEMDATA, nListBoxItem, (LPARAM)lpFrame);
        SendMessage(hWndList, LB_SETSEL, bSelected, nListBoxItem);

        ++nListBoxItem;
      }
      PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MDILIST_LIST, LBN_SELCHANGE), 0);
    }
    else if (LOWORD(wParam) == IDC_MDILIST_ALL ||
             LOWORD(wParam) == IDC_MDILIST_ONLYMODIFIED ||
             LOWORD(wParam) == IDC_MDILIST_ONLYUNMODIFIED)
    {
      if (LOWORD(wParam) - IDC_MDILIST_ALL != nModifyFilter)
      {
        nModifyFilter=LOWORD(wParam) - IDC_MDILIST_ALL;
        FillMdiListListBox(hWndList, FALSE, nModifyFilter, (moCur.dwMdiListOptions & MLO_ONLYNAMES));
        PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MDILIST_LIST, LBN_SELCHANGE), 0);
      }
    }
    else if (LOWORD(wParam) == IDC_MDILIST_INVERTSELECTION)
    {
      int *lpSelItems;
      int nCount;
      int nSelCount;
      int nItem;
      int nSelItem=0;

      if ((nCount=(int)SendMessage(hWndList, LB_GETCOUNT, 0, 0)) > 0)
      {
        nSelCount=GetListBoxSelItems(hWndList, &lpSelItems);
        SendMessage(hWndList, LB_SETSEL, FALSE, -1);

        for (nItem=0; nItem < nCount; ++nItem)
        {
          while (nSelItem < nSelCount && lpSelItems[nSelItem] < nItem)
            ++nSelItem;

          if (nItem != lpSelItems[nSelItem])
            SendMessage(hWndList, LB_SETSEL, TRUE, nItem);
        }
        FreeListBoxSelItems(&lpSelItems);

        PostMessage(hDlg, WM_COMMAND, MAKELONG(IDC_MDILIST_LIST, LBN_SELCHANGE), 0);
      }
    }
    else if (LOWORD(wParam) == IDC_MDILIST_LIST)
    {
      if (HIWORD(wParam) == LBN_SELCHANGE)
      {
        int nCount;
        int nSelCount;

        nCount=(int)SendMessage(hWndList, LB_GETCOUNT, 0, 0);
        nSelCount=(int)SendMessage(hWndList, LB_GETSELCOUNT, 0, 0);
        EnableWindow(hWndUp, nSelCount > 0 && !nModifyFilter);
        EnableWindow(hWndDown, nSelCount > 0 && !nModifyFilter);
        EnableWindow(hWndSort, !nModifyFilter);
        EnableWindow(hWndSave, nSelCount > 0);
        EnableWindow(hWndClose, nSelCount > 0);

        xprintfW(wbuf, L"%d / %d", nSelCount, nCount);
        SetWindowTextWide(hWndStats, wbuf);
      }
      else if (HIWORD(wParam) == LBN_DBLCLK)
      {
        PostMessage(hDlg, WM_COMMAND, IDOK, 0);
      }
    }
    else if (LOWORD(wParam) == IDOK)
    {
      int *lpSelItems;
      int nItemNew;
      int nItemOld;
      INT_PTR nData;

      if (!nModifyFilter && bListChanged)
      {
        for (nItemNew=0;; ++nItemNew)
        {
          if ((nData=SendMessage(hWndList, LB_GETITEMDATA, nItemNew, 0)) == LB_ERR)
            break;
          if ((nItemOld=GetTabItemFromParam(hTab, nData)) != -1)
            MoveTabItem(hTab, nItemOld, nItemNew);
        }
        bListChanged=FALSE;
      }
      if (GetListBoxSelItems(hWndList, &lpSelItems))
      {
        if ((nData=SendMessage(hWndList, LB_GETITEMDATA, lpSelItems[0], 0)) != LB_ERR)
          if ((nItem=GetTabItemFromParam(hTab, nData)) != -1)
            SelectTabItem(hTab, nItem);
        FreeListBoxSelItems(&lpSelItems);
      }
      EndDialog(hDlg, 1);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDCANCEL)
    {
      EndDialog(hDlg, 0);
      return TRUE;
    }
  }
  else if (uMsg == WM_DESTROY)
  {
    //Destroy resources
    DestroyMenu(hMenuList);
  }
  DialogResizeMessages(&drs[0], &rcMdiListMinMaxDialog, &moCur.rcMdiListCurrentDialog, DRM_PAINTSIZEGRIP, hDlg, uMsg, wParam, lParam);

  return FALSE;
}

void FillMdiListListBox(HWND hWnd, BOOL bSort, int nModifyFilter, BOOL bOnlyNames)
{
  TCITEMW tcItem;
  FRAMEDATA *lpFrame;
  int nTabItem=0;
  int nListBoxItem=0;
  int nNew;

  SendMessage(hWnd, LB_RESETCONTENT, 0, 0);

  for (;;)
  {
    tcItem.mask=TCIF_PARAM;
    if (!SendMessage(hTab, TCM_GETITEMW, nTabItem++, (LPARAM)&tcItem))
      break;
    lpFrame=(FRAMEDATA *)tcItem.lParam;

    if (nModifyFilter == MLMF_ALL ||
        (nModifyFilter == MLMF_ONLYMODIFIED && lpFrame->ei.bModified) ||
        (nModifyFilter == MLMF_ONLYUNMODIFIED && !lpFrame->ei.bModified))
    {
      xprintfW(wbuf, L"%s%s", bOnlyNames?GetFileName(lpFrame->wszFile, lpFrame->nFileLen):lpFrame->wszFile, lpFrame->ei.bModified?L" *":L"");

      if (bSort)
        nNew=ListBox_AddStringWide(hWnd, wbuf);
      else
        nNew=ListBox_InsertStringWide(hWnd, nListBoxItem++, wbuf);

      SendMessage(hWnd, LB_SETITEMDATA, nNew, (LPARAM)lpFrame);
    }
  }
}

int MoveListBoxItem(HWND hWnd, int nOldIndex, int nNewIndex)
{
  wchar_t *wpText;
  int nIndex=LB_ERR;
  INT_PTR nData;
  int nTextLen;

  if ((nTextLen=(int)SendMessage(hWnd, LB_GETTEXTLEN, nOldIndex, 0)) != LB_ERR)
  {
    if (wpText=AllocWideStr(nTextLen + 1))
    {
      ListBox_GetTextWide(hWnd, nOldIndex, wpText);
      nData=SendMessage(hWnd, LB_GETITEMDATA, nOldIndex, 0);

      SendMessage(hWnd, LB_DELETESTRING, nOldIndex, 0);
      nIndex=ListBox_InsertStringWide(hWnd, nNewIndex, wpText);
      SendMessage(hWnd, LB_SETITEMDATA, nIndex, nData);
      FreeWideStr(wpText);
    }
  }
  return nIndex;
}

BOOL ShiftListBoxSelItems(HWND hWnd, BOOL bMoveDown)
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

          MoveListBoxItem(hWnd, nOldIndex, nNewIndex);
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

          MoveListBoxItem(hWnd, nOldIndex, nNewIndex);
          SendMessage(hWnd, LB_SETSEL, TRUE, nNewIndex);
          bResult=TRUE;
        }
      }
    }
    FreeListBoxSelItems(&lpSelItems);
  }
  return bResult;
}

BOOL SaveListBoxSelItems(HWND hWnd)
{
  FRAMEDATA *lpFrame;
  int *lpSelItems;
  int nSelCount;
  int nItem;
  int i;
  BOOL bResult=TRUE;

  if (nSelCount=GetListBoxSelItems(hWnd, &lpSelItems))
  {
    for (i=nSelCount - 1; i >= 0; --i)
    {
      if ((INT_PTR)(lpFrame=(FRAMEDATA *)SendMessage(hWnd, LB_GETITEMDATA, lpSelItems[i], 0)) != LB_ERR)
      {
        if ((nItem=GetTabItemFromParam(hTab, (LPARAM)lpFrame)) != -1)
        {
          SelectTabItem(hTab, nItem);

          if (!DoFileSave())
          {
            bResult=FALSE;
            break;
          }
          SendMessage(hWnd, LB_DELETESTRING, lpSelItems[i], 0);
          ListBox_InsertStringWide(hWnd, lpSelItems[i], lpFrameCurrent->wszFile);

          SendMessage(hWnd, LB_SETITEMDATA, lpSelItems[i], (LPARAM)lpFrame);
          SendMessage(hWnd, LB_SETSEL, TRUE, lpSelItems[i]);
        }
      }
    }
    FreeListBoxSelItems(&lpSelItems);
  }
  return bResult;
}

void ArrangeListBoxSelItems(HWND hWnd, int nBar)
{
  FRAMEDATA *lpFrame;
  RECT rcClient;
  int *lpSelItems;
  int nHeight;
  int nWidth;
  int nSelCount;
  int nItem;
  int i;

  if (nMDI == WMD_MDI)
  {
    if (nSelCount=GetListBoxSelItems(hWnd, &lpSelItems))
    {
      GetClientRect(hMdiClient, &rcClient);
      if (nBar == SB_HORZ)
        rcClient.top=rcClient.bottom - rcClient.bottom / nSelCount;
      else if (nBar == SB_VERT)
        rcClient.left=rcClient.right - rcClient.right / nSelCount;

      for (i=nSelCount - 1; i >= 0; --i)
      {
        if ((INT_PTR)(lpFrame=(FRAMEDATA *)SendMessage(hWnd, LB_GETITEMDATA, lpSelItems[i], 0)) != LB_ERR)
        {
          if ((nItem=GetTabItemFromParam(hTab, (LPARAM)lpFrame)) != -1)
          {
            SelectTabItem(hTab, nItem);
            if (GetWindowLongPtrWide(lpFrame->hWndEditParent, GWL_STYLE) & WS_MAXIMIZE)
              SendMessage(hMdiClient, WM_MDIRESTORE, (WPARAM)lpFrame->hWndEditParent, 0);

            MoveWindow(lpFrame->hWndEditParent, rcClient.left, rcClient.top, rcClient.right - rcClient.left, rcClient.bottom - rcClient.top, TRUE);

            if (nBar == SB_HORZ)
            {
              nHeight=rcClient.bottom - rcClient.top;
              rcClient.top-=nHeight;
              rcClient.bottom-=nHeight;
            }
            else if (nBar == SB_VERT)
            {
              nWidth=rcClient.right - rcClient.left;
              rcClient.left-=nWidth;
              rcClient.right-=nWidth;
            }
          }
        }
      }
      FreeListBoxSelItems(&lpSelItems);
    }
  }
}

BOOL CloseListBoxSelItems(HWND hWnd)
{
  FRAMEDATA *lpFrame;
  int *lpSelItems;
  int nSelCount;
  int i;
  BOOL bResult=TRUE;

  if (nSelCount=GetListBoxSelItems(hWnd, &lpSelItems))
  {
    for (i=nSelCount - 1; i >= 0; --i)
    {
      if ((INT_PTR)(lpFrame=(FRAMEDATA *)SendMessage(hWnd, LB_GETITEMDATA, lpSelItems[i], 0)) != LB_ERR)
      {
        if (DestroyFrameWindow(lpFrame) != FWDE_SUCCESS)
        {
          bResult=FALSE;
          break;
        }
        SendMessage(hWnd, LB_DELETESTRING, lpSelItems[i], 0);
      }
    }
    FreeListBoxSelItems(&lpSelItems);
  }
  return bResult;
}

int GetListBoxSelItems(HWND hWnd, int **lpSelItems)
{
  int nSelCount;

  if (lpSelItems)
  {
    nSelCount=(int)SendMessage(hWnd, LB_GETSELCOUNT, 0, 0);

    if (*lpSelItems=(int *)API_HeapAlloc(hHeap, 0, nSelCount * sizeof(int)))
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
    API_HeapFree(hHeap, 0, (LPVOID)*lpSelItems);
    *lpSelItems=NULL;
  }
}

//// About

BOOL CALLBACK AboutDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  if (uMsg == WM_INITDIALOG)
  {
    SendMessage(hDlg, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)hMainIcon);
    SetDlgItemTextWide(hDlg, IDC_ABOUT_VERSION, APP_ABOUT_VERSIONW);
    SetDlgItemTextWide(hDlg, IDC_ABOUT_HOMEPAGE, APP_ABOUT_HOMEPAGEW);
    SetDlgItemTextWide(hDlg, IDC_ABOUT_EMAIL_SHENGALTS, APP_ABOUT_EMAIL_SHENGALTSW);
  }
  else if (uMsg == WM_COMMAND)
  {
    if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
    {
      EndDialog(hDlg, 0);
      return TRUE;
    }
    else if (LOWORD(wParam) == IDC_ABOUT_HOMEPAGE)
    {
      ShellExecuteWide(GetParent(hDlg), L"open", APP_ABOUT_HOMEPAGEW, NULL, NULL, SW_MAXIMIZE);
      return TRUE;
    }
  }
  return FALSE;
}

int GetUserManual(wchar_t *wszManual, int nManualLen)
{
  wchar_t wszFile[MAX_PATH];

  wszFile[0]=L'\0';
  if (PRIMARYLANGID(dwLangModule) == LANG_RUSSIAN)
    xprintfW(wszFile, L"%s\\AkelFiles\\Docs\\AkelHelp-%s.htm", wszExeDir, L"Rus");
  if (!*wszFile || !FileExistsWide(wszFile))
    xprintfW(wszFile, L"%s\\AkelFiles\\Docs\\AkelHelp-%s.htm", wszExeDir, L"Eng");

  if (FileExistsWide(wszFile))
    return (int)xstrcpynW(wszManual, wszFile, nManualLen);
  return 0;
}



//// System-like MessageBox implementation

int MessageBoxCustom(HWND hWndParent, const wchar_t *wpText, const wchar_t *wpCaption, UINT uType, HICON hIcon, BUTTONMESSAGEBOX *bmb)
{
  DIALOGMESSAGEBOX dmb;
  int nResult;

  SendMessage(hMainWnd, AKDN_MESSAGEBOXBEGIN, (WPARAM)hWndParent, 0);
  dmb.hWndParent=hWndParent;
  dmb.wpText=wpText;
  dmb.wpCaption=wpCaption;
  dmb.uType=uType;
  dmb.hIcon=hIcon;
  dmb.bmb=bmb;
  nResult=(int)API_DialogBoxParam(hLangLib, MAKEINTRESOURCEW(IDD_MESSAGEBOX), hWndParent, (DLGPROC)MessageBoxDlgProc, (LPARAM)&dmb);
  SendMessage(hMainWnd, AKDN_MESSAGEBOXEND, (WPARAM)hWndParent, 0);

  return nResult;
}

BOOL CALLBACK MessageBoxDlgProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  static DIALOGMESSAGEBOX *lpDialog;

  if (uMsg == WM_INITDIALOG)
  {
    BUTTONMESSAGEBOX *lpButton;
    RECT rcDlg;
    RECT rcTextOut;
    SIZE sizeString;
    POINT ptCursor;
    HWND hWndIcon;
    HWND hWndText;
    HWND hWndButton;
    HDC hDC;
    HFONT hGuiFont;
    HICON hIcon;
    wchar_t wszString[MAX_PATH];
    TCHAR *pIconIndex;
    DWORD dwIconType;
    DWORD dwStyle;
    int nButtonWidth=0;
    int nButtonHeight;
    int nButtonEdge;
    int nButtonOffset;
    int nButtonsCount=0;
    int nButtonsWidth;
    int nButtonX;
    int nButtonY;
    int nStrLen;
    BOOL bSnapToDefButton;

    lpDialog=(DIALOGMESSAGEBOX *)lParam;
    hDlgMsgBox=hDlg;
    SystemParametersInfoA(SPI_GETSNAPTODEFBUTTON, 0, &bSnapToDefButton, 0);
    hGuiFont=(HFONT)GetStockObject(DEFAULT_GUI_FONT);

    if (hDC=GetDC(hDlg))
    {
      SelectObject(hDC, hGuiFont);

      //Get maximum button width
      for (lpButton=lpDialog->bmb; lpButton->wpButtonStr; ++lpButton)
      {
        if ((INT_PTR)lpButton->wpButtonStr > 0xFFFF)
          nStrLen=(int)xstrcpynW(wszString, lpButton->wpButtonStr, MAX_PATH);
        else
          nStrLen=API_LoadStringW(hLangLib, (UINT)(UINT_PTR)lpButton->wpButtonStr, wszString, MAX_PATH);
        if (nStrLen)
        {
          GetTextExtentPoint32W(hDC, wszString, nStrLen, &sizeString);
          nButtonWidth=max(nButtonWidth, sizeString.cx + GetSystemMetrics(SM_CYFIXEDFRAME) * 2);
        }
        ++nButtonsCount;
      }

      //Get scale factor for ScaleX and ScaleY
      ScaleInit(hDC, NULL);
      nButtonWidth=max(nButtonWidth, ScaleX(75));
      nButtonHeight=ScaleY(23);
      nButtonEdge=ScaleX(6);
      nButtonOffset=ScaleX(16);

      //MessageBox title
      SetWindowTextWide(hDlg, lpDialog->wpCaption);

      //MessageBox icon
      dwIconType=(BYTE)(lpDialog->uType & 0xf0);
      hIcon=lpDialog->hIcon;

      if (!hIcon)
      {
        pIconIndex=NULL;
        if (dwIconType == MB_ICONINFORMATION) //Same as MB_ICONASTERISK
          pIconIndex=IDI_INFORMATION;
        else if (dwIconType == MB_ICONQUESTION)
          pIconIndex=IDI_QUESTION;
        else if (dwIconType == MB_ICONEXCLAMATION) //Same as MB_ICONWARNING
          pIconIndex=IDI_EXCLAMATION;
        else if (dwIconType == MB_ICONSTOP) //Same as MB_ICONERROR
          pIconIndex=IDI_ERROR;
        else if (dwIconType == MB_ICONHAND)
          pIconIndex=IDI_HAND;
        if (pIconIndex)
          hIcon=LoadIcon(NULL, pIconIndex);
      }
      if (hIcon)
      {
        hWndIcon=CreateWindowExWide(0, L"STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_ICON|SS_REALSIZEIMAGE, ScaleX(11), ScaleY(11), 32, 32, hDlg, (HMENU)(UINT_PTR)-1, hInstance, NULL);
        SendMessage(hWndIcon, STM_SETICON, (WPARAM)hIcon, 0);
      }

      //Get buttons width
      nButtonsWidth=nButtonsCount * (nButtonWidth + nButtonEdge) - nButtonEdge;
      nButtonsWidth=(nButtonsWidth / 2) * 2;

      //MessageBox text
      rcTextOut.left=hIcon?ScaleY(60):nButtonOffset;
      rcTextOut.right=max(nButtonsWidth + nButtonOffset * 2, ScaleX(500)) - nButtonOffset;
      rcTextOut.top=0;
      rcTextOut.bottom=0;

      if (bOldWindows)
      {
        WideCharToMultiByte(CP_ACP, 0, lpDialog->wpText, -1, buf, BUFFER_SIZE, NULL, NULL);
        DrawTextA(hDC, buf, -1, &rcTextOut, DT_CALCRECT|DT_WORDBREAK|DT_NOPREFIX|DT_EDITCONTROL);
      }
      else DrawTextW(hDC, lpDialog->wpText, -1, &rcTextOut, DT_CALCRECT|DT_WORDBREAK|DT_NOPREFIX|DT_EDITCONTROL);

      if (hIcon)
        rcTextOut.top=max(ScaleY(11 + 16) - RectH(&rcTextOut) / 2, ScaleY(11));
      else
        rcTextOut.top=ScaleY(11);
      rcTextOut.bottom+=rcTextOut.top;

      ReleaseDC(hDlg, hDC);
    }
    else return FALSE;

    if (hWndText=CreateWindowExWide(0, L"STATIC", NULL, WS_CHILD|WS_VISIBLE|SS_NOPREFIX|SS_EDITCONTROL, rcTextOut.left, rcTextOut.top, RectW(&rcTextOut), RectH(&rcTextOut), hDlg, (HMENU)(UINT_PTR)-1, hInstance, NULL))
    {
      SendMessage(hWndText, WM_SETFONT, (WPARAM)hGuiFont, TRUE);
      SetWindowTextWide(hWndText, lpDialog->wpText);
    }

    //Get button client Y position
    nButtonY=rcTextOut.bottom + ScaleY(17);
    if (hIcon)
    {
      if (ScaleY(11 + 32) > rcTextOut.bottom)
        nButtonY=ScaleY(60);
    }

    //MessageBox position
    rcDlg.right=max(nButtonsWidth + nButtonOffset * 2, rcTextOut.right + nButtonOffset);
    rcDlg.bottom=nButtonY + nButtonHeight + nButtonHeight / 2 + GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYFIXEDFRAME) * 2;
    rcDlg.left=GetSystemMetrics(SM_CXSCREEN) / 2 - rcDlg.right / 2;
    rcDlg.top=GetSystemMetrics(SM_CYSCREEN) / 2 - rcDlg.bottom / 2;
    if (rcDlg.left < 0) rcDlg.left=0;
    if (rcDlg.top < 0) rcDlg.top=0;
    SetWindowPos(hDlg, NULL, rcDlg.left, rcDlg.top, rcDlg.right, rcDlg.bottom, SWP_NOZORDER|SWP_NOACTIVATE);

    //MessageBox buttons
    nButtonX=rcDlg.right / 2 - nButtonsWidth / 2 - GetSystemMetrics(SM_CXFIXEDFRAME);

    for (lpButton=lpDialog->bmb; lpButton->wpButtonStr; ++lpButton)
    {
      dwStyle=WS_CHILD|WS_VISIBLE|WS_TABSTOP;
      if (lpButton == lpDialog->bmb)
        dwStyle|=WS_GROUP;
      if (lpButton->dwFlags & BMB_DEFAULT)
        dwStyle|=BS_DEFPUSHBUTTON;
      if (lpButton->dwFlags & BMB_DISABLED)
        dwStyle|=WS_DISABLED;

      if (hWndButton=CreateWindowExWide(0, L"BUTTON", NULL, dwStyle, nButtonX, nButtonY, nButtonWidth, nButtonHeight, hDlg, (HMENU)(UINT_PTR)lpButton->nButtonControlID, hInstance, NULL))
      {
        SendMessage(hWndButton, WM_SETFONT, (WPARAM)hGuiFont, TRUE);

        if ((INT_PTR)lpButton->wpButtonStr > 0xFFFF)
          xstrcpynW(wszString, lpButton->wpButtonStr, MAX_PATH);
        else
          API_LoadStringW(hLangLib, (UINT)(UINT_PTR)lpButton->wpButtonStr, wszString, MAX_PATH);
        SetWindowTextWide(hWndButton, wszString);

        if (lpButton->dwFlags & BMB_DEFAULT)
        {
          SetFocus(hWndButton);
          if (bSnapToDefButton)
          {
            ptCursor.x=nButtonX + nButtonWidth / 2;
            ptCursor.y=nButtonY + nButtonHeight / 2;
            ClientToScreen(hDlg, &ptCursor);
            SetCursorPos(ptCursor.x, ptCursor.y);
          }
        }
        nButtonX+=nButtonWidth + nButtonEdge;
      }
    }

    MessageBeep(dwIconType);
  }
  else if (uMsg == WM_COMMAND)
  {
    BUTTONMESSAGEBOX *lpButton;

    for (lpButton=lpDialog->bmb; lpButton->wpButtonStr; ++lpButton)
    {
      if (LOWORD(wParam) == lpButton->nButtonControlID)
      {
        EndDialog(hDlg, LOWORD(wParam));
        hDlgMsgBox=NULL;
        return TRUE;
      }
    }
  }
  return FALSE;
}

BOOL ScaleInit(HDC hDC, HWND hWnd)
{
  if (!ptScale.x && !ptScale.y)
  {
    HDC hNewDC=hDC;

    if (!hDC) hNewDC=GetDC(hWnd);

    if (hNewDC)
    {
      ptScale.x=GetDeviceCaps(hNewDC, LOGPIXELSX);
      ptScale.y=GetDeviceCaps(hNewDC, LOGPIXELSY);

      //Align to 16 pixel
      if (ptScale.x % 16) ptScale.x+=16 - ptScale.x % 16;
      if (ptScale.y % 16) ptScale.y+=16 - ptScale.y % 16;
    }
    else return FALSE;

    if (!hDC) ReleaseDC(hWnd, hNewDC);
  }
  return TRUE;
}

int ScaleX(int x)
{
  if (ptScale.x)
    return MulDiv(x, ptScale.x, 96);
  return x;
}

int ScaleY(int y)
{
  if (ptScale.y)
    return MulDiv(y, ptScale.y, 96);
  return y;
}


//// Dockable windows

DOCK* StackDockAdd(HDOCK *hDocks, DOCK *dkData)
{
  DOCK *dkElement=(DOCK *)hDocks->hStack.first;
  int nIndex=1;

  while (dkElement)
  {
    if (dkData->nSide == dkElement->nSide)
    {
      if (dkData->nSide == DKS_LEFT)
      {
        if (dkData->rcSize.left < dkElement->rcSize.left)
          break;
      }
      else if (dkData->nSide == DKS_RIGHT)
      {
        if (dkData->rcSize.left > dkElement->rcSize.left)
          break;
      }
      else if (dkData->nSide == DKS_TOP)
      {
        if (dkData->rcSize.top < dkElement->rcSize.top)
          break;
      }
      else if (dkData->nSide == DKS_BOTTOM)
      {
        if (dkData->rcSize.top > dkElement->rcSize.top)
          break;
      }
    }
    else if (dkData->nSide < dkElement->nSide)
      break;

    ++nIndex;

    dkElement=dkElement->next;
  }

  if (!StackInsertIndex((stack **)&hDocks->hStack.first, (stack **)&hDocks->hStack.last, (stack **)&dkElement, nIndex, sizeof(DOCK)))
  {
    dkElement->dwFlags=dkData->dwFlags;
    dkElement->hWnd=dkData->hWnd;
    dkElement->nSide=dkData->nSide;
    dkElement->rcSize=dkData->rcSize;
    dkElement->rcDragDrop=dkData->rcDragDrop;
  }
  return dkElement;
}

int DockSetSide(HDOCK *hDocks, DOCK *dkData, int nSide)
{
  DOCK *dkElement=(DOCK *)hDocks->hStack.first;
  int nIndex=1;

  while (dkElement)
  {
    if (nSide == dkElement->nSide)
    {
      if (dkData != dkElement)
      {
        if (nSide == DKS_LEFT)
        {
          if (dkData->rcSize.left <= dkElement->rcSize.left)
            break;
        }
        if (nSide == DKS_RIGHT)
        {
          if (dkData->rcSize.left >= dkElement->rcSize.left)
            break;
        }
        else if (nSide == DKS_TOP)
        {
          if (dkData->rcSize.top <= dkElement->rcSize.top)
            break;
        }
        else if (nSide == DKS_BOTTOM)
        {
          if (dkData->rcSize.top >= dkElement->rcSize.top)
            break;
        }
      }
    }
    else
    {
      if (nSide < dkElement->nSide)
        break;
      if (dkData == dkElement)
        --nIndex;
    }
    ++nIndex;

    dkElement=dkElement->next;
  }
  dkData->nSide=nSide;

  return StackMoveIndex((stack **)&hDocks->hStack.first, (stack **)&hDocks->hStack.last, (stack *)dkData, nIndex);
}

DOCK* StackDockFindWindow(HDOCK *hDocks, HWND hWnd, BOOL bChild)
{
  DOCK *dkElement=(DOCK *)hDocks->hStack.first;

  if (hWnd)
  {
    while (dkElement)
    {
      if (dkElement->hWnd)
      {
        if (dkElement->hWnd == hWnd || (bChild && IsChild(dkElement->hWnd, hWnd)))
          return dkElement;
      }
      dkElement=dkElement->next;
    }
  }
  return NULL;
}

HWND StackDockNextWindow(HDOCK *hDocks, DOCK *dkData, BOOL bPrevious)
{
  if (bPrevious)
  {
    if (!dkData)
      dkData=(DOCK *)hDocks->hStack.last;
    else
      dkData=dkData->prev;

    while (dkData)
    {
      if (dkData->hWnd && !(dkData->dwFlags & DKF_HIDDEN))
        return dkData->hWnd;

      dkData=dkData->prev;
    }
  }
  else
  {
    if (!dkData)
      dkData=(DOCK *)hDocks->hStack.first;
    else
      dkData=dkData->next;

    while (dkData)
    {
      if (dkData->hWnd && !(dkData->dwFlags & DKF_HIDDEN))
        return dkData->hWnd;

      dkData=dkData->next;
    }
  }
  return NULL;
}

DOCK* StackDockFromPoint(HDOCK *hDocks, POINT *ptScreen)
{
  DOCK *dkElement=(DOCK *)hDocks->hStack.first;
  RECT rc;

  while (dkElement)
  {
    if (dkElement->hWnd)
    {
      if (GetWindowRect(dkElement->hWnd, &rc))
      {
        if (PtInRect(&rc, *ptScreen))
          return dkElement;
      }
    }
    dkElement=dkElement->next;
  }
  return NULL;
}

void StackDockSize(HDOCK *hDocks, int nSide, NSIZE *ns)
{
  DOCK *dkElement=(DOCK *)hDocks->hStack.first;
  RECT rcDock;

  while (dkElement)
  {
    if (dkElement->nSide > nSide) break;

    if (dkElement->nSide == nSide)
    {
      if (!(dkElement->dwFlags & DKF_HIDDEN))
      {
        if (dkElement->rcSize.right || dkElement->rcSize.bottom)
        {
          rcDock.left=ns->rcCurrent.left;
          rcDock.top=ns->rcCurrent.top;
          rcDock.right=dkElement->rcSize.right;
          rcDock.bottom=dkElement->rcSize.bottom;

          if (dkElement->nSide == DKS_LEFT)
          {
            if (ns->rcCurrent.right - rcDock.right < DOCK_MAINMIN_X)
              rcDock.right=ns->rcCurrent.right - DOCK_MAINMIN_X;
            else if (ns->rcCurrent.right >= DOCK_MAINMIN_X + DOCK_BORDER_2X)
              rcDock.right=max(rcDock.right, DOCK_BORDER_2X);

            ns->rcCurrent.left+=rcDock.right;
            ns->rcCurrent.right-=rcDock.right;
          }
          else if (dkElement->nSide == DKS_TOP)
          {
            if (ns->rcCurrent.bottom - rcDock.bottom < DOCK_MAINMIN_Y)
              rcDock.bottom=ns->rcCurrent.bottom - DOCK_MAINMIN_Y;
            else if (ns->rcCurrent.bottom >= DOCK_MAINMIN_Y + DOCK_BORDER_2X)
              rcDock.bottom=max(rcDock.bottom, DOCK_BORDER_2X);

            ns->rcCurrent.top+=rcDock.bottom;
            ns->rcCurrent.bottom-=rcDock.bottom;
          }
          else if (dkElement->nSide == DKS_RIGHT)
          {
            if (ns->rcCurrent.right - rcDock.right < DOCK_MAINMIN_X)
              rcDock.right=ns->rcCurrent.right - DOCK_MAINMIN_X;
            else if (ns->rcCurrent.right >= DOCK_MAINMIN_X + DOCK_BORDER_2X)
              rcDock.right=max(rcDock.right, DOCK_BORDER_2X);

            ns->rcCurrent.right-=rcDock.right;
            rcDock.left=ns->rcCurrent.left + ns->rcCurrent.right;
          }
          else if (dkElement->nSide == DKS_BOTTOM)
          {
            if (ns->rcCurrent.bottom - rcDock.bottom < DOCK_MAINMIN_Y)
              rcDock.bottom=ns->rcCurrent.bottom - DOCK_MAINMIN_Y;
            else if (ns->rcCurrent.bottom >= DOCK_MAINMIN_Y + DOCK_BORDER_2X)
              rcDock.bottom=max(rcDock.bottom, DOCK_BORDER_2X);

            ns->rcCurrent.bottom-=rcDock.bottom;
            rcDock.top=ns->rcCurrent.top + ns->rcCurrent.bottom;
          }

          if (dkElement->hWnd)
          {
            if (dkElement->nSide == DKS_LEFT ||
                dkElement->nSide == DKS_RIGHT)
            {
              if (hDocksStack.nSizingSide) dkElement->rcSize.left=rcDock.left;
              MoveWindow(dkElement->hWnd, rcDock.left, ns->rcCurrent.top, rcDock.right, ns->rcCurrent.bottom, FALSE);
              RedrawWindow(dkElement->hWnd, NULL, NULL, RDW_INVALIDATE|RDW_ERASE|RDW_ALLCHILDREN);
            }
            else if (dkElement->nSide == DKS_TOP ||
                     dkElement->nSide == DKS_BOTTOM)
            {
              if (hDocksStack.nSizingSide) dkElement->rcSize.top=rcDock.top;
              MoveWindow(dkElement->hWnd, ns->rcCurrent.left, rcDock.top, ns->rcCurrent.right, rcDock.bottom, FALSE);
              RedrawWindow(dkElement->hWnd, NULL, NULL, RDW_INVALIDATE|RDW_ERASE|RDW_ALLCHILDREN);
            }
          }
        }
      }
    }

    dkElement=dkElement->next;
  }
}

void StackDockDelete(HDOCK *hDocks, DOCK *dkData)
{
  StackDelete((stack **)&hDocks->hStack.first, (stack **)&hDocks->hStack.last, (stack *)dkData);
}

void StackDockFree(HDOCK *hDocks)
{
  StackClear((stack **)&hDocks->hStack.first, (stack **)&hDocks->hStack.last);
}

BOOL TranslateMessageDialog(HDOCK *hDocks, LPMSG lpMsg)
{
  if (hDlgModeless && IsDialogMessageWide(hDlgModeless, lpMsg))
  {
    if (lpMsg->message >= WM_KEYFIRST && lpMsg->message <= WM_KEYLAST)
    {
      SendMessageWide(hDlgModeless, lpMsg->message, lpMsg->wParam, lpMsg->lParam);
    }
    return TRUE;
  }
  else
  {
    DOCK *dkElement=(DOCK *)hDocks->hStack.first;

    while (dkElement)
    {
      if (!(dkElement->dwFlags & DKF_OWNTHREAD))
      {
        if (dkElement->hWnd && IsDialogMessageWide(dkElement->hWnd, lpMsg))
        {
          if (lpMsg->message >= WM_KEYFIRST && lpMsg->message <= WM_KEYLAST)
          {
            SendMessageWide(dkElement->hWnd, lpMsg->message, lpMsg->wParam, lpMsg->lParam);
          }
          return TRUE;
        }
      }
      dkElement=dkElement->next;
    }
  }
  return FALSE;
}


//// Procedures stack

int StackProcGet(HSTACK *hStack, int nIndex, WNDPROCDATA **ProcData)
{
  return StackGetElement((stack *)hStack->first, (stack *)hStack->last, (stack **)ProcData, nIndex);
}

int StackProcSet(HSTACK *hStack, WNDPROC NewProc, WNDPROCDATA **NewProcData, WNDPROC *FirstProc)
{
  if (!NewProcData || !*NewProcData)
  {
    WNDPROCDATA *lpElement;

    if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, 1, sizeof(WNDPROCDATA)))
    {
      lpElement->CurProc=NewProc;
      if (lpElement->next)
      {
        lpElement->next->PrevProc=NewProc;
        lpElement->NextProc=lpElement->next->CurProc;
      }
      if (NewProcData) *NewProcData=lpElement;
    }
  }
  else
  {
    if (!NewProc)
    {
      if ((*NewProcData)->prev) (*NewProcData)->prev->NextProc=(*NewProcData)->NextProc;
      if ((*NewProcData)->next) (*NewProcData)->next->PrevProc=(*NewProcData)->PrevProc;
      StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)*NewProcData);
      *NewProcData=NULL;
    }
    else
    {
      (*NewProcData)->CurProc=NewProc;
      if ((*NewProcData)->prev) (*NewProcData)->prev->NextProc=NewProc;
      if ((*NewProcData)->next) (*NewProcData)->next->PrevProc=NewProc;
    }
  }
  if (FirstProc)
  {
    if (hStack->first)
      *FirstProc=((WNDPROCDATA *)(hStack->first))->CurProc;
    else
      *FirstProc=NULL;
  }
  return 0;
}

void StackProcFree(HSTACK *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}


//// Handles stack

PLUGINHANDLE* StackHandleIncrease(HSTACK *hStack, HMODULE hModule)
{
  PLUGINHANDLE *phElement=(PLUGINHANDLE *)hStack->first;

  while (phElement)
  {
    if (phElement->hModule == hModule)
    {
      phElement->nCount+=1;
      return phElement;
    }
    phElement=phElement->next;
  }

  if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&phElement, -1, sizeof(PLUGINHANDLE)))
  {
    phElement->hModule=hModule;
    phElement->nCount=1;
  }
  return phElement;
}

PLUGINHANDLE* StackHandleDecrease(HSTACK *hStack, HMODULE hModule)
{
  PLUGINHANDLE *phElement=(PLUGINHANDLE *)hStack->first;

  while (phElement)
  {
    if (phElement->hModule == hModule)
    {
      phElement->nCount-=1;
      if (phElement->nCount <= 0)
      {
        StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)phElement);
        return NULL;
      }
      return phElement;
    }
    phElement=phElement->next;
  }
  return NULL;
}

void StackHandleFree(HSTACK *hStack)
{
  PLUGINHANDLE *phElement=(PLUGINHANDLE *)hStack->first;

  while (phElement)
  {
    while (phElement->nCount > 0 && FreeLibrary(phElement->hModule))
    {
      phElement->nCount-=1;
    }
    phElement=phElement->next;
  }
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}


//// MDI windows

FRAMEDATA* StackFrameInsert(HSTACK *hStack)
{
  FRAMEDATA *lpElement=NULL;

  StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, -1, sizeof(FRAMEDATA));
  return lpElement;
}

FRAMEDATA* StackFrameGetByIndex(HSTACK *hStack, int nIndex)
{
  FRAMEDATA *lpFrame;

  if (!StackGetElement((stack *)hStack->first, (stack *)hStack->last, (stack **)&lpFrame, nIndex))
    return lpFrame;
  return NULL;
}

FRAMEDATA* StackFrameGetByHandle(HSTACK *hStack, AEHDOC hDocEdit)
{
  FRAMEDATA *lpFrame;

  for (lpFrame=(FRAMEDATA *)hStack->last; lpFrame; lpFrame=lpFrame->prev)
  {
    if (lpFrame->ei.hDocEdit == hDocEdit)
      return lpFrame;
  }
  return NULL;
}

FRAMEDATA* StackFrameGetByName(HSTACK *hStack, const wchar_t *wpFileName, int nFileNameLen)
{
  FRAMEDATA *lpFrame;

  if (nFileNameLen == -1)
    nFileNameLen=(int)xstrlenW(wpFileName);

  if (lpFrameCurrent->ei.hWndEdit)
  {
    if (lpFrameCurrent->nFileLen == nFileNameLen)
    {
      if (!xstrcmpiW(lpFrameCurrent->wszFile, wpFileName))
        return lpFrameCurrent;
    }
  }

  for (lpFrame=(FRAMEDATA *)hStack->first; lpFrame; lpFrame=lpFrame->next)
  {
    if (lpFrame->nFileLen == nFileNameLen)
    {
      if (!xstrcmpiW(lpFrame->wszFile, wpFileName))
        return lpFrame;
    }
  }
  return NULL;
}

FRAMEDATA* StackFrameGetNext(HSTACK *hStack, FRAMEDATA *lpFrame, BOOL bPrev)
{
  if (!bPrev)
  {
    if (lpFrame)
    {
      if (!lpFrame->next)
        return (FRAMEDATA *)hStack->first;
      else
        return lpFrame->next;
    }
  }
  else
  {
    if (lpFrame)
    {
      if (!lpFrame->prev)
        return (FRAMEDATA *)hStack->last;
      else
        return lpFrame->prev;
    }
  }
  return NULL;
}

DWORD StackFrameGetIndex(HSTACK *hStack, FRAMEDATA *lpFramePointer)
{
  FRAMEDATA *lpFrame;
  DWORD dwIndex=1;

  for (lpFrame=(FRAMEDATA *)hStack->first; lpFrame; lpFrame=lpFrame->next)
  {
    if (lpFrame == lpFramePointer)
      return dwIndex;
    ++dwIndex;
  }
  return 0;
}

FRAMEDATA* StackFrameIsValid(HSTACK *hStack, FRAMEDATA *lpFramePointer)
{
  FRAMEDATA *lpFrame;

  for (lpFrame=(FRAMEDATA *)hStack->first; lpFrame; lpFrame=lpFrame->next)
  {
    if (lpFrame == lpFramePointer)
      return lpFrame;
  }
  return NULL;
}

void StackFrameMove(HSTACK *hStack, FRAMEDATA *lpFrame, int nIndex)
{
  StackMoveIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack *)lpFrame, nIndex);
}

void StackFrameDelete(HSTACK *hStack, FRAMEDATA *lpFrame)
{
  StackRecentCaretFree(&lpFrame->hRecentCaretStack);
  if (lpFrame == lpFrameCurrent)
    lpFrameCurrent=&fdDefault;
  StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)lpFrame);
}

void StackFramesFree(HSTACK *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}


//// Owner-drawn buttons

BUTTONDRAWITEM* StackButtonDrawInsert(HSTACK *hStack)
{
  BUTTONDRAWITEM *lpElement=NULL;

  StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, -1, sizeof(BUTTONDRAWITEM));
  return lpElement;
}

BUTTONDRAWITEM* StackButtonDrawGet(HSTACK *hStack, HWND hWnd)
{
  BUTTONDRAWITEM *lpButtonDraw=(BUTTONDRAWITEM *)hStack->last;

  while (lpButtonDraw)
  {
    if (lpButtonDraw->hWnd == hWnd)
      return lpButtonDraw;

    lpButtonDraw=lpButtonDraw->prev;
  }
  return NULL;
}

void StackButtonDrawDelete(HSTACK *hStack, BUTTONDRAWITEM *lpButtonDraw)
{
  StackDelete((stack **)&hStack->first, (stack **)&hStack->last, (stack *)lpButtonDraw);
}

void StackButtonDrawFree(HSTACK *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}

void SetButtonDraw(HWND hWndButton, BUTTONDRAW *bd)
{
  BUTTONDRAWITEM *lpButtonDrawItem;

  if (bd == NULL)
  {
    if (lpButtonDrawItem=StackButtonDrawGet(&hButtonDrawStack, hWndButton))
    {
      SetWindowLongPtrWide(hWndButton, GWLP_WNDPROC, (UINT_PTR)lpButtonDrawItem->OldButtonProc);

      StackButtonDrawDelete(&hButtonDrawStack, lpButtonDrawItem);
    }
  }
  else
  {
    if (!(lpButtonDrawItem=StackButtonDrawGet(&hButtonDrawStack, hWndButton)))
      lpButtonDrawItem=StackButtonDrawInsert(&hButtonDrawStack);

    if (lpButtonDrawItem)
    {
      lpButtonDrawItem->hWnd=hWndButton;
      xmemcpy(&lpButtonDrawItem->bd, bd, sizeof(BUTTONDRAW));

      if (!lpButtonDrawItem->OldButtonProc)
      {
        lpButtonDrawItem->OldButtonProc=(WNDPROC)GetWindowLongPtrWide(hWndButton, GWLP_WNDPROC);
        SetWindowLongPtrWide(hWndButton, GWLP_WNDPROC, (UINT_PTR)NewButtonDrawProc);
      }
    }
  }
}


//// Recent caret

RECENTCARETITEM* StackRecentCaretInsert(STACKRECENTCARET *hStack)
{
  RECENTCARETITEM *lpElement=NULL;

  StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, -1, sizeof(RECENTCARETITEM));
  return lpElement;
}

void StackRecentCaretFree(STACKRECENTCARET *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}

void RecentCaretSet(AENSELCHANGE *aensc)
{
  if (!bRecentCaretMsg)
  {
    if ((aensc->dwType & AESCT_SETSELMESSAGE) ||
        (!(aensc->dwType & AESCT_MOUSECAPTURE) &&
         ((aensc->dwType & AESCT_MOUSESINGLECLK) ||
          //(aensc->dwType & AESCT_MOUSEDOUBLECLK) ||
          //(aensc->dwType & AESCT_MOUSETRIPLECLK) ||
          (aensc->dwType & AESCT_MOUSELEFTMARGIN))))
    {
      if (!lpFrameCurrent->hRecentCaretStack.last || lpFrameCurrent->hRecentCaretStack.last->nCaretOffset != aensc->crRichSel.cpMin)
      {
        RECENTCARETITEM *lpRecentCaret;

        if (lpRecentCaret=StackRecentCaretInsert(&lpFrameCurrent->hRecentCaretStack))
        {
          lpRecentCaret->nCaretOffset=aensc->crRichSel.cpMin;
        }
      }
      lpFrameCurrent->lpCurRecentCaret=NULL;
    }
  }
}


//// Status bar

void SetSelectionStatus(AEHDOC hDocEdit, HWND hWndEdit, AECHARRANGE *cr, AECHARINDEX *ci)
{
  if (hWndEdit && lpFrameCurrent->ei.hDocEdit == hDocEdit)
  {
    wchar_t wszStatus[MAX_PATH];
    BOOL bColumnSel;

    nLoopCase=0;

    if (cr && ci)
    {
      crCurSel.ciMin=cr->ciMin;
      crCurSel.ciMax=cr->ciMax;
      ciCurCaret=*ci;
      bColumnSel=(BOOL)SendMessage(hWndEdit, AEM_GETCOLUMNSEL, 0, 0);
    }
    else
    {
      GetSel(hWndEdit, &crCurSel, &bColumnSel, &ciCurCaret);
      lpFrameCurrent->nSelSubtract=0;
    }

    lpFrameCurrent->nCaretLine=GetLineNumber(lpFrameCurrent, ciCurCaret.nLine);
    lpFrameCurrent->nCaretColumn=GetIndexColumn(lpFrameCurrent, &ciCurCaret);

    if (!AEC_IndexCompare(&crCurSel.ciMin, &crCurSel.ciMax))
    {
      xprintfW(wszStatus, L"%u:%u", lpFrameCurrent->nCaretLine + 1, lpFrameCurrent->nCaretColumn + 1);
      lpFrameCurrent->nSelSubtract=0;
    }
    else
    {
      if (bColumnSel || !lpFrameCurrent->nSelSubtract || mod(lpFrameCurrent->crPrevSel.ciMin.nLine - crCurSel.ciMin.nLine) + mod(lpFrameCurrent->crPrevSel.ciMax.nLine - crCurSel.ciMax.nLine) >= crCurSel.ciMax.nLine - crCurSel.ciMin.nLine)
      {
        lpFrameCurrent->nSelSubtract=IndexSubtract(hWndEdit, &crCurSel.ciMax, &crCurSel.ciMin, AELB_ASOUTPUT, -1);
      }
      else
      {
        lpFrameCurrent->nSelSubtract+=IndexSubtract(hWndEdit, &lpFrameCurrent->crPrevSel.ciMin, &crCurSel.ciMin, AELB_ASOUTPUT, -1);
        lpFrameCurrent->nSelSubtract+=IndexSubtract(hWndEdit, &crCurSel.ciMax, &lpFrameCurrent->crPrevSel.ciMax, AELB_ASOUTPUT, -1);
      }
      xprintfW(wszStatus, L"%u:%u, %Iu", lpFrameCurrent->nCaretLine + 1, lpFrameCurrent->nCaretColumn + 1, lpFrameCurrent->nSelSubtract);
      if (bColumnSel) lpFrameCurrent->nSelSubtract=0;
    }
    lpFrameCurrent->crPrevSel.ciMin=crCurSel.ciMin;
    lpFrameCurrent->crPrevSel.ciMax=crCurSel.ciMax;

    StatusBar_SetTextWide(hStatus, SBP_POSITION, wszStatus);
    UpdateStatusUser(lpFrameCurrent, CSB_CHARHEX|CSB_CHARDEC|CSB_CHARLETTER|CSB_RICHOFFSET|CSB_BYTEOFFSET|CSB_LINEALLCOUNT|CSB_RICHCOUNT|CSB_LINESELCOUNT|CSB_LINESELBEGIN|CSB_LINESELEND);
  }
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void SetModifyStatus(FRAMEDATA *lpFrame, BOOL bState)
{
  if (!lpFrame || lpFrame == lpFrameCurrent)
  {
    //Read only status
    if (lpFrameCurrent->ei.bReadOnly)
    {
      if (!ssStatus.bReadOnly)
      {
        API_LoadStringW(hLangLib, STR_READONLY, wbuf, BUFFER_SIZE);
        StatusBar_SetTextWide(hStatus, SBP_MODIFY, wbuf);
        ssStatus.bReadOnly=TRUE;
      }
    }
    else
    {
      if (ssStatus.bReadOnly)
      {
        if (ssStatus.bModified)
          API_LoadStringW(hLangLib, STR_MODIFIED, wbuf, BUFFER_SIZE);
        else
          wbuf[0]=L'\0';

        StatusBar_SetTextWide(hStatus, SBP_MODIFY, wbuf);
        ssStatus.bReadOnly=FALSE;
      }
    }

    //Modify status
    if (moCur.dwShowModify & SM_STATUSBAR)
    {
      if (ssStatus.bModified != bState)
      {
        if (!ssStatus.bReadOnly)
        {
          if (bState)
            API_LoadStringW(hLangLib, STR_MODIFIED, wbuf, BUFFER_SIZE);
          else
            wbuf[0]=L'\0';

          StatusBar_SetTextWide(hStatus, SBP_MODIFY, wbuf);
        }
        ssStatus.bModified=bState;
      }
    }
    if (!lpFrame) return;

    lpFrameCurrent->ei.bModified=bState;
  }
  else lpFrame->ei.bModified=bState;

  //Set modify flag
  SendMessage(lpFrame->ei.hWndEdit, AEM_SETMODIFY, bState, 0);
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void SetOvertypeStatus(FRAMEDATA *lpFrame, BOOL bState)
{
  if (!lpFrame || lpFrame == lpFrameCurrent)
  {
    if (ssStatus.bOvertypeMode != bState)
    {
      StatusBar_SetTextWide(hStatus, SBP_INSERT, bState?L"Ovr":L"Ins");
      ssStatus.bOvertypeMode=bState;
    }
    if (!lpFrame) return;

    lpFrameCurrent->ei.bOvertypeMode=bState;
  }
  else lpFrame->ei.bOvertypeMode=bState;

  //Set overtype mode
  SendMessage(lpFrame->ei.hWndEdit, AEM_EXSETOPTIONS, bState?AECOOP_OR:AECOOP_XOR, AECOE_OVERTYPE);
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void SetNewLineStatus(FRAMEDATA *lpFrame, int nState, DWORD dwFlags)
{
  if (!lpFrame || lpFrame == lpFrameCurrent)
  {
    if (ssStatus.nNewLine != nState)
    {
      if (nState == NEWLINE_WIN)
        StatusBar_SetTextWide(hStatus, SBP_NEWLINE, STR_NEWLINE_WIN);
      else if (nState == NEWLINE_UNIX)
        StatusBar_SetTextWide(hStatus, SBP_NEWLINE, STR_NEWLINE_UNIX);
      else if (nState == NEWLINE_MAC)
        StatusBar_SetTextWide(hStatus, SBP_NEWLINE, STR_NEWLINE_MAC);
      ssStatus.nNewLine=nState;
    }
    if (!lpFrame) return;

    lpFrameCurrent->ei.nNewLine=nState;
  }
  else lpFrame->ei.nNewLine=nState;

  //Set new line
  SendMessage(lpFrame->ei.hWndEdit, AEM_SETNEWLINE, AENL_INPUT|AENL_OUTPUT, MAKELONG(AELB_ASIS, AELB_ASIS));

  if (nState == NEWLINE_WIN)
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETNEWLINE, dwFlags, MAKELONG(AELB_RN, AELB_RN));
  else if (nState == NEWLINE_UNIX)
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETNEWLINE, dwFlags, MAKELONG(AELB_N, AELB_N));
  else if (nState == NEWLINE_MAC)
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETNEWLINE, dwFlags, MAKELONG(AELB_R, AELB_R));

  lpFrame->nSelSubtract=0;
  SendMessage(lpFrame->ei.hWndEdit, AEM_UPDATESEL, AESELT_COLUMNASIS|AESELT_LOCKSCROLL, 0);
}

void SetCodePageStatus(FRAMEDATA *lpFrame, int nCodePage, BOOL bBOM)
{
  if (!lpFrame || lpFrame == lpFrameCurrent)
  {
    if (ssStatus.nCodePage != nCodePage || ssStatus.bBOM != bBOM)
    {
      GetCodePageName(nCodePage, wbuf, BUFFER_SIZE);

      if (IsCodePageUnicode(nCodePage))
      {
        if (!bBOM)
        {
          API_LoadStringW(hLangLib, STR_WITHOUT, wbuf2, BUFFER_SIZE);
          xprintfW(wbuf, L"%s %s BOM", wbuf, wbuf2);
        }
      }
      StatusBar_SetTextWide(hStatus, SBP_CODEPAGE, wbuf);

      ssStatus.nCodePage=nCodePage;
      ssStatus.bBOM=bBOM;
    }
    if (!lpFrame) return;

    lpFrameCurrent->ei.nCodePage=nCodePage;
    lpFrameCurrent->ei.bBOM=bBOM;
  }
  else
  {
    lpFrame->ei.nCodePage=nCodePage;
    lpFrame->ei.bBOM=bBOM;
  }
}

void UpdateStatusUser(FRAMEDATA *lpFrame, DWORD dwFlags)
{
  if (moCur.wszStatusUserFormat[0])
  {
    if (moCur.dwStatusUserFlags & dwFlags)
    {
      if (lpFrame->ei.hWndEdit)
      {
        if (((moCur.dwStatusUserFlags & CSB_CHARHEX) && (dwFlags & CSB_CHARHEX)) ||
            ((moCur.dwStatusUserFlags & CSB_CHARDEC) && (dwFlags & CSB_CHARDEC)) ||
            ((moCur.dwStatusUserFlags & CSB_CHARLETTER) && (dwFlags & CSB_CHARLETTER)))
        {
          lpFrame->nCaretChar=AEC_CharAtIndex(&ciCurCaret);
          if (lpFrame->nCaretChar <= -AELB_R && lpFrame->nCaretChar >= -AELB_RRN)
            lpFrame->nCaretChar=L'\r';
          else if (lpFrame->nCaretChar == -AELB_EOF)
            lpFrame->nCaretChar=0xFFFF;
        }
        if ((moCur.dwStatusUserFlags & CSB_RICHOFFSET) && (dwFlags & CSB_RICHOFFSET))
          lpFrame->nCaretRichOffset=AkelIndexToRichOffset(lpFrame->ei.hWndEdit, &ciCurCaret);
        if ((moCur.dwStatusUserFlags & CSB_BYTEOFFSET) && (dwFlags & CSB_BYTEOFFSET))
          lpFrame->nCaretByteOffset=-IndexSubtract(lpFrame->ei.hWndEdit, NULL, &ciCurCaret, AELB_ASIS, FALSE);
        if ((moCur.dwStatusUserFlags & CSB_LINEALLCOUNT) && (dwFlags & CSB_LINEALLCOUNT))
        {
          if (!(moCur.dwStatusPosType & SPT_LINEWRAP) && lpFrame->ei.bWordWrap)
            lpFrame->nLineCountAll=(int)SendMessage(lpFrame->ei.hWndEdit, AEM_GETLINENUMBER, AEGL_LINEUNWRAPCOUNT, 0);
          else
            lpFrame->nLineCountAll=(int)SendMessage(lpFrame->ei.hWndEdit, AEM_GETLINENUMBER, AEGL_LINECOUNT, 0);
        }
        if (((moCur.dwStatusUserFlags & CSB_LINESELCOUNT) && (dwFlags & CSB_LINESELCOUNT)) ||
            ((moCur.dwStatusUserFlags & CSB_LINESELBEGIN) && (dwFlags & CSB_LINESELBEGIN)) ||
            ((moCur.dwStatusUserFlags & CSB_LINESELEND) && (dwFlags & CSB_LINESELEND)))
        {
          if (ciCurCaret.nLine == crCurSel.ciMin.nLine)
          {
            lpFrame->nLineSelBegin=lpFrame->nCaretLine + 1;
            lpFrame->nLineSelEnd=GetLineNumber(lpFrameCurrent, crCurSel.ciMax.nLine) + 1;
          }
          else
          {
            lpFrame->nLineSelBegin=GetLineNumber(lpFrameCurrent, crCurSel.ciMin.nLine) + 1;
            lpFrame->nLineSelEnd=lpFrame->nCaretLine + 1;
          }
          if (AEC_IsFirstCharInLine(&crCurSel.ciMax))
            lpFrame->nLineCountSel=lpFrame->nLineSelEnd - lpFrame->nLineSelBegin;
          else
            lpFrame->nLineCountSel=lpFrame->nLineSelEnd - lpFrame->nLineSelBegin + 1;
        }
        if ((moCur.dwStatusUserFlags & CSB_RICHCOUNT) && (dwFlags & CSB_RICHCOUNT))
          lpFrame->nRichCount=GetTextLength(lpFrame->ei.hWndEdit);
        if ((moCur.dwStatusUserFlags & CSB_FONTPOINT) && (dwFlags & CSB_FONTPOINT))
          lpFrame->nFontPoint=(int)SendMessage(lpFrame->ei.hWndEdit, AEM_GETCHARSIZE, AECS_POINTSIZE, (LPARAM)NULL);
      }
      if ((moCur.dwStatusUserFlags & CSB_CAPSLOCK) && (dwFlags & CSB_CAPSLOCK))
        lpFrame->bCapsLock=(GetKeyState(VK_CAPITAL) & 1) == 1;
      if ((moCur.dwStatusUserFlags & CSB_NUMLOCK) && (dwFlags & CSB_NUMLOCK))
        lpFrame->bNumLock=(GetKeyState(VK_NUMLOCK) & 1) == 1;

      //Set status bar parts text
      SetTextStatusUser(lpFrame, dwFlags);
    }
  }
}

void SetTextStatusUser(FRAMEDATA *lpFrame, DWORD dwFlags)
{
  STATUSPART *sp;

  for (sp=hStatusStack.first; sp; sp=sp->next)
  {
    if ((sp->dwFormatFlags & dwFlags) || dwFlags == (DWORD)-1)
    {
      TranslateStatusUser(lpFrame, sp->wpFormat, sp->nFormatLen, wbuf, BUFFER_SIZE);
      StatusBar_SetTextWide(hStatus, sp->nIndex, wbuf);
    }
  }
}

DWORD TranslateStatusUser(FRAMEDATA *lpFrame, const wchar_t *wpString, int nStringLen, wchar_t *wszBuffer, int nBufferSize)
{
  //%[width] - Add status bar delimiter.
  //%Ch - Current character hex code in file codepage (lowercase).
  //%CH - Current character hex code in file codepage (uppercase).
  //%ch - Current character hex code (lowercase).
  //%cH - Current character hex code (uppercase).
  //%cd - Current character decimal code.
  //%cl - Current character letter.
  //%or - Offset in symbols. Any newline breaks counted as one symbol (RichEdit).
  //%ob - Offset in symbols. Newline breaks: \r\r\n - three symbols, \r\n - two symbols, \r - one symbol, \n - one symbol.
  //%al - Count of lines in document.
  //%ar - Count of symbols in document (RichEdit).
  //%lb - Number of first selected line.
  //%le - Number of last selected line.
  //%ls - Count of lines in selection.
  //%f - Font size.
  //%t - Tabulation size.
  //%m - Column marker position.
  //%cap[text] - Text displayed when "Caps Lock" is on.
  //%num[text] - Text displayed when "Num Lock" is on.
  //%se[text] - Text displayed when end of the document reached during search.
  //%r - Replace count after "Replace all".
  //%dc - Count of all documents (MDI/PMDI).
  //%dm - Count of modified documents (MDI/PMDI).
  //%ds - Count of unmodified documents (MDI/PMDI).
  //%di - Active document index (MDI/PMDI).
  //%% - % symbol.

  STATUSPART *sp=NULL;
  const wchar_t *wpStringMax=wpString + nStringLen;
  int nPartIndex=SBP_USER;
  DWORD dwFlags=0;
  DWORD i;

  if (!wszBuffer)
  {
    if (sp=StackStatusPartInsert(&hStatusStack))
    {
      sp->wpFormat=wpString;
      sp->nFormatLen=(int)(wpStringMax - wpString);
      sp->nPartSize=-1;
      sp->nIndex=nPartIndex++;
    }
  }

  for (i=0; wpString < wpStringMax; ++wpString)
  {
    if (*wpString == L'%')
    {
      if (*++wpString == L'%')
      {
        if (lpFrame && wszBuffer) wszBuffer[i]=L'%';
        ++i;
      }
      else if (*wpString == L'[')
      {
        const wchar_t *wpStrStart=wpString - 1;
        const wchar_t *wpStrEnd;
        int nPartSize;

        nPartSize=(int)xatoiW(wpString + 1, &wpString);
        for (wpStrEnd=wpString; *wpStrEnd != L']' && *wpStrEnd; ++wpStrEnd);
        wpString=wpStrEnd;

        if (nPartSize && sp)
        {
          sp->nPartSize=nPartSize;
          sp->nFormatLen=(int)(wpStrStart - sp->wpFormat);
          sp->dwFormatFlags=dwFlags;
          dwFlags=0;

          if (sp=StackStatusPartInsert(&hStatusStack))
          {
            sp->wpFormat=wpString + 1;
            sp->nFormatLen=(int)(wpStringMax - sp->wpFormat);
            sp->nPartSize=-1;
            sp->nIndex=nPartIndex++;
          }
        }
      }
      else if (*wpString == L'C')
      {
        if (*++wpString == L'h' || *wpString == L'H')
        {
          if (lpFrame)
          {
            wchar_t wszChar[10];
            char szChar[20];
            int nCharLen;
            int nCharCount;

            if (lpFrame->nCaretChar == 0xFFFF)
            {
              if (lpFrame->ei.nCodePage == CP_UNICODE_UTF32LE ||
                  lpFrame->ei.nCodePage == CP_UNICODE_UTF32BE)
                xstrcpyW(wszChar, L"FFFFFFFF");
              else
                xstrcpyW(wszChar, L"FFFF");
            }
            else
            {
              if (lpFrame->nCaretChar >= 0x10000) //Unicode surrogate pair
              {
                wszChar[0]=AEC_HighSurrogateFromScalar(lpFrame->nCaretChar);
                wszChar[1]=AEC_LowSurrogateFromScalar(lpFrame->nCaretChar);
                nCharLen=2;
              }
              else
              {
                wszChar[0]=(wchar_t)lpFrame->nCaretChar;
                nCharLen=1;
              }

              if (lpFrame->ei.nCodePage == CP_UNICODE_UTF16LE ||
                  lpFrame->ei.nCodePage == CP_UNICODE_UTF16BE)
              {
                nCharLen*=sizeof(wchar_t);
                xmemcpy(szChar, wszChar, nCharLen);
                if (lpFrame->ei.nCodePage == CP_UNICODE_UTF16BE)
                  ChangeTwoBytesOrder((unsigned char *)szChar, nCharLen, NULL);
              }
              else if (lpFrame->ei.nCodePage == CP_UNICODE_UTF32LE ||
                       lpFrame->ei.nCodePage == CP_UNICODE_UTF32BE)
              {
                nCharLen=(int)UTF16toUTF32(wszChar, nCharLen, NULL, (unsigned long *)szChar, sizeof(szChar)) * sizeof(unsigned long);
                if (lpFrame->ei.nCodePage == CP_UNICODE_UTF32BE)
                  ChangeFourBytesOrder((unsigned char *)szChar, nCharLen, NULL);
              }
              else nCharLen=WideCharToMultiByte(lpFrame->ei.nCodePage, 0, wszChar, nCharLen, szChar, sizeof(szChar), NULL, NULL);

              for (nCharCount=0; nCharCount < nCharLen; ++nCharCount)
                dec2hexW((unsigned char)szChar[nCharCount], wszChar + nCharCount * 2, 2, (*wpString == L'h'));
              wszChar[nCharCount * 2]=L'\0';
            }
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%s", wszChar);
          }
          else dwFlags|=CSB_CHARHEX;
        }
      }
      else if (*wpString == L'c')
      {
        if (*++wpString == L'h' || *wpString == L'H')
        {
          if (lpFrame)
          {
            if (lpFrame->nCaretChar >= 0x10000) //Unicode scalar value of surrogate pair
              i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, (*wpString == L'h')?L"%06x":L"%06X", lpFrame->nCaretChar);
            else
              i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, (*wpString == L'h')?L"%04x":L"%04X", lpFrame->nCaretChar);
          }
          else dwFlags|=CSB_CHARHEX;
        }
        else if (*wpString == L'd')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", (lpFrame->nCaretChar == 0xFFFF)?-1:lpFrame->nCaretChar);
          else
            dwFlags|=CSB_CHARDEC;
        }
        else if (*wpString == L'l')
        {
          if (lpFrame)
          {
            if (lpFrame->nCaretChar >= 0x10000) //Unicode scalar value of surrogate pair
              i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%c%c", AEC_HighSurrogateFromScalar(lpFrame->nCaretChar), AEC_LowSurrogateFromScalar(lpFrame->nCaretChar));
            else
              i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%c", (lpFrame->nCaretChar == 0xFFFF || lpFrame->nCaretChar == L'\r' || lpFrame->nCaretChar == L'\t')?L' ':lpFrame->nCaretChar);
          }
          else dwFlags|=CSB_CHARLETTER;
        }
        else if (*wpString == L'a')
        {
          if (*++wpString == L'p' && *++wpString == L'[')
          {
            const wchar_t *wpStrEnd;

            for (wpStrEnd=++wpString; *wpStrEnd != L']' && *wpStrEnd; ++wpStrEnd);
            if (lpFrame)
              i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%.%ds", lpFrame->bCapsLock?min(wpStrEnd - wpString, 64):0, wpString);
            else
              dwFlags|=CSB_CAPSLOCK;

            wpString=wpStrEnd;
          }
        }
      }
      else if (*wpString == L'n')
      {
        if (*++wpString == L'u')
        {
          if (*++wpString == L'm' && *++wpString == L'[')
          {
            const wchar_t *wpStrEnd;

            for (wpStrEnd=++wpString; *wpStrEnd != L']' && *wpStrEnd; ++wpStrEnd);
            if (lpFrame)
              i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%.%ds", lpFrame->bNumLock?min(wpStrEnd - wpString, 64):0, wpString);
            else
              dwFlags|=CSB_NUMLOCK;

            wpString=wpStrEnd;
          }
        }
      }
      else if (*wpString == L'o')
      {
        if (*++wpString == L'r')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%Id", lpFrame->nCaretRichOffset);
          else
            dwFlags|=CSB_RICHOFFSET;
        }
        else if (*wpString == L'b')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%Id", lpFrame->nCaretByteOffset);
          else
            dwFlags|=CSB_BYTEOFFSET;
        }
      }
      else if (*wpString == L'a')
      {
        if (*++wpString == L'l')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", lpFrame->nLineCountAll);
          else
            dwFlags|=CSB_LINEALLCOUNT;
        }
        else if (*wpString == L'r')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", lpFrame->nRichCount);
          else
            dwFlags|=CSB_RICHCOUNT;
        }
      }
      else if (*wpString == L'l')
      {
        if (*++wpString == L's')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", lpFrame->nLineCountSel);
          else
            dwFlags|=CSB_LINESELCOUNT;
        }
        else if (*wpString == L'b')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", lpFrame->nLineSelBegin);
          else
            dwFlags|=CSB_LINESELBEGIN;
        }
        else if (*wpString == L'e')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", lpFrame->nLineSelEnd);
          else
            dwFlags|=CSB_LINESELEND;
        }
      }
      else if (*wpString == L'f')
      {
        if (lpFrame)
          i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", lpFrame->nFontPoint);
        else
          dwFlags|=CSB_FONTPOINT;
      }
      else if (*wpString == L't')
      {
        if (lpFrame)
          i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", lpFrame->nTabStopSize);
        else
          dwFlags|=CSB_TABSIZE;
      }
      else if (*wpString == L'm')
      {
        if (lpFrame)
          i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", lpFrame->dwMarker);
        else
          dwFlags|=CSB_MARKER;
      }
      else if (*wpString == L's')
      {
        if (*++wpString == L'e' && *++wpString == L'[')
        {
          const wchar_t *wpStrEnd;

          for (wpStrEnd=++wpString; *wpStrEnd != L']' && *wpStrEnd; ++wpStrEnd);
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%.%ds", lpFrame->bReachedEOF?min(wpStrEnd - wpString, 64):0, wpString);
          else
            dwFlags|=CSB_SEARCHENDED;

          wpString=wpStrEnd;
        }
      }
      else if (*wpString == L'r')
      {
        if (lpFrame)
          i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%Id", lpFrame->nReplaceCount);
        else
          dwFlags|=CSB_REPLACECOUNT;
      }
      else if (*wpString == L'd')
      {
        if (*++wpString == L'c')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", nDocumentsCount);
          else
            dwFlags|=CSB_DOCUMENTSCOUNT;
        }
        else if (*wpString == L'm')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", nDocumentsModified);
          else
            dwFlags|=CSB_DOCUMENTSMODIFIED;
        }
        else if (*wpString == L's')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", nDocumentsCount - nDocumentsModified);
          else
            dwFlags|=CSB_DOCUMENTSSAVED;
        }
        else if (*wpString == L'i')
        {
          if (lpFrame)
            i+=(DWORD)xprintfW(wszBuffer?wszBuffer + i:NULL, L"%d", nDocumentIndex + 1);
          else
            dwFlags|=CSB_DOCUMENTINDEX;
        }
      }
      else break;
    }
    else
    {
      if (lpFrame && wszBuffer) wszBuffer[i]=*wpString;
      ++i;
    }
  }
  if (lpFrame && wszBuffer) wszBuffer[i]=L'\0';
  if (sp) sp->dwFormatFlags=dwFlags;
  return i;
}

int* SetStatusParts(STACKSTATUSPART *lpStatusStack)
{
  int *lpSBParts;
  int nStatusParts=SBP_USER + lpStatusStack->nElements;

  if (lpSBParts=(int *)API_HeapAlloc(hHeap, HEAP_ZERO_MEMORY, nStatusParts * sizeof(int)))
  {
    lpSBParts[SBP_POSITION]=110;
    lpSBParts[SBP_MODIFY]=220;
    lpSBParts[SBP_INSERT]=250;
    lpSBParts[SBP_NEWLINE]=280;
    lpSBParts[SBP_CODEPAGE]=-1;

    //Set user parts
    if (lpStatusStack->nElements)
    {
      STATUSPART *sp;
      int nPartIndex=SBP_USER;
      int nPartSize;

      lpSBParts[SBP_CODEPAGE]=560;
      nPartSize=lpSBParts[SBP_CODEPAGE];

      for (sp=lpStatusStack->first; sp; sp=sp->next)
      {
        if (sp->nPartSize == -1)
        {
          lpSBParts[nPartIndex++]=-1;
          break;
        }
        nPartSize+=sp->nPartSize;
        lpSBParts[nPartIndex++]=nPartSize;
      }
    }
    SendMessage(hStatus, SB_SETPARTS, nStatusParts, (LPARAM)lpSBParts);
  }
  return lpSBParts;
}

STATUSPART* StackStatusPartInsert(STACKSTATUSPART *hStack)
{
  STATUSPART *lpElement=NULL;

  if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, -1, sizeof(STATUSPART)))
  {
    ++hStack->nElements;
    return lpElement;
  }
  return NULL;
}

void StackStatusPartFree(STACKSTATUSPART *hStack)
{
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
  hStack->nElements=0;
}


//// Associations

const wchar_t* GetAssociatedIconW(const wchar_t *wpFile, wchar_t *wszIconFile, int *nIconIndex, HICON *phiconLarge, HICON *phiconSmall)
{
  wchar_t wszRoot[MAX_PATH];
  wchar_t wszKey[MAX_PATH];
  wchar_t wszValue[MAX_PATH];
  wchar_t wszTemp[MAX_PATH];
  wchar_t *wpFileName;
  const wchar_t *wpExt;
  HKEY hKey;
  DWORD dwType;
  DWORD dwSize;
  int nIndex=0;
  int i;
  int j;

  if (wpExt=GetFileExt(wpFile, -1))
  {
    if (RegOpenKeyExWide(HKEY_CLASSES_ROOT, wpExt - 1, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
      wszRoot[0]=L'\0';
      dwSize=BUFFER_SIZE * sizeof(wchar_t);
      RegQueryValueExWide(hKey, L"", NULL, &dwType, (LPBYTE)wszRoot, &dwSize);
      RegCloseKey(hKey);

      if (wszRoot[0])
      {
        //Get DefaultIcon string
        xprintfW(wszKey, L"%s\\DefaultIcon", wszRoot);

        if (RegOpenKeyExWide(HKEY_CLASSES_ROOT, wszKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
        {
          dwSize=BUFFER_SIZE * sizeof(wchar_t);
          RegQueryValueExWide(hKey, L"", NULL, &dwType, (LPBYTE)wszValue, &dwSize);
          RegCloseKey(hKey);

          for (i=0, j=0; wszValue[i]; ++i)
          {
            if (wszValue[i] != L'\"')
              wszValue[j++]=wszValue[i];
          }
          wszValue[j]=L'\0';
        }

        //If no DefaultIcon or if it equal to %1
        if (!hKey || !xstrcmpW(wszValue, L"%1"))
        {
          xprintfW(wszKey, L"%s\\shell\\open\\command", wszRoot);

          if (RegOpenKeyExWide(HKEY_CLASSES_ROOT, wszKey, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
          {
            dwSize=BUFFER_SIZE * sizeof(wchar_t);
            RegQueryValueExWide(hKey, L"", NULL, &dwType, (LPBYTE)wszValue, &dwSize);
            RegCloseKey(hKey);

            for (i=0, j=0; wszValue[i]; ++i)
            {
              if (wszValue[i] != L'\"')
                wszValue[j++]=wszValue[i];
              else
                if (i != 0) break;
            }
            wszValue[j]=L'\0';
          }

          if (!hKey || !xstrcmpW(wszValue, L"%1"))
          {
            xstrcpynW(wszValue, wpFile, MAX_PATH);
            wpExt=NULL;
          }
        }

        //Extract icons
        for (i=(int)xstrlenW(wszValue) - 1; i > 0; --i)
        {
          if (wszValue[i] == L',')
          {
            wszValue[i]=L'\0';
            nIndex=(int)xatoiW(wszValue + i + 1, NULL);
            if (nIndex == -1) nIndex=0;
            break;
          }
          else if (wszValue[i] == L'-' || (wszValue[i] >= L'0' && wszValue[i] <= L'9'))
            continue;
          break;
        }
        ExpandEnvironmentStringsWide(wszValue, wszTemp, MAX_PATH);
        if (SearchPathWide(NULL, wszTemp, NULL, MAX_PATH, wszValue, &wpFileName))
        {
          ExtractIconExWide(wszValue, nIndex, phiconLarge, phiconSmall, 1);

          if (wszIconFile) xstrcpynW(wszIconFile, wszValue, MAX_PATH);
          if (nIconIndex) *nIconIndex=nIndex;
          return wpExt;
        }
      }
    }
  }
  if (wszIconFile) wszIconFile[0]=L'\0';
  if (nIconIndex) *nIconIndex=0;
  if (phiconLarge) *phiconLarge=0;
  if (phiconSmall) *phiconSmall=0;
  return NULL;
}

void AssociateFileTypesW(HINSTANCE hInstance, const wchar_t *wpFileTypes, DWORD dwFlags)
{
  wchar_t wszModule[MAX_PATH];
  wchar_t wszAssocKey[MAX_PATH];
  wchar_t wszTypeKey[16];
  wchar_t wszExt[16];
  const wchar_t *wpExtStart=wpFileTypes;
  const wchar_t *wpExtEnd=wpFileTypes;
  HKEY hKey;
  DWORD dwType;
  DWORD dwSize;

  GetModuleFileNameWide(hInstance, wszModule, MAX_PATH);

  if (dwFlags & FTA_OPEN) xstrcpyW(wszTypeKey, L"Open");
  else if (dwFlags & FTA_EDIT) xstrcpyW(wszTypeKey, L"Edit");
  else if (dwFlags & FTA_PRINT) xstrcpyW(wszTypeKey, L"Print");

  for (;;)
  {
    if (wpExtEnd=AKD_wcschr(wpExtStart, L';'))
      xstrcpynW(wbuf, wpExtStart, wpExtEnd - wpExtStart + 1);
    else
      xstrcpynW(wbuf, wpExtStart, MAX_PATH);
    if (!*wbuf) goto NextExt;
    xprintfW(wszExt, L".%s", wbuf);

    if (RegCreateKeyExWide(HKEY_CLASSES_ROOT, wszExt, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
      dwSize=MAX_PATH * sizeof(wchar_t);
      if (RegQueryValueExWide(hKey, L"", NULL, &dwType, (LPBYTE)wszAssocKey, &dwSize) != ERROR_SUCCESS || !*wszAssocKey)
      {
        if (dwFlags & FTA_ASSOCIATE)
        {
          xprintfW(wszAssocKey, L"%sfile", wszExt + 1);
          RegSetValueExWide(hKey, L"", 0, REG_SZ, (LPBYTE)wszAssocKey, BytesInString(wszAssocKey));
        }
        else if (dwFlags & FTA_DEASSOCIATE)
        {
          RegCloseKey(hKey);
          return;
        }
      }
      RegCloseKey(hKey);

      //Associate icon
      if (dwFlags & FTA_OPEN)
      {
        xprintfW(wbuf2, L"%s\\DefaultIcon", wszAssocKey);

        if (RegCreateKeyExWide(HKEY_CLASSES_ROOT, wbuf2, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL) == ERROR_SUCCESS)
        {
          dwSize=BUFFER_SIZE * sizeof(wchar_t);
          if (RegQueryValueExWide(hKey, L"", NULL, &dwType, (LPBYTE)wbuf2, &dwSize) != ERROR_SUCCESS || !*wbuf2)
          {
            if (dwFlags & FTA_ASSOCIATE) RegSetValueExWide(hKey, L"", 0, REG_SZ, (LPBYTE)wszModule, BytesInString(wszModule));
          }
          else if (dwFlags & FTA_DEASSOCIATE)
          {
            if (!xstrcmpiW(wbuf2, wszModule)) RegDeleteValueWide(hKey, L"");
          }
          RegCloseKey(hKey);
        }
      }

      //Associate command
      xprintfW(wbuf2, L"%s\\Shell\\%s\\Command", wszAssocKey, wszTypeKey);
      xprintfW(wbuf, L"\"%s\"%s \"%%1\"", wszModule, (dwFlags & FTA_PRINT)?L" /p":L"");

      if (RegCreateKeyExWide(HKEY_CLASSES_ROOT, wbuf2, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, NULL, &hKey, NULL) == ERROR_SUCCESS)
      {
        if (dwFlags & FTA_ASSOCIATE)
        {
          dwSize=BUFFER_SIZE * sizeof(wchar_t);
          if (RegQueryValueExWide(hKey, L"", NULL, &dwType, (LPBYTE)wbuf2, &dwSize) == ERROR_SUCCESS)
          {
            if (*wbuf2 && xstrcmpiW(wbuf, wbuf2))
            {
              if (RegQueryValueExWide(hKey, L"AkelUndo", NULL, NULL, NULL, NULL) != ERROR_SUCCESS)
                RegSetValueExWide(hKey, L"AkelUndo", 0, dwType, (LPBYTE)wbuf2, dwSize);
            }
          }
          RegSetValueExWide(hKey, L"", 0, REG_SZ, (LPBYTE)wbuf, BytesInString(wbuf));
        }
        else if (dwFlags & FTA_DEASSOCIATE)
        {
          dwSize=BUFFER_SIZE * sizeof(wchar_t);
          if (RegQueryValueExWide(hKey, L"AkelUndo", NULL, &dwType, (LPBYTE)wbuf2, &dwSize) == ERROR_SUCCESS)
          {
            if (*wbuf2) RegSetValueExWide(hKey, L"", 0, dwType, (LPBYTE)wbuf2, dwSize);
            RegDeleteValueWide(hKey, L"AkelUndo");
          }
          dwSize=BUFFER_SIZE * sizeof(wchar_t);
          if (RegQueryValueExWide(hKey, L"", NULL, &dwType, (LPBYTE)wbuf2, &dwSize) == ERROR_SUCCESS)
          {
            if (!*wbuf2 || !xstrcmpiW(wbuf, wbuf2))
            {
              RegCloseKey(hKey);
              xprintfW(wbuf2, L"%s\\Shell\\%s", wszAssocKey, wszTypeKey);
              if (RegOpenKeyExWide(HKEY_CLASSES_ROOT, wbuf2, 0, KEY_ENUMERATE_SUB_KEYS, &hKey) == ERROR_SUCCESS)
              {
                RegDeleteKeyWide(hKey, L"Command");
                RegCloseKey(hKey);
                xprintfW(wbuf2, L"%s\\Shell", wszAssocKey);
                if (RegOpenKeyExWide(HKEY_CLASSES_ROOT, wbuf2, 0, KEY_ENUMERATE_SUB_KEYS, &hKey) == ERROR_SUCCESS)
                {
                  RegDeleteKeyWide(hKey, wszTypeKey);
                  dwSize=BUFFER_SIZE * sizeof(wchar_t);
                  if (RegEnumKeyExW(hKey, 0, wbuf2, &dwSize, NULL, NULL, NULL, NULL) != ERROR_SUCCESS)
                  {
                    RegCloseKey(hKey);
                    if (RegOpenKeyExWide(HKEY_CLASSES_ROOT, wszAssocKey, 0, KEY_ENUMERATE_SUB_KEYS, &hKey) == ERROR_SUCCESS)
                      RegDeleteKeyWide(hKey, L"Shell");
                  }
                }
              }
            }
          }
        }
        RegCloseKey(hKey);
      }
    }

    //Check off "Always open with this program"
    if (dwFlags & FTA_OPEN)
    {
      xprintfW(wbuf, L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\%s", wszExt);

      if (RegOpenKeyExWide(HKEY_CURRENT_USER, wbuf, 0, KEY_QUERY_VALUE|KEY_SET_VALUE, &hKey) == ERROR_SUCCESS)
      {
        if (dwFlags & FTA_ASSOCIATE)
        {
          dwSize=BUFFER_SIZE * sizeof(wchar_t);
          if (RegQueryValueExWide(hKey, L"Application", NULL, &dwType, (LPBYTE)wbuf, &dwSize) == ERROR_SUCCESS)
          {
            if (*wbuf) RegSetValueExWide(hKey, L"AkelUndo", 0, dwType, (LPBYTE)wbuf, dwSize);
            RegDeleteValueWide(hKey, L"Application");
          }
        }
        else if (dwFlags & FTA_DEASSOCIATE)
        {
          dwSize=BUFFER_SIZE * sizeof(wchar_t);
          if (RegQueryValueExWide(hKey, L"AkelUndo", NULL, &dwType, (LPBYTE)wbuf, &dwSize) == ERROR_SUCCESS)
          {
            if (*wbuf) RegSetValueExWide(hKey, L"Application", 0, dwType, (LPBYTE)wbuf, dwSize);
            RegDeleteValueWide(hKey, L"AkelUndo");
          }
        }
        RegCloseKey(hKey);
      }
    }
    NextExt:
    wpExtStart=wpExtEnd + 1;
    if (!wpExtEnd) break;
    if (!*wpExtStart) break;
  }
}

ASSOCICON* StackIconInsert(STACKASSOCICON *hStack, const wchar_t *wpFile, int nFileLen)
{
  ASSOCICON *lpElement=NULL;

  if (nFileLen == -1)
    nFileLen=(int)xstrlenW(wpFile);

  if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, -1, sizeof(ASSOCICON)))
  {
    if (wpFile)
    {
      xstrcpynW(lpElement->wszFile, wpFile, MAX_PATH);
      lpElement->nFileLen=nFileLen;
      lpElement->wpExt=GetAssociatedIconW(lpElement->wszFile, NULL, NULL, NULL, &lpElement->hIcon);
    }
    if (lpElement->hIcon)
    {
      ++hStack->nValidIcons;
      lpElement->nIconIndex=hStack->nValidIcons;
    }
    return lpElement;
  }
  return NULL;
}

ASSOCICON* StackIconGet(STACKASSOCICON *hStack, const wchar_t *wpFile, int nFileLen, const wchar_t *wpExt)
{
  ASSOCICON *lpElement=hStack->first->next;

  if (nFileLen == -1)
    nFileLen=(int)xstrlenW(wpFile);

  while (lpElement)
  {
    if ((lpElement->wpExt && !xstrcmpiW(lpElement->wpExt, wpExt)) ||
        (!lpElement->wpExt && lpElement->nFileLen == nFileLen &&
                              !xstrcmpiW(lpElement->wszFile, wpFile)))
    {
      return lpElement;
    }
    lpElement=lpElement->next;
  }
  return NULL;
}

void StackIconsFree(STACKASSOCICON *hStack)
{
  ASSOCICON *lpElement=(ASSOCICON *)hStack->first->next;

  while (lpElement)
  {
    if (lpElement->hIcon) DestroyIcon(lpElement->hIcon);

    lpElement=lpElement->next;
  }
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
  hStack->nValidIcons=0;
}


//// Fonts

//For WMD_PMDI required: lpFrame == lpFrameCurrent
HFONT SetChosenFont(HWND hWnd, const LOGFONTW *lfFont)
{
  FONTITEM *fi;

  if (!(fi=StackFontItemGet(&hFontsStack, lfFont)))
    fi=StackFontItemInsert(&hFontsStack, lfFont);
  SendMessage(hWnd, WM_SETFONT, (WPARAM)fi->hFont, FALSE);
  return fi->hFont;
}

FONTITEM* StackFontItemInsert(HSTACK *hStack, const LOGFONTW *lfFont)
{
  FONTITEM *lpElement=NULL;

  if (!StackInsertIndex((stack **)&hStack->first, (stack **)&hStack->last, (stack **)&lpElement, -1, sizeof(FONTITEM)))
  {
    xmemcpy(&lpElement->lfFont, lfFont, sizeof(LOGFONTW));
    lpElement->hFont=(HFONT)CreateFontIndirectWide(&lpElement->lfFont);
    return lpElement;
  }
  return NULL;
}

FONTITEM* StackFontItemGet(HSTACK *hStack, const LOGFONTW *lfFont)
{
  FONTITEM *lpElement=(FONTITEM *)hStack->first;

  while (lpElement)
  {
    if (lpElement->lfFont.lfHeight == lfFont->lfHeight &&
        lpElement->lfFont.lfWeight == lfFont->lfWeight &&
        lpElement->lfFont.lfItalic == lfFont->lfItalic &&
        lpElement->lfFont.lfCharSet == lfFont->lfCharSet)
    {
      if (!xstrcmpiW(lpElement->lfFont.lfFaceName, lfFont->lfFaceName))
        return lpElement;
    }
    lpElement=lpElement->next;
  }
  return NULL;
}

void StackFontItemsFree(HSTACK *hStack)
{
  FONTITEM *lpElement=(FONTITEM *)hStack->first;

  while (lpElement)
  {
    if (lpElement->hFont) DeleteObject(lpElement->hFont);

    lpElement=lpElement->next;
  }
  StackClear((stack **)&hStack->first, (stack **)&hStack->last);
}


//// Inter-Process Communication

HANDLE MemCreate(const char *pName, DWORD dwSize)
{
  HANDLE hMem;

  if (!dwSize)
    hMem=OpenFileMappingA(FILE_MAP_ALL_ACCESS, FALSE, pName);
  else
    hMem=CreateFileMappingA(INVALID_HANDLE_VALUE, NULL, PAGE_READWRITE, 0, dwSize, pName);
  if (hMem == INVALID_HANDLE_VALUE)
  {
    //GetLastError();
  }
  return hMem;
}

void* MemMap(HANDLE hMem, DWORD dwSize)
{
  return MapViewOfFile(hMem, FILE_MAP_ALL_ACCESS, 0, 0, dwSize);
}

BOOL MemUnmap(void *lpMem)
{
  return UnmapViewOfFile(lpMem);
}

BOOL MemClose(HANDLE hMem)
{
  return CloseHandle(hMem);
}


//// AkelAdmin

BOOL AkelAdminInit(const wchar_t *wpFile)
{
  if (!wszAkelAdminPipe[0])
    xprintfW(wszAkelAdminExe, L"%s\\AkelFiles\\AkelAdmin.exe", wszExeDir);

  if (FileExistsWide(wszAkelAdminExe))
  {
    //Custom MessageBox
    BUTTONMESSAGEBOX bmb[]={{IDOK,     MAKEINTRESOURCEW(STR_MESSAGEBOX_CONTINUE), BMB_DEFAULT},
                            {IDCANCEL, MAKEINTRESOURCEW(STR_MESSAGEBOX_CANCEL),   0},
                            {0, 0, 0}};
    HMODULE hAdvApi32;
    HMODULE hShell32;

    if (!wszAkelAdminPipe[0])
      xprintfW(wszAkelAdminPipe, L"\\\\.\\pipe\\%s-%d", STR_AKELADMINW, dwProcessId);
    if (!hIconShieldAkelAdmin)
      ExtractIconExWide(wszAkelAdminExe, 1, &hIconShieldAkelAdmin, NULL, 1);

    //Get functions addresses
    if (!SetSecurityInfoPtr || !SetEntriesInAclWPtr)
    {
      hAdvApi32=GetModuleHandleA("advapi32.dll");
      SetSecurityInfoPtr=(DWORD (WINAPI *)(HANDLE, SE_OBJECT_TYPE, SECURITY_INFORMATION, PSID, PSID, PACL, PACL))GetProcAddress(hAdvApi32, "SetSecurityInfo");
      SetEntriesInAclWPtr=(DWORD (WINAPI *)(ULONG, PEXPLICIT_ACCESSW, PACL, PACL *))GetProcAddress(hAdvApi32, "SetEntriesInAclW");
    }
    if (!ShellExecuteExWPtr)
    {
      hShell32=GetModuleHandleA("shell32.dll");
      ShellExecuteExWPtr=(BOOL (WINAPI *)(LPSHELLEXECUTEINFOW))GetProcAddress(hShell32, "ShellExecuteExW");
    }

    API_LoadStringW(hLangLib, MSG_ACCESSDENIED, wbuf, BUFFER_SIZE);
    xprintfW(wszMsg, wbuf, wpFile);
    if (MessageBoxCustom(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_ICONEXCLAMATION, hIconShieldAkelAdmin, &bmb[0]) == IDOK)
    {
      if (!bPipeInitAkelAdmin)
      {
        wchar_t wszParams[MAX_PATH];
        SHELLEXECUTEINFOW sei;
        HANDLE lpHandles[2];
        HANDLE hMutex;

        //Pipe server doesn't exist
        if (hMutex=CreateEventW(NULL, FALSE, FALSE, STR_AKELADMINW))
        {
          //Set security for hMutex. It required under limited user of WinXP.
          if (SetSecurityInfoPtr && SetEntriesInAclWPtr)
          {
            ACL *pNewACL=NULL;
            EXPLICIT_ACCESSW eal[1];
            SID_IDENTIFIER_AUTHORITY SIDAuthWorld=SECURITY_WORLD_SID_AUTHORITY;
            SID *pSIDEveryone=NULL;

            //Specify the DACL to use. Create a SID for the Everyone group.
            if (AllocateAndInitializeSid(&SIDAuthWorld, 1, SECURITY_WORLD_RID, 0, 0, 0, 0, 0, 0, 0, (void **)&pSIDEveryone))
            {
              eal[0].grfAccessPermissions=GENERIC_ALL;
              eal[0].grfAccessMode=SET_ACCESS;
              eal[0].grfInheritance=NO_INHERITANCE;
              eal[0].Trustee.TrusteeForm=TRUSTEE_IS_SID;
              eal[0].Trustee.TrusteeType=TRUSTEE_IS_WELL_KNOWN_GROUP;
              eal[0].Trustee.ptstrName=(wchar_t *)pSIDEveryone;
              eal[0].Trustee.MultipleTrusteeOperation=NO_MULTIPLE_TRUSTEE;
              eal[0].Trustee.pMultipleTrustee=NULL;

              if ((*SetEntriesInAclWPtr)(1, eal, NULL, &pNewACL) == ERROR_SUCCESS)
              {
                (*SetSecurityInfoPtr)(hMutex, SE_KERNEL_OBJECT, DACL_SECURITY_INFORMATION, NULL, NULL, pNewACL, NULL);
                LocalFree(pNewACL);
              }
              FreeSid(pSIDEveryone);
            }
          }

          //Initialize AkelAdmin process
          if (ShellExecuteExWPtr)
          {
            xprintfW(wszParams, L"\"%d\" \"%d\"", AAA_CMDINIT, dwProcessId);
            sei.cbSize=sizeof(SHELLEXECUTEINFOW);
            sei.fMask=SEE_MASK_NOCLOSEPROCESS;
            sei.hwnd=hMainWnd;
            sei.lpVerb=L"runas";
            sei.lpFile=wszAkelAdminExe;
            sei.lpParameters=wszParams;
            sei.lpDirectory=NULL;
            sei.nShow=SW_SHOWDEFAULT;

            if ((*ShellExecuteExWPtr)(&sei))
            {
              lpHandles[0]=hMutex;
              lpHandles[1]=sei.hProcess;

              //Wait for mutex signal or process exit
              if (WaitForMultipleObjects(2, lpHandles, FALSE, INFINITE) == WAIT_OBJECT_0)
              {
                bPipeInitAkelAdmin=TRUE;

                //Return focus if needed
                if (lpFrameCurrent->ei.hWndEdit && GetFocus() != lpFrameCurrent->ei.hWndEdit)
                  SetFocus(lpFrameCurrent->ei.hWndEdit);
              }
              CloseHandle(sei.hProcess);
            }
          }
          CloseHandle(hMutex);
        }
      }
      return bPipeInitAkelAdmin;
    }
  }
  return FALSE;
}

BOOL AkelAdminSend(int nAction, const wchar_t *wpFile)
{
  if (bPipeInitAkelAdmin)
  {
    ADMINPIPE apipe;
    HANDLE hFilePipe;
    UINT_PTR dwBytesRead;
    UINT_PTR dwBytesWritten;

    apipe.dwExitCode=1;
    apipe.nAction=nAction;
    xstrcpynW(apipe.wszFile, wpFile, MAX_PATH);
    apipe.dwLangModule=dwLangModule;

    for (;;)
    {
      //Connect to pipe server, send and receive data.
      if ((hFilePipe=CreateFileW(wszAkelAdminPipe, GENERIC_READ|GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL)) != INVALID_HANDLE_VALUE)
      {
        API_WriteFile(hFilePipe, &apipe, sizeof(ADMINPIPE), &dwBytesWritten, NULL);
        ReadFile64(hFilePipe, &apipe, sizeof(ADMINPIPE), &dwBytesRead, NULL);
        CloseHandle(hFilePipe);
        break;
      }
      else if (GetLastError() == ERROR_PIPE_BUSY)
      {
        //Wait until pipe server became free.
        WaitNamedPipeW(wszAkelAdminPipe, NMPWAIT_WAIT_FOREVER);
      }
      else break;
    }
    return !apipe.dwExitCode;
  }
  return FALSE;
}

void AkelAdminExit()
{
  if (bPipeInitAkelAdmin)
  {
    ADMINPIPE apipe;
    HANDLE hFilePipe;
    UINT_PTR dwBytesRead;
    UINT_PTR dwBytesWritten;

    for (;;)
    {
      if ((hFilePipe=CreateFileW(wszAkelAdminPipe, GENERIC_READ|GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL)) != INVALID_HANDLE_VALUE)
      {
        //Unload AkelAdmin.
        xmemset(&apipe, 0, sizeof(ADMINPIPE));
        apipe.nAction=AAA_EXIT;
        API_WriteFile(hFilePipe, &apipe, sizeof(ADMINPIPE), &dwBytesWritten, NULL);
        ReadFile64(hFilePipe, &apipe, sizeof(ADMINPIPE), &dwBytesRead, NULL);
        CloseHandle(hFilePipe);

        bPipeInitAkelAdmin=FALSE;
        break;
      }
      else if (GetLastError() == ERROR_PIPE_BUSY)
      {
        //Wait until pipe server became free.
        WaitNamedPipeW(wszAkelAdminPipe, NMPWAIT_WAIT_FOREVER);
      }
      else break;
    }
  }
}


//// Command line functions

wchar_t* GetCommandLineParamsWide(const unsigned char *pCmdParams, wchar_t **wppCmdParamsStart, wchar_t **wppCmdParamsEnd)
{
  wchar_t *wpCmd=wszCmdLine;
  wchar_t *wpMaxCmd=wszCmdLine + COMMANDLINE_SIZE;

  if (nCmdLineBeginLen)
  {
    wpCmd+=xstrcpynW(wpCmd, wpCmdLineBegin, wpMaxCmd - wpCmd);
    wpCmd+=xstrcpynW(wpCmd, L" ", wpMaxCmd - wpCmd);
  }

  if (wppCmdParamsStart) *wppCmdParamsStart=wpCmd;
  if (bOldWindows)
    wpCmd+=xprintfW(wpCmd, L"%.%dS", wpMaxCmd - wpCmd, pCmdParams);
  else
    wpCmd+=xprintfW(wpCmd, L"%.%ds", wpMaxCmd - wpCmd, pCmdParams);
  if (wppCmdParamsEnd) *wppCmdParamsEnd=wpCmd;

  if (nCmdLineEndLen)
  {
    wpCmd+=xstrcpynW(wpCmd, L" ", wpMaxCmd - wpCmd);
    wpCmd+=xstrcpynW(wpCmd, wpCmdLineEnd, wpMaxCmd - wpCmd);
  }
  return wszCmdLine;
}

char* GetCommandLineParamsA()
{
  char *lpCmdLine=GetCommandLineA();

  if (*lpCmdLine++ == '\"')
    while (*lpCmdLine != '\"' && *lpCmdLine != '\0') ++lpCmdLine;
  else
    while (*lpCmdLine != ' ' && *lpCmdLine != '\0') ++lpCmdLine;
  if (*lpCmdLine != '\0')
    while (*++lpCmdLine == ' ');

  return lpCmdLine;
}

wchar_t* GetCommandLineParamsW()
{
  wchar_t *lpwCmdLine=GetCommandLineW();

  if (*lpwCmdLine++ == L'\"')
    while (*lpwCmdLine != L'\"' && *lpwCmdLine != L'\0') ++lpwCmdLine;
  else
    while (*lpwCmdLine != L' ' && *lpwCmdLine != L'\0') ++lpwCmdLine;
  if (*lpwCmdLine != L'\0')
    while (*++lpwCmdLine == L' ');

  return lpwCmdLine;
}

int GetCommandLineArg(const wchar_t *wpCmdLine, wchar_t *wszArg, int nArgMax, const wchar_t **wpNextArg, BOOL bParseAsNotepad)
{
  const wchar_t *wpCount=wpCmdLine;
  wchar_t *wpArg=wszArg;
  wchar_t *wpArgMax=wszArg + nArgMax - 1;
  wchar_t wchFirstChar;
  wchar_t wchStopChar;
  wchar_t wchQuoteChar=0;

  while (*wpCount == L' ') ++wpCount;

  wchFirstChar=*wpCount;
  if (*wpCount == L'\"' || *wpCount == L'\'' || *wpCount == L'`')
  {
    wchStopChar=*wpCount++;
    wchQuoteChar=wchStopChar;
  }
  else wchStopChar=L' ';

  if (bParseAsNotepad && wchFirstChar != L'/' && wchStopChar == L' ')
  {
    for (; *wpCount != L'/' && *wpCount != L'\"' && *wpCount != L'\0'; ++wpCount)
    {
      if (wpArg < wpArgMax)
      {
        if (wszArg) *wpArg=*wpCount;
        ++wpArg;
      }
    }
    while (wpArg > wszArg && *(wpArg - 1) == L' ')
    {
      --wpArg;
      if (wszArg) *wpArg=L'\0';
    }
  }
  else
  {
    for (; *wpCount; ++wpCount)
    {
      if (!wchQuoteChar || *wpCount == wchQuoteChar)
      {
        if (*wpCount == wchStopChar)
        {
          if (wchStopChar != wchFirstChar && wchStopChar != L' ')
          {
            if (wpArg < wpArgMax)
            {
              if (wszArg) *wpArg=*wpCount;
              ++wpArg;
            }
          }
          break;
        }
        if (!wchQuoteChar && wchFirstChar == L'/')
        {
          if (*wpCount == L'\"' || *wpCount == L'\'' || *wpCount == L'`')
            wchQuoteChar=*wpCount;
          else if (*wpCount == L'(')
            wchStopChar=L')';
        }
        else wchQuoteChar=0;
      }
      if (wpArg < wpArgMax)
      {
        if (wszArg) *wpArg=*wpCount;
        ++wpArg;
      }
    }
  }
  if (*wpCount == wchStopChar)
    ++wpCount;
  if (wszArg) *wpArg=L'\0';
  if (wpNextArg)
    for (*wpNextArg=wpCount; **wpNextArg == L' '; ++*wpNextArg);

  return (int)(wpArg - wszArg);
}

int ParseCmdLine(const wchar_t **wppCmdLine, int nType)
{
  const wchar_t *wpCmdLine;
  const wchar_t *wpCmdLineNext;
  HWND hWndFriend=NULL;
  int nOpen;
  DWORD dwCallMethod;
  BOOL bFileOpenedSDI=FALSE;
  BOOL bIgnoreNextArg=FALSE;

  if (wppCmdLine && *wppCmdLine)
  {
    wpCmdLine=*wppCmdLine;
    wpCmdLineNext=*wppCmdLine;

    for (; GetCommandLineArg(wpCmdLine, wszCmdArg, COMMANDARG_SIZE, &wpCmdLineNext, !(dwCmdLineOptions & CLO_NONOTEPADCMD)); wpCmdLine=wpCmdLineNext)
    {
      if (bIgnoreNextArg)
      {
        bIgnoreNextArg=FALSE;
        continue;
      }
      if (wszCmdArg[0] == L'/')
      {
        //On load
        if (!xstrcmpiW(wszCmdArg, L"/Z"))
        {
          bIgnoreNextArg=TRUE;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/REASSOC"))
        {
          if (moCur.dwFileTypesAssociated & FTA_OPEN)
            AssociateFileTypesW(hInstance, moCur.wszFileTypesOpen, FTA_OPEN|FTA_ASSOCIATE);
          if (moCur.dwFileTypesAssociated & FTA_EDIT)
            AssociateFileTypesW(hInstance, moCur.wszFileTypesEdit, FTA_EDIT|FTA_ASSOCIATE);
          if (moCur.dwFileTypesAssociated & FTA_PRINT)
            AssociateFileTypesW(hInstance, moCur.wszFileTypesPrint, FTA_PRINT|FTA_ASSOCIATE);
          if (moCur.dwFileTypesAssociated) SHChangeNotify(SHCNE_ASSOCCHANGED, 0, 0, 0);
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/DEASSOC"))
        {
          if (moCur.dwFileTypesAssociated & FTA_OPEN)
            AssociateFileTypesW(hInstance, moCur.wszFileTypesOpen, FTA_OPEN|FTA_DEASSOCIATE);
          if (moCur.dwFileTypesAssociated & FTA_EDIT)
            AssociateFileTypesW(hInstance, moCur.wszFileTypesEdit, FTA_EDIT|FTA_DEASSOCIATE);
          if (moCur.dwFileTypesAssociated & FTA_PRINT)
            AssociateFileTypesW(hInstance, moCur.wszFileTypesPrint, FTA_PRINT|FTA_DEASSOCIATE);
          if (moCur.dwFileTypesAssociated) SHChangeNotify(SHCNE_ASSOCCHANGED, 0, 0, 0);
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/QUIT"))
        {
          return PCLE_QUIT;
        }
        if (!xstrcmpiW(wszCmdArg, L"/END"))
        {
          return PCLE_END;
        }
        if (!xstrcmpiW(wszCmdArg, L"/P"))
        {
          dwCmdLineOptions|=CLO_GLOBALPRINT;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/C+"))
        {
          dwCmdLineOptions&=~CLO_MSGOPENCREATENO;
          dwCmdLineOptions|=CLO_MSGOPENCREATEYES;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/C-"))
        {
          dwCmdLineOptions&=~CLO_MSGOPENCREATEYES;
          dwCmdLineOptions|=CLO_MSGOPENCREATENO;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/C"))
        {
          dwCmdLineOptions&=~CLO_MSGOPENCREATEYES & ~CLO_MSGOPENCREATENO;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/B+"))
        {
          dwCmdLineOptions&=~CLO_MSGOPENBINARYNO;
          dwCmdLineOptions|=CLO_MSGOPENBINARYYES;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/B-"))
        {
          dwCmdLineOptions&=~CLO_MSGOPENBINARYYES;
          dwCmdLineOptions|=CLO_MSGOPENBINARYNO;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/B"))
        {
          dwCmdLineOptions&=~CLO_MSGOPENBINARYYES & ~CLO_MSGOPENBINARYNO;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/L+"))
        {
          dwCmdLineOptions&=~CLO_MSGSAVELOSTSYMBOLSNO;
          dwCmdLineOptions|=CLO_MSGSAVELOSTSYMBOLSYES;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/L-"))
        {
          dwCmdLineOptions&=~CLO_MSGSAVELOSTSYMBOLSYES;
          dwCmdLineOptions|=CLO_MSGSAVELOSTSYMBOLSNO;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/L"))
        {
          dwCmdLineOptions&=~CLO_MSGSAVELOSTSYMBOLSYES & ~CLO_MSGSAVELOSTSYMBOLSNO;
          continue;
        }
        if (!xstrcmpiW(wszCmdArg, L"/X"))
        {
          dwCmdLineOptions|=CLO_NONOTEPADCMD;
          continue;
        }
        if (!xstrcmpinW(L"/Show(", wszCmdArg, (UINT_PTR)-1))
        {
          dwCmdShow=(DWORD)xatoiW(wszCmdArg + 6, NULL);
          if (nType != PCL_ONLOAD)
            ShowWindow(hMainWnd, dwCmdShow);
          continue;
        }
        if (nType == PCL_ONLOAD) return PCLE_ONLOAD;

        //Process actions
        if (dwCallMethod=CallMethod(wszCmdArg, L""))
          return dwCallMethod;
        continue;
      }
      if (!*wszCmdArg) continue;

      //Open file
      if (nMDI == WMD_SDI)
      {
        if (moCur.bSingleOpenFile)
        {
          if (GetFullName(wszCmdArg, wszCmdArg, MAX_PATH, NULL))
          {
            if ((hWndFriend=FindWindowExWide(NULL, NULL, APP_SDI_CLASSW, wszCmdArg)) &&
                (hWndFriend=GetParent(hWndFriend)))
            {
              if (hWndFriend != hMainWnd)
              {
                ActivateWindow(hWndFriend);
                SendMessage(hWndFriend, AKD_SETCMDLINEOPTIONS, dwCmdLineOptions, 0);
                SendCmdLine(hWndFriend, wpCmdLine, TRUE, TRUE);
                return PCLE_QUIT;
              }
            }
          }
        }
        if (nType == PCL_ONLOAD) return PCLE_ONLOAD;

        if (!bFileOpenedSDI)
        {
          if (!SaveChanged(0))
            return PCLE_END;
          nOpen=OpenDocument(NULL, wszCmdArg, OD_ADT_BINARY_ERROR|OD_ADT_REG_CODEPAGE|(*wpCmdLineNext?OD_MULTIFILE:0), 0, FALSE);
          if (nOpen != EOD_SUCCESS && nOpen != EOD_MSGNO && nOpen != EOD_WINDOW_EXIST)
            return PCLE_END;
          bFileOpenedSDI=TRUE;
          continue;
        }
        hWndFriend=DoFileNewWindow(STARTF_NOMUTEX);
        SendMessage(hWndFriend, AKD_SETCMDLINEOPTIONS, dwCmdLineOptions, 0);
        SendCmdLine(hWndFriend, wpCmdLine, TRUE, TRUE);
        return PCLE_END;
      }
      if (nType == PCL_ONLOAD) return PCLE_ONLOAD;

      //nMDI == WMD_MDI || nMDI == WMD_PMDI
      nOpen=OpenDocument(NULL, wszCmdArg, OD_ADT_BINARY_ERROR|OD_ADT_REG_CODEPAGE|(*wpCmdLineNext?OD_MULTIFILE:0), 0, FALSE);
      if (nOpen != EOD_SUCCESS && nOpen != EOD_MSGNO && nOpen != EOD_WINDOW_EXIST)
        return PCLE_END;
    }
  }
  return PCLE_END;
}

void SendCmdLine(HWND hWnd, const wchar_t *wpCmdLine, BOOL bPost, BOOL bQuitAsEnd)
{
  COPYDATASTRUCT cds;
  PARSECMDLINEPOSTW *pclp;

  if (pclp=(PARSECMDLINEPOSTW *)GlobalAlloc(GMEM_FIXED, sizeof(PARSECMDLINEPOSTW)))
  {
    pclp->bPostMessage=bPost;
    pclp->nCmdLineLen=(int)xstrcpynW(pclp->szCmdLine, wpCmdLine, COMMANDLINE_SIZE);
    pclp->nWorkDirLen=GetCurrentDirectoryWide(MAX_PATH, pclp->szWorkDir);
    pclp->bQuitAsEnd=bQuitAsEnd;

    cds.dwData=CD_PARSECMDLINEW;
    cds.cbData=sizeof(PARSECMDLINEPOSTW);
    cds.lpData=(PVOID)pclp;
    SendMessage(hWnd, WM_COPYDATA, (WPARAM)hWnd, (LPARAM)&cds);
    GlobalFree((HGLOBAL)pclp);
  }
}

DWORD CallMethod(const wchar_t *wpMethod, const wchar_t *wpUrlLink)
{
  STACKEXTPARAM hParamStack={0};
  EXTPARAM *lpParameter;
  const wchar_t *wpAction;
  DWORD dwAction=0;

  if (*wpMethod == L'/')
    ++wpMethod;

  if (!xstrcmpinW(L"Show(", wpMethod, (UINT_PTR)-1))
  {
    DWORD dwCmdShow=(DWORD)xatoiW(wpMethod + 5, NULL);
    ShowWindow(hMainWnd, dwCmdShow);
  }
  else if (!xstrcmpinW(L"Command(", wpMethod, (UINT_PTR)-1))
  {
    dwAction=EXTACT_COMMAND;
  }
  else if (!xstrcmpinW(L"Call(", wpMethod, (UINT_PTR)-1))
  {
    dwAction=EXTACT_CALL;
  }
  else if (!xstrcmpinW(L"Exec(", wpMethod, (UINT_PTR)-1))
  {
    dwAction=EXTACT_EXEC;
  }
  else if (!xstrcmpinW(L"OpenFile(", wpMethod, (UINT_PTR)-1))
  {
    dwAction=EXTACT_OPENFILE;
  }
  else if (!xstrcmpinW(L"SaveFile(", wpMethod, (UINT_PTR)-1))
  {
    dwAction=EXTACT_SAVEFILE;
  }
  else if (!xstrcmpinW(L"Font(", wpMethod, (UINT_PTR)-1))
  {
    dwAction=EXTACT_FONT;
  }
  else if (!xstrcmpinW(L"Recode(", wpMethod, (UINT_PTR)-1))
  {
    dwAction=EXTACT_RECODE;
  }
  else if (!xstrcmpinW(L"Insert(", wpMethod, (UINT_PTR)-1))
  {
    dwAction=EXTACT_INSERT;
  }

  if (dwAction)
  {
    for (wpAction=wpMethod; *wpAction != L'('; ++wpAction);
    ParseMethodParameters(&hParamStack, ++wpAction, NULL);

    if (dwAction == EXTACT_COMMAND)
    {
      int nCommand=0;
      LPARAM lParam=0;

      if (lpParameter=GetMethodParameter(&hParamStack, 1))
        nCommand=(int)lpParameter->nNumber;
      if (lpParameter=GetMethodParameter(&hParamStack, 2))
        lParam=lpParameter->nNumber;

      if (nCommand)
      {
        SendMessage(hMainWnd, WM_COMMAND, nCommand, lParam);
      }
    }
    else if (dwAction == EXTACT_CALL)
    {
      PLUGINCALLSENDW pcs;
      unsigned char *lpStruct=NULL;
      int nStructSize;

      ExpandMethodParameters(&hParamStack, lpFrameCurrent->wszFile, wszExeDir, wpUrlLink);

      if (nStructSize=StructMethodParameters(&hParamStack, NULL))
      {
        if (lpStruct=(unsigned char *)GlobalAlloc(GPTR, nStructSize))
        {
          pcs.pFunction=hParamStack.first->wpString;
          if (nStructSize > (INT_PTR)sizeof(INT_PTR))
          {
            pcs.lParam=(LPARAM)lpStruct;
            StructMethodParameters(&hParamStack, lpStruct);
          }
          else pcs.lParam=0;

          CallPluginSend(NULL, &pcs, 0);
          GlobalFree((HGLOBAL)lpStruct);
        }
      }
    }
    else if (dwAction == EXTACT_EXEC)
    {
      STARTUPINFOW si;
      PROCESS_INFORMATION pi;
      wchar_t *wpCmdLine=NULL;
      wchar_t *wpWorkDir=NULL;
      BOOL bWait=FALSE;

      ExpandMethodParameters(&hParamStack, lpFrameCurrent->wszFile, wszExeDir, wpUrlLink);
      if (lpParameter=GetMethodParameter(&hParamStack, 1))
        wpCmdLine=lpParameter->wpExpanded;
      if (lpParameter=GetMethodParameter(&hParamStack, 2))
        wpWorkDir=lpParameter->wpExpanded;
      if (lpParameter=GetMethodParameter(&hParamStack, 3))
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
            //GetExitCodeProcess(pi.hProcess, &dwExitCode);
          }
          CloseHandle(pi.hProcess);
          CloseHandle(pi.hThread);
        }
      }
    }
    else if (dwAction == EXTACT_OPENFILE ||
             dwAction == EXTACT_SAVEFILE)
    {
      wchar_t *wpFile=NULL;
      int nCodePage=-1;
      BOOL bBOM=-1;

      ExpandMethodParameters(&hParamStack, lpFrameCurrent->wszFile, wszExeDir, wpUrlLink);
      if (lpParameter=GetMethodParameter(&hParamStack, 1))
        wpFile=lpParameter->wpExpanded;
      if (lpParameter=GetMethodParameter(&hParamStack, 2))
        nCodePage=(int)lpParameter->nNumber;
      if (lpParameter=GetMethodParameter(&hParamStack, 3))
        bBOM=(BOOL)lpParameter->nNumber;

      if (dwAction == EXTACT_OPENFILE)
      {
        DWORD dwFlags=OD_ADT_BINARY_ERROR;
        int nOpen;

        if (nCodePage == -1)
          dwFlags|=OD_ADT_DETECT_CODEPAGE;
        if (bBOM == -1)
          dwFlags|=OD_ADT_DETECT_BOM;
        nOpen=OpenDocument(NULL, wpFile, dwFlags, nCodePage, bBOM);
        if (nOpen != EOD_SUCCESS && nOpen != EOD_MSGNO && nOpen != EOD_WINDOW_EXIST)
          return PCLE_END;
      }
      else if (dwAction == EXTACT_SAVEFILE)
      {
        if (nCodePage == -1)
          nCodePage=lpFrameCurrent->ei.nCodePage;
        if (bBOM == -1)
          bBOM=lpFrameCurrent->ei.bBOM;
        if (SaveDocument(NULL, wpFile, nCodePage, bBOM, SD_UPDATE) != ESD_SUCCESS)
          return PCLE_END;
      }
    }
    else if (dwAction == EXTACT_FONT)
    {
      wchar_t *wpFaceName=NULL;
      DWORD dwFontStyle=0;
      int nPointSize=0;
      HDC hDC;

      if (lpFrameCurrent->ei.hWndEdit)
      {
        ExpandMethodParameters(&hParamStack, lpFrameCurrent->wszFile, wszExeDir, wpUrlLink);
        if (lpParameter=GetMethodParameter(&hParamStack, 1))
          wpFaceName=lpParameter->wpExpanded;
        if (lpParameter=GetMethodParameter(&hParamStack, 2))
          dwFontStyle=(DWORD)lpParameter->nNumber;
        if (lpParameter=GetMethodParameter(&hParamStack, 3))
          nPointSize=(int)lpParameter->nNumber;

        if (nPointSize)
        {
          if (hDC=GetDC(lpFrameCurrent->ei.hWndEdit))
          {
            lpFrameCurrent->lf.lfHeight=-MulDiv(nPointSize, GetDeviceCaps(hDC, LOGPIXELSY), 72);
            ReleaseDC(lpFrameCurrent->ei.hWndEdit, hDC);
          }
        }
        if (dwFontStyle != FS_NONE)
        {
          lpFrameCurrent->lf.lfWeight=(dwFontStyle == FS_FONTBOLD || dwFontStyle == FS_FONTBOLDITALIC)?FW_BOLD:FW_NORMAL;
          lpFrameCurrent->lf.lfItalic=(dwFontStyle == FS_FONTITALIC || dwFontStyle == FS_FONTBOLDITALIC)?TRUE:FALSE;
        }
        if (*wpFaceName != L'\0')
        {
          xstrcpynW(lpFrameCurrent->lf.lfFaceName, wpFaceName, LF_FACESIZE);
        }
        SetChosenFont(lpFrameCurrent->ei.hWndEdit, &lpFrameCurrent->lf);
        UpdateMappedPrintWidth(lpFrameCurrent);
        UpdateStatusUser(lpFrameCurrent, CSB_FONTPOINT|CSB_MARKER);
      }
    }
    else if (dwAction == EXTACT_RECODE)
    {
      TEXTRECODE tr={0};

      if (lpFrameCurrent->ei.hWndEdit)
      {
        if (lpParameter=GetMethodParameter(&hParamStack, 1))
          tr.nCodePageFrom=(int)lpParameter->nNumber;
        if (lpParameter=GetMethodParameter(&hParamStack, 2))
          tr.nCodePageTo=(int)lpParameter->nNumber;
        RecodeTextW(lpFrameCurrent, NULL, 0, &tr.nCodePageFrom, &tr.nCodePageTo);
      }
    }
    else if (dwAction == EXTACT_INSERT)
    {
      wchar_t *wpText=NULL;
      wchar_t *wpUnescText=NULL;
      int nTextLen=-1;
      int nUnescTextLen;
      DWORD dwCaret=(DWORD)-1;
      BOOL bEscSequences=FALSE;

      if (lpFrameCurrent->ei.hWndEdit && !IsReadOnly(lpFrameCurrent->ei.hWndEdit))
      {
        ExpandMethodParameters(&hParamStack, lpFrameCurrent->wszFile, wszExeDir, wpUrlLink);
        if (lpParameter=GetMethodParameter(&hParamStack, 1))
          wpText=lpParameter->wpExpanded;
        if (lpParameter=GetMethodParameter(&hParamStack, 2))
          bEscSequences=(BOOL)lpParameter->nNumber;

        if (bEscSequences)
        {
          if (nUnescTextLen=(int)TranslateEscapeString(lpFrameCurrent, wpText, NULL, NULL))
          {
            if (wpUnescText=(wchar_t *)GlobalAlloc(GPTR, nUnescTextLen * sizeof(wchar_t)))
            {
              nTextLen=(int)TranslateEscapeString(lpFrameCurrent, wpText, wpUnescText, &dwCaret);
              wpText=wpUnescText;
            }
          }
        }
        if (wpText)
        {
          AECHARINDEX ciInsertPos;

          ReplaceSelW(lpFrameCurrent->ei.hWndEdit, wpText, nTextLen, AELB_ASINPUT, AEREPT_COLUMNASIS, &ciInsertPos, NULL);

          if (dwCaret != (DWORD)-1)
          {
            IndexOffset(lpFrameCurrent->ei.hWndEdit, &ciInsertPos, dwCaret, AELB_ASINPUT);
            SendMessage(lpFrameCurrent->ei.hWndEdit, AEM_EXSETSEL, (WPARAM)&ciInsertPos, (LPARAM)&ciInsertPos);
          }
        }
        if (wpUnescText) GlobalFree((HGLOBAL)wpUnescText);
      }
    }
    FreeMethodParameters(&hParamStack);
  }
  return 0;
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

void ExpandMethodParameters(STACKEXTPARAM *hParamStack, const wchar_t *wpFile, const wchar_t *wpExeDir, const wchar_t *wpUrlLink)
{
  //%f -file, %d -file directory, %a -AkelPad directory, %u -URL address (work in "UrlCommand" manual parameter), %% -%
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

INT_PTR TranslateEscapeString(FRAMEDATA *lpFrame, const wchar_t *wpInput, wchar_t *wszOutput, DWORD *lpdwCaret)
{
  const wchar_t *a=wpInput;
  wchar_t *b=wszOutput;
  wchar_t whex[5];
  int nDec;

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
        if (lpFrame->ei.nNewLine == NEWLINE_MAC)
        {
          if (wszOutput) *b=L'\r';
          ++b;
        }
        else if (lpFrame->ei.nNewLine == NEWLINE_UNIX)
        {
          if (wszOutput) *b=L'\n';
          ++b;
        }
        else if (lpFrame->ei.nNewLine == NEWLINE_WIN)
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
        AECHARRANGE cr;
        wchar_t *wszSelText=NULL;
        INT_PTR nSelTextLen;

        GetSel(lpFrame->ei.hWndEdit, &cr, NULL, NULL);
        if (nSelTextLen=ExGetRangeTextW(lpFrame->ei.hWndEdit, &cr.ciMin, &cr.ciMax, FALSE, wszOutput?&wszSelText:NULL, AELB_ASIS, FALSE))
        {
          if (wszOutput)
          {
            xmemcpy(b, wszSelText, nSelTextLen * sizeof(wchar_t));
            FreeText(wszSelText);
          }
          else --nSelTextLen;

          b+=nSelTextLen;
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

  Error:
  if (wszOutput) *wszOutput=L'\0';
  return 0;
}


//// Other functions

BOOL SetFrameInfo(FRAMEDATA *lpFrame, int nType, UINT_PTR dwData)
{
  if (!lpFrame)
  {
    if (lpFrameCurrent->hWndEditParent)
      lpFrame=lpFrameCurrent;
  }

  if (nType == FIS_TABSTOPSIZE)
  {
    if (lpFrame->nTabStopSize != (int)dwData && (int)dwData > 0)
    {
      lpFrame->nTabStopSize=(int)dwData;
      SetTabStops(lpFrame->ei.hWndEdit, lpFrame->nTabStopSize, TRUE);
      UpdateStatusUser(lpFrame, CSB_TABSIZE);
      return TRUE;
    }
  }
  else if (nType == FIS_TABSTOPASSPACES)
  {
    if (lpFrame->bTabStopAsSpaces != (BOOL)dwData)
    {
      lpFrame->bTabStopAsSpaces=(BOOL)dwData;
      return TRUE;
    }
  }
  else if (nType == FIS_UNDOLIMIT)
  {
    if (lpFrame->nUndoLimit != (int)dwData && (int)dwData >= 0)
    {
      lpFrame->nUndoLimit=(int)dwData;
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETUNDOLIMIT, (WPARAM)lpFrame->nUndoLimit, 0);
      return TRUE;
    }
  }
  else if (nType == FIS_DETAILEDUNDO)
  {
    if (lpFrame->bDetailedUndo != (BOOL)dwData)
    {
      lpFrame->bDetailedUndo=(BOOL)dwData;
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETOPTIONS, lpFrame->bDetailedUndo?AECOOP_OR:AECOOP_XOR, AECO_DETAILEDUNDO);
      return TRUE;
    }
  }
  else if (nType == FIS_WRAP)
  {
    if (lpFrame->dwWrapType != LOWORD(dwData) || lpFrame->dwWrapLimit != HIWORD(dwData))
    {
      lpFrame->dwWrapType=LOWORD(dwData);
      lpFrame->dwWrapLimit=HIWORD(dwData);

      if (lpFrame->ei.bWordWrap)
      {
        UpdateShowHScroll(lpFrame);
        SetWordWrap(lpFrame, lpFrame->dwWrapType, lpFrame->dwWrapLimit);
      }
      return TRUE;
    }
  }
  else if (nType == FIS_MARKER)
  {
    if (lpFrame->dwMarker != dwData && (int)dwData >= -1)
    {
      lpFrame->dwMarker=(DWORD)dwData;
      SetMarker(lpFrame, lpFrame->dwMarker);
      UpdateStatusUser(lpFrame, CSB_MARKER);
      return TRUE;
    }
  }
  else if (nType == FIS_CARETWIDTH)
  {
    if (lpFrame->nCaretWidth != (int)dwData && (int)dwData >= 0)
    {
      lpFrame->nCaretWidth=(int)dwData;

      //Update width
      {
        POINT pt;

        SendMessage(lpFrame->ei.hWndEdit, AEM_GETCARETWIDTH, 0, (LPARAM)&pt);
        pt.x=lpFrame->nCaretWidth;
        SendMessage(lpFrame->ei.hWndEdit, AEM_SETCARETWIDTH, 0, (LPARAM)&pt);
        return TRUE;
      }
    }
  }
  else if (nType == FIS_CARETOPTIONS)
  {
    if (lpFrame->dwCaretOptions != (DWORD)dwData)
    {
      DWORD dwAddOptions;
      DWORD dwCurOptions;

      lpFrame->dwCaretOptions=(DWORD)dwData;

      dwAddOptions=0;
      if (lpFrame->dwCaretOptions & CO_CARETOUTEDGE)
        dwAddOptions|=AECO_CARETOUTEDGE;
      if (lpFrame->dwCaretOptions & CO_CARETVERTLINE)
        dwAddOptions|=AECO_ACTIVECOLUMN;
      if (lpFrame->dwCaretOptions & CO_CARETACTIVELINE)
        dwAddOptions|=AECO_ACTIVELINE;
      if (lpFrame->dwCaretOptions & CO_CARETACTIVELINEBORDER)
        dwAddOptions|=AECO_ACTIVELINEBORDER;
      dwCurOptions=(DWORD)SendMessage(lpFrame->ei.hWndEdit, AEM_GETOPTIONS, 0, 0);
      dwCurOptions&=~AECO_CARETOUTEDGE & ~AECO_ACTIVECOLUMN & ~AECO_ACTIVELINE & ~AECO_ACTIVELINEBORDER;
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETOPTIONS, AECOOP_SET, dwCurOptions|dwAddOptions);

      return TRUE;
    }
  }
  else if (nType == FIS_ALTLINES)
  {
    if (lpFrame->dwAltLineSkip != LOWORD(dwData) || lpFrame->dwAltLineFill != HIWORD(dwData))
    {
      lpFrame->dwAltLineSkip=LOWORD(dwData);
      lpFrame->dwAltLineFill=HIWORD(dwData);
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETALTLINE, MAKELONG(lpFrame->dwAltLineSkip, lpFrame->dwAltLineFill), 0);
      return TRUE;
    }
  }
  else if (nType == FIS_ALTLINEBORDER)
  {
    if (lpFrame->bAltLineBorder != (BOOL)dwData)
    {
      lpFrame->bAltLineBorder=(BOOL)dwData;
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETOPTIONS, lpFrame->bAltLineBorder?AECOOP_OR:AECOOP_XOR, AECO_ALTLINEBORDER);
      return TRUE;
    }
  }
  else if (nType == FIS_RECTMARGINS)
  {
    if (xmemcmp(&lpFrame->rcEditMargins, (RECT *)dwData, sizeof(RECT)))
    {
      if (lpFrame->ei.hWndMaster)
      {
        SetMargins(lpFrame->ei.hWndMaster, (RECT *)dwData, &lpFrame->rcEditMargins);
        if (lpFrame->ei.hWndClone1)
          SetMargins(lpFrame->ei.hWndClone1, (RECT *)dwData, &lpFrame->rcEditMargins);
        if (lpFrame->ei.hWndClone2)
          SetMargins(lpFrame->ei.hWndClone2, (RECT *)dwData, &lpFrame->rcEditMargins);
        if (lpFrame->ei.hWndClone3)
          SetMargins(lpFrame->ei.hWndClone3, (RECT *)dwData, &lpFrame->rcEditMargins);
      }
      else SetMargins(lpFrame->ei.hWndEdit, (RECT *)dwData, &lpFrame->rcEditMargins);

      lpFrame->rcEditMargins=*(RECT *)dwData;
      return TRUE;
    }
  }
  else if (nType == FIS_LINEGAP)
  {
    if (lpFrame->dwLineGap != (DWORD)dwData && (int)dwData >= 0)
    {
      lpFrame->dwLineGap=(DWORD)dwData;
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETLINEGAP, lpFrame->dwLineGap, 0);
      return TRUE;
    }
  }
  else if (nType == FIS_MOUSEOPTIONS)
  {
    if (lpFrame->dwMouseOptions != (DWORD)dwData)
    {
      DWORD dwAddOptions;
      DWORD dwCurOptions;

      lpFrame->dwMouseOptions=(DWORD)dwData;

      dwAddOptions=0;
      if (!(lpFrame->dwMouseOptions & MO_LEFTMARGINSELECTION))
        dwAddOptions|=AECO_NOMARGINSEL;
      if (lpFrame->dwMouseOptions & MO_RICHEDITMOUSE)
        dwAddOptions|=AECO_LBUTTONUPCONTINUECAPTURE;
      if (!(lpFrame->dwMouseOptions & MO_MOUSEDRAGGING))
        dwAddOptions|=AECO_DISABLEDRAG;
      if (lpFrame->dwMouseOptions & MO_RCLICKMOVECARET)
        dwAddOptions|=AECO_RBUTTONDOWNMOVECARET;
      if (lpFrame->dwMouseOptions & MO_NONEWLINEMOUSESELECT)
        dwAddOptions|=AECO_NONEWLINEMOUSESELECT;
      if (lpFrame->dwMouseOptions & MO_MARGINSELUNWRAPLINE)
        dwAddOptions|=AECO_MARGINSELUNWRAPLINE;
      if (lpFrame->dwMouseOptions & MO_MBUTTONDOWNNOSCROLL)
        dwAddOptions|=AECO_MBUTTONDOWNNOSCROLL;
      dwCurOptions=(DWORD)SendMessage(lpFrame->ei.hWndEdit, AEM_GETOPTIONS, 0, 0);
      dwCurOptions&=~AECO_NOMARGINSEL & ~AECO_LBUTTONUPCONTINUECAPTURE & ~AECO_DISABLEDRAG & ~AECO_RBUTTONDOWNMOVECARET & ~AECO_NONEWLINEMOUSESELECT & ~AECO_MARGINSELUNWRAPLINE & ~AECO_MBUTTONDOWNNOSCROLL;
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETOPTIONS, AECOOP_SET, dwCurOptions|dwAddOptions);

      return TRUE;
    }
  }
  else if (nType == FIS_SHOWURL)
  {
    if (lpFrame->bShowURL != (BOOL)dwData)
    {
      lpFrame->bShowURL=(BOOL)dwData;
      SendMessage(lpFrame->ei.hWndEdit, AEM_EXSETOPTIONS, lpFrame->bShowURL?AECOOP_OR:AECOOP_XOR, AECOE_DETECTURL);
      return TRUE;
    }
  }
  else if (nType == FIS_URLPREFIXESENABLE)
  {
    if (lpFrame->bUrlPrefixesEnable != (BOOL)dwData)
    {
      lpFrame->bUrlPrefixesEnable=(BOOL)dwData;

      if (lpFrame->bUrlPrefixesEnable >= 0)
      {
        if (lpFrame->bUrlPrefixesEnable)
          SetUrlPrefixes(lpFrame->ei.hWndEdit, lpFrame->wszUrlPrefixes);
        else
          SetUrlPrefixes(lpFrame->ei.hWndEdit, NULL);
      }
      return TRUE;
    }
  }
  else if (nType == FIS_URLPREFIXES)
  {
    if (xstrcmpW(lpFrame->wszUrlPrefixes, (wchar_t *)dwData))
    {
      xstrcpynW(lpFrame->wszUrlPrefixes, (wchar_t *)dwData, URL_PREFIXES_SIZE);

      if (lpFrame->bUrlPrefixesEnable >= 0)
      {
        if (lpFrame->bUrlPrefixesEnable)
          SetUrlPrefixes(lpFrame->ei.hWndEdit, lpFrame->wszUrlPrefixes);
        else
          SetUrlPrefixes(lpFrame->ei.hWndEdit, NULL);
      }
      return TRUE;
    }
  }
  else if (nType == FIS_URLDELIMITERSENABLE)
  {
    if (lpFrame->bUrlDelimitersEnable != (BOOL)dwData)
    {
      lpFrame->bUrlDelimitersEnable=(BOOL)dwData;

      if (lpFrame->bUrlDelimitersEnable >= 0)
      {
        if (lpFrame->bUrlDelimitersEnable)
        {
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLLEFTDELIMITERS, 0, (LPARAM)lpFrame->wszUrlLeftDelimiters);
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLRIGHTDELIMITERS, 0, (LPARAM)lpFrame->wszUrlRightDelimiters);
        }
        else
        {
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLLEFTDELIMITERS, 0, (LPARAM)NULL);
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLRIGHTDELIMITERS, 0, (LPARAM)NULL);
        }
      }
      return TRUE;
    }
  }
  else if (nType == FIS_URLLEFTDELIMITERS)
  {
    if (xstrcmpW(lpFrame->wszUrlLeftDelimiters, (wchar_t *)dwData))
    {
      xstrcpynW(lpFrame->wszUrlLeftDelimiters, (wchar_t *)dwData, URL_DELIMITERS_SIZE);

      if (lpFrame->bUrlDelimitersEnable >= 0)
      {
        if (lpFrame->bUrlDelimitersEnable)
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLLEFTDELIMITERS, 0, (LPARAM)lpFrame->wszUrlLeftDelimiters);
        else
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLLEFTDELIMITERS, 0, (LPARAM)NULL);
      }
      return TRUE;
    }
  }
  else if (nType == FIS_URLRIGHTDELIMITERS)
  {
    if (xstrcmpW(lpFrame->wszUrlRightDelimiters, (wchar_t *)dwData))
    {
      xstrcpynW(lpFrame->wszUrlRightDelimiters, (wchar_t *)dwData, URL_DELIMITERS_SIZE);

      if (lpFrame->bUrlDelimitersEnable >= 0)
      {
        if (lpFrame->bUrlDelimitersEnable)
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLRIGHTDELIMITERS, 0, (LPARAM)lpFrame->wszUrlRightDelimiters);
        else
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETURLRIGHTDELIMITERS, 0, (LPARAM)NULL);
      }
      return TRUE;
    }
  }
  else if (nType == FIS_WORDDELIMITERSENABLE)
  {
    if (lpFrame->bWordDelimitersEnable != (BOOL)dwData)
    {
      lpFrame->bWordDelimitersEnable=(BOOL)dwData;

      if (lpFrame->bWordDelimitersEnable >= 0)
      {
        if (lpFrame->bWordDelimitersEnable)
        {
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDBREAK, moCur.dwWordBreakCustom, 0);
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDDELIMITERS, 0, (LPARAM)lpFrame->wszWordDelimiters);
        }
        else
        {
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDBREAK, dwWordBreakDefault, 0);
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDDELIMITERS, 0, (LPARAM)NULL);
        }
      }
      return TRUE;
    }
  }
  else if (nType == FIS_WORDDELIMITERS)
  {
    if (xstrcmpW(lpFrame->wszWordDelimiters, (wchar_t *)dwData))
    {
      xstrcpynW(lpFrame->wszWordDelimiters, (wchar_t *)dwData, WORD_DELIMITERS_SIZE);

      if (lpFrame->bWordDelimitersEnable >= 0)
      {
        if (lpFrame->bWordDelimitersEnable)
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDDELIMITERS, 0, (LPARAM)lpFrame->wszWordDelimiters);
        else
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDDELIMITERS, 0, (LPARAM)NULL);
      }
      return TRUE;
    }
  }
  else if (nType == FIS_WRAPDELIMITERSENABLE)
  {
    if (lpFrame->bWrapDelimitersEnable != (BOOL)dwData)
    {
      lpFrame->bWrapDelimitersEnable=(BOOL)dwData;

      if (lpFrame->bWrapDelimitersEnable >= 0)
      {
        if (lpFrame->bWrapDelimitersEnable)
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWRAPDELIMITERS, 0, (LPARAM)lpFrame->wszWrapDelimiters);
        else
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWRAPDELIMITERS, 0, (LPARAM)NULL);
      }
      return TRUE;
    }
  }
  else if (nType == FIS_WRAPDELIMITERS)
  {
    if (xstrcmpW(lpFrame->wszWrapDelimiters, (wchar_t *)dwData))
    {
      xstrcpynW(lpFrame->wszWrapDelimiters, (wchar_t *)dwData, WRAP_DELIMITERS_SIZE);

      if (lpFrame->bWrapDelimitersEnable >= 0)
      {
        if (lpFrame->bWrapDelimitersEnable)
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWRAPDELIMITERS, 0, (LPARAM)lpFrame->wszWrapDelimiters);
        else
          SendMessage(lpFrame->ei.hWndEdit, AEM_SETWRAPDELIMITERS, 0, (LPARAM)NULL);
      }
      return TRUE;
    }
  }
  else if (nType == FIS_CLICKURL)
  {
    if (lpFrame->nClickURL != (int)dwData)
    {
      lpFrame->nClickURL=(int)dwData;
      return TRUE;
    }
  }
  else if (nType == FIS_LOCKINHERIT)
  {
    if (lpFrame->dwLockInherit != (DWORD)dwData)
    {
      lpFrame->dwLockInherit=(DWORD)dwData;
      return TRUE;
    }
  }
  else if (nType == FIS_COLORS)
  {
    AECOLORS *aec=(AECOLORS *)dwData;

    if (xmemcmp(&lpFrame->aec, aec, sizeof(AECOLORS)))
    {
      SendMessage(lpFrame->ei.hWndEdit, AEM_SETCOLORS, 0, (LPARAM)aec);
      lpFrame->aec.dwFlags=aec->dwFlags;
      SendMessage(lpFrame->ei.hWndEdit, AEM_GETCOLORS, 0, (LPARAM)&lpFrame->aec);
      lpFrame->aec.dwFlags=AECLR_ALL;
      return TRUE;
    }
  }
  else if (nType == FIS_BKIMAGE)
  {
    BKIMAGE *bki=(BKIMAGE *)dwData;

    if (xstrcmpiW(lpFrame->wszBkImageFile, bki->wpFile) || lpFrame->nBkImageAlpha != bki->nAlpha)
    {
      SetBkImage(lpFrame, bki->wpFile, bki->nAlpha);
      return TRUE;
    }
  }

  return FALSE;
}

BOOL GetEditInfo(HWND hWnd, EDITINFO *ei)
{
  if (lpFrameCurrent->ei.hWndEdit)
  {
    if (!hWnd || IsEditActive(hWnd))
    {
      if (ei) xmemcpy(ei, &lpFrameCurrent->ei, sizeof(EDITINFO));
      return TRUE;
    }
    else
    {
      FRAMEDATA *lpFrame;

      if (lpFrame=GetFrameDataFromEditWindow(hWnd))
      {
        if (ei) xmemcpy(ei, &lpFrame->ei, sizeof(EDITINFO));
        return TRUE;
      }
    }
  }
  return FALSE;
}

DWORD IsEditActive(HWND hWnd)
{
  if (!hWnd)
    return 0;
  if (hWnd == lpFrameCurrent->ei.hWndEdit)
    return CN_EDIT;

  if (lpFrameCurrent->ei.hWndMaster)
  {
    if (hWnd == lpFrameCurrent->ei.hWndMaster)
      return CN_MASTER;
    if (lpFrameCurrent->ei.hWndClone1 && hWnd == lpFrameCurrent->ei.hWndClone1)
      return CN_CLONE1;
    if (lpFrameCurrent->ei.hWndClone2 && hWnd == lpFrameCurrent->ei.hWndClone2)
      return CN_CLONE2;
    if (lpFrameCurrent->ei.hWndClone3 && hWnd == lpFrameCurrent->ei.hWndClone3)
      return CN_CLONE3;
  }
  return 0;
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void UpdateShowHScroll(FRAMEDATA *lpFrame)
{
  if (!(moCur.dwPaintOptions & PAINT_HIDENOSCROLL))
  {
    BOOL bShowScroll=TRUE;

    if (lpFrame->ei.bWordWrap && !lpFrame->dwWrapLimit && !lpFrame->ei.hWndClone1 && !lpFrame->ei.hWndClone3)
      bShowScroll=FALSE;

    if (lpFrame->ei.hWndMaster)
    {
      SendMessage(lpFrame->ei.hWndMaster, AEM_SHOWSCROLLBAR, SB_HORZ, bShowScroll);
      if (lpFrame->ei.hWndClone1)
        SendMessage(lpFrame->ei.hWndClone1, AEM_SHOWSCROLLBAR, SB_HORZ, bShowScroll);
      if (lpFrame->ei.hWndClone2)
        SendMessage(lpFrame->ei.hWndClone2, AEM_SHOWSCROLLBAR, SB_HORZ, bShowScroll);
      if (lpFrame->ei.hWndClone3)
        SendMessage(lpFrame->ei.hWndClone3, AEM_SHOWSCROLLBAR, SB_HORZ, bShowScroll);
    }
    else SendMessage(lpFrame->ei.hWndEdit, AEM_SHOWSCROLLBAR, SB_HORZ, bShowScroll);
  }
}

int SaveLineScroll(HWND hWnd)
{
  SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
  return (int)SendMessage(hWnd, AEM_GETLINENUMBER, AEGL_FIRSTVISIBLELINE, 0);
}

void RestoreLineScroll(HWND hWnd, int nBeforeLine)
{
  if ((int)SendMessage(hWnd, AEM_GETLINENUMBER, AEGL_FIRSTVISIBLELINE, 0) != nBeforeLine)
  {
    POINT64 ptScrollPos;

    ptScrollPos.x=-1;
    ptScrollPos.y=SendMessage(hWnd, AEM_VPOSFROMLINE, AECT_GLOBAL, nBeforeLine);
    SendMessage(hWnd, AEM_SETSCROLLPOS, 0, (LPARAM)&ptScrollPos);
  }
  SendMessage(hWnd, AEM_UPDATECARET, 0, 0);
  SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
  InvalidateRect(hWnd, NULL, TRUE);
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

BOOL SelectColorDialog(HWND hWndOwner, COLORREF *crColor)
{
  COLORREF crReserve=*crColor;

  if (bOldWindows)
  {
    CHOOSECOLORA ccA;

    xmemset(&ccA, 0, sizeof(CHOOSECOLORA));
    ccA.lStructSize =sizeof(CHOOSECOLORA);
    ccA.lpCustColors=crCustColors;
    ccA.Flags       =CC_FULLOPEN|CC_RGBINIT;
    ccA.hwndOwner   =hWndOwner;
    ccA.rgbResult   =*crColor;

    if (ChooseColorA(&ccA))
    {
      *crColor=ccA.rgbResult;
      return TRUE;
    }
  }
  else
  {
    CHOOSECOLORW ccW;

    xmemset(&ccW, 0, sizeof(CHOOSECOLORW));
    ccW.lStructSize =sizeof(CHOOSECOLORW);
    ccW.lpCustColors=crCustColors;
    ccW.Flags       =CC_FULLOPEN|CC_RGBINIT;
    ccW.hwndOwner   =hWndOwner;
    ccW.rgbResult   =*crColor;

    if (ChooseColorW(&ccW))
    {
      *crColor=ccW.rgbResult;
      return TRUE;
    }
  }
  *crColor=crReserve;
  return FALSE;
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void SetMarker(FRAMEDATA *lpFrame, DWORD dwPos)
{
  if (dwPos == (DWORD)-1)
  {
    if (!lpFrame->dwMappedPrintWidth)
      lpFrame->dwMappedPrintWidth=GetMappedPrintWidth(lpFrame->ei.hWndEdit);
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETMARKER, AEMT_PIXEL, lpFrame->dwMappedPrintWidth);
  }
  else SendMessage(lpFrame->ei.hWndEdit, AEM_SETMARKER, AEMT_SYMBOL, dwPos);
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void SetWordWrap(FRAMEDATA *lpFrame, DWORD dwType, DWORD dwLimit)
{
  if (dwLimit == (DWORD)-1)
  {
    if (!lpFrame->dwMappedPrintWidth)
      lpFrame->dwMappedPrintWidth=GetMappedPrintWidth(lpFrame->ei.hWndEdit);
    SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDWRAP, dwType|AEWW_LIMITPIXEL, lpFrame->dwMappedPrintWidth);
  }
  else SendMessage(lpFrame->ei.hWndEdit, AEM_SETWORDWRAP, dwType|AEWW_LIMITSYMBOL, dwLimit);
}

void SetMargins(HWND hWnd, const RECT *lprcNewMargins, const RECT *lprcOldMargins)
{
  RECT rcEditMargins;
  RECT rcCurMargins;

  SendMessage(hWnd, AEM_GETRECT, AERC_MARGINS, (LPARAM)&rcCurMargins);
  rcEditMargins.left=rcCurMargins.left - (lprcOldMargins?lprcOldMargins->left:0) + lprcNewMargins->left;
  rcEditMargins.top=rcCurMargins.top - (lprcOldMargins?lprcOldMargins->top:0) + lprcNewMargins->top;
  rcEditMargins.right=rcCurMargins.right - (lprcOldMargins?lprcOldMargins->right:0) + lprcNewMargins->right;
  rcEditMargins.bottom=rcCurMargins.bottom - (lprcOldMargins?lprcOldMargins->bottom:0) + lprcNewMargins->bottom;
  SendMessage(hWnd, AEM_SETRECT, AERC_MARGINS|AERC_UPDATE, (LPARAM)&rcEditMargins);
}

void SetTabStops(HWND hWnd, int nTabStops, BOOL bSetRedraw)
{
  SendMessage(hWnd, AEM_SETTABSTOP, nTabStops, bSetRedraw);
}

BOOL InsertTabStop(HWND hWnd)
{
  wchar_t wszSpaces[MAX_PATH];
  int nSpaces;
  int i;

  if (IsReadOnly(hWnd)) return FALSE;

  if (lpFrameCurrent->bTabStopAsSpaces)
  {
    for (i=min(crCurSel.ciMin.lpLine->nLineLen, crCurSel.ciMin.nCharInLine) - 1; i > 0; --i)
    {
      if (crCurSel.ciMin.lpLine->wpLine[i] == L'\t')
      {
        ++i;
        break;
      }
    }
    i=max(i, 0);
    nSpaces=lpFrameCurrent->nTabStopSize - (crCurSel.ciMin.nCharInLine - i) % lpFrameCurrent->nTabStopSize;
    nSpaces=min(nSpaces, MAX_PATH - 1);

    for (i=0; i < nSpaces; ++i)
      wszSpaces[i]=L' ';
    wszSpaces[i]=L'\0';
    ReplaceSelW(hWnd, wszSpaces, -1, AELB_ASINPUT, AEREPT_COLUMNASIS, NULL, NULL);
  }
  else SendMessage(hWnd, AEM_INSERTCHAR, VK_TAB, 0);

  return TRUE;
}

BOOL IndentTabStop(HWND hWnd, int nAction)
{
  wchar_t wszSpaces[MAX_PATH];
  int nSpaces;
  int i;

  if (IsReadOnly(hWnd)) return FALSE;

  if (lpFrameCurrent->bTabStopAsSpaces)
  {
    nSpaces=min(lpFrameCurrent->nTabStopSize, MAX_PATH - 1);

    for (i=0; i < nSpaces; ++i)
      wszSpaces[i]=L' ';
    wszSpaces[i]=L'\0';
    DoEditModifyStringInSelection(hWnd, nAction, wszSpaces);
  }
  else DoEditModifyStringInSelection(hWnd, nAction, L"\t");

  return TRUE;
}

BOOL AutoIndent(HWND hWnd, AECHARRANGE *cr)
{
  AECHARINDEX ciChar=cr->ciMin;
  wchar_t *wpText;

  if (IsReadOnly(hWnd)) return FALSE;

  if (!(moCur.dwStatusPosType & SPT_LINEWRAP) && lpFrameCurrent->ei.bWordWrap)
    SendMessage(hWnd, AEM_GETINDEX, AEGI_WRAPLINEBEGIN, (LPARAM)&ciChar);

  //Calculate spaces
  for (ciChar.nCharInLine=0; AEC_IndexCompare(&ciChar, &cr->ciMin) < 0 &&
                             (ciChar.lpLine->wpLine[ciChar.nCharInLine] == L' ' ||
                              ciChar.lpLine->wpLine[ciChar.nCharInLine] == L'\t'); ++ciChar.nCharInLine);

  if (ciChar.nCharInLine)
  {
    //Insert spaces
    if (wpText=AllocWideStr(ciChar.nCharInLine + 2))
    {
      wpText[0]=L'\n';

      for (ciChar.nCharInLine=0; AEC_IndexCompare(&ciChar, &cr->ciMin) < 0 &&
                                 (ciChar.lpLine->wpLine[ciChar.nCharInLine] == L' ' ||
                                  ciChar.lpLine->wpLine[ciChar.nCharInLine] == L'\t'); ++ciChar.nCharInLine)
      {
        wpText[ciChar.nCharInLine + 1]=ciChar.lpLine->wpLine[ciChar.nCharInLine];
      }
      wpText[ciChar.nCharInLine + 1]=L'\0';

      ReplaceSelW(hWnd, wpText, -1, AELB_ASINPUT, 0, NULL, NULL);
      FreeWideStr(wpText);
      return TRUE;
    }
  }
  return FALSE;
}

int SetUrlPrefixes(HWND hWnd, const wchar_t *wpPrefixes)
{
  wchar_t wszBuffer[URL_PREFIXES_SIZE];
  int nUrlPrefixMax=URL_PREFIXES_SIZE - 2;
  int a=0;
  int b=0;

  if (wpPrefixes)
  {
    while (wpPrefixes[a] == L' ') ++a;

    for (; b < nUrlPrefixMax; ++a, ++b)
    {
      if (wpPrefixes[a])
      {
        if (wpPrefixes[a] == L' ')
        {
          while (wpPrefixes[a + 1] == L' ') ++a;
          wszBuffer[b]=L'\0';
        }
        else wszBuffer[b]=wpPrefixes[a];
      }
      else break;
    }
    wszBuffer[b]=L'\0';
    wszBuffer[++b]=L'\0';

    return (int)SendMessage(hWnd, AEM_SETURLPREFIXES, 0, (LPARAM)wszBuffer);
  }
  else
  {
    return (int)SendMessage(hWnd, AEM_SETURLPREFIXES, 0, (LPARAM)NULL);
  }
}

BOOL IsReadOnly(HWND hWnd)
{
  FRAMEDATA *lpFrame;

  if (lpFrame=GetFrameDataFromEditWindow(hWnd))
    if (lpFrame->ei.bReadOnly)
      return TRUE;
  return FALSE;
}

int IsFile(const wchar_t *wpFile)
{
  DWORD dwAttr;

  dwAttr=GetFileAttributesWide(wpFile);
  if (dwAttr == INVALID_FILE_ATTRIBUTES)
    return ERROR_INVALID_HANDLE;
  if (dwAttr & FILE_ATTRIBUTE_DIRECTORY)
    return ERROR_DIRECTORY;
  return ERROR_SUCCESS;
}

BOOL IsPathFull(const wchar_t *wpPath)
{
  if (wpPath[0] == L'\\' && wpPath[1] == L'\\') return TRUE;
  if (wpPath[0] != L'\0' && wpPath[1] == L':') return TRUE;
  return FALSE;
}

BOOL GetFileWin32Data(const wchar_t *wpFile, WIN32_FIND_DATAW *wfd)
{
  HANDLE hFile;

  if ((hFile=FindFirstFileWide(wpFile, wfd)) != INVALID_HANDLE_VALUE)
  {
    FindClose(hFile);
    return TRUE;
  }
  wfd->dwFileAttributes=INVALID_FILE_ATTRIBUTES;
  return FALSE;
}

int GetFileDir(const wchar_t *wpFile, int nFileLen, wchar_t *wszFileDir, DWORD dwFileDirMax)
{
  const wchar_t *wpCount;

  if (nFileLen == -1) nFileLen=(int)xstrlenW(wpFile);
  if (wszFileDir) wszFileDir[0]=L'\0';

  for (wpCount=wpFile + nFileLen - 1; wpCount >= wpFile; --wpCount)
  {
    if (*wpCount == L'\\')
      return (int)xstrcpynW(wszFileDir, wpFile, min(dwFileDirMax, (DWORD)(wpCount - wpFile) + 1));
  }
  return 0;
}

BOOL GetFullName(const wchar_t *wpFile, wchar_t *wszFileFullName, int nFileMax, int *lpnFileLen)
{
  wchar_t wszFileBuf[MAX_PATH];
  wchar_t *wpFileName;
  int nFileLen=0;
  BOOL bResult=FALSE;

  if (GetFullPathNameWide(wpFile, MAX_PATH, wszFileBuf, &wpFileName))
  {
    if (bResult=FileExistsWide(wszFileBuf))
    {
      if (nFileLen=GetLongPathNameWide(wszFileBuf, wszFileFullName, nFileMax))
        goto End;

      //GetLongPathName can fail on x86 build when open system file on x64,
      //but we return TRUE to allow AkelAdmin.exe execution.
    }
    nFileLen=(int)xstrcpynW(wszFileFullName, wszFileBuf, nFileMax);
  }
  else wszFileFullName[0]=L'\0';

  End:
  if (lpnFileLen) *lpnFileLen=nFileLen;
  return bResult;
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

int GetBaseName(const wchar_t *wpFile, wchar_t *wszBaseName, int nBaseNameMaxLen)
{
  int nFileLen=(int)xstrlenW(wpFile);
  int nEndOffset=-1;
  int i;

  for (i=nFileLen - 1; i >= 0; --i)
  {
    if (wpFile[i] == L'\\')
      break;

    if (nEndOffset == -1)
    {
      if (wpFile[i] == L'.')
        nEndOffset=i;
    }
  }
  ++i;
  if (nEndOffset == -1) nEndOffset=nFileLen;
  nBaseNameMaxLen=min(nEndOffset - i + 1, nBaseNameMaxLen);
  xstrcpynW(wszBaseName, wpFile + i, nBaseNameMaxLen);

  return nBaseNameMaxLen;
}

const wchar_t* GetFileExt(const wchar_t *wpFile, int nFileLen)
{
  const wchar_t *wpCount;

  if (nFileLen == -1) nFileLen=(int)xstrlenW(wpFile);

  for (wpCount=wpFile + nFileLen - 1; wpCount >= wpFile; --wpCount)
  {
    if (*wpCount == L'.') return wpCount + 1;
    if (*wpCount == L'\\') break;
  }
  return NULL;
}

int GetFileStreamOffset(const wchar_t *wpFile, int nFileLen)
{
  const wchar_t *wpCount;

  if (nFileLen == -1) nFileLen=(int)xstrlenW(wpFile);

  for (wpCount=wpFile + nFileLen - 1; wpCount > wpFile; --wpCount)
  {
    if (*wpCount == L':') return (int)(wpCount - wpFile);
    if (*wpCount == L'\\') break;
  }
  return 0;
}

const wchar_t* GetNewLineString(int nNewLine)
{
  if (nNewLine == NEWLINE_WIN) return L"\r\n";
  if (nNewLine == NEWLINE_UNIX) return L"\n";
  if (nNewLine == NEWLINE_MAC) return L"\r";
  return L"";
}

void GetTimeString(const wchar_t *wpFormat, wchar_t *wszOutput, BOOL bWithoutSeconds)
{
  if (!*wpFormat)
  {
    wchar_t wszTime[128];
    wchar_t wszDate[128];

    GetTimeFormatWide(LOCALE_USER_DEFAULT, bWithoutSeconds?TIME_NOSECONDS:0, NULL, NULL, wszTime, 128);
    GetDateFormatWide(LOCALE_USER_DEFAULT, 0, NULL, NULL, wszDate, 128);
    xprintfW(wszOutput, L"%s %s", wszTime, wszDate);
  }
  else
  {
    wchar_t wszBuffer[MAX_PATH];
    wchar_t wszAMPM[128];
    INT_PTR nChanges=0;

    nChanges+=StrReplaceW(wpFormat, -1, L"tt", -1, L"\x0002", 1, FRF_MATCHCASE, wszBuffer, NULL, NULL, NULL, NULL);
    nChanges+=StrReplaceW(wszBuffer, -1, L"t", -1, L"\x0001", 1, FRF_MATCHCASE, wszOutput, NULL, NULL, NULL, NULL);
    GetTimeFormatWide(LOCALE_USER_DEFAULT, 0, NULL, wszOutput, wszBuffer, 128);
    GetDateFormatWide(LOCALE_USER_DEFAULT, 0, NULL, wszBuffer, wszOutput, 128);
    if (nChanges)
    {
      GetTimeFormatWide(LOCALE_USER_DEFAULT, 0, NULL, L"tt", wszAMPM, 128);
      StrReplaceW(wszOutput, -1, L"\x0002", 1, wszAMPM, -1, FRF_MATCHCASE, wszBuffer, NULL, NULL, NULL, NULL);

      GetTimeFormatWide(LOCALE_USER_DEFAULT, 0, NULL, L"t", wszAMPM, 128);
      StrReplaceW(wszBuffer, -1, L"\x0001", 1, wszAMPM, -1, FRF_MATCHCASE, wszOutput, NULL, NULL, NULL, NULL);
    }
  }
}

BOOL GetFileWriteTimeWide(const wchar_t *wpFile, FILETIME *ft)
{
  HANDLE hFile;

  //FindFirstFile not equal to GetFileTime if time zone was changed,
  //because FindFirstFile gets file time from cache.
  //See "File Times and Daylight Saving Time" on MSDN.
  {
    WIN32_FIND_DATAW wfd;

    if ((hFile=FindFirstFileWide(wpFile, &wfd)) != INVALID_HANDLE_VALUE)
    {
      *ft=wfd.ftLastWriteTime;
      FindClose(hFile);
      return TRUE;
    }
  }

  //if ((hFile=CreateFileWide(wpFile, FILE_READ_ATTRIBUTES, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL)) != INVALID_HANDLE_VALUE)
  //{
  //  GetFileTime(hFile, NULL, NULL, ft);
  //  CloseHandle(hFile);
  //  return TRUE;
  //}
  return FALSE;
}

BOOL GetFileVersionA(const char *pFile, int *nMajor, int *nMinor, int *nRelease, int *nBuild, DWORD *dwLanguage)
{
  struct LANGANDCODEPAGE {
    WORD wLanguage;
    WORD wCodePage;
  } *lpTranslate;

  VS_FIXEDFILEINFO ffi;
  VS_FIXEDFILEINFO *pffi=&ffi;
  void *pVerBuf;
  DWORD dwHandle;
  DWORD dwVerSize;
  UINT uLen;
  BOOL bResult=FALSE;

  *nMajor=*nMinor=*nRelease=*nBuild=0;

  if (dwVerSize=GetFileVersionInfoSizeA((char *)pFile, &dwHandle))
  {
    if (pVerBuf=API_HeapAlloc(hHeap, 0, dwVerSize))
    {
      if (GetFileVersionInfoA((char *)pFile, dwHandle, dwVerSize, pVerBuf))
      {
        if (VerQueryValueA(pVerBuf, "\\", (void **)&pffi, &uLen))
        {
          *nMajor=pffi->dwFileVersionMS / 0x00010000;
          *nMinor=pffi->dwFileVersionMS & 0x0000FFFF;
          *nRelease=pffi->dwFileVersionLS / 0x00010000;
          *nBuild=pffi->dwFileVersionLS & 0x0000FFFF;
          bResult=TRUE;
        }
        if (dwLanguage)
        {
          if (VerQueryValueA(pVerBuf, "\\VarFileInfo\\Translation", (void **)&lpTranslate, &uLen))
          {
            *dwLanguage=lpTranslate->wLanguage;
          }
        }
      }
      API_HeapFree(hHeap, 0, pVerBuf);
    }
  }
  return bResult;
}

BOOL GetFileVersionW(const wchar_t *wpFile, int *nMajor, int *nMinor, int *nRelease, int *nBuild, DWORD *dwLanguage)
{
  struct LANGANDCODEPAGE {
    WORD wLanguage;
    WORD wCodePage;
  } *lpTranslate;

  VS_FIXEDFILEINFO ffi;
  VS_FIXEDFILEINFO *pffi=&ffi;
  void *pVerBuf;
  DWORD dwHandle;
  DWORD dwVerSize;
  UINT uLen;
  BOOL bResult=FALSE;

  *nMajor=*nMinor=*nRelease=*nBuild=0;

  if (dwVerSize=GetFileVersionInfoSizeW((wchar_t *)wpFile, &dwHandle))
  {
    if (pVerBuf=API_HeapAlloc(hHeap, 0, dwVerSize))
    {
      if (GetFileVersionInfoW((wchar_t *)wpFile, dwHandle, dwVerSize, pVerBuf))
      {
        if (VerQueryValueW(pVerBuf, L"\\", (void **)&pffi, &uLen))
        {
          *nMajor=pffi->dwFileVersionMS / 0x00010000;
          *nMinor=pffi->dwFileVersionMS & 0x0000FFFF;
          *nRelease=pffi->dwFileVersionLS / 0x00010000;
          *nBuild=pffi->dwFileVersionLS & 0x0000FFFF;
          bResult=TRUE;
        }
        if (dwLanguage)
        {
          if (VerQueryValueW(pVerBuf, L"\\VarFileInfo\\Translation", (void **)&lpTranslate, &uLen))
          {
            *dwLanguage=lpTranslate->wLanguage;
          }
        }
      }
      API_HeapFree(hHeap, 0, pVerBuf);
    }
  }
  return bResult;
}

int VersionCompare(DWORD dwVersion1, DWORD dwVersion2)
{
  if (LOBYTE(dwVersion1) != LOBYTE(dwVersion2))
    return LOBYTE(dwVersion1) - LOBYTE(dwVersion2);
  if (HIBYTE(dwVersion1) != HIBYTE(dwVersion2))
    return HIBYTE(dwVersion1) - HIBYTE(dwVersion2);
  if (LOBYTE(HIWORD(dwVersion1)) != LOBYTE(HIWORD(dwVersion2)))
    return LOBYTE(HIWORD(dwVersion1)) - LOBYTE(HIWORD(dwVersion2));
  if (HIBYTE(HIWORD(dwVersion1)) != HIBYTE(HIWORD(dwVersion2)))
    return HIBYTE(HIWORD(dwVersion1)) - HIBYTE(HIWORD(dwVersion2));
  return 0;
}

int TranslateFileString(const wchar_t *wpString, wchar_t *wszBuffer, int nBufferSize)
{
  //%f -file, %d -file directory, %a -AkelPad directory, %% -%
  wchar_t *wpFile=lpFrameCurrent->wszFile;
  wchar_t *wpExeDir=wszExeDir;
  wchar_t *wszSource;
  wchar_t *wpSource;
  wchar_t *wpTarget=wszBuffer;
  wchar_t *wpTargetMax=wszBuffer + (wszBuffer?nBufferSize:0x7FFFFFFF);
  int nStringLen;

  //Expand environment strings
  nStringLen=ExpandEnvironmentStringsWide(wpString, NULL, 0);

  if (wszSource=(wchar_t *)GlobalAlloc(GPTR, nStringLen * sizeof(wchar_t)))
  {
    ExpandEnvironmentStringsWide(wpString, wszSource, nStringLen);

    //Expand plugin variables
    for (wpSource=wszSource; *wpSource && wpTarget < wpTargetMax;)
    {
      if (*wpSource == L'%')
      {
        if (*++wpSource == L'%')
        {
          ++wpSource;
          if (wszBuffer) *wpTarget=L'%';
          ++wpTarget;
        }
        else if (*wpSource == L'f' || *wpSource == L'F')
        {
          ++wpSource;
          if (*wpFile) wpTarget+=xstrcpynW(wszBuffer?wpTarget:NULL, wpFile, wpTargetMax - wpTarget) - !wszBuffer;
        }
        else if (*wpSource == L'd' || *wpSource == L'D')
        {
          ++wpSource;
          if (*wpFile) wpTarget+=GetFileDir(wpFile, -1, wszBuffer?wpTarget:NULL, (DWORD)(wpTargetMax - wpTarget)) - !wszBuffer;
        }
        else if (*wpSource == L'a' || *wpSource == L'A')
        {
          ++wpSource;
          wpTarget+=xstrcpynW(wszBuffer?wpTarget:NULL, wpExeDir, wpTargetMax - wpTarget) - !wszBuffer;
        }
      }
      else
      {
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

void TrimModifyState(wchar_t *wszFile, int nFileLen)
{
  wchar_t *wpCount;

  if (nFileLen == -1) nFileLen=(int)xstrlenW(wszFile);

  for (wpCount=wszFile + nFileLen - 1; wpCount >= wszFile; --wpCount)
  {
    if (*wpCount == L'*' || *wpCount == L' ')
      *wpCount=L'\0';
    else
      break;
  }
}

void UpdateAsterisk(FRAMEDATA *lpFrame, BOOL bModified)
{
  if (nMDI == WMD_SDI)
  {
    if (moCur.dwShowModify & SM_MAINTITLE_SDI)
    {
      wchar_t *wpTitle=AllocWideStr(BUFFER_SIZE);

      xprintfW(wpTitle, L"%s%s - %s", bModified?L"* ":L"", GetFileName(lpFrame->wszFile, lpFrame->nFileLen), APP_MAIN_TITLEW);
      SetWindowTextWide(hMainWnd, wpTitle);

      FreeWideStr(wpTitle);
    }
  }
  if (nMDI)
  {
    if (moCur.dwShowModify & SM_TABTITLE_MDI)
    {
      int nItem;

      if ((nItem=GetTabItemFromParam(hTab, (LPARAM)lpFrame)) != -1)
      {
        wchar_t *wpTitle=AllocWideStr(BUFFER_SIZE);
        TCITEMW tcItem;

        tcItem.mask=TCIF_TEXT;
        tcItem.pszText=wpTitle;
        tcItem.cchTextMax=MAX_PATH;
        TabCtrl_GetItemWide(hTab, nItem, &tcItem);

        if (bModified)
          xprintfW(wpTitle, L"%s *", wpTitle);
        else
          TrimModifyState(wpTitle, -1);
        TabCtrl_SetItemWide(hTab, nItem, &tcItem);

        FreeWideStr(wpTitle);
      }
    }
    if (moCur.dwShowModify & SM_FRAMETITLE_MDI)
    {
      wchar_t *wpTitle=AllocWideStr(BUFFER_SIZE);

      if (nMDI == WMD_MDI)
      {
        xprintfW(wpTitle, L"%s%s", lpFrame->wszFile, bModified?L" *":L"");
        SetWindowTextWide(lpFrame->hWndEditParent, wpTitle);
      }
      else if (nMDI == WMD_PMDI)
      {
        if (lpFrame->wszFile[0] || bModified)
        {
          xprintfW(wpTitle, L"%s - [%s%s]", APP_MAIN_TITLEW, lpFrame->wszFile, bModified?L" *":L"");
          SetWindowTextWide(hMainWnd, wpTitle);
        }
        else SetWindowTextWide(hMainWnd, APP_MAIN_TITLEW);
      }
      FreeWideStr(wpTitle);
    }
  }
}

void SetMouseCapture(HWND hWnd, DWORD dwType)
{
  if (!dwMouseCapture)
    SetCapture(hWnd);
  dwMouseCapture|=dwType;
}

void ReleaseMouseCapture(DWORD dwType)
{
  if (dwMouseCapture & dwType)
  {
    dwMouseCapture&=~dwType;
    if (!dwMouseCapture)
      ReleaseCapture();
  }
}

BOOL SwitchLayout(HWND hWndEdit, AECHARINDEX *lpciCaret)
{
  AECHARINDEX ciPrevChar;
  HKL *lpList;
  HKL dwFirstNonLatinLocale=0;
  HKL dwLatinLocale=0;
  HKL dwCurLocale;
  DWORD dwPrimaryLocale;
  int nListCount;
  int nChar;
  int nCharLayout=CHARLAYOUT_NONE;
  int i;
  BOOL bResult=FALSE;

  if (hWndEdit && GetFocus() == hWndEdit)
  {
    if (AEC_PrevCharInLineEx(lpciCaret, &ciPrevChar))
    {
      if ((nChar=AEC_CharAtIndex(&ciPrevChar)) >= 0)
        nCharLayout=DetectCharLayout(nChar);
    }
    if (nCharLayout == CHARLAYOUT_NONE)
    {
      if ((nChar=AEC_CharAtIndex(lpciCaret)) >= 0)
        nCharLayout=DetectCharLayout(nChar);
    }

    if (nCharLayout != CHARLAYOUT_NONE)
    {
      if ((nListCount=GetKeyboardLayoutList(0, NULL)) > 1)
      {
        if (lpList=(HKL *)GlobalAlloc(GPTR, nListCount * sizeof(UINT_PTR)))
        {
          GetKeyboardLayoutList(nListCount, lpList);

          //Find Latin locale
          for (i=0; i < nListCount; ++i)
          {
            if (PRIMARYLANGID((UINT_PTR)lpList[i]) == LANG_ENGLISH)
              dwLatinLocale=lpList[i];
            else if (!dwFirstNonLatinLocale)
              dwFirstNonLatinLocale=lpList[i];
          }
          if (dwLatinLocale)
          {
            //dwCurLocale=(HKL)GetKeyboardLayout(0);
            dwCurLocale=(HKL)SendMessage(hWndEdit, AEM_INPUTLANGUAGE, 0, 0);
            dwPrimaryLocale=PRIMARYLANGID((UINT_PTR)dwCurLocale);

            if (nCharLayout == CHARLAYOUT_ENGLISH)
            {
              if (dwPrimaryLocale != LANG_ENGLISH)
              {
                ActivateKeyboardLayout(dwLatinLocale, 0);
                bResult=TRUE;
              }
            }
            else if (nCharLayout == CHARLAYOUT_NONENGLISH)
            {
              if (dwPrimaryLocale == LANG_ENGLISH)
              {
                ActivateKeyboardLayout(dwFirstNonLatinLocale, 0);
                bResult=TRUE;
              }
            }
          }
          GlobalFree((HGLOBAL)lpList);
        }
      }
    }
  }
  return bResult;
}

int DetectCharLayout(int nChar)
{
  if ((nChar >= L'A' && nChar <= L'Z') || (nChar >= L'a' && nChar <= L'z'))
    return CHARLAYOUT_ENGLISH;
  else if (nChar > 0x80)
    return CHARLAYOUT_NONENGLISH;
  return CHARLAYOUT_NONE;
}

void ActivateKeyboard(HKL dwInputLocale)
{
  DWORD dwLangID=LOWORD(dwInputLocale);

  if (dwInputLocale != (HKL)(UINT_PTR)-1)
  {
    if (bWindowsNT4)
    {
      DWORD dwLangIDInit=LOWORD(GetKeyboardLayout(0));
      DWORD dwLangIDCount=dwLangIDInit;

      while (dwLangIDCount != dwLangID)
      {
        ActivateKeyboardLayout((HKL)(UINT_PTR)HKL_NEXT, 0);
        dwLangIDCount=LOWORD(GetKeyboardLayout(0));
        if (dwLangIDCount == dwLangIDInit) break;
      }
    }
    else
    {
      if (LOWORD(GetKeyboardLayout(0)) != dwLangID)
        ActivateKeyboardLayout(dwInputLocale, 0);
    }
  }
}

void ActivateWindow(HWND hWnd)
{
  DWORD dwStyle=(DWORD)GetWindowLongPtrWide(hWnd, GWL_STYLE);

  if (dwStyle & WS_VISIBLE)
  {
    if (dwStyle & WS_MINIMIZE)
    {
      ShowWindow(hWnd, SW_RESTORE);
      SendMessage(hWnd, WM_COMMAND, IDM_INTERNAL_UPDATEMAINCHILDREN, 0);
    }
    else
    {
      if (GetForegroundWindow() != hWnd)
      {
        SetForegroundWindow(hWnd); //BringWindowToTop(hWnd);
        SendMessage(hWnd, WM_COMMAND, IDM_INTERNAL_UPDATEMAINCHILDREN, 0);
      }
    }
  }
  SendMessage(hWnd, AKDN_ACTIVATE, 0, 0);
}

HWND FindAkelCopy()
{
  HWND hWndFriend=NULL;

  if (moCur.dwSingleOpenProgram & SOP_SAMEEXE)
    EnumWindows(EnumAkelCopyProc, (LPARAM)&hWndFriend);
  else
    hWndFriend=FindWindowExWide(NULL, NULL, APP_MAIN_CLASSW, NULL);
  return hWndFriend;
}

BOOL CALLBACK EnumAkelCopyProc(HWND hWnd, LPARAM lParam)
{
  HWND *hWndFriend=(HWND *)lParam;

  GetClassNameWide(hWnd, wbuf, BUFFER_SIZE);
  if (!xstrcmpW(wbuf, APP_MAIN_CLASSW))
  {
    if (GetAkelPadExe(hWnd, wbuf, BUFFER_SIZE))
    {
      if (!xstrcmpiW(wszExeFile, wbuf))
      {
        *hWndFriend=hWnd;
        return FALSE;
      }
    }
  }
  return TRUE;
}

int GetAkelPadExe(HWND hWnd, wchar_t *szExeFile, int nExeFileMax)
{
  HANDLE hMemRemote;
  HANDLE hMemLocal;
  wchar_t *wszMemRemote;
  wchar_t *wszMemLocal;
  DWORD dwMemSize=nExeFileMax * sizeof(wchar_t);
  int nResult=0;

  //Current AkelPad process must have privileges to send AKD_* messages
  if (hMemRemote=(HANDLE)SendMessage(hWnd, AKD_MEMCREATE, (WPARAM)"Global\\AkelPad", dwMemSize))
  {
    if (wszMemRemote=(wchar_t *)SendMessage(hWnd, AKD_MEMMAP, (WPARAM)hMemRemote, dwMemSize))
    {
      SendMessage(hWnd, AKD_GETMAININFO, MI_AKELEXEW, (WPARAM)wszMemRemote);

      //Read data from other process
      if (hMemLocal=(HANDLE)MemCreate("Global\\AkelPad", 0))
      {
        if (wszMemLocal=(wchar_t *)MemMap(hMemLocal, dwMemSize))
        {
          nResult=(int)xstrcpynW(szExeFile, wszMemLocal, nExeFileMax);
          MemUnmap(wszMemLocal);
        }
        MemClose(hMemLocal);
      }
      SendMessage(hWnd, AKD_MEMUNMAP, (WPARAM)wszMemRemote, 0);
    }
    SendMessage(hWnd, AKD_MEMCLOSE, (WPARAM)hMemRemote, 0);
  }
  else
  {
    API_LoadStringW(hLangLib, MSG_ACCESSDENIED, wbuf, BUFFER_SIZE);
    xprintfW(wszMsg, wbuf, L"SingleOpenProgram=3");
    API_MessageBox(hMainWnd?hMainWnd:NULL, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONEXCLAMATION);
  }
  return nResult;
}

HWND NextDialog(BOOL bPrevious)
{
  static ENUMDLG ed;
  HWND hWndNext=NULL;
  HWND hWndToFind;

  //Modeless dialogs
  if (!(hWndToFind=GetActiveWindow()))
    return NULL;
  if (hWndToFind == hMainWnd)
    hWndToFind=NULL;

  if (!hWndToFind)
  {
    ed.hWndToFind=NULL;
    ed.bNextNext=FALSE;
  }
  else if (hWndToFind == ed.hWndResult)
  {
    if (!ed.hWndNextResult)
    {
      SetActiveWindow(hMainWnd);
      return hMainWnd;
    }
    ed.hWndToFind=ed.hWndNextResult;
    ed.bNextNext=FALSE;
  }
  else
  {
    ed.hWndToFind=hWndToFind;
    ed.bNextNext=TRUE;
  }
  ed.nCount=0;
  ed.hWndResult=NULL;
  ed.hWndNextResult=NULL;
  ed.hWndNextNextResult=NULL;
  ed.bFound=FALSE;
  EnumWindows(EnumDialogsProc, (LPARAM)&ed);

  if (ed.nCount)
  {
    if (ed.bNextNext)
    {
      hWndNext=ed.hWndResult=ed.hWndNextResult;
      ed.hWndNextResult=ed.hWndNextNextResult;
    }
    else hWndNext=ed.hWndResult;

    if (!hWndNext) hWndNext=hMainWnd;
    SetActiveWindow(hWndNext);
  }
  else
  {
    //Dockable windows
    DOCK *dkElement=NULL;
    HWND hWndFocus=GetFocus();

    if (hWndFocus)
    {
      if (!IsEditActive(hWndFocus) && hWndFocus != hMdiClient)
        dkElement=StackDockFindWindow(&hDocksStack, hWndFocus, TRUE);

      hWndNext=StackDockNextWindow(&hDocksStack, dkElement, bPrevious);
    }
    if (!hWndNext) hWndNext=hMainWnd;
    SetFocus(hWndNext);
  }
  return hWndNext;
}

BOOL CALLBACK EnumDialogsProc(HWND hWnd, LPARAM lParam)
{
  ENUMDLG *ped=(ENUMDLG *)lParam;

  if (!ped->bFound)
  {
    if (ped->hWndToFind == hWnd)
      ped->bFound=TRUE;
  }
  if ((HWND)GetWindowLongPtrWide(hWnd, GWLP_HWNDPARENT) == hMainWnd && IsWindowVisible(hWnd))
  {
    if (!ped->bFound)
    {
      if (!ped->hWndToFind)
        ped->bFound=TRUE;
    }
    ped->nCount+=1;

    if (ped->bFound)
    {
      if (!ped->hWndResult)
      {
        ped->hWndResult=hWnd;
      }
      else if (!ped->hWndNextResult)
      {
        ped->hWndNextResult=hWnd;
        if (!ped->bNextNext) return FALSE;
      }
      else if (!ped->hWndNextNextResult)
      {
        ped->hWndNextNextResult=hWnd;
        return FALSE;
      }
    }
  }
  return TRUE;
}

HWND NextClone(BOOL bPrevious)
{
  HWND lpClones[]={lpFrameCurrent->ei.hWndMaster, lpFrameCurrent->ei.hWndClone1, lpFrameCurrent->ei.hWndClone2, lpFrameCurrent->ei.hWndClone3};
  int nCloneCount=sizeof(lpClones) / sizeof(HWND);
  int i;

  if (lpFrameCurrent->ei.hWndMaster && lpFrameCurrent->ei.hWndEdit)
  {
    for (i=0; i < nCloneCount; ++i)
    {
      if (lpFrameCurrent->ei.hWndEdit == lpClones[i])
        break;
    }

    while (i < nCloneCount)
    {
      if (bPrevious)
      {
        if (--i < 0)
          i=nCloneCount - 1;
      }
      else
      {
        if (++i >= nCloneCount)
          i=0;
      }

      if (lpClones[i])
      {
        SetFocus(lpClones[i]);
        return lpClones[i];
      }
    }
  }
  return NULL;
}

void UpdateSize()
{
  int nTabHeight;
  int nEditHeight;
  BOOL bStatusBar=FALSE;

  if (!hDocksStack.bSizing)
  {
    hDocksStack.bSizing=TRUE;

    GetClientRect(hMainWnd, &nsSize.rcInitial);
    if (moCur.bStatusBar && IsWindowVisible(hStatus))
    {
      nsSize.rcInitial.bottom-=nStatusHeight;
      bStatusBar=TRUE;
    }
    nsSize.rcCurrent=nsSize.rcInitial;
    SendMessage(hMainWnd, AKDN_SIZE_ONSTART, 0, (LPARAM)&nsSize);

    //Docks
    StackDockSize(&hDocksStack, (hDocksStack.nSizingSide == DKS_LEFT)?DKS_RIGHT:DKS_LEFT, &nsSize);
    StackDockSize(&hDocksStack, (hDocksStack.nSizingSide == DKS_LEFT)?DKS_LEFT:DKS_RIGHT, &nsSize);
    StackDockSize(&hDocksStack, (hDocksStack.nSizingSide == DKS_TOP)?DKS_BOTTOM:DKS_TOP, &nsSize);
    StackDockSize(&hDocksStack, (hDocksStack.nSizingSide == DKS_TOP)?DKS_TOP:DKS_BOTTOM, &nsSize);
    hDocksStack.nSizingSide=0;

    //Edits
    if (!nMDI || (moCur.dwTabOptionsMDI & TAB_VIEW_NONE) || !IsWindowVisible(hTab))
      nTabHeight=0;
    else
      nTabHeight=TAB_HEIGHT;
    nEditHeight=nsSize.rcCurrent.bottom - nTabHeight;

    if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
    {
      fdDefault.rcEditWindow.left=nsSize.rcCurrent.left;
      fdDefault.rcEditWindow.top=nsSize.rcCurrent.top + ((moCur.dwTabOptionsMDI & TAB_VIEW_TOP)?nTabHeight:0);
      fdDefault.rcEditWindow.right=nsSize.rcCurrent.right;
      fdDefault.rcEditWindow.bottom=nEditHeight;
      ResizeEditWindow(lpFrameCurrent, 0);
    }
    if (nMDI)
    {
      if (nTabHeight && ((moCur.dwTabOptionsMDI & TAB_VIEW_TOP) || (moCur.dwTabOptionsMDI & TAB_VIEW_BOTTOM)))
      {
        MoveWindow(hTab, nsSize.rcCurrent.left, nsSize.rcCurrent.top + ((moCur.dwTabOptionsMDI & TAB_VIEW_BOTTOM)?nEditHeight:0), nsSize.rcCurrent.right, nTabHeight, FALSE);
        InvalidateRect(hTab, NULL, TRUE);
        if (bStatusBar && nMDI == WMD_PMDI)
          InvalidateRect(hStatus, NULL, TRUE);
      }
      if (nMDI == WMD_MDI)
        MoveWindow(hMdiClient, nsSize.rcCurrent.left, nsSize.rcCurrent.top + ((moCur.dwTabOptionsMDI & TAB_VIEW_TOP)?nTabHeight:0), nsSize.rcCurrent.right, nEditHeight, TRUE);
    }
    SendMessage(hMainWnd, AKDN_SIZE_ONFINISH, 0, (LPARAM)&nsSize);
    hDocksStack.bSizing=FALSE;
  }
}

BOOL DialogResizeMessages(DIALOGRESIZE *drs, RECT *rcMinMax, RECT *rcCurrent, DWORD dwFlags, HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
  if (uMsg == WM_INITDIALOG)
  {
    RECT rcTemplate;
    RECT rcControl;
    DWORD dwFlags=SWP_NOMOVE;
    int i;

    rcTemplate=*rcCurrent;
    GetWindowPos(hDlg, NULL, rcCurrent);

    for (i=0; drs[i].lpWnd; ++i)
    {
      if (*drs[i].lpWnd)
      {
        GetWindowPos(*drs[i].lpWnd, hDlg, &rcControl);
        if (drs[i].dwType & DRS_SIZE)
        {
          if (drs[i].dwType & DRS_X)
            drs[i].nOffset=rcCurrent->right - (rcControl.left + rcControl.right);
          else if (drs[i].dwType & DRS_Y)
            drs[i].nOffset=rcCurrent->bottom - (rcControl.top + rcControl.bottom);
        }
        else if (drs[i].dwType & DRS_MOVE)
        {
          if (drs[i].dwType & DRS_X)
            drs[i].nOffset=rcCurrent->right - rcControl.left;
          else if (drs[i].dwType & DRS_Y)
            drs[i].nOffset=rcCurrent->bottom - rcControl.top;
        }
      }
    }

    if (rcTemplate.right && rcTemplate.bottom)
    {
      if (GetWindowLongPtrWide(hDlg, GWL_STYLE) & DS_CENTER)
      {
        rcTemplate.left=rcCurrent->left + (rcCurrent->right - rcTemplate.right) / 2;
        rcTemplate.top=rcCurrent->top + (rcCurrent->bottom - rcTemplate.bottom) / 2;
        dwFlags&=~SWP_NOMOVE;
      }
      SetWindowPos(hDlg, 0, rcTemplate.left, rcTemplate.top, rcTemplate.right, rcTemplate.bottom, dwFlags|SWP_NOZORDER|SWP_NOACTIVATE);
    }
  }
  else if (uMsg == WM_GETMINMAXINFO)
  {
    if (rcMinMax)
    {
      MINMAXINFO *mmi=(MINMAXINFO *)lParam;

      if (rcMinMax->left)
        mmi->ptMinTrackSize.x=rcMinMax->left;
      if (rcMinMax->top)
        mmi->ptMinTrackSize.y=rcMinMax->top;
      if (rcMinMax->right)
        mmi->ptMaxTrackSize.x=rcMinMax->right;
      if (rcMinMax->bottom)
        mmi->ptMaxTrackSize.y=rcMinMax->bottom;
    }
  }
  else if (uMsg == WM_MOVE)
  {
    if (!(GetWindowLongPtrWide(hDlg, GWL_STYLE) & DS_CENTER))
    {
      RECT rcTemplate;

      GetWindowPos(hDlg, NULL, &rcTemplate);
      rcCurrent->left=rcTemplate.left;
      rcCurrent->top=rcTemplate.top;
      return TRUE;
    }
  }
  else if (uMsg == WM_SIZE)
  {
    if (lParam)
    {
      RECT rcControl;
      DWORD dwFlags;
      int i;

      GetWindowPos(hDlg, NULL, rcCurrent);

      for (i=0; drs[i].lpWnd; ++i)
      {
        if (*drs[i].lpWnd)
        {
          dwFlags=0;
          if (drs[i].dwType & DRS_SIZE)
            dwFlags|=SWP_NOMOVE;
          else if (drs[i].dwType & DRS_MOVE)
            dwFlags|=SWP_NOSIZE;
          else
            continue;

          GetWindowPos(*drs[i].lpWnd, hDlg, &rcControl);
          SetWindowPos(*drs[i].lpWnd, 0, (drs[i].dwType & DRS_X)?(rcCurrent->right - drs[i].nOffset):rcControl.left,
                                         (drs[i].dwType & DRS_Y)?(rcCurrent->bottom - drs[i].nOffset):rcControl.top,
                                         (drs[i].dwType & DRS_X)?(rcCurrent->right - rcControl.left - drs[i].nOffset):rcControl.right,
                                         (drs[i].dwType & DRS_Y)?(rcCurrent->bottom - rcControl.top - drs[i].nOffset):rcControl.bottom,
                                          dwFlags|SWP_NOZORDER|SWP_NOACTIVATE);
        }
      }
      InvalidateRect(hDlg, NULL, TRUE);
      return TRUE;
    }
  }
  else if (uMsg == WM_PAINT)
  {
    if (dwFlags & DRM_PAINTSIZEGRIP)
    {
      PAINTSTRUCT ps;
      RECT rcGrip;
      HDC hDC;

      if (hDC=BeginPaint(hDlg, &ps))
      {
        GetClientRect(hDlg, &rcGrip);
        rcGrip.left=rcGrip.right - GetSystemMetrics(SM_CXVSCROLL);
        rcGrip.top=rcGrip.bottom - GetSystemMetrics(SM_CYVSCROLL);
        DrawFrameControl(hDC, &rcGrip, DFC_SCROLL, DFCS_SCROLLSIZEGRIP);
        EndPaint(hDlg, &ps);
      }
    }
  }
  return FALSE;
}

void GetMovingRect(DOCK *dkData, POINT *ptScreen, MINMAXINFO *mmi, RECT *rcScreen)
{
  RECT rcInitial=nsSize.rcInitial;
  RECT rcCurrent=nsSize.rcCurrent;
  HWND hWndParent=GetParent(dkData->hWnd);

  GetWindowRect(dkData->hWnd, rcScreen);
  ClientToScreenRect(hWndParent, &rcInitial);
  ClientToScreenRect(hWndParent, &rcCurrent);

  if (dkData->nSide == DKS_LEFT)
  {
    rcScreen->right=max(rcScreen->left + mmi->ptMinTrackSize.x, ptScreen->x);
    rcScreen->right=min(rcScreen->left + mmi->ptMaxTrackSize.x, rcScreen->right);
    rcScreen->right=min(rcInitial.right, rcScreen->right);
    rcScreen->right=min(rcCurrent.right - DOCK_MAINMIN_X, rcScreen->right);
    if (rcScreen->right - rcScreen->left < DOCK_BORDER_2X)
      rcScreen->right=rcScreen->left + DOCK_BORDER_2X;
  }
  else if (dkData->nSide == DKS_TOP)
  {
    rcScreen->bottom=max(rcScreen->top + mmi->ptMinTrackSize.y, ptScreen->y);
    rcScreen->bottom=min(rcScreen->top + mmi->ptMaxTrackSize.y, rcScreen->bottom);
    rcScreen->bottom=min(rcInitial.bottom, rcScreen->bottom);
    rcScreen->bottom=min(rcCurrent.bottom - DOCK_MAINMIN_Y, rcScreen->bottom);
    if (rcScreen->bottom - rcScreen->top < DOCK_BORDER_2X)
      rcScreen->bottom=rcScreen->top + DOCK_BORDER_2X;
  }
  else if (dkData->nSide == DKS_RIGHT)
  {
    rcScreen->left=min(rcScreen->right - mmi->ptMinTrackSize.x, ptScreen->x);
    rcScreen->left=max(rcScreen->right - mmi->ptMaxTrackSize.x, rcScreen->left);
    rcScreen->left=max(rcInitial.left, rcScreen->left);
    rcScreen->left=max(rcCurrent.left + DOCK_MAINMIN_X, rcScreen->left);
    if (rcScreen->right - rcScreen->left < DOCK_BORDER_2X)
      rcScreen->left=rcScreen->right - DOCK_BORDER_2X;
  }
  else if (dkData->nSide == DKS_BOTTOM)
  {
    rcScreen->top=min(rcScreen->bottom - mmi->ptMinTrackSize.y, ptScreen->y);
    rcScreen->top=max(rcScreen->bottom - mmi->ptMaxTrackSize.y, rcScreen->top);
    rcScreen->top=max(rcInitial.top, rcScreen->top);
    rcScreen->top=max(rcCurrent.top + DOCK_MAINMIN_Y, rcScreen->top);
    if (rcScreen->bottom - rcScreen->top < DOCK_BORDER_2X)
      rcScreen->top=rcScreen->bottom - DOCK_BORDER_2X;
  }
}

void DrawMovingRect(RECT *rcScreen)
{
  const WORD DotPattern[]={0x00aa, 0x0055, 0x00aa, 0x0055, 0x00aa, 0x0055, 0x00aa, 0x0055};
  RECT rcClient=*rcScreen;
  HDC hDC;
  HBITMAP hBitmap;
  HBRUSH hBrush;
  HBRUSH hBrushOld;

  if (!ScreenToClient(hMainWnd, (POINT *)&rcClient.left))
    return;
  if (!ScreenToClient(hMainWnd, (POINT *)&rcClient.right))
    return;

  if (hDC=GetDC(hMainWnd))
  {
    hBitmap=CreateBitmap(8, 8, 1, 1, DotPattern);
    hBrush=CreatePatternBrush(hBitmap);
    hBrushOld=(HBRUSH)SelectObject(hDC, hBrush);

    PatBlt(hDC, rcClient.left, rcClient.top, DOCK_BORDER_1X, rcClient.bottom - rcClient.top - DOCK_BORDER_1X, PATINVERT);
    PatBlt(hDC, rcClient.left + DOCK_BORDER_1X, rcClient.top, rcClient.right - rcClient.left - DOCK_BORDER_1X, DOCK_BORDER_1X, PATINVERT);
    PatBlt(hDC, rcClient.right - DOCK_BORDER_1X, rcClient.top + DOCK_BORDER_1X, DOCK_BORDER_1X, rcClient.bottom - rcClient.top - DOCK_BORDER_1X, PATINVERT);
    PatBlt(hDC, rcClient.left, rcClient.bottom - DOCK_BORDER_1X, rcClient.right - rcClient.left - DOCK_BORDER_1X, DOCK_BORDER_1X, PATINVERT);

    SelectObject(hDC, hBrushOld);
    DeleteObject(hBrush);
    DeleteObject(hBitmap);
    ReleaseDC(hMainWnd, hDC);
  }
}

int GetMouseEdge(HWND hWnd, POINT *ptScreen)
{
  RECT rc;

  GetWindowRect(hWnd, &rc);

  if (PtInRect(&rc, *ptScreen))
  {
    if (ptScreen->x >= rc.left && ptScreen->x <= rc.left + DOCK_BORDER_1X) return DKS_LEFT;
    if (ptScreen->y >= rc.top && ptScreen->y <= rc.top + DOCK_BORDER_1X) return DKS_TOP;
    if (ptScreen->x <= rc.right && ptScreen->x >= rc.right - DOCK_BORDER_1X) return DKS_RIGHT;
    if (ptScreen->y <= rc.bottom && ptScreen->y >= rc.bottom - DOCK_BORDER_1X) return DKS_BOTTOM;
  }
  return 0;
}

void SetEdgeCursor(int nEdge)
{
  if (nEdge == DKS_LEFT || nEdge == DKS_RIGHT)
    SetCursor(LoadCursor(NULL, IDC_SIZEWE));
  else if (nEdge == DKS_TOP || nEdge == DKS_BOTTOM)
    SetCursor(LoadCursor(NULL, IDC_SIZENS));
  else
    SetCursor(LoadCursor(NULL, IDC_SIZEALL));
}

int GetOppEdge(int nEdge)
{
  if (nEdge == DKS_LEFT) return DKS_RIGHT;
  if (nEdge == DKS_TOP) return DKS_BOTTOM;
  if (nEdge == DKS_RIGHT) return DKS_LEFT;
  if (nEdge == DKS_BOTTOM) return DKS_TOP;
  return 0;
}

BOOL IsCursorOnWindow(HWND hWnd)
{
  RECT rcScreen;
  POINT ptScreen;

  GetCursorPos(&ptScreen);
  GetWindowRect(hWnd, &rcScreen);
  return PtInRect(&rcScreen, ptScreen);
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

BOOL ScreenToClientRect(HWND hWnd, RECT *rc)
{
  if (!ScreenToClient(hWnd, (POINT *)&rc->left))
    return FALSE;
  if (!ScreenToClient(hWnd, (POINT *)&rc->right))
    return FALSE;

  rc->right-=rc->left;
  rc->bottom-=rc->top;
  return TRUE;
}

BOOL ClientToScreenRect(HWND hWnd, RECT *rc)
{
  rc->right+=rc->left;
  rc->bottom+=rc->top;

  if (!ClientToScreen(hWnd, (POINT *)&rc->left))
    return FALSE;
  if (!ClientToScreen(hWnd, (POINT *)&rc->right))
    return FALSE;
  return TRUE;
}

BOOL EnsureWindowInMonitor(RECT *rcWindow)
{
  MONITORINFO mi;
  HMONITOR hMonitor=NULL;
  RECT rcNewWindow;

  if ((DWORD)rcWindow->left == (DWORD)CW_USEDEFAULT)
    return TRUE;

  //Size of the work area on the primary display monitor
  xmemset(&mi, 0, sizeof(MONITORINFO));
  SystemParametersInfoA(SPI_GETWORKAREA, 0, &mi.rcWork, 0);

  //Not in primary monitor
  if (!PtInRect(&mi.rcWork, *(POINT *)&rcWindow->left))
  {
    //Get monitor handle
    if (MonitorFromPointPtr)
      hMonitor=(*MonitorFromPointPtr)(*(POINT *)&rcWindow->left, MONITOR_DEFAULTTONEAREST);

    //Get monitor info
    mi.cbSize=sizeof(MONITORINFO);

    if (GetMonitorInfoAPtr)
      (*GetMonitorInfoAPtr)(hMonitor, &mi);
  }
  else return TRUE;

  //Not in any monitor
  if (!PtInRect(&mi.rcWork, *(POINT *)&rcWindow->left))
  {
    //Make corrections
    if (rcWindow->left >= mi.rcWork.left && rcWindow->left < mi.rcWork.right)
      rcNewWindow.left=rcWindow->left;
    else
      rcNewWindow.left=mi.rcWork.left;
    if (rcWindow->top >= mi.rcWork.top && rcWindow->top < mi.rcWork.bottom)
      rcNewWindow.top=rcWindow->top;
    else
      rcNewWindow.top=mi.rcWork.top;
    rcNewWindow.right=rcWindow->right;
    rcNewWindow.bottom=rcWindow->bottom;
    *rcWindow=rcNewWindow;
    return FALSE;
  }
  else return TRUE;
}

//For WMD_PMDI required: lpFrame == lpFrameCurrent
void UpdateTitle(FRAMEDATA *lpFrame)
{
  const wchar_t *wpFileName;
  int nFileLen;

  //Get file name without path
  wpFileName=GetFileName(lpFrame->wszFile, lpFrame->nFileLen);

  if (nMDI == WMD_SDI || nMDI == WMD_PMDI)
  {
    if (nMDI == WMD_SDI)
    {
      //Set caption of synchronizing window
      SetWindowTextWide(hDummyWindow, lpFrame->wszFile);
    }

    //Set caption of main window
    if (*wpFileName)
    {
      if (nMDI == WMD_SDI)
      {
        xprintfW(wbuf, L"%s - %s", wpFileName, APP_MAIN_TITLEW);
        SetWindowTextWide(hMainWnd, wbuf);
      }
      else if (nMDI == WMD_PMDI)
      {
        xprintfW(wbuf, L"%s - [%s]", APP_MAIN_TITLEW, lpFrame->wszFile);
        SetWindowTextWide(hMainWnd, wbuf);
      }
    }
    else SetWindowTextWide(hMainWnd, APP_MAIN_TITLEW);
  }

  if (nMDI == WMD_MDI || nMDI == WMD_PMDI)
  {
    ASSOCICON *ai;
    TCITEMW tcItem;
    const wchar_t *wpExt;
    HICON hIcon=NULL;
    int nIconIndex=0;
    int nItem;

    if (nMDI == WMD_MDI)
    {
      //Set caption of frame window
      SetWindowTextWide(lpFrame->hWndEditParent, lpFrame->wszFile);
    }

    //Find file icon
    nFileLen=lpFrame->nFileLen;
    if (lpFrame->nStreamOffset)
    {
      lpFrame->wszFile[lpFrame->nStreamOffset]=L'\0';
      nFileLen=lpFrame->nStreamOffset;
    }
    if (wpExt=GetFileExt(lpFrame->wszFile, nFileLen))
    {
      if (!(ai=StackIconGet(&hIconsStack, lpFrame->wszFile, nFileLen, wpExt)))
        if (ai=StackIconInsert(&hIconsStack, lpFrame->wszFile, nFileLen))
          if (ai->hIcon) ImageList_AddIcon(hImageList, ai->hIcon);
      hIcon=ai->hIcon;
      nIconIndex=ai->nIconIndex;
    }
    if (!hIcon)
    {
      hIcon=hIconEmpty;
      nIconIndex=0;
    }
    if (lpFrame->nStreamOffset) lpFrame->wszFile[lpFrame->nStreamOffset]=L':';

    if ((nItem=GetTabItemFromParam(hTab, (LPARAM)lpFrame)) != -1)
    {
      wchar_t wszTabName[MAX_PATH];
      wchar_t wszTabNameAmp[MAX_PATH];

      if (*moCur.wszTabNameFind)
      {
        PATREPLACE pr;

        pr.wpStr=wpFileName;
        pr.wpMaxStr=pr.wpStr + xstrlenW(pr.wpStr);
        pr.wpText=pr.wpStr;
        pr.wpMaxText=pr.wpMaxStr;
        pr.wpPat=moCur.wszTabNameFind;
        pr.wpMaxPat=pr.wpPat + xstrlenW(pr.wpPat);
        pr.wpRep=moCur.wszTabNameRep;
        pr.wpMaxRep=pr.wpRep + xstrlenW(pr.wpRep);
        pr.dwOptions=RESE_ISMATCH;
        pr.wpDelim=NULL;
        pr.wpNewLine=NULL;
        pr.wszResult=wszTabName;
        PatReplace(&pr);

        if (!pr.nReplaceCount)
          xstrcpynW(wszTabName, wpFileName, MAX_PATH);
      }
      else xstrcpynW(wszTabName, wpFileName, MAX_PATH);

      //Replace "&" with "&&"
      FixAmpW(wszTabName, wszTabNameAmp, MAX_PATH);

      //Set tab text
      tcItem.mask=TCIF_TEXT|TCIF_IMAGE;
      tcItem.pszText=wszTabNameAmp;
      tcItem.iImage=nIconIndex;
      TabCtrl_SetItemWide(hTab, nItem, &tcItem);
    }

    //Set frame info
    lpFrame->hIcon=hIcon;
    lpFrame->nIconIndex=nIconIndex;

    if (nMDI == WMD_MDI)
    {
      SendMessage(lpFrame->hWndEditParent, WM_SETICON, ICON_SMALL, (LPARAM)hIcon);
      DrawMenuBar(hMainWnd);
    }
  }
}

void UpdateTabs(HWND hWnd)
{
  static HWND hTabSpin;
  int nPos;

  if (!hTabSpin)
    hTabSpin=GetWindow(hWnd, GW_CHILD);

  if (hTabSpin)
  {
    if (nDocumentsCount < 30000)
    {
      if (nPos=LOWORD(SendMessage(hTabSpin, UDM_GETPOS, 0, 0)))
      {
        SendMessage(hWnd, WM_HSCROLL, MAKELONG(SB_THUMBPOSITION, nPos), 0);
        SendMessage(hWnd, WM_HSCROLL, MAKELONG(SB_ENDSCROLL, 0), 0);
      }
    }
    else
    {
      //Destroy Up-Down control to prevent crash after WM_HSCROLL when many thousands of tabs appeared ~30000.
      DestroyWindow(hTabSpin);
      hTabSpin=NULL;
    }
  }
}

int InsertTabItem(HWND hWnd, int nIndex, LPARAM lParam)
{
  TCITEMW tcItem;
  int nItemCount;
  int nResult=-1;

  if (hWnd)
  {
    nItemCount=(int)SendMessage(hWnd, TCM_GETITEMCOUNT, 0, 0);
    if (nIndex == -1) nIndex=nItemCount;

    tcItem.mask=TCIF_TEXT|TCIF_IMAGE|TCIF_PARAM;
    tcItem.pszText=(wchar_t *)L"";
    tcItem.iImage=0;
    tcItem.lParam=lParam;
    nResult=TabCtrl_InsertItemWide(hWnd, nIndex, &tcItem);

    if (!nItemCount || nIndex < nItemCount) InvalidateRect(hWnd, NULL, TRUE);
  }
  return nResult;
}

LPARAM GetTabParamFromItem(HWND hWnd, int nItem)
{
  TCITEMW tcItem;

  tcItem.mask=TCIF_PARAM;
  if (TabCtrl_GetItemWide(hWnd, nItem, &tcItem))
    return tcItem.lParam;
  return 0;
}

int GetTabItemFromParam(HWND hWnd, LPARAM lParam)
{
  TCITEMW tcItem;
  int nCurSel;
  int i;

  nCurSel=(int)SendMessage(hWnd, TCM_GETCURSEL, 0, 0);
  tcItem.mask=TCIF_PARAM;
  TabCtrl_GetItemWide(hWnd, nCurSel, &tcItem);
  if (tcItem.lParam == lParam)
    return nCurSel;

  for (i=0; TabCtrl_GetItemWide(hWnd, i, &tcItem); ++i)
  {
    if (i != nCurSel)
    {
      if (tcItem.lParam == lParam)
        return i;
    }
  }
  return -1;
}

int GetTabItemFromCursorPos(HWND hWnd)
{
  TCHITTESTINFO thti;

  GetCursorPos(&thti.pt);
  ScreenToClient(hWnd, &thti.pt);
  return (int)SendMessage(hWnd, TCM_HITTEST, 0, (LPARAM)&thti);
}

int GetTabItemForDrop(HWND hWnd, POINT *rcScreen)
{
  TCHITTESTINFO thti;
  RECT rcTab;
  BOOL bMinus=TRUE;
  BOOL bBetween=FALSE;
  int nTabClientX;
  int nCurSel;
  int i=-1;

  GetWindowRect(hWnd, &rcTab);

  if (rcScreen->x > rcTab.left && rcScreen->y > rcTab.top &&
      rcScreen->x < rcTab.right && rcScreen->y < rcTab.bottom)
  {
    nTabClientX=rcScreen->x - rcTab.left;
    rcTab.right-=rcTab.left;
    rcTab.bottom-=rcTab.top;
    thti.pt.x=nTabClientX;
    thti.pt.y=rcTab.bottom / 2;

    if (SendMessage(hWnd, TCM_HITTEST, 0, (LPARAM)&thti) == -1)
      bBetween=TRUE;

    for (;;)
    {
      if ((i=(int)SendMessage(hWnd, TCM_HITTEST, 0, (LPARAM)&thti)) != -1) break;

      if (bMinus)
      {
        if (--thti.pt.x < 0)
        {
          thti.pt.x=nTabClientX + 1;
          bMinus=FALSE;
        }
      }
      else if (++thti.pt.x > rcTab.right) break;
    }

    nCurSel=(int)SendMessage(hWnd, TCM_GETCURSEL, 0, 0);
    if (i < nCurSel && bBetween && bMinus) ++i;
  }
  return i;
}

int SelectTabItem(HWND hWnd, int nIndex)
{
  NMHDR nmhdr={hWnd, ID_TAB, TCN_SELCHANGE};
  int nCurSel;

  nCurSel=(int)SendMessage(hWnd, TCM_GETCURSEL, 0, 0);

  if (nIndex != nCurSel)
  {
    if (SendMessage(hWnd, TCM_SETCURSEL, nIndex, 0) != -1)
    {
      SendMessage(GetParent(hWnd), WM_NOTIFY, ID_TAB, (LPARAM)&nmhdr);
    }
    else return -1;
  }
  return nCurSel;
}

int MoveTabItem(HWND hWnd, int nIndexOld, int nIndexNew)
{
  TCITEMW tcItem;
  wchar_t wszItemText[MAX_PATH];

  tcItem.mask=TCIF_TEXT|TCIF_IMAGE|TCIF_PARAM;
  tcItem.pszText=wszItemText;
  tcItem.cchTextMax=MAX_PATH;
  tcItem.iImage=0;
  tcItem.lParam=0;

  if (TabCtrl_GetItemWide(hWnd, nIndexOld, &tcItem))
  {
    if (nIndexOld == nIndexNew)
      return nIndexNew;
    SendMessage(hWnd, TCM_DELETEITEM, nIndexOld, 0);
    nIndexNew=TabCtrl_InsertItemWide(hWnd, nIndexNew, &tcItem);

    if (nIndexOld == nDocumentIndex && nIndexNew >= 0)
    {
      nDocumentIndex=nIndexNew;
      UpdateStatusUser(lpFrameCurrent, CSB_DOCUMENTINDEX);
    }
    return nIndexNew;
  }
  return -1;
}

BOOL DeleteTabItem(HWND hWnd, int nIndex)
{
  if (SendMessage(hWnd, TCM_DELETEITEM, nIndex, 0))
  {
    if (nIndex == nDocumentIndex)
    {
      if ((nDocumentIndex=(int)SendMessage(hWnd, TCM_GETCURSEL, 0, 0)) == -1)
        nDocumentIndex=0;
      UpdateStatusUser(lpFrameCurrent, CSB_DOCUMENTINDEX);
    }
    return TRUE;
  }
  return FALSE;
}

void FreeMemorySearch()
{
  if (wszFindText)
  {
    FreeWideStr(wszFindText);
    wszFindText=NULL;
  }
  if (wszReplaceText)
  {
    FreeWideStr(wszReplaceText);
    wszReplaceText=NULL;
  }
}

int BytesInString(const wchar_t *wpString)
{
  return (int)(xstrlenW(wpString) + 1) * sizeof(wchar_t);
}

const wchar_t* FindArrayByIndex(const wchar_t *wpString, int nIndex)
{
  int nCount=1;

  if (nCount < nIndex)
  {
    for (;;)
    {
      if (*wpString == L'\0')
      {
        if (*++wpString == L'\0')
          return NULL;
        if (++nCount >= nIndex)
          break;
      }
      ++wpString;
    }
  }
  return wpString;
}

char* AKD_strchr(const char *s, int c)
{
  if (c == '\r' || c == '\n')
  {
    while (*s != '\0' && *s != '\r' && *s != '\n')
      ++s;
    if (*s != '\0')
      return ((char *)s);
    return NULL;
  }
  else
  {
    while (*s != '\0' && *s != (char)c)
      ++s;
    if (*s != '\0')
      return ((char *)s);
    return NULL;
  }
}

wchar_t* AKD_wcschr(const wchar_t *s, wchar_t c)
{
  if (c == L'\r' || c == L'\n')
  {
    while (*s != L'\0' && *s != L'\r' && *s != L'\n')
      ++s;
    if (*s != L'\0')
      return ((wchar_t *)s);
    return NULL;
  }
  else
  {
    while (*s != L'\0' && *s != c)
      ++s;
    if (*s != L'\0')
      return ((wchar_t *)s);
    return NULL;
  }
}


//// API functions replacement

HMENU API_LoadMenuA(HINSTANCE hLoadInstance, char *lpMenuName)
{
  HMENU hResult;

  if (!(hResult=LoadMenuA(hLoadInstance, lpMenuName)))
    if (hLoadInstance != hInstance)
      hResult=LoadMenuA(hInstance, lpMenuName);

  return hResult;
}

HMENU API_LoadMenuW(HINSTANCE hLoadInstance, wchar_t *lpMenuName)
{
  HMENU hResult;

  if (!(hResult=LoadMenuW(hLoadInstance, lpMenuName)))
    if (hLoadInstance != hInstance)
      hResult=LoadMenuW(hInstance, lpMenuName);

  return hResult;
}

HACCEL API_LoadAcceleratorsA(HINSTANCE hLoadInstance, char *lpTableName)
{
  HACCEL hResult;

  if (!(hResult=LoadAcceleratorsA(hLoadInstance, lpTableName)))
    if (hLoadInstance != hInstance)
      hResult=LoadAcceleratorsA(hInstance, lpTableName);

  return hResult;
}

HACCEL API_LoadAcceleratorsW(HINSTANCE hLoadInstance, wchar_t *lpTableName)
{
  HACCEL hResult;

  if (!(hResult=LoadAcceleratorsW(hLoadInstance, lpTableName)))
    if (hLoadInstance != hInstance)
      hResult=LoadAcceleratorsW(hInstance, lpTableName);

  return hResult;
}

HICON API_LoadIconA(HINSTANCE hLoadInstance, char *lpIconName)
{
  HICON hResult;

  if (!(hResult=LoadIconA(hLoadInstance, lpIconName)))
    if (hLoadInstance != hInstance)
      hResult=LoadIconA(hInstance, lpIconName);

  return hResult;
}

HICON API_LoadIconW(HINSTANCE hLoadInstance, wchar_t *lpIconName)
{
  HICON hResult;

  if (!(hResult=LoadIconW(hLoadInstance, lpIconName)))
    if (hLoadInstance != hInstance)
      hResult=LoadIconW(hInstance, lpIconName);

  return hResult;
}

HANDLE API_LoadImageA(HINSTANCE hLoadInstance, char *lpszName, UINT uType, int cxDesired, int cyDesired, UINT fuLoad)
{
  HANDLE hResult;

  if (!(hResult=LoadImageA(hLoadInstance, lpszName, uType, cxDesired, cyDesired, fuLoad)))
    if (hLoadInstance != hInstance)
      hResult=LoadImageA(hInstance, lpszName, uType, cxDesired, cyDesired, fuLoad);

  return hResult;
}

HANDLE API_LoadImageW(HINSTANCE hLoadInstance, wchar_t *lpszName, UINT uType, int cxDesired, int cyDesired, UINT fuLoad)
{
  HANDLE hResult;

  if (!(hResult=LoadImageW(hLoadInstance, lpszName, uType, cxDesired, cyDesired, fuLoad)))
    if (hLoadInstance != hInstance)
      hResult=LoadImageW(hInstance, lpszName, uType, cxDesired, cyDesired, fuLoad);

  return hResult;
}

int API_LoadStringA(HINSTANCE hLoadInstance, UINT uID, char *lpBuffer, int nBufferMax)
{
  int nResult;

  if (!(nResult=LoadStringA(hLoadInstance, uID, lpBuffer, nBufferMax)))
    if (hLoadInstance != hInstance)
      if (!(nResult=LoadStringA(hInstance, uID, lpBuffer, nBufferMax)))
        lpBuffer[0]='\0';

  return nResult;
}

int API_LoadStringW(HINSTANCE hLoadInstance, UINT uID, wchar_t *lpBuffer, int nBufferMax)
{
  int nResult;

  if (!(nResult=LoadStringWide(hLoadInstance, uID, lpBuffer, nBufferMax)))
    if (hLoadInstance != hInstance)
      if (!(nResult=LoadStringWide(hInstance, uID, lpBuffer, nBufferMax)))
        lpBuffer[0]=L'\0';

  return nResult;
}

int API_MessageBox(HWND hWnd, const wchar_t *lpText, const wchar_t *lpCaption, UINT uType)
{
  HWND hWndParent=hWnd;
  DWORD dwStyle;
  int nResult;

  if (hMainWnd) SendMessage(hMainWnd, AKDN_MESSAGEBOXBEGIN, (WPARAM)hWnd, 0);

  if (hWnd)
  {
    dwStyle=(DWORD)GetWindowLongPtrWide(hWnd, GWL_STYLE);
    if (/*(dwStyle & WS_DISABLED) || */!(dwStyle & WS_VISIBLE))
      hWndParent=NULL;
  }
  nResult=MessageBoxW(hWndParent, lpText, lpCaption, uType);

  if (hMainWnd) SendMessage(hMainWnd, AKDN_MESSAGEBOXEND, (WPARAM)hWnd, 0);

  return nResult;
}

HWND API_CreateDialog(HINSTANCE hLoadInstance, wchar_t *lpTemplateName, HWND hWndParent, DLGPROC lpDialogFunc)
{
  HWND hResult;

  if (!(hResult=CreateDialogWide(hLoadInstance, lpTemplateName, hWndParent, lpDialogFunc)))
    if (hLoadInstance != hInstance)
      hResult=CreateDialogWide(hInstance, lpTemplateName, hWndParent, lpDialogFunc);

  return hResult;
}

INT_PTR API_DialogBox(HINSTANCE hLoadInstance, wchar_t *lpTemplateName, HWND hWndParent, DLGPROC lpDialogFunc)
{
  INT_PTR nResult;

  if ((nResult=DialogBoxWide(hLoadInstance, lpTemplateName, hWndParent, lpDialogFunc)) == -1)
    if (hLoadInstance != hInstance)
      nResult=DialogBoxWide(hInstance, lpTemplateName, hWndParent, lpDialogFunc);

  return nResult;
}

INT_PTR API_DialogBoxParam(HINSTANCE hLoadInstance, wchar_t *lpTemplateName, HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
  INT_PTR nResult;

  if ((nResult=DialogBoxParamWide(hLoadInstance, lpTemplateName, hWndParent, lpDialogFunc, dwInitParam)) == -1)
    if (hLoadInstance != hInstance)
      nResult=DialogBoxParamWide(hInstance, lpTemplateName, hWndParent, lpDialogFunc, dwInitParam);

  return nResult;
}

HANDLE API_CreateFileA(const char *lpFileName, DWORD dwDesiredAccess, DWORD dwShareMode, LPSECURITY_ATTRIBUTES lpSecurityAttributes, DWORD dwCreationDisposition, DWORD dwFlagsAndAttributes, HANDLE hTemplateFile)
{
  wchar_t wszFile[MAX_PATH];
  HANDLE hResult;

  if ((hResult=CreateFileA(lpFileName, dwDesiredAccess, dwShareMode, lpSecurityAttributes, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)) == INVALID_HANDLE_VALUE)
  {
    MultiByteToWideChar(CP_ACP, 0, lpFileName, -1, wszFile, MAX_PATH);
    API_LoadStringW(hLangLib, MSG_CANNOT_OPEN_FILE, wbuf, BUFFER_SIZE);
    xprintfW(wszMsg, wbuf, wszFile);
    API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
  }

  return hResult;
}

HANDLE API_CreateFileW(const wchar_t *lpFileName, DWORD dwDesiredAccess, DWORD dwShareMode, LPSECURITY_ATTRIBUTES lpSecurityAttributes, DWORD dwCreationDisposition, DWORD dwFlagsAndAttributes, HANDLE hTemplateFile)
{
  HANDLE hResult;

  if ((hResult=CreateFileWide(lpFileName, dwDesiredAccess, dwShareMode, lpSecurityAttributes, dwCreationDisposition, dwFlagsAndAttributes, hTemplateFile)) == INVALID_HANDLE_VALUE)
  {
    API_LoadStringW(hLangLib, MSG_CANNOT_OPEN_FILE, wbuf, BUFFER_SIZE);
    xprintfW(wszMsg, wbuf, lpFileName);
    API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
  }

  return hResult;
}

BOOL API_WriteFile(HANDLE hFile, LPCVOID lpBuffer, UINT_PTR nNumberOfBytesToWrite, UINT_PTR *lpNumberOfBytesWritten, LPOVERLAPPED lpOverlapped)
{
  if (!WriteFile64(hFile, lpBuffer, nNumberOfBytesToWrite, lpNumberOfBytesWritten, lpOverlapped))
  {
    API_LoadStringW(hLangLib, MSG_ERROR_IO, wszMsg, BUFFER_SIZE);
    API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
    return FALSE;
  }
  return TRUE;
}

LPVOID API_HeapAlloc(HANDLE hHeap, DWORD dwFlags, SIZE_T dwBytes)
{
  LPVOID lpResult;

  if (!(lpResult=HeapAlloc(hHeap, dwFlags, dwBytes)))
  {
    API_LoadStringW(hLangLib, MSG_ERROR_NOT_ENOUGH_MEMORY, wszMsg, BUFFER_SIZE);
    API_MessageBox(hMainWnd, wszMsg, APP_MAIN_TITLEW, MB_OK|MB_ICONERROR);
  }

  return lpResult;
}

BOOL API_HeapFree(HANDLE hHeap, DWORD dwFlags, LPVOID lpMem)
{
  BOOL bResult;

  bResult=HeapFree(hHeap, dwFlags, lpMem);

  return bResult;
}

wchar_t* AllocWideStr(SIZE_T dwSize)
{
  return (wchar_t *)API_HeapAlloc(hHeap, 0, dwSize * sizeof(wchar_t));
}

BOOL FreeWideStr(wchar_t *wpVar)
{
  if (wpVar)
    return API_HeapFree(hHeap, 0, (LPVOID)wpVar);
  return FALSE;
}
