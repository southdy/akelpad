#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <activscp.h>
#include <stddef.h>
#include "Scripts.h"


//Global variables
const IActiveScriptSiteVtbl MyIActiveScriptSiteVtbl={
  QueryInterface,
  AddRef,
  Release,
  GetLCID,
  GetItemInfo,
  GetDocVersionString,
  OnScriptTerminate,
  OnStateChange,
  OnScriptError,
  OnEnterScript,
  OnLeaveScript
};

const IActiveScriptSiteWindowVtbl MyIActiveScriptSiteWindowVtbl={
  SiteWindow_QueryInterface,
  SiteWindow_AddRef,
  SiteWindow_Release,
  GetSiteWindow,
  EnableModeless
};

#ifdef _WIN64
  IActiveScriptSiteDebug64Vtbl
#else
  IActiveScriptSiteDebug32Vtbl
#endif
MyIActiveScriptSiteDebugVtbl={
   SiteDebug_QueryInterface,
   SiteDebug_AddRef,
   SiteDebug_Release,
   SiteDebug_GetDocumentContextFromPosition,
   SiteDebug_GetApplication,
   SiteDebug_GetRootApplicationNode,
   SiteDebug_OnScriptErrorDebug
 };

HHOOK hHookMessageBox;


//// Script text execution

HRESULT ExecScriptText(void *lpScriptThread, GUID *guidEngine)
{
  SCRIPTTHREAD *st=(SCRIPTTHREAD *)lpScriptThread;
  HRESULT nCoInit;
  HRESULT nResult=E_FAIL;
  DWORD dwDebugApplicationCookie=0;
  MYDWORD_PTR dwDebugSourceContext=0;
  BOOL bInitDebugJIT=FALSE;

  nCoInit=CoInitialize(0);

  if (st->dwDebugJIT & JIT_DEBUG)
  {
    if (CoCreateInstance(&CLSID_ProcessDebugManager, 0, CLSCTX_INPROC_SERVER, &IID_IProcessDebugManager, (void **)&st->objProcessDebugManager) == S_OK)
    {
      if (st->objProcessDebugManager->lpVtbl->CreateApplication(st->objProcessDebugManager, &st->objDebugApplication) == S_OK)
      {
        if (st->objDebugApplication->lpVtbl->SetName(st->objDebugApplication, L"AkelPad Scripts") == S_OK)
        {
          if (st->objProcessDebugManager->lpVtbl->AddApplication(st->objProcessDebugManager, st->objDebugApplication, &dwDebugApplicationCookie) == S_OK)
          {
            //Helper
            if (st->objProcessDebugManager->lpVtbl->CreateDebugDocumentHelper(st->objProcessDebugManager, NULL, &st->objDebugDocumentHelper) == S_OK)
            {
              if (st->objDebugDocumentHelper->lpVtbl->Init(st->objDebugDocumentHelper, st->objDebugApplication, st->wszScriptName, st->wszScriptName, TEXT_DOC_ATTR_READONLY) == S_OK)
              {
                if (st->objDebugDocumentHelper->lpVtbl->Attach(st->objDebugDocumentHelper, NULL) == S_OK)
                {
                  if (st->objDebugDocumentHelper->lpVtbl->SetDocumentAttr(st->objDebugDocumentHelper, TEXT_DOC_ATTR_READONLY) == S_OK)
                  {
                    bInitDebugJIT=TRUE;
                  }
                }
              }
            }
          }
        }
      }
    }
    if (!bInitDebugJIT)
      MessageBoxW(hMainWnd, GetLangStringW(wLangModule, STRID_DEBUG_ERROR), wszPluginTitle, MB_OK|MB_ICONEXCLAMATION);
  }

  if (CoCreateInstance(guidEngine, 0, CLSCTX_INPROC_SERVER, &IID_IActiveScript, (void **)&st->objActiveScript) == S_OK)
  {
    if (st->objActiveScript->lpVtbl->QueryInterface(st->objActiveScript, &IID_IActiveScriptParse, (void **)&st->objActiveScriptParse) == S_OK)
    {
      if (st->objActiveScriptParse->lpVtbl->InitNew(st->objActiveScriptParse) == S_OK)
      {
        st->MyActiveScriptSite.lpVtbl=&MyIActiveScriptSiteVtbl;
        st->MyActiveScriptSite.dwCount=1;
        st->MyActiveScriptSite.lpScriptThread=st;
        st->MyActiveScriptSiteWindow.lpVtbl=&MyIActiveScriptSiteWindowVtbl;
        st->MyActiveScriptSiteWindow.dwCount=1;
        st->MyActiveScriptSiteWindow.lpScriptThread=st;
        st->MyActiveScriptSiteDebug.lpVtbl=&MyIActiveScriptSiteDebugVtbl;
        st->MyActiveScriptSiteDebug.dwCount=1;
        st->MyActiveScriptSiteDebug.lpScriptThread=st;

        if (st->objActiveScript->lpVtbl->SetScriptSite(st->objActiveScript, (IActiveScriptSite *)&st->MyActiveScriptSite) == S_OK)
        {
          if (st->objActiveScript->lpVtbl->AddNamedItem(st->objActiveScript, L"WScript", SCRIPTITEM_ISVISIBLE|SCRIPTITEM_NOCODE) == S_OK &&
              st->objActiveScript->lpVtbl->AddNamedItem(st->objActiveScript, L"AkelPad", SCRIPTITEM_ISVISIBLE|SCRIPTITEM_NOCODE) == S_OK &&
              st->objActiveScript->lpVtbl->AddNamedItem(st->objActiveScript, L"Constants", SCRIPTITEM_GLOBALMEMBERS|SCRIPTITEM_ISVISIBLE|SCRIPTITEM_NOCODE) == S_OK)
          {
            if (bInitDebugJIT)
            {
              if (st->objDebugDocumentHelper->lpVtbl->AddUnicodeText(st->objDebugDocumentHelper, st->wpScriptText) == S_OK)
              {
                if (st->objDebugDocumentHelper->lpVtbl->DefineScriptBlock(st->objDebugDocumentHelper, 0, (ULONG)st->nScriptTextLen, st->objActiveScript, FALSE, &dwDebugSourceContext) == S_OK)
                {
                  st->bInitDebugJIT=TRUE;
                  if (st->dwDebugJIT & JIT_FROMSTART)
                    st->objDebugApplication->lpVtbl->CauseBreak(st->objDebugApplication);
                }
              }
            }
            if (st->objActiveScriptParse->lpVtbl->ParseScriptText(st->objActiveScriptParse, st->wpScriptText, NULL, NULL, NULL, dwDebugSourceContext, 0, st->bInitDebugJIT?SCRIPTTEXT_HOSTMANAGESSOURCE:0, NULL, NULL) == S_OK)
            {
              st->objActiveScript->lpVtbl->SetScriptState(st->objActiveScript, SCRIPTSTATE_CONNECTED);
              nResult=S_OK;
            }
          }
        }
        st->MyActiveScriptSite.lpVtbl->Release((IActiveScriptSite *)&st->MyActiveScriptSite);
        st->MyActiveScriptSiteWindow.lpVtbl->Release((IActiveScriptSiteWindow *)&st->MyActiveScriptSiteWindow);
        st->MyActiveScriptSiteDebug.lpVtbl->Release((IActiveScriptSiteDebug *)&st->MyActiveScriptSiteDebug);
      }
      st->objActiveScriptParse->lpVtbl->Release(st->objActiveScriptParse);
    }
    st->objActiveScript->lpVtbl->Close(st->objActiveScript);
    st->objActiveScript->lpVtbl->Release(st->objActiveScript);
  }
  if (st->objDebugDocumentHelper)
  {
    st->objDebugDocumentHelper->lpVtbl->Detach(st->objDebugDocumentHelper);
    st->objDebugDocumentHelper->lpVtbl->Release(st->objDebugDocumentHelper);
  }
  if (st->objDebugApplication)
    st->objDebugApplication->lpVtbl->Close(st->objDebugApplication);
  if (st->objProcessDebugManager)
    st->objProcessDebugManager->lpVtbl->Release(st->objProcessDebugManager);

  if (nCoInit == S_OK)
    CoUninitialize();

  return nResult;
}

HRESULT GetScriptEngineA(char *szExt, GUID *guidEngine)
{
  char szValue[MAX_PATH];
  wchar_t wszValue[MAX_PATH];
  HKEY hKey;
  HKEY hSubKey;
  DWORD dwType;
  DWORD dwSize;
  int nQuery;
  HRESULT nResult=E_FAIL;

  if (RegOpenKeyExA(HKEY_CLASSES_ROOT, szExt, 0, KEY_QUERY_VALUE|KEY_READ, &hKey) == ERROR_SUCCESS)
  {
    dwSize=sizeof(szValue);
    nQuery=RegQueryValueExA(hKey, NULL, 0, &dwType, (LPBYTE)szValue, &dwSize);
    RegCloseKey(hKey);

    if (nQuery == ERROR_SUCCESS)
    {
      Loop:

      if (RegOpenKeyExA(HKEY_CLASSES_ROOT, (LPCSTR)szValue, 0, KEY_QUERY_VALUE|KEY_READ, &hKey) == ERROR_SUCCESS)
      {
        if (RegOpenKeyExA(hKey, "CLSID", 0, KEY_QUERY_VALUE|KEY_READ, &hSubKey) == ERROR_SUCCESS)
        {
          dwSize=sizeof(szValue);
          nQuery=RegQueryValueExA(hSubKey, 0, 0, &dwType, (LPBYTE)szValue, &dwSize);
          RegCloseKey(hSubKey);

          MultiByteToWideChar(CP_ACP, 0, szValue, -1, wszValue, MAX_PATH);
          if (CLSIDFromString(wszValue, guidEngine) == NOERROR)
            nResult=S_OK;
        }
        else if (szExt)
        {
          if (RegOpenKeyExA(hKey, "ScriptEngine", 0, KEY_QUERY_VALUE|KEY_READ, &hSubKey) == ERROR_SUCCESS)
          {
            dwSize=sizeof(szValue);
            nQuery=RegQueryValueExA(hSubKey, 0, 0, &dwType, (LPBYTE)szValue, &dwSize);
            RegCloseKey(hSubKey);

            if (nQuery == ERROR_SUCCESS)
            {
              RegCloseKey(hKey);
              szExt=NULL;
              goto Loop;
            }
          }
        }
      }
      RegCloseKey(hKey);
    }
  }
  return nResult;
}

//MessageBox with custom button text
int CBTMessageBox(HWND hWnd, const wchar_t *wpText, const wchar_t *wpCaption, UINT uType)
{
  hHookMessageBox=SetWindowsHookEx(WH_CBT, &CBTMessageBoxProc, 0, GetCurrentThreadId());
  return MessageBoxW(hWnd, wpText, wpCaption, uType);
}

LRESULT CALLBACK CBTMessageBoxProc(INT nCode, WPARAM wParam, LPARAM lParam)
{
  if (nCode == HCBT_ACTIVATE)
  {
    if (GetDlgItem((HWND)wParam, IDYES))
      SetDlgItemTextWide((HWND)wParam, IDYES, GetLangStringW(wLangModule, STRID_STOP));
    if (GetDlgItem((HWND)wParam, IDNO))
      SetDlgItemTextWide((HWND)wParam, IDNO, GetLangStringW(wLangModule, STRID_EDIT));
    if (GetDlgItem((HWND)wParam, IDCANCEL))
      SetDlgItemTextWide((HWND)wParam, IDCANCEL, GetLangStringW(wLangModule, STRID_CONTINUE));
    UnhookWindowsHookEx(hHookMessageBox);
    hHookMessageBox=NULL;
    return 0;
  }
  else return CallNextHookEx(hHookMessageBox, nCode, wParam, lParam);
}


//// IActiveScriptSite implementation

HRESULT STDMETHODCALLTYPE QueryInterface(IActiveScriptSite *this, REFIID riid, void **ppv)
{
  SCRIPTTHREAD *lpScriptThread;

  if (!ppv) return E_POINTER;
  *ppv=NULL;

  if (AKD_IsEqualIID(riid, &IID_IUnknown) || AKD_IsEqualIID(riid, &IID_IActiveScriptSite))
  {
    *ppv=this;
    AddRef(*ppv);
  }
  else if (AKD_IsEqualIID(riid, &IID_IActiveScriptSiteWindow))
  {
    lpScriptThread=(SCRIPTTHREAD *)((IRealActiveScriptSite *)this)->lpScriptThread;
    *ppv=&lpScriptThread->MyActiveScriptSiteWindow;
    SiteWindow_AddRef(*ppv);
  }
  else if (AKD_IsEqualIID(riid, &IID_IActiveScriptSiteDebug))
  {
    lpScriptThread=(SCRIPTTHREAD *)((IRealActiveScriptSite *)this)->lpScriptThread;
    if (lpScriptThread->dwDebugJIT & JIT_DEBUG)
    {
      *ppv=&lpScriptThread->MyActiveScriptSiteDebug;
      SiteDebug_AddRef(*ppv);
    }
  }
  if (!*ppv)
    return E_NOINTERFACE;
  return S_OK;
}

ULONG STDMETHODCALLTYPE AddRef(IActiveScriptSite *this)
{
  return ++((IRealActiveScriptSite *)this)->dwCount;
}

ULONG STDMETHODCALLTYPE Release(IActiveScriptSite *this)
{
  return --((IRealActiveScriptSite *)this)->dwCount;
}

HRESULT STDMETHODCALLTYPE GetItemInfo(IActiveScriptSite *this, LPCOLESTR objectName, DWORD dwReturnMask, IUnknown **objPtr, ITypeInfo **typeInfo)
{
  HRESULT hr=E_FAIL;

  if (dwReturnMask & SCRIPTINFO_IUNKNOWN) *objPtr=0;
  if (dwReturnMask & SCRIPTINFO_ITYPEINFO) *typeInfo=0;

  if (!xstrcmpiW(objectName, L"WScript"))
  {
    if (dwReturnMask & SCRIPTINFO_IUNKNOWN)
      hr=Class_CreateInstance(NULL, NULL, &IID_IWScript, (void **)objPtr);
    if (dwReturnMask & SCRIPTINFO_ITYPEINFO)
      hr=WScript_GetTypeInfo(NULL, 0, 0, typeInfo);
  }
  else if (!xstrcmpiW(objectName, L"AkelPad"))
  {
    if (dwReturnMask & SCRIPTINFO_IUNKNOWN)
      hr=Class_CreateInstance(NULL, NULL, &IID_IDocument, (void **)objPtr);
    if (dwReturnMask & SCRIPTINFO_ITYPEINFO)
      hr=Document_GetTypeInfo(NULL, 0, 0, typeInfo);
  }
  else if (!xstrcmpiW(objectName, L"Constants"))
  {
    if (dwReturnMask & SCRIPTINFO_IUNKNOWN)
      hr=Class_CreateInstance(NULL, NULL, &IID_IConstants, (void **)objPtr);
    if (dwReturnMask & SCRIPTINFO_ITYPEINFO)
      hr=Constants_GetTypeInfo(NULL, 0, 0, typeInfo);
  }
  return hr;
}

HRESULT STDMETHODCALLTYPE OnScriptError(IActiveScriptSite *this, IActiveScriptError *scriptError)
{
  SCRIPTTHREAD *lpScriptThread=(SCRIPTTHREAD *)((IRealActiveScriptSite *)this)->lpScriptThread;
  wchar_t wszScriptFile[MAX_PATH];
  INCLUDEITEM *lpIncludeItem=NULL;
  DWORD dwIncludeIndex=0;
  ULONG nLine;
  LONG nChar;
  EXCEPINFO ei;

  //Error info
  xmemset(&ei, 0, sizeof(EXCEPINFO));
  scriptError->lpVtbl->GetSourcePosition(scriptError, &dwIncludeIndex, &nLine, &nChar);
  scriptError->lpVtbl->GetExceptionInfo(scriptError, &ei);

  if (lpScriptThread->bQuiting)
  {
    //Just in case - SCRIPT_E_PROPAGATE must disconnect without enter here.
    lpScriptThread->objActiveScript->lpVtbl->SetScriptState(lpScriptThread->objActiveScript, SCRIPTSTATE_DISCONNECTED);
    return S_OK;
  }

  //Message text
  if (!*wszErrorMsg)
  {
    if (ei.bstrDescription)
      xstrcpynW(wszErrorMsg, ei.bstrDescription, MAX_PATH);
  }
  if (dwIncludeIndex)
    lpIncludeItem=StackGetInclude(&lpScriptThread->hIncludesStack, dwIncludeIndex);
  if (lpIncludeItem)
    xstrcpynW(wszScriptFile, lpIncludeItem->wszInclude, MAX_PATH);
  else
    xstrcpynW(wszScriptFile, lpScriptThread->wszScriptFile, MAX_PATH);

  xprintfW(wszBuffer, GetLangStringW(wLangModule, STRID_SCRIPTERROR),
                      wszScriptFile,
                      nLine + 1,
                      nChar + 1,
                      wszErrorMsg,
                      ei.scode,
                      ei.bstrSource?ei.bstrSource:L"");

  wszErrorMsg[0]='\0';
  SysFreeString(ei.bstrSource);
  SysFreeString(ei.bstrDescription);
  SysFreeString(ei.bstrHelpFile);

  //Show message
  {
    EDITINFO ei;
    int nChoice;
    int nOpenResult;

    nChoice=CBTMessageBox(IsWindowEnabled(hMainWnd)?hMainWnd:NULL, wszBuffer, wszPluginTitle, (lpScriptThread->hDialogCallbackStack.first?MB_YESNOCANCEL:MB_YESNO)|MB_ICONERROR);

    if (nChoice == IDYES || //"Stop"
        nChoice == IDNO)    //"Edit"
    {
      lpScriptThread->bStopped=TRUE;

      if (!IsWindowEnabled(hMainWnd))
        EnableWindow(hMainWnd, TRUE);

      //Stop script
      //CloseScriptWindows(lpScriptThread);
      if (lpScriptThread->dwMessageLoop)
        PostQuitMessage(0);
      lpScriptThread->objActiveScript->lpVtbl->Close(lpScriptThread->objActiveScript);

      if (nChoice == IDNO) //"Edit"
      {
        Document_OpenFile(NULL, wszScriptFile, OD_ADT_BINARY_ERROR|OD_ADT_DETECT_CODEPAGE|OD_ADT_DETECT_BOM, 0, 0, &nOpenResult);

        if (nOpenResult == EOD_SUCCESS || (nMDI != WMD_SDI && nOpenResult == EOD_WINDOW_EXIST))
        {
          if (SendMessage(hMainWnd, AKD_GETEDITINFO, (WPARAM)NULL, (LPARAM)&ei))
          {
            CHARRANGE64 cr;
            int nLockScroll;

            if (bAkelEdit)
            {
              if ((nLockScroll=(int)SendMessage(ei.hWndEdit, AEM_LOCKSCROLL, (WPARAM)-1, 0)) == -1)
                SendMessage(ei.hWndEdit, AEM_LOCKSCROLL, SB_BOTH, TRUE);
            }
            nLine=(int)SendMessage(ei.hWndEdit, AEM_GETWRAPLINE, (WPARAM)nLine, (LPARAM)NULL);
            cr.cpMin=SendMessage(ei.hWndEdit, EM_LINEINDEX, (WPARAM)nLine, 0) + nChar;
            cr.cpMax=cr.cpMin;
            SendMessage(ei.hWndEdit, EM_EXSETSEL64, 0, (LPARAM)&cr);
            if (bAkelEdit && nLockScroll == -1)
            {
              SendMessage(ei.hWndEdit, AEM_LOCKSCROLL, SB_BOTH, FALSE);
              ScrollCaret(ei.hWndEdit);
            }
          }
        }
      }
      InvalidateRect(hMainWnd, NULL, FALSE);
    }
    else if (nChoice == IDCANCEL) //"Continue"
    {
    }
  }
  return S_OK;
}

HRESULT STDMETHODCALLTYPE GetLCID(IActiveScriptSite *this, LCID *lcid)
{
  *lcid=LOCALE_USER_DEFAULT;
  return S_OK;
}

HRESULT STDMETHODCALLTYPE GetDocVersionString(IActiveScriptSite *this, BSTR *version)
{
  *version=0;
  return S_OK;
}

HRESULT STDMETHODCALLTYPE OnScriptTerminate(IActiveScriptSite *this, const VARIANT *pvr, const EXCEPINFO *pei)
{
  return S_OK;
}

HRESULT STDMETHODCALLTYPE OnStateChange(IActiveScriptSite *this, SCRIPTSTATE state)
{
  return S_OK;
}

HRESULT STDMETHODCALLTYPE OnEnterScript(IActiveScriptSite *this)
{
  return S_OK;
}

HRESULT STDMETHODCALLTYPE OnLeaveScript(IActiveScriptSite *this)
{
  return S_OK;
}


//// IActiveScriptSiteWindow implementation

HRESULT STDMETHODCALLTYPE SiteWindow_QueryInterface(IActiveScriptSiteWindow *this, REFIID riid, void **ppv)
{
  SCRIPTTHREAD *lpScriptThread=(SCRIPTTHREAD *)((IRealActiveScriptSiteWindow *)this)->lpScriptThread;

  return QueryInterface((IActiveScriptSite *)&lpScriptThread->MyActiveScriptSite, riid, ppv);
}

ULONG STDMETHODCALLTYPE SiteWindow_AddRef(IActiveScriptSiteWindow *this)
{
  return ++((IRealActiveScriptSiteWindow *)this)->dwCount;
}

ULONG STDMETHODCALLTYPE SiteWindow_Release(IActiveScriptSiteWindow *this)
{
  return --((IRealActiveScriptSiteWindow *)this)->dwCount;
}

HRESULT STDMETHODCALLTYPE GetSiteWindow(IActiveScriptSiteWindow *this, HWND *phwnd)
{
  *phwnd=hMainWnd;
  return S_OK;
}

HRESULT STDMETHODCALLTYPE EnableModeless(IActiveScriptSiteWindow *this, BOOL enable)
{
  return S_OK;
}


//// IActiveScriptSiteDebug implementation

HRESULT STDMETHODCALLTYPE SiteDebug_QueryInterface(IActiveScriptSiteDebug *this, REFIID riid, void **ppv)
{
  SCRIPTTHREAD *lpScriptThread=(SCRIPTTHREAD *)((IRealActiveScriptSiteDebug *)this)->lpScriptThread;

  return QueryInterface((IActiveScriptSite *)&lpScriptThread->MyActiveScriptSite, riid, ppv);
}

ULONG STDMETHODCALLTYPE SiteDebug_AddRef(IActiveScriptSiteDebug *this)
{
  return ++((IRealActiveScriptSiteDebug *)this)->dwCount;
}

ULONG STDMETHODCALLTYPE SiteDebug_Release(IActiveScriptSiteDebug *this)
{
  return --((IRealActiveScriptSiteDebug *)this)->dwCount;
}

HRESULT STDMETHODCALLTYPE SiteDebug_GetDocumentContextFromPosition(IActiveScriptSiteDebug *this, MYDWORD_PTR dwSourceContext, ULONG uCharacterOffset, ULONG uNumChars, IDebugDocumentContext **ppsc)
{
  SCRIPTTHREAD *lpScriptThread=(SCRIPTTHREAD *)((IRealActiveScriptSiteDebug *)this)->lpScriptThread;
  ULONG ulStartPos=0;
  HRESULT hr;

  if (!ppsc) return E_POINTER;
  *ppsc=NULL;

  if (lpScriptThread->objDebugDocumentHelper)
  {
    if ((hr=lpScriptThread->objDebugDocumentHelper->lpVtbl->GetScriptBlockInfo(lpScriptThread->objDebugDocumentHelper, dwSourceContext, NULL, &ulStartPos, NULL)) == S_OK)
      hr=lpScriptThread->objDebugDocumentHelper->lpVtbl->CreateDebugDocumentContext(lpScriptThread->objDebugDocumentHelper, ulStartPos + uCharacterOffset, uNumChars, ppsc);
  }
  else hr=E_NOTIMPL;

  return hr;
}

HRESULT STDMETHODCALLTYPE SiteDebug_GetApplication(IActiveScriptSiteDebug *this, IDebugApplication **ppda)
{
  SCRIPTTHREAD *lpScriptThread=(SCRIPTTHREAD *)((IRealActiveScriptSiteDebug *)this)->lpScriptThread;

  if (!ppda) return E_POINTER;
  *ppda=lpScriptThread->objDebugApplication;
  return S_OK;
}

HRESULT STDMETHODCALLTYPE SiteDebug_GetRootApplicationNode(IActiveScriptSiteDebug *this, IDebugApplicationNode **ppdanRoot)
{
  SCRIPTTHREAD *lpScriptThread=(SCRIPTTHREAD *)((IRealActiveScriptSiteDebug *)this)->lpScriptThread;

  if (!ppdanRoot) return E_POINTER;
  return lpScriptThread->objDebugDocumentHelper->lpVtbl->GetDebugApplicationNode(lpScriptThread->objDebugDocumentHelper, ppdanRoot);
}

HRESULT STDMETHODCALLTYPE SiteDebug_OnScriptErrorDebug(IActiveScriptSiteDebug *this, IActiveScriptErrorDebug *pErrorDebug, BOOL *pfEnterDebugger, BOOL *pfCallOnScriptErrorWhenContinuing)
{
  SCRIPTTHREAD *lpScriptThread=(SCRIPTTHREAD *)((IRealActiveScriptSiteDebug *)this)->lpScriptThread;

  if (pfEnterDebugger)
  {
    if (lpScriptThread->bQuiting)
      *pfEnterDebugger=FALSE;
    else
      *pfEnterDebugger=TRUE;
  }
  if (pfCallOnScriptErrorWhenContinuing)
    *pfCallOnScriptErrorWhenContinuing=TRUE;
  return S_OK;
}
