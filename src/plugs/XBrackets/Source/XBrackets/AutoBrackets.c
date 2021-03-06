#include "AutoBrackets.h"
#include "Plugin.h"


// ClearType definitions... >>>
#ifndef SPI_GETCLEARTYPE
  #define SPI_GETCLEARTYPE 0x1048
#endif

#ifndef SPI_GETFONTSMOOTHING
  #define SPI_GETFONTSMOOTHING 0x004A
#endif

#ifndef SPI_GETFONTSMOOTHINGTYPE
  #define SPI_GETFONTSMOOTHINGTYPE 0x200A
#endif

#ifndef FE_FONTSMOOTHINGCLEARTYPE
  #define FE_FONTSMOOTHINGCLEARTYPE 0x0002
#endif
// <<< ...ClearType definitions


//#define _verify_ab_
#undef _verify_ab_

//#define _verify_hl_
#undef _verify_hl_

enum TFileType {
  tftNone = 0,
  tftText,
  tftC_Cpp,
  tftH_Hpp,
  tftPas
};

enum TFileType2 {
  tfmNone           = 0x00,
  tfmComment1       = 0x01,
  tfmHtmlCompatible = 0x02,
  tfmEscaped1       = 0x04
};

enum TBracketType {
  tbtNone = 0,
  tbtBracket,  //  (
  tbtSquare,   //  [
  tbtBrace,    //  {
  tbtDblQuote, //  "
  tbtSglQuote, //  '
  tbtTag,      //  <
  tbtTag2,

  tbtCount,

  tbtUser
};

static const char* strBracketsA[tbtCount - 1] = {
  "()",
  "[]",
  "{}",
  "\"\"",
  "\'\'",
  "<>",
  "</>"
};

static const WCHAR* strBracketsW[tbtCount - 1] = {
  L"()",
  L"[]",
  L"{}",
  L"\"\"",
  L"\'\'",
  L"<>",
  L"</>"
};

#define  MAX_EXT             16
#define  MAX_ESCAPED_PREFIX  20
#define  HIGHLIGHT_INDEXES    6
// variables
INT_X IndexesHighlighted[HIGHLIGHT_INDEXES] = { -1, -1, -1, -1, -1, -1 };
INT_X IndexesToHighlight[HIGHLIGHT_INDEXES] = { -1, -1, -1, -1, -1, -1 };
INT_X prevIndexesToHighlight[HIGHLIGHT_INDEXES] = { -1, -1, -1, -1, -1, -1 };
INT_X CurrentBracketsIndexes[2] = { -1, -1 };

#define XBR_FONTSTYLE_UNDEFINED 0xF0F0F0FF
#define XBR_FONTCOLOR_UNDEFINED ((DWORD)(-1))

#define XBR_STATEFLAG_DONOTHL     0x0001
#define XBR_STATEFLAG_INSELECTION 0x0002

typedef struct sCharacterHighlightData {
    DWORD dwState;           // 0 or combination of XBR_STATEFLAG_*
    DWORD dwFontStyle;       // a copy of AEHLPAINT.dwFontStyle (AEHLS_*)
    DWORD dwActiveTextColor; // a copy of AEHLPAINT.dwActiveText
    DWORD dwActiveBkColor;   // a copy of AEHLPAINT.dwActiveBk
} tCharacterHighlightData;

typedef struct sCharacterInfo {
    INT_X nIndex;
    tCharacterHighlightData chd;
} tCharacterInfo;

tCharacterInfo  hgltCharacterInfo[HIGHLIGHT_INDEXES];
tCharacterInfo  prevCharacterInfo[HIGHLIGHT_INDEXES];

int             nCurrentFileType = tftNone;
int             nCurrentFileType2 = tfmNone;
HWND            hCurrentEditWnd = NULL; // Can be NULL! Use hActualEditWnd.
HWND            hActualEditWnd = NULL; //currentEdit;
BOOL            bBracketsInternalRepaint = FALSE;

#if use_aen_paint
UINT            nAenPaintWanted = 0x00;
TEXTMETRICW     AenPaint_tmW;
TEXTMETRICA     AenPaint_tmA;
#endif

INT_X           nAutoRightBracketPos = -1;
int             nAutoRightBracketType = tbtNone;

extern HWND     g_hMainWnd;
extern BOOL     g_bOldWindows;
extern BOOL     g_bOldRichEdit; // TRUE means Rich Edit 2.0
extern BOOL     g_bAkelEdit;

extern UINT     uBracketsHighlight;
extern BOOL     bBracketsHighlightVisibleArea;
extern BOOL     bBracketsRightExistsOK;
extern BOOL     bBracketsDoSingleQuote;
extern BOOL     bBracketsDoTag;
extern BOOL     bBracketsDoTag2;
extern BOOL     bBracketsDoTagIf;
extern BOOL     bBracketsSkipEscaped;
extern BOOL     bBracketsSkipComment1;
extern COLORREF bracketsColourHighlight[2];
extern char     strHtmlFileExtsA[STR_FILEEXTS_SIZE];
extern wchar_t  strHtmlFileExtsW[STR_FILEEXTS_SIZE];
extern char     strEscaped1FileExtsA[STR_FILEEXTS_SIZE];
extern wchar_t  strEscaped1FileExtsW[STR_FILEEXTS_SIZE];
extern char     strComment1FileExtsA[STR_FILEEXTS_SIZE];
extern wchar_t  strComment1FileExtsW[STR_FILEEXTS_SIZE];

extern DWORD    g_dwOptions[OPT_DWORD_COUNT];

wchar_t         strUserBracketsW[MAX_USER_BRACKETS + 1][4];
char            strUserBracketsA[MAX_USER_BRACKETS + 1][4];
wchar_t         strNextCharOkW[MAX_PREV_NEXT_CHAR_OK_SIZE];
wchar_t         strPrevCharOkW[MAX_PREV_NEXT_CHAR_OK_SIZE];


// CharacterHighlightData useful functions
static void CharacterHighlightData_Clear(tCharacterHighlightData* chd)
{
  chd->dwState = 0;
  chd->dwFontStyle = XBR_FONTSTYLE_UNDEFINED;
  chd->dwActiveTextColor = XBR_FONTCOLOR_UNDEFINED;
  chd->dwActiveBkColor = XBR_FONTCOLOR_UNDEFINED;
}
static void CharacterHighlightData_Copy(tCharacterHighlightData* chdDst, const tCharacterHighlightData* chdSrc)
{
  chdDst->dwState = chdSrc->dwState;
  chdDst->dwFontStyle = chdSrc->dwFontStyle;
  chdDst->dwActiveTextColor = chdSrc->dwActiveTextColor;
  chdDst->dwActiveBkColor = chdSrc->dwActiveBkColor;
}

// CharacterInfo useful functions
static void CharacterInfo_ClearItem(tCharacterInfo* ciItem)
{
  ciItem->nIndex = -1;
  CharacterHighlightData_Clear(&ciItem->chd);
}
static void CharacterInfo_ClearAll(tCharacterInfo* ciSet)
{
  int i;
  for (i = 0; i < HIGHLIGHT_INDEXES; i++)
  {
    CharacterInfo_ClearItem(&ciSet[i]);
  }
}
static void CharacterInfo_Copy(tCharacterInfo* ciSetDst, const tCharacterInfo* ciSetSrc)
{
  int i;
  for (i = 0; i < HIGHLIGHT_INDEXES; i++)
  {
    ciSetDst[i].nIndex = ciSetSrc[i].nIndex;
    CharacterHighlightData_Copy(&ciSetDst[i].chd, &ciSetSrc[i].chd);
  }
}
static void CharacterInfo_Add(tCharacterInfo* ciSet, INT_X nIndex, const tCharacterHighlightData* chd)
{
  if (nIndex >= 0)
  {
    int i;
    for (i = 0; i < HIGHLIGHT_INDEXES; i++)
    {
      if ((ciSet[i].nIndex < 0) /*|| (ci[i].nIndex == nIndex)*/) // empty
      {
        ciSet[i].nIndex = nIndex;
        CharacterHighlightData_Copy(&ciSet[i].chd, chd);
        break;
      }
    }
  }
}
static tCharacterHighlightData* CharacterInfo_GetHighlightData(tCharacterInfo* ciSet, INT_X nIndex)
{
  if (nIndex >= 0)
  {
    int i;
    for (i = 0; i < HIGHLIGHT_INDEXES; i++)
    {
      if (ciSet[i].nIndex == nIndex)
        return &ciSet[i].chd;
    }
  }
  return (tCharacterHighlightData *) 0;
}
static const tCharacterHighlightData* CharacterInfo_GetHighlightDataConst(const tCharacterInfo* ciSet, INT_X nIndex)
{
  if (nIndex >= 0)
  {
    int i;
    for (i = 0; i < HIGHLIGHT_INDEXES; i++)
    {
      if (ciSet[i].nIndex == nIndex)
        return &ciSet[i].chd;
    }
  }
  return (const tCharacterHighlightData *) 0;
}

// functions prototypes
static WCHAR char2wchar(const char ch);
static void  getEscapedPrefixPos(const INT_X nOffset, INT_X* pnPos, INT* pnLen);
static BOOL  isEscapedPrefixW(const wchar_t* strW, int len);
static void  remove_duplicate_indexes_and_sort(INT_X* indexes, const INT size /* = HIGHLIGHT_INDEXES */);
static BOOL  AutoBracketsFunc(MSGINFO* pmsgi, int nBracketType, BOOL bOverwriteMode);
static BOOL  GetHighlightIndexes(const int nHighlightIndex, const INT_X nCharacterPosition, const CHARRANGE_X* pSelection);
static void  GetPosFromChar(HWND hEd, const INT_X nCharacterPosition, POINTL* lpPos);
static BOOL  IsClearTypeEnabled();
static void  CopyMemory1(void* dst, const void* src, unsigned int size);

enum eHighlightFlags {
  HF_DOHIGHLIGHT = 0x01,
  HF_CLEARTYPE = 0x02,

  HF_UNINITIALIZED = 0xFFFF
};
static void  HighlightCharacter(const INT_X nCharacterPosition, const unsigned int uHighlightFlags);


/*
int main(void)
{
    // C++ initialization
    currentEdit.m_hWnd = NULL;
    return 0;
}
*/

static void* sys_memalloc(SIZE_T nBytes)
{
  return (void *) GlobalAlloc(GMEM_FIXED, nBytes);
}

static void sys_memfree(void* p)
{
  GlobalFree( (HGLOBAL) p );
}

static const wchar_t* getBracketsPairW(int nBracketType)
{
  if ( nBracketType < tbtUser )
    return strBracketsW[nBracketType - tbtBracket];

  return strUserBracketsW[nBracketType - tbtUser];
}

static const char* getBracketsPairA(int nBracketType)
{
  if ( nBracketType < tbtUser )
    return strBracketsA[nBracketType - tbtBracket];

  return strUserBracketsA[nBracketType - tbtUser];
}

static int getLeftBracketType(const wchar_t wch)
{
  int nLeftBracketType = tbtNone;

  switch (wch)
  {
    case L'(' :
      nLeftBracketType = tbtBracket;
      break;
    case L'[' :
      nLeftBracketType = tbtSquare;
      break;
    case L'{' :
      nLeftBracketType = tbtBrace;
      break;
    case L'\"' :
      nLeftBracketType = tbtDblQuote;
      break;
    case L'<' :
      if (bBracketsDoTag)
        nLeftBracketType = tbtTag;
      break;
    case L'\'' :
      if (bBracketsDoSingleQuote)
      {
        nLeftBracketType = tbtSglQuote;
        break;
      } // otherwise go to default (no break here)
    default :
      if (wch != 0)
      {
        int i = 0;
        while (i < MAX_USER_BRACKETS)
        {
          if (strUserBracketsW[i][0] == wch)
          {
            nLeftBracketType = tbtUser + i;
            i = MAX_USER_BRACKETS; // break condition
          }
          else if (strUserBracketsW[i][0] == 0)
          {
            i = MAX_USER_BRACKETS; // break condition
          }
          else
            ++i;
        }
      }
      break;
  }

  return nLeftBracketType;
}

static int getRightBracketType(const wchar_t wch)
{
  int nRightBracketType = tbtNone;

  switch (wch)
  {
    case L')' :
      nRightBracketType = tbtBracket;
      break;
    case L']' :
      nRightBracketType = tbtSquare;
      break;
    case L'}' :
      nRightBracketType = tbtBrace;
      break;
    case L'\"' :
      nRightBracketType = tbtDblQuote;
      break;
    case L'>' :
      if (bBracketsDoTag)
        nRightBracketType = tbtTag; 
      // no break here
    case L'/' :
      if (bBracketsDoTag2)
        nRightBracketType = tbtTag2;
      break;
    case L'\'' :
      if (bBracketsDoSingleQuote)
      {
        nRightBracketType = tbtSglQuote;
        break;
      } // otherwise go to default (no break here)
    default :
      if (wch != 0)
      {
        int i = 0;
        while (i < MAX_USER_BRACKETS)
        {
          if (strUserBracketsW[i][1] == wch)
          {
            nRightBracketType = tbtUser + i;
            i = MAX_USER_BRACKETS; // break condition
          }
          else if (strUserBracketsW[i][1] == 0)
          {
            i = MAX_USER_BRACKETS; // break condition
          }
          else
            ++i;
        }
      }
      break;
  }

  return nRightBracketType;
}

static BOOL isSepOrOneOfW(const wchar_t wch, const wchar_t* pwszChars)
{
  switch (wch)
  {
    case L'\x0D' :
    case L'\x0A' :
    case L'\x00' :
    case L' ' :
    case L'\t' :
      return TRUE;
    default:
      while (*pwszChars)
      {
        if (*pwszChars != wch)
          ++pwszChars;
        else
          return TRUE;
      }
  }
  return FALSE;
}

void OnEditCharPressed(MSGINFO* pmsgi)
{
  WCHAR    wch;
  int      nBracketType;
  EDITINFO ei;

  if (SendMessage(g_hMainWnd, AKD_GETEDITINFO, (WPARAM) NULL, (LPARAM) &ei) == 0)
  {
    nAutoRightBracketPos = -1;
    nAutoRightBracketType = tbtNone;
    return;
  }

  if (ei.bReadOnly)
  {
    // don't process read-only file
    nAutoRightBracketPos = -1;
    nAutoRightBracketType = tbtNone;
    return;
  }

  if (ei.bOvertypeMode)
  {
    if (g_dwOptions[OPT_DWORD_AUTOCOMPLETE_OVR_AUTOBR] == 0)
    {
      nAutoRightBracketPos = -1;
      nAutoRightBracketType = tbtNone;
      return;
    }
  }
  
  // verifying if a typed character is a bracket
  if (g_bOldWindows)
  {
    char ch = (char) pmsgi->wParam;
    wch = char2wchar(ch);
  }
  else
  {
    wch = (WCHAR) pmsgi->wParam;
  }
  
  if (pmsgi->hWnd != hActualEditWnd)
  {
    nAutoRightBracketPos = -1;
    nAutoRightBracketType = tbtNone;
  }

  if (nAutoRightBracketPos >= 0)
  {
    // the right bracket has been just added (automatically)
    // but you may duplicate it manually
    INT_X nEditPos;
    INT_X nEditEndPos;

    if (g_bOldWindows)
      AnyRichEdit_ExGetSelPos(pmsgi->hWnd, &nEditPos, &nEditEndPos);
    else
      AnyRichEdit_ExGetSelPosW(pmsgi->hWnd, &nEditPos, &nEditEndPos);

    if (nEditPos == nAutoRightBracketPos)
    {
      wchar_t next_wch;

      if (g_bOldWindows)
      {
        char next_ch = (char) AnyRichEdit_GetCharAt(pmsgi->hWnd, nEditPos);
        next_wch = char2wchar(next_ch);
      }
      else
      {
        next_wch = AnyRichEdit_GetCharAtW(pmsgi->hWnd, nEditPos);
      }

      if (getRightBracketType(next_wch) == nAutoRightBracketType)
      {
        if (getRightBracketType(wch) == nAutoRightBracketType)
        {
          // annul pressed character
          pmsgi->wParam = 0;
          ++nEditPos;
          if (nAutoRightBracketType == tbtTag2)
            ++nEditPos;
          if (g_bOldWindows)
            AnyRichEdit_ExSetSelPos(pmsgi->hWnd, nEditPos, nEditPos);
          else
            AnyRichEdit_ExSetSelPosW(pmsgi->hWnd, nEditPos, nEditPos);
          nAutoRightBracketPos = -1;
          nAutoRightBracketType = tbtNone;
          return;
        }

        // one character has been typed, 
        // so the auto-closed right bracket position increased
        ++nAutoRightBracketPos; 
      }
      else
      {
        // not auto-closed right bracket type
        nAutoRightBracketPos = -1;
        nAutoRightBracketType = tbtNone;
      }
    }
    else
    {
      // not auto-closed right bracket position
      nAutoRightBracketPos = -1;
      nAutoRightBracketType = tbtNone;
    }
  }

  nBracketType = getLeftBracketType(wch);
  if (nBracketType != tbtNone)
  {
    // a typed character is a bracket
    if (!AutoBracketsFunc(pmsgi, nBracketType, ei.bOvertypeMode))
    {
      if (nBracketType == tbtTag)
      {
        if (bBracketsDoTag2)
          nBracketType = tbtTag2;
      }
      if (nBracketType == nAutoRightBracketType)
      {
        nAutoRightBracketPos = -1;
        nAutoRightBracketType = tbtNone;
      }
    }
  }
}

void OnEditGetActiveBrackets(MSGINFO* pmsgi, const BOOL bAndHighlight)
{
  INT         i;
  BOOL        bHighlighted;
  INT_X       nEditPos;
  INT_X       nEditEndPos;
  CHARRANGE_X crSelection;

  if (bBracketsInternalRepaint)
    return;
  
  if (!pmsgi->hWnd)
    return;

  /*
  TCHAR str[256];
  wsprintf(str, "hWnd = %08X, prev = %08X", pmsgi->hWnd, hCurrentEditWnd);
  MessageBox(NULL, str, "OnEditGetActiveBrackets", MB_OK);
  */

  if (hActualEditWnd != pmsgi->hWnd)
  {
    nAutoRightBracketPos = -1;
    nAutoRightBracketType = tbtNone;
  }
  hActualEditWnd = pmsgi->hWnd;
  if (hActualEditWnd != hCurrentEditWnd)
  {
    nCurrentFileType = getFileType(&nCurrentFileType2);
  }

  // getting current position and selection
  if (g_bOldWindows)
    AnyRichEdit_ExGetSelPos(hActualEditWnd, &nEditPos, &nEditEndPos);
  else
    AnyRichEdit_ExGetSelPosW(hActualEditWnd, &nEditPos, &nEditEndPos);

  if (pmsgi->uMsg == WM_LBUTTONUP)
  {
    if (nEditEndPos == nEditPos)
    {
      // it's just simple WM_LBUTTONUP right after WM_LBUTTONDOWN;
      // active brackets have been highlighted already
      return;
    }
    // nEditEndPos != nEditPos,
    // some text has been selected with a mouse
  }

  crSelection.cpMin = nEditPos;
  crSelection.cpMax = nEditEndPos;

  bHighlighted = FALSE;
  for (i = 0; i < HIGHLIGHT_INDEXES; i++)
  {
    IndexesToHighlight[i] = -1;
  }
  CurrentBracketsIndexes[0] = -1;
  CurrentBracketsIndexes[1] = -1;

  if (g_bAkelEdit)
  {
    CharacterInfo_ClearAll(hgltCharacterInfo);
  }

  if (GetHighlightIndexes(0, nEditEndPos-1, &crSelection)) // left character
    bHighlighted = TRUE;

  if (g_dwOptions[OPT_DWORD_HIGHLIGHT_HLT_BOTHBR] || !bHighlighted)
  {
    if (GetHighlightIndexes(3, nEditEndPos, &crSelection))   // right character
      bHighlighted = TRUE;
  }

  if ((!bHighlighted) && (nEditEndPos != nEditPos))
  {
    if (GetHighlightIndexes(0, nEditPos-1, &crSelection)) // left character
      bHighlighted = TRUE;

    if (g_dwOptions[OPT_DWORD_HIGHLIGHT_HLT_BOTHBR] || !bHighlighted)
    {
      if (GetHighlightIndexes(3, nEditPos, &crSelection))   // right character
        bHighlighted = TRUE;
    }
  }
  
  if (bHighlighted)
  {
    for (i = 0; i < HIGHLIGHT_INDEXES; i++)
    {
      prevIndexesToHighlight[i] = IndexesToHighlight[i];
    }
    remove_duplicate_indexes_and_sort(IndexesToHighlight, HIGHLIGHT_INDEXES);

    if (g_bAkelEdit)
    {
      CharacterInfo_Copy(prevCharacterInfo, hgltCharacterInfo);
    }
  }
  
  if (bAndHighlight)
  {
    // highlight
    OnEditHighlightActiveBrackets();
  }
  
}

void OnEditHighlightActiveBrackets(void)
{
  unsigned int uHighlightFlags;
  int   i, j;
  INT_X index;
  INT_X indexesToRemoveHighlight[HIGHLIGHT_INDEXES];

  if (bBracketsInternalRepaint)
    return;

  uHighlightFlags = HF_UNINITIALIZED;

  for (i = 0; i < HIGHLIGHT_INDEXES; i++)
  {
    index = IndexesHighlighted[i];
    for (j = 0; j < HIGHLIGHT_INDEXES; j++)
    {
      if (index == IndexesToHighlight[j])
        break;
    }
    if (j < HIGHLIGHT_INDEXES)
      indexesToRemoveHighlight[i] = -1;
    else
      indexesToRemoveHighlight[i] = index;
  }

  for (i = 0; i < HIGHLIGHT_INDEXES; i++)
  {
    index = indexesToRemoveHighlight[i];
    if (index >= 0)
    {
      IndexesHighlighted[i] = -1; // prevents from WM_PAINT infinite loop
      if (uHighlightFlags == HF_UNINITIALIZED)
      {
        uHighlightFlags = 0;
        if (IsClearTypeEnabled())
          uHighlightFlags |= HF_CLEARTYPE;
      }
      HighlightCharacter(index, uHighlightFlags);
    }
  }

  for (i = 0; i < HIGHLIGHT_INDEXES; i++)
  {
    index = IndexesToHighlight[i];
    IndexesHighlighted[i] = index; // prevents from WM_PAINT infinite loop
    if (index >= 0)
    {
      if (uHighlightFlags == HF_UNINITIALIZED)
      {
        uHighlightFlags = 0;
        if (IsClearTypeEnabled())
          uHighlightFlags |= HF_CLEARTYPE;
      }
      HighlightCharacter(index, uHighlightFlags | HF_DOHIGHLIGHT);
    }
  }

}

/*static*/ BOOL AutoBracketsFunc(MSGINFO* pmsgi, int nBracketType, BOOL bOverwriteMode)
{
  INT_X   nEditPos;
  INT_X   nEditEndPos;
  WCHAR   next_wch;
  BOOL    bPrevCharOK = TRUE;
  BOOL    bNextCharOK = FALSE;

  #ifdef _verify_ab_
    TCHAR str[128];
  #endif

  hActualEditWnd = pmsgi->hWnd;
  if (hActualEditWnd != hCurrentEditWnd)
  {
    nCurrentFileType = getFileType(&nCurrentFileType2);
  }

  if (nBracketType == tbtTag)
  {
    if (bBracketsDoTagIf && !(nCurrentFileType2 & tfmHtmlCompatible))
      return FALSE;
  }

  // getting current position and selection
  if (g_bOldWindows)
    AnyRichEdit_ExGetSelPos(hActualEditWnd, &nEditPos, &nEditEndPos);
  else
    AnyRichEdit_ExGetSelPosW(hActualEditWnd, &nEditPos, &nEditEndPos);

  // is something selected?
  if (nEditEndPos != nEditPos)
  {
    if (g_dwOptions[OPT_DWORD_AUTOCOMPLETE_SEL_AUTOBR] == 0)
    {
      // removing selection
      if (g_bOldWindows)
        AnyRichEdit_ReplaceSelText(hActualEditWnd, "", TRUE);
      else
        AnyRichEdit_ReplaceSelTextW(hActualEditWnd, L"", TRUE);
    }
    else
    {
      if (nBracketType == tbtTag)
      {
        if (bBracketsDoTag2)
          nBracketType = tbtTag2;
      }
        
      if (g_bOldWindows)
      {
        INT_X       nSelLen;
        const char* pBrPairA;
        char*       pTextA;
        int         nBrPairLen;

        nSelLen = nEditEndPos - nEditPos;
        pBrPairA = getBracketsPairA(nBracketType);
        nBrPairLen = lstrlenA(pBrPairA);
        pTextA = (char*)sys_memalloc(sizeof(char)*(nSelLen + nBrPairLen + 2));
        if (pTextA)
        {
          pTextA[0] = pBrPairA[0];
          nSelLen = AnyRichEdit_GetTextAt(hActualEditWnd, nEditPos, nSelLen, pTextA + 1);
          if (g_dwOptions[OPT_DWORD_AUTOCOMPLETE_SEL_AUTOBR] == 2)
          {
            if ((pTextA[1] == pBrPairA[0]) && 
                (lstrcmpA(pTextA + nSelLen, pBrPairA + 1) == 0))
            {
              // already in brackets/quotes; exclude them
              pTextA[nSelLen] = 0;
              AnyRichEdit_ReplaceSelText(hActualEditWnd, pTextA + 2, TRUE);
              nEditEndPos -= nBrPairLen;
            }
            else
            {
              // enclose in brackets/quotes
              lstrcpyA(pTextA + nSelLen + 1, pBrPairA + 1);
              AnyRichEdit_ReplaceSelText(hActualEditWnd, pTextA, TRUE);
              nEditEndPos += nBrPairLen;
            }
          }
          else // (g_dwOptions[OPT_DWORD_AUTOCOMPLETE_SEL_AUTOBR] == 1)
          {
            lstrcpyA(pTextA + nSelLen + 1, pBrPairA + 1);
            AnyRichEdit_ReplaceSelText(hActualEditWnd, pTextA, TRUE);
            ++nEditPos;
            ++nEditEndPos;
          }
          sys_memfree(pTextA);
          AnyRichEdit_ExSetSelPos(hActualEditWnd, nEditPos, nEditEndPos);
        }
      }
      else
      {
        INT_X          nSelLen;
        const wchar_t* pBrPairW;
        wchar_t*       pTextW;
        int            nBrPairLen;
          
        nSelLen = nEditEndPos - nEditPos;
        pBrPairW = getBracketsPairW(nBracketType);
        nBrPairLen = lstrlenW(pBrPairW);
        pTextW = (wchar_t*)sys_memalloc(sizeof(wchar_t)*(nSelLen + nBrPairLen + 2));
        if (pTextW)
        {
          pTextW[0] = pBrPairW[0];
          nSelLen = AnyRichEdit_GetTextAtW(hActualEditWnd, nEditPos, nSelLen, pTextW + 1);
          if (g_dwOptions[OPT_DWORD_AUTOCOMPLETE_SEL_AUTOBR] == 2)
          {
            if ((pTextW[1] == pBrPairW[0]) && 
                (lstrcmpW(pTextW + nSelLen, pBrPairW + 1) == 0))
            {
              // already in brackets/quotes; exclude them
              pTextW[nSelLen] = 0;
              AnyRichEdit_ReplaceSelTextW(hActualEditWnd, pTextW + 2, TRUE);
              nEditEndPos -= nBrPairLen;
            }
            else
            {
              // enclose in brackets/quotes
              lstrcpyW(pTextW + nSelLen + 1, pBrPairW + 1);
              AnyRichEdit_ReplaceSelTextW(hActualEditWnd, pTextW, TRUE);
              nEditEndPos += nBrPairLen;
            }
          }
          else // (g_dwOptions[OPT_DWORD_AUTOCOMPLETE_SEL_AUTOBR] == 1)
          {
            lstrcpyW(pTextW + nSelLen + 1, pBrPairW + 1);
            AnyRichEdit_ReplaceSelTextW(hActualEditWnd, pTextW, TRUE);
            ++nEditPos;
            ++nEditEndPos;
          }
          sys_memfree(pTextW);
          AnyRichEdit_ExSetSelPosW(hActualEditWnd, nEditPos, nEditEndPos);
        }
      }

      // annul pressed character
      pmsgi->wParam = 0;
      // no auto bracket position
      nAutoRightBracketPos = -1;
      nAutoRightBracketType = tbtNone;
      return TRUE;
    }
  }

  // getting next character
  if (g_bOldWindows)
  {
    char next_ch = (char) AnyRichEdit_GetCharAt(hActualEditWnd, nEditPos);
    next_wch = char2wchar(next_ch);
  }
  else
  {
    next_wch = AnyRichEdit_GetCharAtW(hActualEditWnd, nEditPos);
  }

  #ifdef _verify_ab_
    wsprintf(str, "next_ch = %04X", next_wch);
    MessageBox(NULL, str, "AutoBracketsFunc", MB_OK);
  #endif

  if ((g_dwOptions[OPT_DWORD_AUTOCOMPLETE_ALL_AUTOBR] & 0x01) ||
      isSepOrOneOfW(next_wch, strNextCharOkW))
  {
    int nBrType = getRightBracketType(next_wch);
    if (nBrType == tbtNone)
    {
      bNextCharOK = TRUE;
    }
    else
    {
      if (nBrType == tbtTag2)
      {
        nBrType = tbtTag;
      }
      if (bBracketsRightExistsOK || (nBrType != nBracketType))
      {
        bNextCharOK = TRUE;
      }
    }
    if (bNextCharOK && !bBracketsRightExistsOK)
    {
      if ((next_wch == L'\'') &&
          (getBracketsPairW(nBracketType)[1] == L'\''))
      {
        bNextCharOK = FALSE;
      }
    }
  }

  #ifdef _verify_ab_
    wsprintf(str, "bNextCharOK = %lu", bNextCharOK);
    MessageBox(NULL, str, "AutoBracketsFunc", MB_OK);
  #endif

  if ( bNextCharOK && 
       ((nBracketType == tbtDblQuote) || 
        (nBracketType == tbtSglQuote) ||
        ((nBracketType >= tbtUser) && 
         (getBracketsPairW(nBracketType)[0] == getBracketsPairW(nBracketType)[1]))
       ) 
     )
  {
    wchar_t prev_wch;

    bPrevCharOK = FALSE;
    // getting previous character
    if (g_bOldWindows)
    {
      char prev_ch = (char) AnyRichEdit_GetCharAt(hActualEditWnd, nEditPos-1);
      prev_wch = char2wchar(prev_ch);
    }
    else
    {
      prev_wch = AnyRichEdit_GetCharAtW(hActualEditWnd, nEditPos-1);
    }

    #ifdef _verify_ab_
      wsprintf(str, "prev_ch = %04X", prev_wch);
      MessageBox(NULL, str, "AutoBracketsFunc", MB_OK);
    #endif

    if ((g_dwOptions[OPT_DWORD_AUTOCOMPLETE_ALL_AUTOBR] & 0x02) ||
        isSepOrOneOfW(prev_wch, strPrevCharOkW))
    {
      int nBrType = getLeftBracketType(prev_wch);
      if (nBrType == tbtNone)
      {
        bPrevCharOK = TRUE;
      }
      else
      {
        if (nBrType != nBracketType)
        {
          //if (getBracketsPairW(nBrType)[0] != getBracketsPairW(nBrType)[1])
          {
            bPrevCharOK = TRUE;
          }
        }
      }
    }
  }

  if (bPrevCharOK && bNextCharOK && 
      bBracketsSkipEscaped && (nCurrentFileType2 & tfmEscaped1))
  {
    wchar_t szPrefixW[MAX_ESCAPED_PREFIX + 2];
    INT_X   pos;
    int     len;

    getEscapedPrefixPos(nEditPos, &pos, &len);

    if (g_bOldWindows)
    {
      char szPrefixA[MAX_ESCAPED_PREFIX + 2];
    
      len = (int) AnyRichEdit_GetTextAt(hActualEditWnd, pos, len, szPrefixA);
      MultiByteToWideChar(CP_ACP, 0, szPrefixA, -1, szPrefixW, MAX_ESCAPED_PREFIX + 1);
    }
    else
    {
      len = (int) AnyRichEdit_GetTextAtW(hActualEditWnd, pos, len, szPrefixW);
    }

    if (isEscapedPrefixW(szPrefixW, len))
      bPrevCharOK = FALSE;
  }

  if (bPrevCharOK && bNextCharOK)
  {
    // annul pressed character
    pmsgi->wParam = 0;

    if (nBracketType == tbtTag)
    {
      if (bBracketsDoTag2)
        nBracketType = tbtTag2;
    }

    // AkelEdit: the caret position can be out of line length
    if (g_bAkelEdit)
    {
      AECHARINDEX ci;

      ci.lpLine = NULL;
      SendMessage( hActualEditWnd, AEM_GETINDEX, AEGI_CARETCHAR, (LPARAM)&ci );
      if (ci.lpLine)
      {
        INT_X nLineIndex;
        
        if (g_bOldWindows)
          nLineIndex = AnyRichEdit_LineIndex(hActualEditWnd, ci.nLine);
        else
          nLineIndex = AnyRichEdit_LineIndexW(hActualEditWnd, ci.nLine);
        
        nEditPos = nLineIndex + ci.nCharInLine;
      }
    }

    if (bOverwriteMode)
    {
      if (g_dwOptions[OPT_DWORD_AUTOCOMPLETE_OVR_AUTOBR] == 2)
      {
        // overwrite mode: right bracket overwrites existing symbol
        int     nLen;
        int     i;
        wchar_t wsz[8];

        if (g_bOldWindows)
        {
          nLen = lstrlenA(getBracketsPairA(nBracketType));
          wsz[0] = 0;
          AnyRichEdit_GetTextAt(hActualEditWnd, nEditPos, nLen, (char*)wsz);
          for (i = 0; i < nLen; i++)
          {
            switch ( ((char*)wsz)[i] )
            {
              case '\r':
              case '\n':
              case 0:
                nLen = i;
                break;
            }
          }
          AnyRichEdit_ExSetSelPos(hActualEditWnd, nEditPos, nEditPos + nLen);
        }
        else
        {
          nLen = lstrlenW(getBracketsPairW(nBracketType));
          wsz[0] = 0;
          AnyRichEdit_GetTextAtW(hActualEditWnd, nEditPos, nLen, wsz);
          for (i = 0; i < nLen; i++)
          {
            switch ( wsz[i] )
            {
              case L'\r':
              case L'\n':
              case 0:
                nLen = i;
                break;
            }
          }
          AnyRichEdit_ExSetSelPosW(hActualEditWnd, nEditPos, nEditPos + nLen);
        }
      }
    }

    // inserting brackets
    if (g_bOldWindows)
      AnyRichEdit_ReplaceSelText(hActualEditWnd, getBracketsPairA(nBracketType), TRUE);
    else
      AnyRichEdit_ReplaceSelTextW(hActualEditWnd, getBracketsPairW(nBracketType), TRUE);

    // placing cursor between brackets
    nEditPos++;
    if (g_bOldWindows)
      AnyRichEdit_ExSetSelPos(hActualEditWnd, nEditPos, nEditPos);
    else
      AnyRichEdit_ExSetSelPosW(hActualEditWnd, nEditPos, nEditPos);

    nAutoRightBracketPos = nEditPos;
    nAutoRightBracketType = nBracketType;
    return TRUE;
  }

  return FALSE;
}

static void updateCurrentBracketsIndexes(INT_X ind1, INT_X ind2)
{
  if (ind1 < ind2)
  {
    if ( (CurrentBracketsIndexes[0] < 0) ||
         (ind1 == CurrentBracketsIndexes[0] - 1) ||
         (ind2 == CurrentBracketsIndexes[1] + 1)
       )
    {
      CurrentBracketsIndexes[0] = ind1; // left bracket
      CurrentBracketsIndexes[1] = ind2; // right bracket
    }
  }
  else
  {
    if ( (CurrentBracketsIndexes[0] < 0) ||
         (ind2 == CurrentBracketsIndexes[0] - 1) ||
         (ind1 == CurrentBracketsIndexes[1] + 1)
       )
    {
      CurrentBracketsIndexes[0] = ind2; // left bracket
      CurrentBracketsIndexes[1] = ind1; // right bracket
    }
  }
}

#define DP_NONE          0x00
#define DP_FORWARD       0x01
#define DP_BACKWARD      0x02
#define DP_DETECT        0x10
#define DP_MAYBEBACKWARD (DP_DETECT | DP_BACKWARD)

static BOOL isSentenceEndChar(const wchar_t wch)
{
  switch (wch)
  {
    case L'.':
    case L'?':
    case L'!':
    case 0x2026: // ellipsis (...)
      return TRUE;
  }
  return FALSE;
}

static int getDuplicatedPairDirection(const INT_X nCharacterPosition, const wchar_t curr_wch)
{
  static HWND hLocalEditWnd = NULL;
  static WCHAR szLocalDelimitersW[128] = { 0 };
  WCHAR prev_wch;
  WCHAR next_wch;

  if (nCharacterPosition <= 0)
    return DP_FORWARD; // no char before, search forward

  if (g_bOldWindows)
  {
    char ch;

    ch = AnyRichEdit_GetCharAt(hActualEditWnd, nCharacterPosition - 1);
    prev_wch = char2wchar(ch);
    ch = AnyRichEdit_GetCharAt(hActualEditWnd, nCharacterPosition + 1);
    next_wch = char2wchar(ch);
  }
  else
  {
    prev_wch = AnyRichEdit_GetCharAtW(hActualEditWnd, nCharacterPosition - 1);
    next_wch = AnyRichEdit_GetCharAtW(hActualEditWnd, nCharacterPosition + 1);
  }

  if (prev_wch == 0 || 
      prev_wch == '\r' || 
      prev_wch == '\n' || 
      next_wch == curr_wch)
    return DP_FORWARD; // previous char is EOL or pair found, search forward

  if (next_wch == 0 || 
      next_wch == '\r' || 
      next_wch == '\n' || 
      prev_wch == curr_wch)
    return DP_BACKWARD; // next char is EOL or pair found, search backward

  if ( hLocalEditWnd != hActualEditWnd )
  {
    szLocalDelimitersW[0] = 0;
    SendMessage( hActualEditWnd, AEM_GETWORDDELIMITERS, 127, (LPARAM) szLocalDelimitersW );
    hLocalEditWnd = hActualEditWnd;
  }

  if (isSepOrOneOfW(prev_wch, szLocalDelimitersW))
  {
    // previous char is a separator
    if (!isSepOrOneOfW(next_wch, szLocalDelimitersW))
      return DP_FORWARD; // next char is not a separator, search forward

    if (isSentenceEndChar(prev_wch))
    {
      if (!isSentenceEndChar(next_wch)) // e.g.  ."-
        return DP_MAYBEBACKWARD;
    }
  }
  else if (isSepOrOneOfW(next_wch, szLocalDelimitersW))
  {
    // next char is a separator
    if (!isSepOrOneOfW(prev_wch, szLocalDelimitersW))
      return DP_BACKWARD; // previous char is not a separator, search backward

    if (isSentenceEndChar(prev_wch))
    {
      if (!isSentenceEndChar(next_wch)) // e.g.  ."-
        return DP_MAYBEBACKWARD;
    }
  }

  return DP_DETECT;
}

typedef struct sGetHighlightIndexesCookie {
    INT_X pos1;
    INT_X pos2;
    int   nResult;
    int   nBracketType;  // -1 or one of TBracketType
    tCharacterHighlightData chd;
} tGetHighlightIndexesCookie;

enum eGetHighLightResult {
    ghlrNone = 0,
    ghlrSingleChar,
    ghlrPair,
    ghlrDoNotHighlight
};

static DWORD CALLBACK GetAkelEditHighlightCallback(UINT_PTR dwCookie, AECHARRANGE *crAkelRange, CHARRANGE64 *crRichRange, AEHLPAINT *hlp)
{
  tGetHighlightIndexesCookie* cookie;
  
  cookie = (tGetHighlightIndexesCookie *) dwCookie;
  cookie->nResult = ghlrNone;

  /*
  if (hlp->dwActiveText != hlp->dwDefaultText)
  {
    // text color is different than default
  }

  if (hlp->dwActiveBk != hlp->dwDefaultBk)
  {
    // background color is different than default
  }
  */

  /* (hlp->dwPaintType);
     AEHPT_QUOTE
     AEHPT_FOLD   */

  if (hlp->qm.lpQuote)
  {
    // quote item
    AECHARINDEX aeCh;
    int n;
    wchar_t wch;

    if (cookie->nBracketType >= 0)
    {
      wch = getBracketsPairW(cookie->nBracketType)[0]; // left
      CopyMemory1(&aeCh, &hlp->qm.crQuoteStart.ciMin, sizeof(AECHARINDEX));
      if (aeCh.nLine == hlp->qm.crQuoteStart.ciMax.nLine)
        n = hlp->qm.crQuoteStart.ciMax.nCharInLine;
      else
        n = aeCh.lpLine->nLineLen;
      while ( aeCh.nCharInLine < n )
      {
        if ( aeCh.lpLine->wpLine[aeCh.nCharInLine] == wch )
        {
          cookie->pos1 = (INT_X) SendMessage(hActualEditWnd, AEM_INDEXTORICHOFFSET, 
                                   0, (LPARAM) &aeCh);
          break;
        }
        ++aeCh.nCharInLine;
      }
    }
    /*else
    {
      cookie->pos1 = (INT_X) SendMessage(hActualEditWnd, AEM_INDEXTORICHOFFSET, 
                               0, (LPARAM) &hlp->qm.crQuoteStart.ciMin);
    }*/

    if (cookie->pos1 >= 0)
    {
      if (AEC_IndexCompare(&hlp->qm.crQuoteEnd.ciMin, &hlp->qm.crQuoteEnd.ciMax) == 0)
      {
        // no ending quote symbol
        cookie->pos2 = -1;
      }
      else
      {
        if (cookie->nBracketType >= 0)
        {
          wch = getBracketsPairW(cookie->nBracketType)[1]; // right
          CopyMemory1(&aeCh, &hlp->qm.crQuoteEnd.ciMax, sizeof(AECHARINDEX));
          if (aeCh.nLine == hlp->qm.crQuoteEnd.ciMin.nLine)
            n = hlp->qm.crQuoteEnd.ciMin.nCharInLine;
          else
            n = 0;
          while ( --aeCh.nCharInLine >= n )
          {
            if ( aeCh.lpLine->wpLine[aeCh.nCharInLine] == wch )
            {
              cookie->pos2 = (INT_X) SendMessage(hActualEditWnd, AEM_INDEXTORICHOFFSET, 
                                       0, (LPARAM) &aeCh);
              break;
            }
          }
        }
        /*else
        {
          cookie->pos2 = (INT_X) SendMessage(hActualEditWnd, AEM_INDEXTORICHOFFSET, 
                                   0, (LPARAM) &hlp->qm.crQuoteEnd.ciMax);
          --cookie->pos2; // because the range is [ciMin; ciMax)
        }*/
      }
    }
  }

  if ((cookie->pos1 >= 0) && (cookie->pos2 >= 0))
    cookie->nResult = ghlrPair;
  else
    cookie->nResult = ghlrSingleChar;

  cookie->chd.dwFontStyle = hlp->dwFontStyle;
  if (hlp->dwActiveText != (DWORD)(-1))
    cookie->chd.dwActiveTextColor = hlp->dwActiveText;
  if (hlp->dwActiveBk != (DWORD)(-1))
    cookie->chd.dwActiveBkColor = hlp->dwActiveBk;

  if (hlp->mrm.lpMarkRange)
  {
    // mark range found
    if (hlp->mrm.lpMarkRange->crText != (DWORD)(-1))
      cookie->chd.dwActiveTextColor = hlp->mrm.lpMarkRange->crText;
    if (hlp->mrm.lpMarkRange->crBk != (DWORD)(-1))
      cookie->chd.dwActiveBkColor = hlp->mrm.lpMarkRange->crBk;
    if ((g_dwOptions[OPT_DWORD_HIGHLIGHT_HLT_STYLE] & XBR_HSF_REDRAWCODER) == 0)
    {
      cookie->nResult = ghlrDoNotHighlight;
    }
  }

  return 0;
}

static void GetHighlightDataFromAkelEdit(const INT_X nCharacterPosition, tGetHighlightIndexesCookie* pCookieHL)
{
  AEGETHIGHLIGHT aehl;
    
  pCookieHL->pos1 = -1;
  pCookieHL->pos2 = -1;
  pCookieHL->nResult = ghlrNone;
  CharacterHighlightData_Clear(&pCookieHL->chd);

  aehl.dwCookie = (UINT_PTR) pCookieHL;
  aehl.dwError = 0;
  aehl.dwFlags = 0/*AEGHF_NOSELECTION | AEGHF_NOACTIVELINETEXT | AEGHF_NOACTIVELINEBK*/;
  aehl.lpCallback = GetAkelEditHighlightCallback;
  SendMessage(hActualEditWnd, AEM_RICHOFFSETTOINDEX, 
    (WPARAM) nCharacterPosition, (LPARAM) &aehl.crText.ciMin);
  AEC_NextCharEx(&aehl.crText.ciMin, &aehl.crText.ciMax);
  // the range is [ciMin, ciMax)
  SendMessage(hActualEditWnd, AEM_HLGETHIGHLIGHT, 0, (LPARAM)&aehl);
}

static DWORD GetInSelectionState(const INT_X nCharacterPosition, const CHARRANGE_X* pSelection)
{
  if ((nCharacterPosition >= pSelection->cpMin) && 
      (nCharacterPosition < pSelection->cpMax))
  {
    return XBR_STATEFLAG_INSELECTION;
  }
  else
    return 0;
}

static int GetAkelEditHighlightInfo(const int nHighlightIndex, const INT_X nCharacterPosition, 
             const int nBracketType, const BOOL bRightBracket, const CHARRANGE_X* pSelection)
{
  tGetHighlightIndexesCookie cookie;
  tGetHighlightIndexesCookie cookieHL;
  AEFINDFOLD ff;
  BOOL bTag2;
    
  bTag2 = FALSE;
  cookie.pos1 = -1;
  cookie.pos2 = -1;
  cookie.nResult = ghlrNone;
  cookie.nBracketType = nBracketType;
  CharacterHighlightData_Clear(&cookie.chd);
  cookie.chd.dwState |= GetInSelectionState(nCharacterPosition, pSelection);

  // first we check for a fold...
  if (bRightBracket)
  {
    ff.dwFlags = AEFF_FINDOFFSET | AEFF_FOLDSTART;
    ff.dwFindIt = nCharacterPosition - 1;
  }
  else
  {
    ff.dwFlags = AEFF_FINDOFFSET | AEFF_FOLDEND;
    ff.dwFindIt = nCharacterPosition + 1;
  }
  ff.lpParent = NULL;
  ff.lpPrevSubling = NULL;

  SendMessage(hActualEditWnd, AEM_FINDFOLD, (WPARAM) &ff, 0);
  if (ff.lpParent)
  {
    // fold found
    if ((ff.lpParent->lpMaxPoint->nPointLen > 0) &&
        (ff.lpParent->lpMinPoint->nPointLen > 0))
    {
      const wchar_t* p;
      int n;
      int i;
      wchar_t wch;

      if (bRightBracket)
      {
        //  ... }
        cookie.pos2 = (INT_X) ff.lpParent->lpMaxPoint->nPointOffset;
        if (cookie.pos2 >= 0)
          cookie.pos2 += (ff.lpParent->lpMaxPoint->nPointLen - 1);

        if (cookie.pos2 == nCharacterPosition)
        {
          const AEPOINT* aePt = ff.lpParent->lpMinPoint;
          p = aePt->ciPoint.lpLine->wpLine + aePt->ciPoint.nCharInLine;
          n = aePt->ciPoint.lpLine->nLineLen - aePt->ciPoint.nCharInLine;
          if (n > aePt->nPointLen)
            n = aePt->nPointLen;
          i = 0;
          wch = getBracketsPairW(nBracketType)[0]; // left bracket

          while (i < n)
          {
            if (p[i] == wch)
            {
              cookie.pos1 = (INT_X) aePt->nPointOffset;
              cookie.pos1 += i;
              break;
            }
            ++i;
          }

          if (cookie.pos1 >= 0)
          {
            cookie.nResult = ghlrPair;
            cookie.chd.dwFontStyle = ff.lpParent->dwFontStyle;
            if (ff.lpParent->crText != (DWORD)(-1))
              cookie.chd.dwActiveTextColor = ff.lpParent->crText;
            if (ff.lpParent->crBk != (DWORD)(-1))
              cookie.chd.dwActiveBkColor = ff.lpParent->crBk;
          }
          else
            cookie.pos2 = -1;
        }
        else
          cookie.pos2 = -1;
      }
      else
      {
        //  { ...
        cookie.pos1 = (INT_X) ff.lpParent->lpMinPoint->nPointOffset;
        if (cookie.pos1 == nCharacterPosition)
        {
          const AEPOINT* aePt = ff.lpParent->lpMaxPoint;
          p = aePt->ciPoint.lpLine->wpLine + aePt->ciPoint.nCharInLine;
          n = aePt->ciPoint.lpLine->nLineLen - aePt->ciPoint.nCharInLine;
          if (n > aePt->nPointLen)
            n = aePt->nPointLen;
          wch = getBracketsPairW(nBracketType)[1]; // right bracket

          while (--n >= 0)
          {
            if (p[n] == wch)
            {
              cookie.pos2 = (INT_X) aePt->nPointOffset;
              cookie.pos2 += n;
              break;
            }
          }

          if (cookie.pos2 >= 0)
          {
            cookie.nResult = ghlrPair;
            cookie.chd.dwFontStyle = ff.lpParent->dwFontStyle;
            if (ff.lpParent->crText != (DWORD)(-1))
              cookie.chd.dwActiveTextColor = ff.lpParent->crText;
            if (ff.lpParent->crBk != (DWORD)(-1))
              cookie.chd.dwActiveBkColor = ff.lpParent->crBk;
          }
          else
            cookie.pos1 = -1;
        }
        else
          cookie.pos1 = -1;
      }
    }

    if (cookie.pos2 >= 0)
    {
      if (nBracketType == tbtTag)
      {
        // < ... >
        WCHAR wch;

        if (bRightBracket)
        {
          if (ff.lpParent->lpMaxPoint->nPointLen > 1)
          {
            // maybe < ... [<]...>
            const AECHARINDEX* aeCh = &ff.lpParent->lpMaxPoint->ciPoint;
            wch = aeCh->lpLine->wpLine[aeCh->nCharInLine];
            if (wch == L'<')
            {
              // yes, < ... [<]...>
              cookie.pos1 = (INT_X) ff.lpParent->lpMaxPoint->nPointOffset;
            }
          }
        }
        else
        {
          const AECHARINDEX* aeCh = &ff.lpParent->lpMaxPoint->ciPoint;
          wch = aeCh->lpLine->wpLine[aeCh->nCharInLine];
          if (wch == L'<')
          {
            // < ... <
            if (ff.lpParent->lpMinPoint->nPointLen > 1)
            {
              // maybe <...[>] ... <
              cookie.pos2 = (INT_X) ff.lpParent->lpMinPoint->nPointOffset;
              cookie.pos2 += (ff.lpParent->lpMinPoint->nPointLen /*- 1*/);

              if (g_bOldWindows)
              {
                const char ch = AnyRichEdit_GetCharAt(hActualEditWnd, cookie.pos2);
                wch = char2wchar(ch);
              }
              else
                wch = AnyRichEdit_GetCharAtW(hActualEditWnd, cookie.pos2);

              if (wch != L'>')
              {
                // not <...[>] ... <
                cookie.pos2 = -1;
                cookie.nResult = ghlrSingleChar;
              }
            }
            else
            {  
              cookie.pos2 = -1;
              cookie.nResult = ghlrSingleChar;
            }
          }
          /*else
            cookie.pos2 += (ff.lpParent->lpMaxPoint->nPointLen - 1);*/
        }

        if (cookie.pos2 >= 0)
        {
          if (g_bOldWindows)
          {
            const char ch = AnyRichEdit_GetCharAt(hActualEditWnd, cookie.pos2 - 1);
            wch = char2wchar(ch);
          }
          else
            wch = AnyRichEdit_GetCharAtW(hActualEditWnd, cookie.pos2 - 1);

          if (wch == L'/')
          {
            // < ... />
            bTag2 = TRUE;
          }
        }
      }
      /*else
      {
        if (!bRightBracket)
          cookie.pos2 += (ff.lpParent->lpMaxPoint->nPointLen - 1);
      }*/
    }
  }

  // ...then we check for a highlight info...
  cookieHL.nBracketType = nBracketType;
  GetHighlightDataFromAkelEdit(nCharacterPosition, &cookieHL);

  if (cookie.nResult == ghlrPair)
  {
    // fold was found initially
    if (cookieHL.nResult == ghlrDoNotHighlight)
    {
      // do not highlight
      cookie.nResult = ghlrDoNotHighlight;
      cookie.chd.dwState |= XBR_STATEFLAG_DONOTHL;
    }
    else
    {
      cookie.chd.dwFontStyle = cookieHL.chd.dwFontStyle;
      if (cookieHL.chd.dwActiveTextColor != XBR_FONTCOLOR_UNDEFINED)
        cookie.chd.dwActiveTextColor = cookieHL.chd.dwActiveTextColor;
      if (cookieHL.chd.dwActiveBkColor != XBR_FONTCOLOR_UNDEFINED)
        cookie.chd.dwActiveBkColor = cookieHL.chd.dwActiveBkColor;
    }
  }
  else
  {
    if ((nCharacterPosition == cookieHL.pos1) ||
        (nCharacterPosition == cookieHL.pos2))
    {
      cookie.pos1 = cookieHL.pos1;
      cookie.pos2 = cookieHL.pos2;
      cookie.nResult = cookieHL.nResult;
    }
    else
    {
      cookie.nResult = ghlrSingleChar;
    }
    cookie.chd.dwFontStyle = cookieHL.chd.dwFontStyle;
    if (cookieHL.chd.dwActiveTextColor != XBR_FONTCOLOR_UNDEFINED)
      cookie.chd.dwActiveTextColor = cookieHL.chd.dwActiveTextColor;
    if (cookieHL.chd.dwActiveBkColor != XBR_FONTCOLOR_UNDEFINED)
      cookie.chd.dwActiveBkColor = cookieHL.chd.dwActiveBkColor;
  }

  // ...finally...
  switch (cookie.nResult)
  {
    case ghlrNone:
      break;
    case ghlrSingleChar:
    {
      CharacterInfo_Add(hgltCharacterInfo, nCharacterPosition, &cookie.chd);
      break;
    }
    case ghlrPair:
    case ghlrDoNotHighlight:
    {
      if (bTag2)
      {
        BOOL bReadyHL = FALSE;

        // tbtTag2
        IndexesToHighlight[nHighlightIndex] = cookie.pos1;
        IndexesToHighlight[nHighlightIndex+1] = cookie.pos2 - 1;
        IndexesToHighlight[nHighlightIndex+2] = cookie.pos2;

        if ((cookie.chd.dwState & XBR_STATEFLAG_DONOTHL) ||
            ((cookie.chd.dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(cookie.pos1, pSelection)))
        {
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos1, &cookie.chd);
        }
        else
        {
          bReadyHL = TRUE;
          cookieHL.nBracketType = -1;
          GetHighlightDataFromAkelEdit(cookie.pos1, &cookieHL);
          cookieHL.chd.dwState |= GetInSelectionState(cookie.pos1, pSelection);
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos1, &cookieHL.chd);
        }

        if ((cookie.chd.dwState & XBR_STATEFLAG_DONOTHL) ||
            ((cookie.chd.dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(cookie.pos2 - 1, pSelection)))
        {
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos2 - 1, &cookie.chd);
        }
        else
        {
          if (!bReadyHL)
          {
            bReadyHL = TRUE;
            cookieHL.nBracketType = -1;
            GetHighlightDataFromAkelEdit(cookie.pos2 - 1, &cookieHL);
            cookieHL.chd.dwState |= GetInSelectionState(cookie.pos2 - 1, pSelection);
          }
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos2 - 1, &cookieHL.chd);
        }
        
        if ((cookie.chd.dwState & XBR_STATEFLAG_DONOTHL) ||
            ((cookie.chd.dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(cookie.pos2, pSelection)))
        {
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos2, &cookie.chd);
        }
        else
        {
          if (!bReadyHL)
          {
            bReadyHL = TRUE;
            cookieHL.nBracketType = -1;
            GetHighlightDataFromAkelEdit(cookie.pos2, &cookieHL);
            cookieHL.chd.dwState |= GetInSelectionState(cookie.pos2, pSelection);
          }
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos2, &cookieHL.chd);
        }
      }
      else
      {
        BOOL bReadyHL = FALSE;

        IndexesToHighlight[nHighlightIndex] = cookie.pos1;
        IndexesToHighlight[nHighlightIndex+1] = cookie.pos2;

        if ((cookie.chd.dwState & XBR_STATEFLAG_DONOTHL) ||
            ((cookie.chd.dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(cookie.pos1, pSelection)))
        {
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos1, &cookie.chd);
        }
        else
        {
          bReadyHL = TRUE;
          cookieHL.nBracketType = -1;
          GetHighlightDataFromAkelEdit(cookie.pos1, &cookieHL);
          cookieHL.chd.dwState |= GetInSelectionState(cookie.pos1, pSelection);
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos1, &cookieHL.chd);
        }

        if ((cookie.chd.dwState & XBR_STATEFLAG_DONOTHL) ||
            ((cookie.chd.dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(cookie.pos2, pSelection)))
        {
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos2, &cookie.chd);
        }
        else
        {
          if (!bReadyHL)
          {
            bReadyHL = TRUE;
            cookieHL.nBracketType = -1;
            GetHighlightDataFromAkelEdit(cookie.pos2, &cookieHL);
            cookieHL.chd.dwState |= GetInSelectionState(cookie.pos2, pSelection);
          }
          CharacterInfo_Add(hgltCharacterInfo, cookie.pos2, &cookieHL.chd);
        }
      }

      updateCurrentBracketsIndexes(cookie.pos1, cookie.pos2);
      break;
    }
  }

  return cookie.nResult;
}

/*static*/ BOOL GetHighlightIndexes(const int nHighlightIndex, const INT_X nCharacterPosition, const CHARRANGE_X* pSelection)
{
  int   nBracketType;
  int   nDuplicatedPairDirection; // left bracket is the same as right
  int   nGetHighlightResult;
  BOOL  bRightBracket;
  WCHAR wch;

  if ( nCharacterPosition < 0 )
    return FALSE;

  /*
  // this does not allow to work with new created file
  if ( nCurrentFileType == tftNone )
    return FALSE;
  */
  
  if ( (prevIndexesToHighlight[0] >= 0) || (prevIndexesToHighlight[3] >= 0) )
  {
    int i;

    for (i = 0; i < HIGHLIGHT_INDEXES; i++)
    {
      if (nCharacterPosition == prevIndexesToHighlight[i])
      {
        INT_X ind1, ind2, ind3;
        int k = (i < 3) ? 0 : 3;
        IndexesToHighlight[nHighlightIndex + 0] = prevIndexesToHighlight[k + 0];
        IndexesToHighlight[nHighlightIndex + 1] = prevIndexesToHighlight[k + 1];
        IndexesToHighlight[nHighlightIndex + 2] = prevIndexesToHighlight[k + 2];
        ind1 = IndexesToHighlight[nHighlightIndex + 0];
        if (IndexesToHighlight[nHighlightIndex + 2] < 0)
        {
          ind2 = IndexesToHighlight[nHighlightIndex + 1];
          ind3 = IndexesToHighlight[nHighlightIndex + 2];
        }
        else
        {
          ind2 = IndexesToHighlight[nHighlightIndex + 2];
          ind3 = IndexesToHighlight[nHighlightIndex + 1];
        }

        updateCurrentBracketsIndexes(ind1, ind2);

        if (g_bAkelEdit)
        {
          tCharacterHighlightData* pchd;
          tGetHighlightIndexesCookie cookieHL;

          pchd = CharacterInfo_GetHighlightData(prevCharacterInfo, ind1);
          if (pchd)
          {
            if ((pchd->dwState & XBR_STATEFLAG_DONOTHL) ||
                ((pchd->dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(ind1, pSelection)))
            {
              CharacterInfo_Add(hgltCharacterInfo, ind1, pchd);
            }
            else
            {
              cookieHL.nBracketType = -1;
              GetHighlightDataFromAkelEdit(ind1, &cookieHL);
              cookieHL.chd.dwState |= GetInSelectionState(ind1, pSelection);
              CharacterInfo_Add(hgltCharacterInfo, ind1, &cookieHL.chd);
              CharacterHighlightData_Copy(pchd, &cookieHL.chd);
            }
          }
          pchd = CharacterInfo_GetHighlightData(prevCharacterInfo, ind2);
          if (pchd)
          {
            if ((pchd->dwState & XBR_STATEFLAG_DONOTHL) ||
                ((pchd->dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(ind2, pSelection)))
            {
              CharacterInfo_Add(hgltCharacterInfo, ind2, pchd);
            }
            else
            {
              cookieHL.nBracketType = -1;
              GetHighlightDataFromAkelEdit(ind2, &cookieHL);
              cookieHL.chd.dwState |= GetInSelectionState(ind2, pSelection);
              CharacterInfo_Add(hgltCharacterInfo, ind2, &cookieHL.chd);
              CharacterHighlightData_Copy(pchd, &cookieHL.chd);
            }
          }
          pchd = CharacterInfo_GetHighlightData(prevCharacterInfo, ind3);
          if (pchd)
          {
            if ((pchd->dwState & XBR_STATEFLAG_DONOTHL) ||
                ((pchd->dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(ind3, pSelection)))
            {
              CharacterInfo_Add(hgltCharacterInfo, ind3, pchd);
            }
            else
            {
              cookieHL.nBracketType = -1;
              GetHighlightDataFromAkelEdit(ind3, &cookieHL);
              cookieHL.chd.dwState |= GetInSelectionState(ind3, pSelection);
              CharacterInfo_Add(hgltCharacterInfo, ind3, &cookieHL.chd);
              CharacterHighlightData_Copy(pchd, &cookieHL.chd);
            }
          }
        }
        
        return TRUE;
      }
    }
  }    
  
  if (g_bOldWindows)
  {
    char ch = AnyRichEdit_GetCharAt(hActualEditWnd, nCharacterPosition);
    wch = char2wchar(ch);
  }
  else
  {
    wch = AnyRichEdit_GetCharAtW(hActualEditWnd, nCharacterPosition);
  }

  nDuplicatedPairDirection = DP_NONE;
  nGetHighlightResult = ghlrNone;
  bRightBracket = FALSE;
  nBracketType = getLeftBracketType(wch);
  if (nBracketType == tbtNone)
  {
    bRightBracket = TRUE;
    nBracketType = getRightBracketType(wch);
    if (nBracketType != tbtNone)
    {
      if (nBracketType == tbtTag2)
      {
        if (wch == L'>')
          nBracketType = tbtTag;
        else
          nBracketType = tbtNone;
      }
    }
  }

  if (nBracketType != tbtNone)
  {
    if (g_bAkelEdit)
    {
      // check AkelEdit's highlight info
      nGetHighlightResult = GetAkelEditHighlightInfo(nHighlightIndex, nCharacterPosition, nBracketType, bRightBracket, pSelection);
      switch (nGetHighlightResult)
      {
        case ghlrDoNotHighlight:
          return TRUE; // processed by AkelEdit, do not highlight by XBrackets
        case ghlrPair:
          return TRUE; // pair found: nothing to process additionally
      }
    }

    // check for duplicated pair (e.g. "" quotes)
    if (getBracketsPairW(nBracketType)[0] == getBracketsPairW(nBracketType)[1])
    {
      if (g_dwOptions[OPT_DWORD_HIGHLIGHT_QUOTE_MAX_LINES] == 0)
        return FALSE;

      nDuplicatedPairDirection = getDuplicatedPairDirection(nCharacterPosition, wch);
      switch (nDuplicatedPairDirection)
      {
        case DP_FORWARD:
          bRightBracket = FALSE; // search from left to right (forward)
          break;
        case DP_BACKWARD:
        case DP_DETECT:
        case DP_MAYBEBACKWARD:
          bRightBracket = TRUE; // search from right to left (backward)
          break;
        /*default:
          nBracketType = tbtNone;
          break;*/
      }

      if (g_dwOptions[OPT_DWORD_HIGHLIGHT_QUOTE_DETECT_LINES] == 0)
      {
        if (nDuplicatedPairDirection == DP_DETECT)
          return FALSE;
      }
    }
  }

  if (nBracketType == tbtTag)
  {
    if (bBracketsDoTagIf && !(nCurrentFileType2 & tfmHtmlCompatible))
      return FALSE;
  }

  if ((nBracketType != tbtNone) && 
      bBracketsSkipEscaped && (nCurrentFileType2 & tfmEscaped1) && 
      (nCharacterPosition > 0))
  {
    wchar_t szPrefixW[MAX_ESCAPED_PREFIX + 2];
    INT_X pos;
    int len;

    getEscapedPrefixPos(nCharacterPosition, &pos, &len);

    if (g_bOldWindows)
    {
      char szPrefixA[MAX_ESCAPED_PREFIX + 2];
      
      len = (int) AnyRichEdit_GetTextAt(hActualEditWnd, pos, len, szPrefixA);
      MultiByteToWideChar(CP_ACP, 0, szPrefixA, -1, szPrefixW, MAX_ESCAPED_PREFIX + 1);
    }
    else
    {
      len = (int) AnyRichEdit_GetTextAtW(hActualEditWnd, pos, len, szPrefixW);
    }

    if (isEscapedPrefixW(szPrefixW, len))
      return FALSE;
  }

  if (nBracketType != tbtNone)
  {
    // we don't want these arrays in the stack ;)
    static TCHAR szLine[0x10000];
    static WCHAR wszLine[0x10000];

    INT_X i;
    INT   iStep, nLen;
    INT   nLine, nMinLine, nMaxLine, nLineStep, nStartLine;
    BOOL  bFound;
    BOOL  bComment;
    WCHAR wchOK, wchFail;
    INT   nFailReferences;
    
    const WCHAR* pcwszLine;
    AECHARINDEX ci;
        
    if (g_bAkelEdit)
      ci.lpLine = NULL;
    else
      pcwszLine = wszLine;

    // the character to search for (another bracket)
    wchOK = getBracketsPairW(nBracketType)[bRightBracket ? 0 : 1];
    if (nDuplicatedPairDirection == DP_NONE)
      wchFail = getBracketsPairW(nBracketType)[bRightBracket ? 1 : 0];
    else
      wchFail = 0;
    nFailReferences = 0;

    if (g_bOldWindows)
    {
      nLine = AnyRichEdit_ExLineFromChar(hActualEditWnd, nCharacterPosition);
      nMaxLine = (bBracketsHighlightVisibleArea ?
        AnyRichEdit_LastVisibleLine(hActualEditWnd) : (AnyRichEdit_GetLineCount(hActualEditWnd) - 1));
      nMinLine = (bBracketsHighlightVisibleArea ?
        AnyRichEdit_FirstVisibleLine(hActualEditWnd) : 0);
    }
    else
    {
      nLine = AnyRichEdit_ExLineFromCharW(hActualEditWnd, nCharacterPosition);
      nMaxLine = (bBracketsHighlightVisibleArea ?
        AnyRichEdit_LastVisibleLineW(hActualEditWnd) : (AnyRichEdit_GetLineCountW(hActualEditWnd) - 1));
      nMinLine = (bBracketsHighlightVisibleArea ?
        AnyRichEdit_FirstVisibleLineW(hActualEditWnd) : 0);
    }

    if (nDuplicatedPairDirection != DP_NONE)
    {
      const int maxLines = (nDuplicatedPairDirection == DP_DETECT) ? 
        g_dwOptions[OPT_DWORD_HIGHLIGHT_QUOTE_DETECT_LINES] : 
          g_dwOptions[OPT_DWORD_HIGHLIGHT_QUOTE_MAX_LINES];

      if (nMaxLine > nLine + maxLines - 1)
        nMaxLine = nLine + maxLines - 1;

      if (nMinLine < nLine - maxLines + 1)
        nMinLine = nLine - maxLines + 1;
    }
    else
    {
      const int maxBrLines = g_dwOptions[OPT_DWORD_HIGHLIGHT_BR_MAX_LINES];
      if (maxBrLines > 0)
      {
        if (nMaxLine > nLine + maxBrLines - 1)
          nMaxLine = nLine + maxBrLines - 1;

        if (nMinLine < nLine - maxBrLines + 1)
          nMinLine = nLine - maxBrLines + 1;
      }
    }

    nStartLine = nLine;
    bFound = FALSE;
    bComment = FALSE;
    nLineStep = bRightBracket ? (-1) : 1; // forward or backward
    iStep = bRightBracket ? (-1) : 1;     // forward or backward

    while ( bRightBracket ? (nLine >= nMinLine) : (nLine <= nMaxLine) )
    {
      if (g_bAkelEdit)
      {
        if (ci.lpLine)
        {
          ci.lpLine = bRightBracket ? ci.lpLine->prev : ci.lpLine->next;
        }
        else
        {
          SendMessage( hActualEditWnd, AEM_GETINDEX, AEGI_CARETCHAR, (LPARAM)&ci );
          if (ci.nLine != nLine)
          {
            if (ci.lpLine && (ci.nLine + 1 == nLine))
            {
              if ((ci.nCharInLine == ci.lpLine->nLineLen) && 
                  (ci.lpLine->nLineBreak == AELB_WRAP))
              {
                // a hack: don't process a position at the line-wrap
                break;
              }
            }

            SendMessage( hActualEditWnd, AEM_GETINDEX, AEGI_FIRSTSELCHAR, (LPARAM)&ci );
            if (ci.nLine != nLine)
              SendMessage( hActualEditWnd, AEM_GETINDEX, AEGI_LASTSELCHAR, (LPARAM)&ci );
          }
        }

        if (ci.lpLine)
        {
          pcwszLine = ci.lpLine->wpLine;
          nLen = ci.lpLine->nLineLen;
        }
        else
          break;
      }
      else
      {
        if (g_bOldWindows)
        {
          nLen = AnyRichEdit_GetLine(hActualEditWnd, nLine, szLine, 0x10000-1);
          // CAnyRichEdit::GetLine sets szLine[nLen] = 0
          MultiByteToWideChar(CP_ACP, 0, szLine, nLen, wszLine, nLen);
          wszLine[nLen] = 0;
        }
        else
        {
          nLen = AnyRichEdit_GetLineW(hActualEditWnd, nLine, wszLine, 0x10000-1);
          // AnyRichEdit_GetLineW sets szLineW[nLen] = 0
        }
      }

      if ( bBracketsSkipComment1 && (nCurrentFileType2 & tfmComment1) )
      {
        int i;

        // pre-processing current line
        for (i = 0; i < nLen; i++)
        {
          wch = pcwszLine[i];
          if (wch == L'/')
          {
            if ((i+1 < nLen) && (pcwszLine[i+1] == L'/'))
            {
              // this is "//" comment
              if ((i == 0) || (pcwszLine[i-1] != L':')) // not "://"
              {
                //wszLine[i] = 0;
                nLen = i;
                bComment = TRUE;
                break;
              }
            }
          }
        }

      }

      i = bRightBracket ? (nLen - 1) : 0;
      if (nStartLine == nLine)
      {
        INT_X nLinePosition;

        if (g_bOldWindows)
          nLinePosition = AnyRichEdit_LineIndex(hActualEditWnd, nLine);
        else
          nLinePosition = AnyRichEdit_LineIndexW(hActualEditWnd, nLine);

        nLinePosition = nCharacterPosition - nLinePosition;

        if (bComment) // start line contains a comment...
        {
          if (nLinePosition >= nLen) // ...and current bracket is commented
          {
            //MessageBoxA(NULL,"break","for ( ... nLine ... )",MB_OK);
            break; // break outside the cycle for nLine
          }
        }

        i = bRightBracket ? (nLinePosition - 1) : (nLinePosition + 1);
      }

      for ( ; bRightBracket ? (i >= 0) : (i < nLen); i += iStep)
      {
        wch = pcwszLine[i];
        if (wch == wchOK)
        {
          if (bBracketsSkipEscaped && (nCurrentFileType2 & tfmEscaped1) && (i > 0))
          {
            INT_X pos;
            INT   len;

            getEscapedPrefixPos(i, &pos, &len);
            if (!isEscapedPrefixW(pcwszLine + pos, len))
              --nFailReferences;
          }
          else
            --nFailReferences;
          
          if (nFailReferences < 0)
          {
            bFound = TRUE;
            break;
          }
        }
        else if (wch == wchFail)
        {
          if (bBracketsSkipEscaped && (nCurrentFileType2 & tfmEscaped1) && (i > 0))
          {
            INT_X pos;
            INT   len;

            getEscapedPrefixPos(i, &pos, &len);
            if (!isEscapedPrefixW(pcwszLine + pos, len))
              ++nFailReferences;
          }
          else
            ++nFailReferences;
        }
      }

      if (bFound)
      {
        INT_X pos1;

        if (nDuplicatedPairDirection == DP_NONE)
        {
          // not duplicated: the pair bracket is found definitely
          break;
        }

        if (g_bOldWindows)
          pos1 = i + AnyRichEdit_LineIndex(hActualEditWnd, nLine);
        else
          pos1 = i + AnyRichEdit_LineIndexW(hActualEditWnd, nLine);

        if (bRightBracket)
        {
          const int nDirection2 = getDuplicatedPairDirection(pos1, wchOK);
          if (nDirection2 != DP_BACKWARD)
          {
            if (nDuplicatedPairDirection == DP_BACKWARD)
            {
              // when search direction is known, let's assume
              // the result is OK within g_dwOptions[OPT_DWORD_HIGHLIGHT_QUOTE_DETECT_LINES]
              if (nStartLine - nLine < (int) g_dwOptions[OPT_DWORD_HIGHLIGHT_QUOTE_DETECT_LINES])
                break;
            }
          }
          if (nDirection2 != DP_FORWARD)
          {
            if (nDuplicatedPairDirection == DP_DETECT)
            {
              // searching backward failed, let's try forward...
              nLine = nMinLine;
            }
            else
            {
              // backward assumption failed, stop searching
              bFound = FALSE;
              break;
            }
          }
          else
            break;
        }
        else
        {
          const int nDirection2 = getDuplicatedPairDirection(pos1, wchOK);
          if (nDirection2 != DP_FORWARD)
          {
            if (nDuplicatedPairDirection == DP_FORWARD)
            {
              // when search direction is known, let's assume
              // the result is OK within g_dwOptions[OPT_DWORD_HIGHLIGHT_QUOTE_DETECT_LINES]
              if (nLine - nStartLine < (int) g_dwOptions[OPT_DWORD_HIGHLIGHT_QUOTE_DETECT_LINES])
                break;
            }
          }
          if (nDirection2 != DP_BACKWARD)
          {
            if ((nDirection2 != DP_MAYBEBACKWARD) || 
                (nDuplicatedPairDirection == DP_DETECT))
            {
              // fail anyway
              bFound = FALSE;
            }
          }
          break;
        }
      }

      if (nDuplicatedPairDirection == DP_DETECT)
      {
        if (bRightBracket)
        {
          if (nLine == nMinLine)
          {
            bRightBracket = FALSE;
            bFound = FALSE;
            bComment = FALSE;
            nLine = nStartLine - 1;
            nLineStep = 1;
            iStep = 1;
            nFailReferences = 0;
            
            if (g_bAkelEdit)
              ci.lpLine = NULL;
            else
              pcwszLine = wszLine;
          }
        }
      }

      nLine += nLineStep;
    }

    if (bFound && (wch == wchOK))
    {
      if (nBracketType == tbtTag)
      {
        WCHAR prev_wch = 0;
        if (bRightBracket)
        {
          if (nCharacterPosition > 0)
          {
            if (g_bOldWindows)
            {
              char ch = AnyRichEdit_GetCharAt(hActualEditWnd, nCharacterPosition-1);
              prev_wch = char2wchar(ch);
            }
            else
            {
              prev_wch = AnyRichEdit_GetCharAtW(hActualEditWnd, nCharacterPosition-1);
            }
          }
        }
        else
        {
          if (i > 0)
          {
            prev_wch = pcwszLine[i-1];
          }
        }
        if (prev_wch == L'/')
        {
          nBracketType = tbtTag2;
        }
      }

      if (g_bOldWindows)
      {
        i += AnyRichEdit_LineIndex(hActualEditWnd, nLine);
      }
      else
      {
        i += AnyRichEdit_LineIndexW(hActualEditWnd, nLine);
      }

      if (nBracketType != tbtTag2)
      {
        if (nGetHighlightResult == ghlrSingleChar)
        {
          const tCharacterHighlightData* pchd;

          pchd = CharacterInfo_GetHighlightDataConst(hgltCharacterInfo, nCharacterPosition);
          if (pchd)
          {
            tGetHighlightIndexesCookie cookieHL;

            if ((pchd->dwState & XBR_STATEFLAG_DONOTHL) ||
                ((pchd->dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(i, pSelection)))
            {
              CharacterInfo_Add(hgltCharacterInfo, i, pchd);
            }
            else
            {
              cookieHL.nBracketType = -1;
              GetHighlightDataFromAkelEdit(i, &cookieHL);
              cookieHL.chd.dwState |= GetInSelectionState(i, pSelection);
              CharacterInfo_Add(hgltCharacterInfo, i, &cookieHL.chd);
            }
          }
        }
        IndexesToHighlight[nHighlightIndex] = i;
        IndexesToHighlight[nHighlightIndex+1] = nCharacterPosition;
      }
      else
      {
        INT_X i1;
        
        i1 = bRightBracket ? (nCharacterPosition - 1) : (i - 1);

        if (nGetHighlightResult == ghlrSingleChar)
        {
          const tCharacterHighlightData* pchd;

          pchd = CharacterInfo_GetHighlightDataConst(hgltCharacterInfo, nCharacterPosition);
          if (pchd)
          {
            tGetHighlightIndexesCookie cookieHL;
            BOOL bReadyHL;

            bReadyHL = FALSE;
            if ((pchd->dwState & XBR_STATEFLAG_DONOTHL) ||
                ((pchd->dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(i, pSelection)))
            {
              CharacterInfo_Add(hgltCharacterInfo, i, pchd);
            }
            else
            {
              bReadyHL = TRUE;
              cookieHL.nBracketType = -1;
              GetHighlightDataFromAkelEdit(i, &cookieHL);
              cookieHL.chd.dwState |= GetInSelectionState(i, pSelection);
              CharacterInfo_Add(hgltCharacterInfo, i, &cookieHL.chd);
            }

            if ((pchd->dwState & XBR_STATEFLAG_DONOTHL) ||
                ((pchd->dwState & XBR_STATEFLAG_INSELECTION) == GetInSelectionState(i1, pSelection)))
            {
              CharacterInfo_Add(hgltCharacterInfo, i1, pchd);
            }
            else
            {
              if (!bReadyHL)
              {
                bReadyHL = TRUE;
                cookieHL.nBracketType = -1;
                GetHighlightDataFromAkelEdit(i1, &cookieHL);
                cookieHL.chd.dwState |= GetInSelectionState(i1, pSelection);
              }
              CharacterInfo_Add(hgltCharacterInfo, i1, &cookieHL.chd);
            }
          }
        }
        IndexesToHighlight[nHighlightIndex] = i;
        IndexesToHighlight[nHighlightIndex+1] = i1;
        IndexesToHighlight[nHighlightIndex+2] = nCharacterPosition;
      }

      updateCurrentBracketsIndexes(i, nCharacterPosition);

      return TRUE;
    }
    else
    {
      //MessageBoxA(NULL,"no indexes","GetHighlightIndexes",MB_OK);
    }

    //return TRUE;
  }

  return FALSE;
}

/*static*/ void GetPosFromChar(HWND hEd, const INT_X nCharacterPosition, POINTL* lpPos)
{
  if (g_bOldRichEdit)
  {
    // RichEdit 2.0
    LRESULT coord;

    if (g_bOldWindows)
      coord = SendMessageA(hEd, EM_POSFROMCHAR, (WPARAM) nCharacterPosition, 0);
    else
      coord = SendMessageW(hEd, EM_POSFROMCHAR, (WPARAM) nCharacterPosition, 0);
    lpPos->x = LOWORD(coord);
    lpPos->y = HIWORD(coord);
  }
  else
  {
    if (g_bOldWindows)
      SendMessageA(hEd, EM_POSFROMCHAR, (WPARAM) lpPos, (LPARAM) nCharacterPosition);
    else
      SendMessageW(hEd, EM_POSFROMCHAR, (WPARAM) lpPos, (LPARAM) nCharacterPosition);
  }
}

/*static*/ void CopyMemory1(void* dst, const void* src, unsigned int size)
{
  unsigned char* pdst;
  const unsigned char* psrc;
    
  pdst = (unsigned char *) dst;
  psrc = (const unsigned char *) src;
    
  while ( size-- )
  {
    *(pdst++) = *(psrc++);
  }
}

static void adjustFont(const DWORD dwFontStyle, const DWORD dwFontFlags, 
                       BYTE* plfItalic, LONG* plfWeight)
{
  switch (dwFontStyle)
  {
    case AEHLS_FONTNORMAL:
      if (!(dwFontFlags & AEHLO_IGNOREFONTNORMAL))
      {
        *plfItalic = FALSE;
        *plfWeight = FW_NORMAL;
      }
      break;
    case AEHLS_FONTBOLD:
      //if (!(dwFontFlags & AEHLO_IGNOREFONTNORMAL))
      //{
        *plfItalic = FALSE;
      //}
      if (!(dwFontFlags & AEHLO_IGNOREFONTBOLD))
      {
        *plfWeight = FW_BOLD;
      }
      break;
    case AEHLS_FONTITALIC:
      if (!(dwFontFlags & AEHLO_IGNOREFONTITALIC))
      {
        *plfItalic = TRUE;
      }
      //if (!(dwFontFlags & AEHLO_IGNOREFONTNORMAL))
      //{
        *plfWeight = FW_NORMAL;
      //}
      break;
    case AEHLS_FONTBOLDITALIC:
      if (!(dwFontFlags & AEHLO_IGNOREFONTITALIC))
      {
        *plfItalic = TRUE;
      }
      if (!(dwFontFlags & AEHLO_IGNOREFONTBOLD))
      {
        *plfWeight = FW_BOLD;
      }
      break;
  }
}

BOOL IsClearTypeEnabled()
{
  BOOL bClearType;
  BOOL bFontSmoothing;
  UINT nFontSmoothingType;

  bClearType = FALSE;
  bFontSmoothing = FALSE;
  nFontSmoothingType = 0;

  if (g_bOldWindows)
  {
    SystemParametersInfoA(SPI_GETFONTSMOOTHING, 0, &bFontSmoothing, 0);
    if (bFontSmoothing)
    {
      SystemParametersInfoA(SPI_GETFONTSMOOTHINGTYPE, 0, &nFontSmoothingType, 0);
    }
  }
  else
  {
    SystemParametersInfoW(SPI_GETFONTSMOOTHING, 0, &bFontSmoothing, 0);
    if (bFontSmoothing)
    {
      SystemParametersInfoW(SPI_GETFONTSMOOTHINGTYPE, 0, &nFontSmoothingType, 0);
    }
  }

  if (bFontSmoothing && (nFontSmoothingType == FE_FONTSMOOTHINGCLEARTYPE))
  {
    bClearType = TRUE;
  }

  return bClearType;
}

#define mix_color(a, b) (BYTE)((((unsigned int)(a))*3)/8 + (((unsigned int)(b))*5)/8)

/*static*/ void HighlightCharacter(const INT_X nCharacterPosition, const unsigned int uHighlightFlags)
{
  POINTL ptBegin, ptEnd;
  tCharacterHighlightData chd;
  const tCharacterHighlightData* pchd;

  if (!hActualEditWnd)
    return;

  pchd = CharacterInfo_GetHighlightDataConst(hgltCharacterInfo, nCharacterPosition);
  if (pchd && (pchd->dwState & XBR_STATEFLAG_DONOTHL))
    return;

  if (pchd)
  {
    CharacterHighlightData_Copy(&chd, pchd);
  }
  else
  {
    CharacterHighlightData_Clear(&chd);
    pchd = &chd;
  }
  // contents of *pchd and chd are equal now

  if (g_bOldWindows)
  {
    if ((nCharacterPosition < AnyRichEdit_FirstVisibleCharIndex(hActualEditWnd)) ||
        (nCharacterPosition > AnyRichEdit_LastVisibleCharIndex(hActualEditWnd)))
    {
      return;
    }
  }
  else
  {
    if ((nCharacterPosition < AnyRichEdit_FirstVisibleCharIndexW(hActualEditWnd)) ||
        (nCharacterPosition > AnyRichEdit_LastVisibleCharIndexW(hActualEditWnd)))
    {
      return;
    }
  }

  if (g_bAkelEdit)
  {
    AEFOLD* lpFold;
    INT     nLine;

    nLine = AnyRichEdit_ExLineFromChar(hActualEditWnd, nCharacterPosition);
    lpFold = (AEFOLD *) SendMessage(hActualEditWnd, AEM_ISLINECOLLAPSED, nLine, 0);
    if (lpFold)
    {
      // the line with this character is collapsed
      return;
    }
  }

  GetPosFromChar(hActualEditWnd, nCharacterPosition, &ptBegin);
  // returns left-top point of the character's rectangle

  if ((ptBegin.x >= 0) && (ptBegin.y >= 0))
  {
    HDC          hDC;
    RECT         EdRect;
    AECHARINDEX  aeci;
    AECHARCOLORS aecc;
    BOOL         bInSelection;
    DWORD        dwFontFlags;

    AnyRichEdit_GetEditRect(hActualEditWnd, &EdRect);

    if (ptBegin.x < EdRect.left)
    {
      // left side of the character can be invisible
      // ptBegin.x = EdRect.left;
      return;
    }

    if (ptBegin.y < EdRect.top)
    {
      // top side of the character can be invisible
      // ptBegin.y = EdRect.top;
      return;
    }

    GetPosFromChar(hActualEditWnd, nCharacterPosition+1, &ptEnd);
    // left-top point of the next character's rectangle

    if ((ptEnd.y != ptBegin.y) || (ptEnd.x < 0))
    {
      ptEnd.x = EdRect.right;
    }

    //cc.nCharPos = nCharacterPosition;
    //bInSelection = (BOOL) SendMessage(g_hMainWnd, AKD_GETCHARCOLOR, (WPARAM) hActualEditWnd, (LPARAM) &cc);
    SendMessage(hActualEditWnd, AEM_RICHOFFSETTOINDEX, (WPARAM) nCharacterPosition, (LPARAM) &aeci);
    aecc.dwFlags = 0;
    aecc.dwFontStyle = 0;
    bInSelection = (BOOL) SendMessage(hActualEditWnd, AEM_GETCHARCOLORS, (WPARAM) &aeci, (LPARAM) &aecc);
    dwFontFlags = (DWORD) SendMessage(hActualEditWnd, AEM_HLGETOPTIONS, 0, 0);

    hDC = GetDC(hActualEditWnd);
    if (hDC)
    {
      HRGN         hRgn, hRgnOld;
      HFONT        hFont, hFontOld;
      RECT         rect;
      COLORREF     textColor, bkColor;
      LOGFONTA     lfA;
      LOGFONTW     lfW;
      //INT         sel1, sel2;
      int          nBkModePrev;
      char         strA[2];
      wchar_t      strW[2];

      bBracketsInternalRepaint = TRUE;

      // to repaint a field under the caret
      HideCaret(hActualEditWnd);

      // at first we select a font...
      if (g_bOldWindows)
      {
        HFONT hf;

        hf = (HFONT) SendMessageA(g_hMainWnd, AKD_GETFONTA, (WPARAM) hActualEditWnd, (LPARAM) &lfA);
        if (!hf)
          hf = (HFONT) SendMessageA(g_hMainWnd, AKD_GETFONTA, (WPARAM) NULL, (LPARAM) &lfA);

        // adjust font style
        adjustFont(chd.dwFontStyle, dwFontFlags, &lfA.lfItalic, &lfA.lfWeight);

        hFont = CreateFontIndirectA(&lfA);
      }
      else
      {
        HFONT hf;

        hf = (HFONT) SendMessageW(g_hMainWnd, AKD_GETFONTW, (WPARAM) hActualEditWnd, (LPARAM) &lfW);
        if (!hf)
          hf = (HFONT) SendMessageW(g_hMainWnd, AKD_GETFONTW, (WPARAM) NULL, (LPARAM) &lfW);

        // adjust font style
        adjustFont(chd.dwFontStyle, dwFontFlags, &lfW.lfItalic, &lfW.lfWeight);

        hFont = CreateFontIndirectW(&lfW);
      }
      hFontOld = (HFONT) SelectObject(hDC, hFont);

      // ...then we call GetTextMetrics for this font...
      if (g_bOldWindows)
      {
        TEXTMETRICA tmA;

        GetTextMetricsA(hDC, &tmA);
        ptEnd.y = ptBegin.y + tmA.tmHeight;
      }
      else
      {
        TEXTMETRICW tmW;

        GetTextMetricsW(hDC, &tmW);
        ptEnd.y = ptBegin.y + tmW.tmHeight;
      }

      ptEnd.y *= 5; // multiply height by (5/4)
      ptEnd.y /= 4; // for "tall" characters
      
      // ...finally we can select a region
      ptEnd.x += 3*(ptEnd.x - ptBegin.x); // for some cursive fonts
      if (ptEnd.x > EdRect.right)
      {
        ptEnd.x = EdRect.right;
      }
      if (ptEnd.y > EdRect.bottom)
      {
        ptEnd.y = EdRect.bottom;
      }

      hRgn = CreateRectRgn(ptBegin.x, ptBegin.y, ptEnd.x, ptEnd.y);
      hRgnOld = (HRGN) SelectObject(hDC, hRgn);

      nBkModePrev = SetBkMode(hDC, TRANSPARENT);
      //bkColor = GetBkColor(hDC);

      //if (g_bOldWindows)
      //  currentEdit.ExGetSelPos(&sel1, &sel2);
      //else
      //  currentEdit.ExGetSelPosW(&sel1, &sel2);

      if (!bInSelection)
      {
        if (chd.dwActiveTextColor != XBR_FONTCOLOR_UNDEFINED)
          aecc.crText = chd.dwActiveTextColor;
        if (chd.dwActiveBkColor != XBR_FONTCOLOR_UNDEFINED)
          aecc.crBk = chd.dwActiveBkColor;
      }
      bkColor = aecc.crBk;
      textColor = aecc.crText;
      if (uHighlightFlags & HF_DOHIGHLIGHT)
      {
        //textColor = RGB(16,112,192);
        if (uBracketsHighlight & BRHLF_TEXT)
          textColor = bracketsColourHighlight[0];
        if (uBracketsHighlight & BRHLF_BKGND)
          bkColor = bracketsColourHighlight[1];

        if (bInSelection)
        {
          if (uBracketsHighlight & BRHLF_BKGND)
          {
            BYTE r, g, b;
            
            r = mix_color( GetRValue(bkColor), GetRValue(aecc.crBk) );
            g = mix_color( GetGValue(bkColor), GetGValue(aecc.crBk) );
            b = mix_color( GetBValue(bkColor), GetBValue(aecc.crBk) );

            bkColor = RGB(r, g, b);
          }
        }
      }

      SetTextColor(hDC, aecc.crBk);
      rect.left   = ptBegin.x;
      rect.top    = ptBegin.y;
      rect.right  = ptEnd.x;
      rect.bottom = ptEnd.y;
      if (g_bOldWindows)
      {
        if (uHighlightFlags & HF_DOHIGHLIGHT)
        {
          COLORREF crOldBk = 0;

          if (uHighlightFlags & HF_CLEARTYPE)
          {
            SetBkMode(hDC, nBkModePrev);
            crOldBk = SetBkColor(hDC, aecc.crBk);
          }

          strA[0] = AnyRichEdit_GetCharAt(hActualEditWnd, nCharacterPosition);
          strA[1] = 0;
          DrawTextA(hDC, strA, 1, &rect, 0);

          if (hFontOld)  SelectObject(hDC, hFontOld);
          DeleteObject(hFont);

          if (uHighlightFlags & HF_CLEARTYPE)
          {
            SetBkColor(hDC, crOldBk);
            SetBkMode(hDC, TRANSPARENT);
          }

          if (chd.dwFontStyle == XBR_FONTSTYLE_UNDEFINED)
          {
            /*if (!lfA.lfItalic)
            {
              // hack: draw italic-style bracket character with bkgnd color
              lfA.lfItalic = TRUE;
              hFont = CreateFontIndirectA(&lfA);
              hFontOld = (HFONT) SelectObject(hDC, hFont);
              DrawTextA(hDC, strA, 1, &rect, 0);
              if (hFontOld)  SelectObject(hDC, hFontOld);
              DeleteObject(hFont);
              lfA.lfItalic = FALSE;
            }*/

            lfA.lfWeight = (g_dwOptions[OPT_DWORD_HIGHLIGHT_HLT_STYLE] & XBR_HSF_BOLDFONT) ? FW_BOLD : FW_NORMAL;
          }
          else
          {
            if (g_dwOptions[OPT_DWORD_HIGHLIGHT_HLT_STYLE] & XBR_HSF_BOLDFONT)
            {
              lfA.lfWeight = FW_BOLD;
            }
          }

          hFont = CreateFontIndirectA(&lfA);
          hFontOld = (HFONT) SelectObject(hDC, hFont);
          if (uBracketsHighlight & BRHLF_BKGND)
          {
            SetBkMode(hDC, OPAQUE);
            SetBkColor(hDC, bkColor);
          }
          if (uBracketsHighlight & BRHLF_TEXT)
          {
            SetTextColor(hDC, textColor);
          }
          else
          {
            SetTextColor(hDC, aecc.crText);
          }
          
          DrawTextA(hDC, strA, 1, &rect, 0);
        }
        else
        {
          InvalidateRgn(hActualEditWnd, hRgn, TRUE);
        }
      }
      else
      {
        if (uHighlightFlags & HF_DOHIGHLIGHT)
        {
          COLORREF crOldBk = 0;

          if (uHighlightFlags & HF_CLEARTYPE)
          {
            SetBkMode(hDC, nBkModePrev);
            crOldBk = SetBkColor(hDC, aecc.crBk);
          }

          strW[0] = AnyRichEdit_GetCharAtW(hActualEditWnd, nCharacterPosition);
          strW[1] = 0;
          DrawTextW(hDC, strW, 1, &rect, 0);

          if (hFontOld)  SelectObject(hDC, hFontOld);
          DeleteObject(hFont);

          if (uHighlightFlags & HF_CLEARTYPE)
          {
            SetBkColor(hDC, crOldBk);
            SetBkMode(hDC, TRANSPARENT);
          }

          if (chd.dwFontStyle == XBR_FONTSTYLE_UNDEFINED)
          {
            /*if (!lfW.lfItalic)
            {
              // hack: draw italic-style bracket character with bkgnd color
              lfW.lfItalic = TRUE;
              hFont = CreateFontIndirectW(&lfW);
              hFontOld = (HFONT) SelectObject(hDC, hFont);
              DrawTextW(hDC, strW, 1, &rect, 0);
              if (hFontOld)  SelectObject(hDC, hFontOld);
              DeleteObject(hFont);
              lfW.lfItalic = FALSE;
            }*/

            lfW.lfWeight = (g_dwOptions[OPT_DWORD_HIGHLIGHT_HLT_STYLE] & XBR_HSF_BOLDFONT) ? FW_BOLD : FW_NORMAL;
          }
          else
          {
            if (g_dwOptions[OPT_DWORD_HIGHLIGHT_HLT_STYLE] & XBR_HSF_BOLDFONT)
            {
              lfW.lfWeight = FW_BOLD;
            }
          }

          hFont = CreateFontIndirectW(&lfW);
          hFontOld = (HFONT) SelectObject(hDC, hFont);
          if (uBracketsHighlight & BRHLF_BKGND)
          {
            SetBkMode(hDC, OPAQUE);
            SetBkColor(hDC, bkColor);
          }
          if (uBracketsHighlight & BRHLF_TEXT)
          {
            SetTextColor(hDC, textColor);
          }
          else
          {
            SetTextColor(hDC, aecc.crText);
          }
          
          DrawTextW(hDC, strW, 1, &rect, 0);
        }
        else
        {
          InvalidateRgn(hActualEditWnd, hRgn, TRUE);
        }
      }
      if (hFontOld)  SelectObject(hDC, hFontOld);
      DeleteObject(hFont);
      if (hRgnOld)  SelectObject(hDC, hRgnOld);
      DeleteObject(hRgn);

      ShowCaret(hActualEditWnd);

      ReleaseDC(hActualEditWnd, hDC);

      bBracketsInternalRepaint = FALSE;
    }
  }

}

void RemoveAllHighlightInfo(const BOOL bRepaint)
{
  int i;

  if (bBracketsInternalRepaint)
    return;
  
  for (i = 0; i < HIGHLIGHT_INDEXES; i++)
  {
    IndexesToHighlight[i] = -1; // set no items to highlight
    prevIndexesToHighlight[i] = -1;
  }
  CurrentBracketsIndexes[0] = -1;
  CurrentBracketsIndexes[1] = -1;

  if (g_bAkelEdit)
  {
    CharacterInfo_ClearAll(hgltCharacterInfo);
    CharacterInfo_ClearAll(prevCharacterInfo);
  }

  if (bRepaint)
  {
    for (i = 0; i < HIGHLIGHT_INDEXES; i++)
    {
      if (IndexesHighlighted[i] >= 0)  // is any item highlighted?
        break;
    }
    if (i < HIGHLIGHT_INDEXES)
    {
      // removing all highlighting
      OnEditHighlightActiveBrackets();
    } // else there is no highlighted items
  }
  else
  {
    for (i = 0; i < HIGHLIGHT_INDEXES; i++)
    {
      IndexesHighlighted[i] = -1;
    }
  }
}

//---------------------------------------------------------------------------

/*static*/ WCHAR char2wchar(const char ch)
{
  char  str[2] = {0, 0};
  WCHAR wstr[2] = {0, 0};

  str[0] = ch;
  MultiByteToWideChar(CP_ACP, 0, str, 1, wstr, 1);
  return wstr[0];
}

/*static*/ void getEscapedPrefixPos(const INT_X nOffset, INT_X* pnPos, INT* pnLen)
{
  if (nOffset > MAX_ESCAPED_PREFIX)
  {
    *pnPos = nOffset - MAX_ESCAPED_PREFIX;
    *pnLen = MAX_ESCAPED_PREFIX;
  }
  else
  {
    *pnPos = 0;
    *pnLen = (INT) nOffset;
  }
}

/*static*/ BOOL isEscapedPrefixW(const wchar_t* strW, int len)
{
  int k = 0;
  while ((len > 0) && (strW[--len] == L'\\'))
  {
    ++k;
  }
  return (k % 2) ? TRUE : FALSE;
}

/*static*/ void remove_duplicate_indexes_and_sort(INT_X* indexes, const INT size)
{
  INT   i, j;
  INT_X ind;

  // removing duplicates
  for (i = 0; i < size; i++)
  {
    ind = indexes[i];
    if (ind >= 0)
    {
      for (j = 0; j < size; j++)
      {
        if ((i != j) && (ind == indexes[j]))
        {
          indexes[j] = -1;
        }
      }
    }
  }

  // sorting
  for (i = 0; i < size; i++)
  {
    ind = indexes[i];
    for (j = i+1; j < size; j++)
    {
      if (indexes[j] > ind)
      {
        indexes[i] = indexes[j];
        indexes[j] = ind;  // old value
        ind = indexes[i];  // new value
      }
    }
  }
}

static const char* getFileExtA(const char* cszFileNameA)
{
  if (cszFileNameA)
  {
    int pos = lstrlenA(cszFileNameA) - 1;
    if (pos > 0)
    {
      while ((pos >= 0) && (cszFileNameA[pos] != '.'))
      {
        --pos;
      }
      return (cszFileNameA + 1 + pos);
    }
  }
  return NULL;
}

static const wchar_t* getFileExtW(const wchar_t* cszFileNameW)
{
  if (cszFileNameW)
  {
    int pos = lstrlenW(cszFileNameW) - 1;
    if (pos > 0)
    {
      while ((pos >= 0) && (cszFileNameW[pos] != L'.'))
      {
        --pos;
      }
      return (cszFileNameW + 1 + pos);
    }
  }
  return NULL;
}

// IMPORTANT!!!  wstr1 and wstr2 must be valid strings!
static int wstr_unsafe_cmp(const wchar_t* wstr1, const wchar_t* wstr2)
{
  while ( (*wstr1) && (*wstr1 == *wstr2) )
  {
    ++wstr1;
    ++wstr2;
  }
  return (*wstr1 - *wstr2);
}

// IMPORTANT!!!  wstr and wsubstr must be valid strings!
static int wstr_unsafe_subcmp(const wchar_t* wstr, const wchar_t* wsubstr)
{
  while ( (*wstr) && (*wstr == *wsubstr) )
  {
    ++wstr;
    ++wsubstr;
  }
  return (*wsubstr) ? (*wstr - *wsubstr) : 0;
}

static BOOL wstr_is_listed_ext(const wchar_t* szExtW, 
  wchar_t* szExtListW, char* szExtListA)
{
  if ( szExtW && szExtW[0] )
  {
    int      len, i, n;
    wchar_t  szW[MAX_EXT];

    if (g_bOldWindows)
    {
      len = lstrlenA(szExtListA);
      MultiByteToWideChar(CP_ACP, 0, szExtListA, len,
        szExtListW, STR_FILEEXTS_SIZE - 1);
      szExtListW[len] = 0;
    }
    else
    {
      len = lstrlenW(szExtListW);
    }

    n = 0;
    i = 0;
    while (i <= len)
    {
      if ( (szExtListW[i]) &&
           (szExtListW[i] != L';') &&
           (szExtListW[i] != L',') &&
           (szExtListW[i] != L' ') )
      {
        szW[n++] = szExtListW[i];
      }
      else
      {
        if (n > 0)
        {
          szW[n] = 0;
          if (wstr_unsafe_cmp(szExtW, szW) == 0)
            return TRUE;
        }
        n = 0;
      }
      ++i;
    }
  }
  return FALSE;
}

static BOOL wstr_is_comment1_ext(const wchar_t* szExtW)
{
  return wstr_is_listed_ext(szExtW, strComment1FileExtsW, strComment1FileExtsA);
}

static BOOL wstr_is_escaped1_ext(const wchar_t* szExtW)
{
  if (g_bOldWindows)
  {
    if (strEscaped1FileExtsA[0] == 0)
      return TRUE;
  }
  else
  {
    if (strEscaped1FileExtsW[0] == 0)
      return TRUE;
  }
  return wstr_is_listed_ext(szExtW, strEscaped1FileExtsW, strEscaped1FileExtsA);
}

static BOOL wstr_is_html_compatible(const wchar_t* szExtW)
{
  if ( szExtW && szExtW[0] )
  {
    int      len, i, n;
    wchar_t  szW[MAX_EXT];

    if (g_bOldWindows)
    {
      len = lstrlenA(strHtmlFileExtsA);
      MultiByteToWideChar(CP_ACP, 0, strHtmlFileExtsA, len,
        strHtmlFileExtsW, STR_FILEEXTS_SIZE - 1);
      strHtmlFileExtsW[len] = 0;
    }
    else
    {
      len = lstrlenW(strHtmlFileExtsW);
    }

    n = 0;
    i = 0;
    while (i <= len)
    {
      if ( (strHtmlFileExtsW[i]) &&
           (strHtmlFileExtsW[i] != L';') &&
           (strHtmlFileExtsW[i] != L',') &&
           (strHtmlFileExtsW[i] != L' ') )
      {
        szW[n++] = strHtmlFileExtsW[i];
      }
      else
      {
        if (n > 0)
        {
          szW[n] = 0;
          n = 0;
          while (szExtW[n])
          {
            if (wstr_unsafe_subcmp(szExtW + n, szW) == 0)
              return TRUE;
            else
              ++n;
          }
        }
        n = 0;
      }
      ++i;
    }
  }
  return FALSE;
}

int getFileType(int* pnCurrentFileType2)
{
  EDITINFO    ei;
  wchar_t     szExtW[MAX_EXT];
  const void* p = NULL;
  int         nType = tftNone;
  int         nType2 = tfmNone;

  if (SendMessage(g_hMainWnd, AKD_GETEDITINFO, (WPARAM) hActualEditWnd, (LPARAM) &ei) != 0)
  {
    if (g_bOldWindows)
    {
      p = getFileExtA( (const char*) ei.pFile ); // file extension ptr (char *)
      if (p)
      {
        char szExtA[MAX_EXT];
        int  len = 0;
        while ( (len < MAX_EXT - 1) &&
                ((szExtA[len] = ((const char*) p)[len]) != 0) )
        {
          ++len;
        }
        szExtA[len] = 0;
        CharLowerA(szExtA); // MustDie9x does not have CharLowerW !
        MultiByteToWideChar(CP_ACP, 0, szExtA, len, szExtW, MAX_EXT - 1);
        szExtW[len] = 0; // file extension (wchar_t *)
      }
    }
    else
    {
      p = getFileExtW( (const wchar_t*) ei.pFile ); // file extension ptr (wchar_t *)
      if (p)
      {
        int len = 0;
        while ( (len < MAX_EXT - 1) &&
                ((szExtW[len] = ((const wchar_t*) p)[len]) != 0) )
        {
          ++len;
        }
        szExtW[len] = 0; // file extension (wchar_t *)
        CharLowerW(szExtW);
      }
    }
  }

  if (p)
  {
    hCurrentEditWnd = hActualEditWnd;

    //MessageBoxW(NULL, szExtW, L"ext", MB_OK);

    if ( (wstr_unsafe_cmp(szExtW, L"c") == 0)   ||
         (wstr_unsafe_cmp(szExtW, L"cc") == 0) ||
         (wstr_unsafe_cmp(szExtW, L"cpp") == 0) ||
         (wstr_unsafe_cmp(szExtW, L"cxx") == 0) )
    {
      nType = tftC_Cpp;
      nType2 = tfmComment1 | tfmEscaped1;
    }
    else if ( (wstr_unsafe_cmp(szExtW, L"h") == 0) ||
              (wstr_unsafe_cmp(szExtW, L"hh") == 0) ||
              (wstr_unsafe_cmp(szExtW, L"hpp") == 0) ||
              (wstr_unsafe_cmp(szExtW, L"hxx") == 0) )
    {
      nType = tftH_Hpp;
      nType2 = tfmComment1 | tfmEscaped1;
    }
    else if ( wstr_unsafe_cmp(szExtW, L"pas") == 0 )
    {
      nType = tftPas;
      nType2 = tfmComment1;
    }
    else
    {
      nType = tftText;
      if ( wstr_is_comment1_ext(szExtW) )
      {
        nType2 |= tfmComment1;
      }
      if ( wstr_is_escaped1_ext(szExtW) )
      {
        nType2 |= tfmEscaped1;
      }
      if ( wstr_is_html_compatible(szExtW) )
      {
        nType2 |= tfmHtmlCompatible;
      }
    }

  }
  else
  {
    hCurrentEditWnd = NULL;
  }

  *pnCurrentFileType2 = nType2;
  return nType;
}

void AutoBrackets_Uninitialize(void)
{
  int i;
    
  nCurrentFileType = tftNone;
  nCurrentFileType2 = tfmNone;
  hCurrentEditWnd = NULL;
  hActualEditWnd = NULL;
  nAutoRightBracketPos = -1;
  nAutoRightBracketType = tbtNone;

  RemoveAllHighlightInfo(TRUE);
  
  for (i = 0; i < HIGHLIGHT_INDEXES; i++)
  {
    IndexesHighlighted[i] = -1;
    IndexesToHighlight[i] = -1;
    prevIndexesToHighlight[i] = -1;
  }
  CurrentBracketsIndexes[0] = -1;
  CurrentBracketsIndexes[1] = -1;
}

static void copyUniqueCharsOkW(wchar_t* pDst, const wchar_t* pSrc)
{
  int n = 0;

  if (pSrc)
  {
    int     i;
    wchar_t wch;
      
    while ((wch = *pSrc) != 0)
    {
      if ((wch != L' ') && 
          (wch != L'\t') && 
          (wch != L'\x0D') && 
          (wch != L'\x0A'))
      {
        i = 0;
        while ((i < n) && (pDst[i] != wch))
        {
          ++i;
        }
        if (i == n)
        {
          pDst[n] = wch;
          ++n;
        }
      }
      ++pSrc;
    }
  }

  pDst[n] = 0;
}

void setNextCharOkW(const wchar_t* cszNextCharOkW)
{
  copyUniqueCharsOkW(strNextCharOkW, cszNextCharOkW);
}

void setPrevCharOkW(const wchar_t* cszPrevCharOkW)
{
  copyUniqueCharsOkW(strPrevCharOkW, cszPrevCharOkW);
}

void setUserBracketsA(const char* cszUserBracketsA)
{
  strUserBracketsA[0][0] = 0;
  strUserBracketsA[0][1] = 0;
  strUserBracketsA[0][2] = 0;
  strUserBracketsA[0][3] = 0;

  if (cszUserBracketsA && cszUserBracketsA[0])
  {
    int n = 0;
    int i = 0;
    const char* p = cszUserBracketsA;

    while (n < MAX_USER_BRACKETS)
    {
      if ((*p) && (*p != ' ') && (*p != '\t'))
      {
        if (i < 2)
        {
          // brackets pair: 1st or 2nd character
          strUserBracketsA[n][i] = *p;
        }
        else if (i == 2)
        {
          // error: more than 2 characters
          strUserBracketsA[n][0] = 0;
          strUserBracketsA[n][1] = 0;
          strUserBracketsA[n][2] = 0;
          strUserBracketsA[n][3] = 0;
        }
        ++i;
      }
      else
      {
        if (i >= 1)
        {
          if (i == 2)
          {
            // brackets pair completed
            strUserBracketsA[n][2] = 0;
            strUserBracketsA[n][3] = 0;
            ++n;
          }
          // erase incomplete brackets pair
          strUserBracketsA[n][0] = 0;
          strUserBracketsA[n][1] = 0;
          strUserBracketsA[n][2] = 0;
          strUserBracketsA[n][3] = 0;
        }
        i = 0;
        if (*p == 0)
          break;
      }
      ++p;
    }
  }
}

void setUserBracketsW(const wchar_t* cszUserBracketsW)
{
  strUserBracketsW[0][0] = 0;
  strUserBracketsW[0][1] = 0;
  strUserBracketsW[0][2] = 0;
  strUserBracketsW[0][3] = 0;

  if (cszUserBracketsW && cszUserBracketsW[0])
  {
    int n = 0;
    int i = 0;
    const wchar_t* p = cszUserBracketsW;

    while (n < MAX_USER_BRACKETS)
    {
      if ((*p) && (*p != L' ') && (*p != L'\t'))
      {
        if (i < 2)
        {
          // brackets pair: 1st or 2nd character
          strUserBracketsW[n][i] = *p;
        }
        else if (i == 2)
        {
          // error: more than 2 characters
          strUserBracketsW[n][0] = 0;
          strUserBracketsW[n][1] = 0;
          strUserBracketsW[n][2] = 0;
          strUserBracketsW[n][3] = 0;
        }
        ++i;
      }
      else
      {
        if (i >= 1)
        {
          if (i == 2)
          {
            // brackets pair completed
            strUserBracketsW[n][2] = 0;
            strUserBracketsW[n][3] = 0;
            ++n;
          }
          // erase incomplete brackets pair
          strUserBracketsW[n][0] = 0;
          strUserBracketsW[n][1] = 0;
          strUserBracketsW[n][2] = 0;
          strUserBracketsW[n][3] = 0;
        }
        i = 0;
        if (*p == 0)
          break;
      }
      ++p;
    }
  }
}

const char* getCurrentBracketsPairA(void)
{
  static char szBracketsPair[4];
  int i;

  for (i = 0; i < 4; i++)
    szBracketsPair[i] = 0;

  if (CurrentBracketsIndexes[0] >= 0)
  {
    char ch;

    szBracketsPair[0] = AnyRichEdit_GetCharAt(hActualEditWnd, CurrentBracketsIndexes[0]);
    ch = AnyRichEdit_GetCharAt(hActualEditWnd, CurrentBracketsIndexes[1]);
    if (ch == '>')
    {
      // "<...>" or "<.../>"
      ch = AnyRichEdit_GetCharAt(hActualEditWnd, CurrentBracketsIndexes[1] - 1);
      if (ch == '/')
      {
        szBracketsPair[1] = '/';
        szBracketsPair[2] = '>';
      }
      else
        szBracketsPair[1] = '>';
    }
    else
      szBracketsPair[1] = ch;
  }

  return szBracketsPair;
}

const wchar_t* getCurrentBracketsPairW(void)
{
  static wchar_t szBracketsPair[4];
  int i;

  for (i = 0; i < 4; i++)
    szBracketsPair[i] = 0;

  if (CurrentBracketsIndexes[0] >= 0)
  {
    wchar_t wch;

    szBracketsPair[0] = AnyRichEdit_GetCharAtW(hActualEditWnd, CurrentBracketsIndexes[0]);
    wch = AnyRichEdit_GetCharAtW(hActualEditWnd, CurrentBracketsIndexes[1]);
    if (wch == L'>')
    {
      // "<...>" or "<.../>"
      wch = AnyRichEdit_GetCharAtW(hActualEditWnd, CurrentBracketsIndexes[1] - 1);
      if (wch == L'/')
      {
        szBracketsPair[1] = L'/';
        szBracketsPair[2] = L'>';
      }
      else
        szBracketsPair[1] = L'>';
    }
    else
      szBracketsPair[1] = wch;
  }

  return szBracketsPair;
}
