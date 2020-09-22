Attribute VB_Name = "mGeneral"
Option Explicit

'/----------------------\
'|   API DECLARATIONS   |
'\----------------------/

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Declare Function SendMessageLng Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function MapDialogRect Lib "user32" (ByVal hDlg As Long, lpRect As RECT) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'Window Messages ---------------------
Public Const WM_VSCROLL = &H115
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_CHAR = &H102
Public Const WM_PASTE = &H302
Public Const WM_COPY = &H301
Public Const WM_CUT = &H300
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_REFRESH_LINE_NUMBERS = WM_USER
'-------------------------------------

' RichTextBox SendMessage Messages
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_LINESCROLL = &HB6
Public Const EM_GETRECT = &HB2
Public Const EM_POSFROMCHAR = &HD6
Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_CANUNDO = &HC6
Public Const EM_SETTEXTEX = (WM_USER + 97)
Public Const EM_GETTEXTEX = (WM_USER + 94)
Public Const EM_GETSELTEXT = (WM_USER + 62)
Public Const EM_REPLACESEL = &HC2
Public Const EM_STREAMOUT = (WM_USER + 74)
Public Const EM_STREAMIN = (WM_USER + 73)
Public Const EM_SCROLLCARET = &HB7
Public Const EM_FINDTEXTEX = (WM_USER + 79)



' --------------------------------

' Consts for various API calls
' --------------------------------
Public Const ST_DEFAULT = 0
Public Const ST_KEEPUNDO = 1
Public Const ST_SELECTION = 2

Public Const GT_DEFAULT = 0
Public Const GT_USECRLF = 1
Public Const GT_SELECTION = 2

Public Const CP_ACP = 0  '  default to ANSI code page
' --------------------------------

' Type Definitions
' --------------------------------
Public Type GETTEXTEX
    cb As Long
    flags As Long
    codepage As Long
    lpDefaultChar As String
    lpUsedDefChar As Boolean
End Type

Public Type SETTEXTEX
    flags As Long
    codepage As Long
End Type

Public Type EDITSTREAM
    dwCookie As Long
    dwError As Long
    pfnCallback As Long
End Type

Public Type CharRange
    cpMin As Long
    cpMax As Long
End Type

Public Type FindTextEx
    chrg As CharRange
    lpstrText As String
    chrgText As CharRange
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Enum eTextType
    SF_TEXT = &H1
    SF_RTF = &H2
End Enum

Public Enum eReplace
    SFF_ALL = &H0
    SFF_SELECTION = &H8000
End Enum

Public Enum eFindText
    FR_DOWN = &H1
    FR_MATCHCASE = &H4
    FR_FINDNEXT = &H8
    FR_WHOLEWORD = &H2
End Enum
' --------------------------------

Public Const TWIPS_PER_PIXEL As Long = 15

Public g_bNoRaiseSelChange As Boolean
Public g_bIgnoreNextKeyPress As Boolean

' Variables used for Streaming Text to/from RTB
Private m_sRTFText As String
Private m_lLen As Long
Private m_lPos As Long

'Get the High and Low Word of a given number
Public Function LoWord(lNumber As Long) As Long
    LoWord = lNumber And &HFFFF&
End Function

Public Function HiWord(lNumber As Long) As Long
    HiWord = (lNumber And &HFFFF0000) / &HFFFF&
End Function

'Gets/Sets the selection for the RichTextBox with the Window Handle hWnd
Public Sub SetSelection(ByVal hWnd As Long, ByVal lSelStart As Long, ByVal lSelEnd As Long)
    g_bNoRaiseSelChange = True
    Call SendMessage(hWnd, EM_SETSEL, lSelStart, lSelEnd)
    g_bNoRaiseSelChange = False
End Sub

Public Sub GetSelection(ByVal hWnd As Long, lSelStart As Long, lSelEnd As Long)
    Call SendMessage(hWnd, EM_GETSEL, VarPtr(lSelStart), VarPtr(lSelEnd))
End Sub

' Gets the location of the first and last character of a given line
' if lLineIndex is -1 then it uses the current line
Public Sub GetLinePosition(ByVal hWnd As Long, ByVal lLineIndex As Long, lLineStart As Long, lLineEnd As Long)
    Dim lIndex As Long
    
    If lLineIndex = -1 Then
        lIndex = SendMessage(hWnd, EM_LINEFROMCHAR, -1, 0&)
    Else
        lIndex = lLineIndex
    End If
    
    lLineStart = SendMessage(hWnd, EM_LINEINDEX, lIndex, 0&)
    lLineEnd = lLineStart + SendMessage(hWnd, EM_LINELENGTH, lLineStart, 0&)
End Sub

' Gets/Sets the RTF/Text for the RichTextBox (Selection or Entire)
Public Function GetText(hWnd As Long, ttTextType As eTextType, rSelection As eReplace) As String

    Dim esRTFText As EDITSTREAM
    esRTFText.dwCookie = 0
    esRTFText.pfnCallback = FncPtr(AddressOf GetRTFCallBack)
    esRTFText.dwError = 0
    
    m_sRTFText = ""
    SendMessage hWnd, EM_STREAMOUT, ttTextType Or rSelection, VarPtr(esRTFText)
    
    GetText = m_sRTFText
    m_sRTFText = ""
End Function

Public Sub SetText(ByVal hWnd As Long, ttTextType As eTextType, rSelection As eReplace, sText As String)
    Dim esRTFText As EDITSTREAM
    Dim lR As Long

    esRTFText.dwCookie = 0
    esRTFText.pfnCallback = FncPtr(AddressOf SetRTFCallBack)
    esRTFText.dwError = 0
    m_sRTFText = sText
    m_lPos = 1
    m_lLen = Len(m_sRTFText)
   ' The text will be streamed in though the LoadCallback function:
   lR = SendMessage(hWnd, EM_STREAMIN, ttTextType Or rSelection, VarPtr(esRTFText))
End Sub

Public Function Occurances(sCheck As String, sLookFor As String) As Long
    Dim i As Integer
    Dim lCount As Long
    Dim lLenCheck As Long
    Dim lLenLookFor As Long
    
    lLenCheck = Len(sCheck)
    lLenLookFor = Len(sLookFor)
    For i = 1 To lLenCheck - lLenLookFor + 1
        If Mid(sCheck, i, lLenLookFor) = sLookFor Then
            lCount = lCount + 1
        End If
    Next i
    Occurances = lCount
End Function

' Wrapper for AddressOf operator
Public Function FncPtr(pFunction As Long) As Long
    FncPtr = pFunction
End Function

' Callback function for GetText
Public Function GetRTFCallBack(ByVal dwCookie As Long, _
                               ByVal lPtrPbBuff As Long, _
                               ByVal cb As Long, _
                               ByVal pcb As Long) As Long
Dim sBuf As String
Dim bData() As Byte
Dim lLen As Long

    lLen = cb
    
    If (lLen > 0) Then
        ReDim bData(0 To lLen - 1) As Byte
        CopyMemory bData(0), ByVal lPtrPbBuff, lLen
        sBuf = StrConv(bData, vbUnicode)
        CopyMemory ByVal pcb, lLen, 4
        m_sRTFText = m_sRTFText & sBuf
    End If
    GetRTFCallBack = 0
End Function

' Callback function for Set Text
Public Function SetRTFCallBack(ByVal dwCookie As Long, _
                               ByVal lPtrPbBuff As Long, _
                               ByVal cb As Long, _
                               ByVal pcb As Long) As Long
Dim sBuf As String
Dim bData() As Byte
Dim lLen As Long
Dim lRead As Long

    CopyMemory lRead, ByVal pcb, 4
    If (m_lLen - m_lPos >= 0) Then
        If (m_lLen - m_lPos < cb) Then
            ReDim bData(0 To (m_lLen - m_lPos)) As Byte
            bData = StrConv(Mid$(m_sRTFText, m_lPos), vbFromUnicode)
            lRead = m_lLen - m_lPos + 1
            CopyMemory ByVal lPtrPbBuff, bData(0), lRead
            m_lPos = m_lLen + 1
        Else
            ReDim bData(0 To cb - 1) As Byte
            bData = StrConv(Mid$(m_sRTFText, m_lPos, cb), vbFromUnicode)
            CopyMemory ByVal lPtrPbBuff, bData(0), cb
            m_lPos = m_lPos + cb
            lRead = cb
        End If
                    
        CopyMemory ByVal pcb, lRead, 4
        SetRTFCallBack = 0
    Else
        lRead = 0
        CopyMemory ByVal pcb, lRead, 4
        SetRTFCallBack = 0
    End If
    
End Function
