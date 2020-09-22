VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSyntaxHighlighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' --------------------------------------------------------
' Name:     cSyntaxHighlighter
' Author:   Chris Thomas
' Company:  Symbiote Software
' Date      March 12, 2003
'
' Purpose:  To provide a RichTextBox with Syntax Highlighting
'           capabilities
' Based On: Syntax Highlighter by Gary Noble
' --------------------------------------------------------
Option Explicit

' Good 'ol APIs
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

' Color Tags for RichText
Private Const TAG_KEYWORD As String = "\cf2 "
Private Const TAG_COMMENT As String = "\cf3 "
Private Const TAG_STRING As String = "\cf4 "
Private Const TAG_DEFAULT As String = "\cf1 "

' String of keywords seperated by m_sListSeperator
Private m_sKeywords As String

' List Seperator for the keywords passed by the user
Private m_sListSeperator As String

' Language Syntax variables
Private m_sLanguageName As String
Private m_lKeywordColor As Long
Private m_lStringColor As Long
Private m_lCommentColor As Long
Private m_lDefaultColor As Long

Private m_sStringChar As String
Private m_sLineCommentChar As String
Private m_sBlockCommentStartChar As String
Private m_sBlockCommentEndChar As String
Private m_bSupportsBlockComment As Boolean

Private m_fntFont As IFontDisp

' The state of the RTB after the last Coloring.
' Used to check if the RTB needs to be Colored again
Private m_sAfterLastColor As String

' RTF Header
Private m_sRTFHeader As String

' Handle to the RTB
Private m_hCodeBoxWnd As Long

' Set the Handle for the RTB
Friend Property Let CodeBoxWnd(hWnd As Long)
    m_hCodeBoxWnd = hWnd
End Property

'Checks to see if the passed char(s) is a terminal char
Public Function IsTerminalCharacter(ByVal sText As String) As Boolean
    Dim sCharacter As String
    Dim lPos As Long
    Dim bFound As Boolean

    If m_bSupportsBlockComment Then
        If Right$(sText, Len(m_sBlockCommentStartChar)) = m_sBlockCommentStartChar Then
            IsTerminalCharacter = True
            Exit Function
        End If

        If Right$(sText, Len(m_sBlockCommentEndChar)) = m_sBlockCommentEndChar Then
            IsTerminalCharacter = True
            Exit Function
        End If
    End If
    
    If Right$(sText, Len(m_sLineCommentChar)) = m_sLineCommentChar Then
        IsTerminalCharacter = True
        Exit Function
    End If
    
    If Right$(sText, Len(m_sStringChar)) = m_sStringChar Then
        IsTerminalCharacter = True
        Exit Function
    End If
    
    sCharacter = Right(sText, 1)
    If sCharacter = " " Then IsTerminalCharacter = True: Exit Function
    If sCharacter = "." Then IsTerminalCharacter = True: Exit Function
    If sCharacter = vbTab Then IsTerminalCharacter = True: Exit Function
    If sCharacter = vbLf Then IsTerminalCharacter = True: Exit Function
    If sCharacter = vbCr Then IsTerminalCharacter = True: Exit Function
    
    IsTerminalCharacter = False
End Function

Public Function FormatWord(ByVal sWord As String) As String
    Dim lPos As Long
    lPos = InStr(1, LCase(m_sKeywords), LCase(m_sListSeperator & sWord & m_sListSeperator))
    If lPos <> 0 Then
        'FormatWord = TAG_KEYWORD & Mid$(m_sKeywords, lPos + 1, Len(sWord)) & TAG_DEFAULT
        FormatWord = TAG_KEYWORD & sWord & TAG_DEFAULT
    Else
        FormatWord = sWord
    End If
End Function

Friend Sub KeyPressed(ByVal iKeyPressed As Integer)
    
    Dim sText As String
    Dim lSelStart As Long
    Dim lSelEnd As Long

    If iKeyPressed <> vbKeyBack And iKeyPressed <> vbKeyDelete Then
        LockWindowUpdate m_hCodeBoxWnd
        If iKeyPressed = vbKeyReturn Then
                            
            GetSelection m_hCodeBoxWnd, lSelStart, lSelEnd
            SetText m_hCodeBoxWnd, SF_RTF, SFF_SELECTION, m_sRTFHeader & "\par }"
            
            SetSelection m_hCodeBoxWnd, lSelStart, lSelStart
            Colorize True
            
            SetSelection m_hCodeBoxWnd, lSelStart + 2, lSelStart + 2
            Colorize True

        Else
            GetSelection m_hCodeBoxWnd, lSelStart, lSelEnd
            SetText m_hCodeBoxWnd, SF_TEXT, SFF_SELECTION, Chr(iKeyPressed)
            Colorize True
        End If
        LockWindowUpdate 0
        g_bIgnoreNextKeyPress = True
        
    End If
    
End Sub

Public Sub Colorize(Optional bCurrentLineOnly As Boolean = False, Optional bOverrideChangeCheck As Boolean = False)
    Dim sCode As String
    Dim sRet As String
    
    Dim lStartOfLine As Long
    Dim lEndOfLine As Long
    
    Dim lSelStart As Long
    Dim lSelEnd As Long
    Dim lLastCommentStart As Long
    Dim lLastCommentEnd As Long
    Dim bInBlockComment As Boolean
    
    LockWindowUpdate m_hCodeBoxWnd
    If bCurrentLineOnly = False Then
        If (m_sAfterLastColor <> GetText(m_hCodeBoxWnd, SF_RTF, SFF_ALL)) Or bOverrideChangeCheck Then

            GetSelection m_hCodeBoxWnd, lSelStart, lSelEnd

            sCode = Trim(GetText(m_hCodeBoxWnd, SF_TEXT, SFF_ALL))
            sCode = ProcessHighlight(sCode)

            SetText m_hCodeBoxWnd, SF_RTF, SFF_ALL, sCode
            
            SetSelection m_hCodeBoxWnd, lSelStart, lSelEnd
            m_sAfterLastColor = sCode
        End If
    Else

        GetSelection m_hCodeBoxWnd, lSelStart, lSelEnd

        GetLinePosition m_hCodeBoxWnd, -1, lStartOfLine, lEndOfLine
        
        
        'If lStartOfLine <> 0 Then
            SetSelection m_hCodeBoxWnd, 0, lStartOfLine
            sCode = GetText(m_hCodeBoxWnd, SF_TEXT, SFF_SELECTION)
            lLastCommentStart = InStrRev(sCode, m_sBlockCommentStartChar, , vbTextCompare)
            lLastCommentEnd = InStrRev(sCode, m_sBlockCommentEndChar, , vbTextCompare)
            bInBlockComment = lLastCommentStart > lLastCommentEnd
            SetSelection m_hCodeBoxWnd, lStartOfLine, lEndOfLine
        'Else
        '    bInBlockComment = False
        'End If
        
        sCode = GetText(m_hCodeBoxWnd, SF_TEXT, SFF_SELECTION)
        
        sRet = ProcessHighlight(sCode, bInBlockComment)
        
        SetText m_hCodeBoxWnd, SF_RTF, SFF_SELECTION, sRet

        SetSelection m_hCodeBoxWnd, lSelStart, lSelStart
        
        m_sAfterLastColor = GetText(m_hCodeBoxWnd, SF_RTF, SFF_ALL)

    End If
    SendMessage m_hCodeBoxWnd, EM_SCROLLCARET, 0&, 0&
    SendMessage m_hCodeBoxWnd, WM_REFRESH_LINE_NUMBERS, 0&, 0&
    LockWindowUpdate 0
End Sub

Private Function ProcessHighlight(sUnformatted As String, Optional bAlreadyInBlock As Boolean = False) As String
        
    Dim sctAppend As New cStringConcatenator
    
    Dim sText As String
    Dim sChar As String
    
    Dim lCount As Long
    Dim lMax As Long
    
    'Dim bSupportBlockComments As Boolean
    
    Dim bInString As Boolean
    Dim bInLineComment As Boolean
    Dim bInBlockComment As Boolean
    
    lMax = Len(sUnformatted)
   
    bInString = False
    bInLineComment = False
    bInBlockComment = bAlreadyInBlock
       
    sctAppend.Append m_sRTFHeader                 'Append RTF Header

    
    If bInBlockComment Then
        sctAppend.Append TAG_COMMENT
    End If
    
    For lCount = 1 To lMax
        sChar = Mid(sUnformatted, lCount, 1)
        sText = sText & sChar
        
        If sChar = "\" Then     '-- Richtext Doesnot Support Single '\' Characters
            sChar = "\\"        '-- We Must Place A Extra '\' For It To Show
            If Len(sText) > 1 Then
                sctAppend.Append FormatWord(Mid(sText, 1, Len(sText) - 1))
            End If
            sctAppend.Append sChar
            sText = ""
            sChar = ""
        ElseIf sChar = "{" Then
            sChar = "\{"
            If Len(sText) > 1 Then
                sctAppend.Append FormatWord(Mid(sText, 1, Len(sText) - 1))
            End If
            sctAppend.Append sChar
            sText = ""
            sChar = ""
        ElseIf sChar = "}" Then
            sChar = "\}"
            If Len(sText) > 1 Then
                sctAppend.Append FormatWord(Mid(sText, 1, Len(sText) - 1))
            End If
            sctAppend.Append sChar
            sText = ""
            sChar = ""
        End If
                
        If bInString Then
            'sText = sText & sChar
            If sChar = m_sStringChar Then
                bInString = False
                sctAppend.Append sText
                sctAppend.Append TAG_DEFAULT
                sText = ""
                sChar = ""
            ElseIf sChar = vbCr Then
                sctAppend.Append Mid(sText, 1, Len(sText) - 1)
                sctAppend.Append "\par "
                sText = ""
                sChar = ""
            End If
        ElseIf bInBlockComment Then
            'sText = sText & sChar
            If Right(sText, Len(m_sBlockCommentEndChar)) = m_sBlockCommentEndChar Then
                bInBlockComment = False
                sctAppend.Append sText
                sctAppend.Append TAG_DEFAULT
                sText = ""
                sChar = ""
            End If
            If sChar = vbCr Then
                sctAppend.Append Mid(sText, 1, Len(sText) - 1)
                sctAppend.Append "\par "
                sctAppend.Append TAG_COMMENT
                sText = ""
                sChar = ""
            End If
        ElseIf bInLineComment Then
            'sText = sText & sChar
            If sChar = vbCr Then
                bInLineComment = False
                If Len(sText) > 0 Then
                    sctAppend.Append Mid(sText, 1, Len(sText) - 1)
                End If
                sctAppend.Append "\par "
                sctAppend.Append TAG_DEFAULT
                sText = ""
                sChar = ""
            End If
        End If
        
        If IsTerminalCharacter(sText) And Not (bInString Or bInLineComment Or bInBlockComment) Then
           
           '-- Open a String
           If sChar = m_sStringChar Then
                bInString = True
                If Len(sText) > 0 Then
                    sctAppend.Append Left(sText, Len(sText) - 1)
                    sctAppend.Append TAG_STRING
                    sctAppend.Append sChar
                    sText = ""
                    sChar = ""
                Else
                    sctAppend.Append TAG_STRING
                    sctAppend.Append sText
                    sText = sChar
                End If
                        
            '-- Comment Text
            ElseIf Right$(sText, Len(m_sLineCommentChar)) = m_sLineCommentChar Then
                bInLineComment = True

                If Len(sText) > Len(m_sLineCommentChar) Then
                    sctAppend.Append FormatWord(Mid(sText, 1, Len(sText) - 1))
                End If
                sctAppend.Append TAG_COMMENT
                sctAppend.Append sText
                sText = ""
                sChar = ""

            '-- Terminal Text
            ElseIf m_bSupportsBlockComment And Right(sText, Len(m_sBlockCommentStartChar)) = m_sBlockCommentStartChar Then
                bInBlockComment = True
                If Len(sText) > 0 Then
                    sctAppend.Append TAG_COMMENT
                    sctAppend.Append sText
                    sText = ""
                    sChar = ""
                Else
                    sctAppend.Append TAG_COMMENT
                    sctAppend.Append sText
                    sText = sChar
                End If
            Else
                If sChar = vbCr Then
                    sChar = TAG_DEFAULT & "\par "
                ElseIf sChar = vbLf Then
                    sChar = ""
                ElseIf sChar = Chr(0) Then
                    sChar = ""
                End If
                sctAppend.Append FormatWord(Mid(sText, 1, Len(sText) - 1))
                sctAppend.Append sChar
                sText = ""
            End If
        End If
    Next lCount
    
    If Not (bInString Or bInLineComment Or bInBlockComment) Then
        sctAppend.Append FormatWord(sText)
    Else
        sctAppend.Append sText
    End If
    
    sctAppend.Append TAG_DEFAULT
    sctAppend.Append "}"
    ProcessHighlight = sctAppend.Value
    
    sctAppend.Clear
    Set sctAppend = Nothing
End Function

'Create the RTF Header
Private Function CreateRTFHeader() As String
    Dim sHeader As String
    
    sHeader = "{\rtf1\ansi\ansicpg1252\deff0\deflang3081\deflangfe1033" & _
              "{\fonttbl{\f0\fnil\fcharset" & m_fntFont.Charset & " " & _
              m_fntFont.Name & "\fs" & Trim(Str(Int(m_fntFont.Size * 2))) & ";}}" & _
              GenerateColorTable & _
              "\lang2056\viewkind4\uc1\pard" & TAG_DEFAULT & "\fs" & Trim(Str(Int(m_fntFont.Size * 2))) & "\cf0 "
    m_sRTFHeader = sHeader
End Function

Private Function GenerateColorTable() As String
    
    Dim sTable As String
    Dim sColor As String
    Dim iRed As Integer
    Dim iGreen As Integer
    Dim iBlue As Integer
    
    sTable = "{\colortbl ;"
    
    ' Default Color -------------------------------------------------------
    iRed = m_lDefaultColor Mod 256
    iGreen = m_lDefaultColor \ 256 Mod 256
    iBlue = m_lDefaultColor \ 256 \ 256 Mod 256
    
    sColor = "{\red" & iRed & "\green" & iGreen & "\blue" & iBlue & ";}"
    
    sTable = sTable & sColor
    
    ' Keyword Color -------------------------------------------------------
    iRed = m_lKeywordColor Mod 256
    iGreen = m_lKeywordColor \ 256 Mod 256
    iBlue = m_lKeywordColor \ 256 \ 256 Mod 256
    
    sColor = "{\red" & iRed & "\green" & iGreen & "\blue" & iBlue & ";}"
    
    sTable = sTable & sColor
    
    ' Comment Color -------------------------------------------------------
    iRed = m_lCommentColor Mod 256
    iGreen = m_lCommentColor \ 256 Mod 256
    iBlue = m_lCommentColor \ 256 \ 256 Mod 256
    
    sColor = "{\red" & iRed & "\green" & iGreen & "\blue" & iBlue & ";}"
    
    sTable = sTable & sColor
    
    ' String Color --------------------------------------------------------
    iRed = m_lStringColor Mod 256
    iGreen = m_lStringColor \ 256 Mod 256
    iBlue = m_lStringColor \ 256 \ 256 Mod 256
    
    sColor = "{\red" & iRed & "\green" & iGreen & "\blue" & iBlue & ";}"
    sTable = sTable & sColor
    
    sTable = sTable & "}"
    
    GenerateColorTable = sTable
End Function

' Properties ---------------------------------------------
Public Property Get LineCommentChar() As String
    LineCommentChar = m_sLineCommentChar
End Property

Public Property Let LineCommentChar(sChar As String)
    m_sLineCommentChar = sChar
End Property

Public Property Let BlockCommentStartChar(sChar As String)
    m_sBlockCommentStartChar = sChar
    m_bSupportsBlockComment = (Len(m_sBlockCommentStartChar) <> 0) And (Len(m_sBlockCommentEndChar) <> 0)
End Property

Public Property Get BlockCommentStartChar() As String
    BlockCommentStartChar = m_sBlockCommentStartChar
End Property

Public Property Let BlockCommentEndChar(sChar As String)
    m_sBlockCommentEndChar = sChar
    m_bSupportsBlockComment = (Len(m_sBlockCommentStartChar) <> 0) And (Len(m_sBlockCommentEndChar) <> 0)
End Property

Public Property Get BlockCommentEndChar() As String
    BlockCommentEndChar = m_sBlockCommentEndChar
End Property

Public Property Get StringChar() As String
    StringChar = m_sStringChar
End Property

Public Property Let StringChar(sChar As String)
    m_sStringChar = sChar
End Property

Public Property Get Keywords() As String
    Keywords = m_sKeywords
End Property

Public Property Let Keywords(sKeywords As String)
    m_sKeywords = sKeywords
    m_sListSeperator = Left(m_sKeywords, 1)
End Property

Public Property Get LanguageName() As String
    LanguageName = m_sLanguageName
End Property

Public Property Let LanguageName(sLanguage As String)
    m_sLanguageName = sLanguage
End Property
    
Public Property Get CommentColor() As Long
    CommentColor = m_lCommentColor
End Property

Public Property Let CommentColor(lColor As Long)
    m_lCommentColor = lColor
    CreateRTFHeader
End Property

Public Property Get StringColor() As Long
    StringColor = m_lStringColor
End Property

Public Property Let StringColor(lColor As Long)
    m_lStringColor = lColor
    CreateRTFHeader
End Property

Public Property Get KeywordColor() As Long
    KeywordColor = m_lKeywordColor
End Property

Public Property Let KeywordColor(lColor As Long)
    m_lKeywordColor = lColor
    CreateRTFHeader
End Property

Public Property Get DefaultColor() As Long
    DefaultColor = m_lDefaultColor
End Property

Public Property Let DefaultColor(lColor As Long)
    m_lDefaultColor = lColor
    CreateRTFHeader
End Property

Public Property Get Font() As IFontDisp
    Set Font = m_fntFont
End Property

Public Property Set Font(fntFont As IFontDisp)
    Set m_fntFont = fntFont
    CreateRTFHeader
End Property
