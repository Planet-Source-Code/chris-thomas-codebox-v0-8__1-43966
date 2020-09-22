VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.UserControl CodeBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   PropertyPages   =   "CodeBox.ctx":0000
   ScaleHeight     =   2865
   ScaleWidth      =   3615
   ToolboxBitmap   =   "CodeBox.ctx":002B
   Begin VB.PictureBox picLineNumbers 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   585
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   580
   End
   Begin RichTextLib.RichTextBox rtfCodeBox 
      Height          =   2175
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3836
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      DisableNoScroll =   -1  'True
      Appearance      =   0
      RightMargin     =   32000
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"CodeBox.ctx":033D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line linSeperator 
      X1              =   585
      X2              =   585
      Y1              =   0
      Y2              =   2640
   End
End
Attribute VB_Name = "CodeBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

' --------------------------------------------------------
' Name:     CodeBox v1.0
' Author:   Chris Thomas
' Company:  Symbiote Software
' Date      March 12, 2003
' --------------------------------------------------------
'
' All of this code is my own except for one noteable exceptions
'   cSuperClass / iSuperClass by Paul Caton (Paul_Caton@hotmail.com)
'   (see cSuperClass for more information on these two modules)
'
' Known Bugs: When a Ctrl + (letter) combination is pressed a character with a ASCII value
'               below 33 gets entered into the CodeBox. I have temporarily fixed this by
'               filtering out any ascii char below 33 (besides enter, space, backspace, delete)
'               If anyone can figure out why it is doing this please let me know

' Questions / Comments should be directed to Chris at cwthomas@connectmail.carleton.ca


'Event Declarations:
Event Change()
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event SelChange()

'Default Property Values:
Private Const DEFAULT_TABSIZE = 4
Private Const DEFAULT_LINENUMBERFORECOLOR = 0
Private Const DEFAULT_CURRENTLINENUMBERCOLOR = 0

Private Const DEFAULT_LINECOLOR = 0
Private Const DEFAULT_LINENUMBERS = True
Private Const DEFAULT_FORECOLOR = 0
Private Const DEFAULT_APPEARANCE = 1
Private Const DEFAULT_BORDERSTYLE = 1
Private Const DEFAULT_KEYWORDCOLOR = 8388608
Private Const DEFAULT_COMMENTCOLOR = 32768
Private Const DEFAULT_STRINGCOLOR = 128


'Property Variables:
Private m_LineNumbers As Boolean
Private m_Appearance As AppearanceConstants
Private m_BorderStyle As RichTextLib.BorderStyleConstants

Private m_LineColor As OLE_COLOR
Private m_LineNumberForeColor As OLE_COLOR
Private m_CurrentLineNumberColor As OLE_COLOR
Private m_TabSize As Integer
'----------------------------------------------------------------


Private m_syhHighlighter As cSyntaxHighlighter
Private m_scSubClass As cSubclass
Private m_hkHook As cHook

' Holds the hWnd of the RichTextBox
Private m_hWndCodeBox As Long

' Used to align the Line Numbers
Private m_lLineHeight As Long
Private m_lMaxVisibleLines As Long
Private m_lLineNumberX As Long

' Used to correct the caret position when selecting
Private m_lLastCaretPos As Long
Private m_lSelPivot As Long

Implements WinSubHook.iSubclass
Implements WinSubHook.iHook

'-- Public Methods -----------------------------------------------------------

' Gets the number of lines in the CodeBox
Public Function LineCount() As Long
    Dim lLines As Long
    lLines = SendMessage(m_hWndCodeBox, EM_GETLINECOUNT, 0&, 0&)
    LineCount = lLines
End Function

' Gets the current row of the caret
Public Function CurrentRow() As Long
    Dim lRow As Long
    lRow = SendMessage(m_hWndCodeBox, EM_LINEFROMCHAR, -1, 0&) + 1
    CurrentRow = lRow
End Function

' Gets the current col of the caret
Public Function CurrentCol() As Long
    Dim lCol As Long
    
    lCol = SendMessage(m_hWndCodeBox, EM_LINEINDEX, -1, 0&)
    CurrentCol = m_lLastCaretPos - lCol + 1
End Function

' Force the CodeBox to Re-Colorize
Public Sub Colorize()
    m_syhHighlighter.Colorize False, True
End Sub

' -- Private Methods / Events ------------------------------------------------------
Private Sub rtfCodeBox_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lSelStart As Long
    Dim lSelEnd As Long
    If Shift = 0 Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
            m_syhHighlighter.Colorize True
        End If
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub rtfCodeBox_KeyPress(KeyAscii As Integer)
    ' Pass the ASCII value of the key pressed to the Highlighter
    m_syhHighlighter.KeyPressed KeyAscii

    ' Raise the KeyPress event
    RaiseEvent KeyPress(KeyAscii)

    If KeyAscii = vbKeyReturn Then
        Call rtfCodeBox_SelChange
    End If

    ' Some keys need to destroyed before being sent to the CodeBox
    If g_bIgnoreNextKeyPress Then
        g_bIgnoreNextKeyPress = False
        KeyAscii = 0
    End If
    
End Sub

' Fixes the current row/col so that it will be at the leading edge of the selection
' (ie if selecting from pos 4 to 7 then the current col will be 7,
'     if selecting from pos 7 to 4 then the current col will be 4.
Private Sub rtfCodeBox_SelChange()
    Dim lRet As Long
    Dim lSelStart As Long
    Dim lSelEnd As Long
            
    If g_bNoRaiseSelChange Then Exit Sub
    lRet = SendMessage(m_hWndCodeBox, EM_GETSEL, VarPtr(lSelStart), VarPtr(lSelEnd))
    
    If lSelStart <> lSelEnd Then
        If m_lSelPivot = 0 Then
            m_lSelPivot = m_lLastCaretPos
        End If
        
        If m_lSelPivot = lSelStart Then
            m_lLastCaretPos = lSelEnd
        ElseIf m_lSelPivot = lSelEnd Then
            m_lLastCaretPos = lSelStart
        End If
    ElseIf lSelStart = lSelEnd Then
        m_lSelPivot = 0
        m_lLastCaretPos = lSelStart
    End If
    
    'Refresh the line numbers
    RefreshLineNumbers
    
    RaiseEvent SelChange
End Sub

Public Sub UpdateLastCaretPos()

End Sub

Private Sub UserControl_Initialize()
    Set m_scSubClass = New cSubclass
    Set m_hkHook = New cHook
    
    'Subclass the RichTextBox
    With m_scSubClass
        .AddMsg WM_VSCROLL, MSG_AFTER
        .AddMsg WM_REFRESH_LINE_NUMBERS, MSG_AFTER
        .AddMsg WM_PASTE, MSG_AFTER
        .AddMsg WM_CHAR, MSG_AFTER
        .AddMsg WM_CHAR, MSG_BEFORE
        .AddMsg WM_KEYDOWN, MSG_AFTER
        '.AddMsg WM_KEYUP, True
        .AddMsg WM_MOUSEWHEEL, MSG_BEFORE
        .AddMsg WM_CUT, MSG_BEFORE
        .AddMsg WM_COPY, MSG_BEFORE
        m_hWndCodeBox = rtfCodeBox.hWnd
        .Subclass m_hWndCodeBox, Me
    End With
    
    With m_hkHook
        .Hook WH_KEYBOARD, Me
    End With
    
    If Not m_syhHighlighter Is Nothing Then
        Set m_syhHighlighter = Nothing
    End If
    
    Set m_syhHighlighter = New cSyntaxHighlighter
    m_syhHighlighter.CodeBoxWnd = m_hWndCodeBox
    
End Sub

Private Sub UserControl_Terminate()
    'Unsubclass the CodeBox
    m_scSubClass.UnSubclass
    
    'Unhook the CodeBox
    m_hkHook.UnHook
    
    'Release the Subclass Class from memory
    Set m_scSubClass = Nothing
    Set m_hkHook = Nothing
    Set m_syhHighlighter = Nothing
End Sub

Private Sub UserControl_Resize()
    'Refresh controls appearance
    SetupControl
End Sub

Private Sub InitLineNumbers()
    Dim rTextArea As RECT
    Dim lRet As Long
    Dim i As Integer
    
    If m_LineNumbers Then
        'The PictureBox font has to be the same as the TextBox Font to get the
        'height of one line of text
        Set picLineNumbers.Font = rtfCodeBox.Font
        
        'Determine the height of one line of text
        m_lLineHeight = picLineNumbers.TextHeight("0")
        
        'Get the Text Area for the TextBox
        lRet = SendMessage(m_hWndCodeBox, EM_GETRECT, 0&, VarPtr(rTextArea))
        
        'Get the maximum number of lines that will fit in the current TextBox
        m_lMaxVisibleLines = (rTextArea.Bottom * TWIPS_PER_PIXEL) \ m_lLineHeight + 1
        
        'Center the LineNumbers
        m_lLineNumberX = (picLineNumbers.Width - picLineNumbers.TextWidth("0000")) / 2
        
        'Are in Run-Time?
        If UserControl.Ambient.UserMode Then
            'Refresh the LineNumbers
            RefreshLineNumbers
        Else
            'We are in Design-Time and the TextBox is not subclassed
            'So we can not use the RefreshLineNumbers Method
            'Just draw some numbers to people can get a glimpse of how
            'the control will look at runtime
            With picLineNumbers
                'Set the first line to the CurrentLineNumberColor
                .Cls
                .ForeColor = m_CurrentLineNumberColor
                .CurrentY = 0
                Do
                    i = i + 1
                    .CurrentX = m_lLineNumberX
                    'Print LineNumber
                    picLineNumbers.Print Format(i, "0000")
                    'Reset color to the normal LineNumber color
                    .ForeColor = m_LineNumberForeColor
                Loop Until .CurrentY > .Height
            End With
        End If
    End If
End Sub

Private Sub SetupControl()
    Dim rTextArea As RECT
    Dim i As Integer
    
    'Errors.... Who needs em..
    On Error Resume Next
    'Eliminates flickers
    LockWindowUpdate UserControl.hWnd
    
    'Set height of rtfCodeBox to that of the control
    rtfCodeBox.Height = UserControl.Height - 60
    linSeperator.Y2 = UserControl.Height
    
    'Do are the line numbers visible?
    If m_LineNumbers Then
        'Move and size controls into place
        rtfCodeBox.Left = 595
        rtfCodeBox.Width = UserControl.Width - 655
                        
        'Get the Text Area of the TextBox
        SendMessage rtfCodeBox.hWnd, EM_GETRECT, 0&, VarPtr(rTextArea)
        
        'Resize LineNumbers Pic Box
        picLineNumbers.Top = rTextArea.Top * TWIPS_PER_PIXEL
        picLineNumbers.Height = rTextArea.Bottom * TWIPS_PER_PIXEL
        
        'Show the LineNumber Controls
        linSeperator.Visible = True
        picLineNumbers.Visible = True
        
        'Initialize and draw the line numbers
        InitLineNumbers
    Else
        'Move and resize the TextBox to fill the UserControl
        rtfCodeBox.Left = 0
        rtfCodeBox.Width = UserControl.Width - rtfCodeBox.Left - 60
        
        'Hide the LineNumbers
        linSeperator.Visible = False
        picLineNumbers.Visible = False
    End If
    LockWindowUpdate 0
End Sub

Private Sub RefreshLineNumbers()
    'Redraws the line numbers
    SendMessage m_hWndCodeBox, WM_REFRESH_LINE_NUMBERS, 0&, 0&
End Sub

' Window Procs ----------------------------------------------------------
Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hWnd As Long, uMsg As WinSubHook.eMsg, wParam As Long, lParam As Long)
    Select Case uMsg
    Case WM_MOUSEWHEEL
        lReturn = 0
        bHandled = True
    Case WM_CHAR
        Select Case wParam
        Case 22         ' Paste
            m_syhHighlighter.Colorize False, True
            Call rtfCodeBox_SelChange
            lReturn = 1
            bHandled = True
        Case Else
            ' Handles cases where the control key is pressed and eliminates
            ' any extra characters that sometimes get generated.
            If GetAsyncKeyState(vbKeyControl) <> 0 Then
                lReturn = 1
                bHandled = True
                Exit Sub
            End If
            bHandled = False
            lReturn = 0
        End Select
    End Select
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hWnd As Long, ByVal uMsg As WinSubHook.eMsg, ByVal wParam As Long, ByVal lParam As Long)
    Dim lFirstVisibleLine As Long
    Dim lFirstChar As Long
    Dim lCurrentLine As Long
    Dim lSelStart As Long
    Dim lSelEnd As Long
    
    Dim lRet As Long
    Dim i As Integer
    
    Select Case uMsg
    Case WM_PASTE
        m_syhHighlighter.Colorize False, True
        Call rtfCodeBox_SelChange
        lReturn = 0
    Case WM_KEYDOWN
        ' Shift + Insert = Paste
        If wParam = vbKeyInsert And GetAsyncKeyState(vbKeyShift) <> 0 Then
            m_syhHighlighter.Colorize False, True
            Call rtfCodeBox_SelChange
            lReturn = 0
        End If
    Case WM_VSCROLL, WM_REFRESH_LINE_NUMBERS
        'Clear the Line Number PictureBox
        picLineNumbers.Cls

        'Get the index of the line that contains the caret
        lCurrentLine = SendMessage(hWnd, EM_LINEFROMCHAR, -1, 0&)
        
        'Get the line index of the first visible line
        lFirstVisibleLine = SendMessage(hWnd, EM_GETFIRSTVISIBLELINE, 0&, 0&)

        'get the position of the first character in the line
        lFirstChar = SendMessage(hWnd, EM_LINEINDEX, lFirstVisibleLine, 0&)

        'Get the location of that character relative to the TextBox
        lRet = SendMessage(hWnd, EM_POSFROMCHAR, lFirstChar, ByVal 0&)

        'Set where to start writing the line numbers
        picLineNumbers.CurrentY = ((HiWord(lRet) * TWIPS_PER_PIXEL) Mod m_lLineHeight)

        'Prints the Line numbers on the picturebox
        picLineNumbers.ForeColor = m_LineNumberForeColor
        For i = 1 To m_lMaxVisibleLines + 1
            picLineNumbers.CurrentX = m_lLineNumberX
            
            If lFirstVisibleLine + i = lCurrentLine + 1 Then
                picLineNumbers.ForeColor = m_CurrentLineNumberColor
                picLineNumbers.Print Format(lFirstVisibleLine + i, "0000")
                picLineNumbers.ForeColor = m_LineNumberForeColor
            Else
                picLineNumbers.Print Format(lFirstVisibleLine + i, "0000")
            End If
        Next i
        picLineNumbers.Refresh
        lReturn = 0
    End Select

End Sub

'Hook Procedures
Private Sub iHook_After(lReturn As Long, ByVal nCode As WinSubHook.eHookCode, ByVal wParam As Long, ByVal lParam As Long)

End Sub

Private Sub iHook_Before(bHandled As Boolean, lReturn As Long, nCode As WinSubHook.eHookCode, wParam As Long, lParam As Long)
    Dim lCol As Long
    Dim lNumSpaces As Long
    If nCode = HC_ACTION Then
        If GetFocus = m_hWndCodeBox Then
            If wParam = vbKeyTab And ((lParam And &H80000000) <> &H80000000) Then
                lCol = Me.CurrentCol - 1
                lNumSpaces = m_TabSize - (lCol Mod m_TabSize)
                SetText m_hWndCodeBox, SF_TEXT, SFF_SELECTION, Space(lNumSpaces)
                lReturn = 1
                bHandled = True
            End If
        End If
    End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Appearance = DEFAULT_APPEARANCE
    m_BorderStyle = DEFAULT_BORDERSTYLE
    m_LineNumbers = DEFAULT_LINENUMBERS
    m_LineColor = DEFAULT_LINECOLOR
    m_CurrentLineNumberColor = DEFAULT_CURRENTLINENUMBERCOLOR
    m_LineNumberForeColor = DEFAULT_LINENUMBERFORECOLOR

    Set m_syhHighlighter.Font = rtfCodeBox.Font
    m_TabSize = DEFAULT_TABSIZE
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    rtfCodeBox.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    picLineNumbers.BackColor = PropBag.ReadProperty("LineNumberBackColor", &H8000000F)
    UserControl.BackColor = picLineNumbers.BackColor
    
    Set rtfCodeBox.Font = PropBag.ReadProperty("Font", Font)
    Set picLineNumbers.Font = rtfCodeBox.Font
    
    rtfCodeBox.AutoVerbMenu = PropBag.ReadProperty("AutoVerbMenu", False)
    rtfCodeBox.HideSelection = PropBag.ReadProperty("HideSelection", True)
    rtfCodeBox.Enabled = PropBag.ReadProperty("Enabled", True)
    
    rtfCodeBox.Locked = PropBag.ReadProperty("Locked", False)
    
    rtfCodeBox.Text = PropBag.ReadProperty("Text", "")
      
    '-----------------------------------
    m_LineNumbers = PropBag.ReadProperty("LineNumbers", DEFAULT_LINENUMBERS)
    m_CurrentLineNumberColor = PropBag.ReadProperty("CurrentLineNumberColor", DEFAULT_CURRENTLINENUMBERCOLOR)
    m_LineNumberForeColor = PropBag.ReadProperty("LineNumberForeColor", DEFAULT_LINENUMBERFORECOLOR)
    m_LineColor = PropBag.ReadProperty("LineColor", DEFAULT_LINECOLOR)
    m_Appearance = PropBag.ReadProperty("Appearance", DEFAULT_APPEARANCE)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", DEFAULT_BORDERSTYLE)
    m_TabSize = PropBag.ReadProperty("TabSize", DEFAULT_TABSIZE)
    
    With m_syhHighlighter
        Set .Font = rtfCodeBox.Font
        .KeywordColor = PropBag.ReadProperty("KeywordColor", DEFAULT_KEYWORDCOLOR)
        .CommentColor = PropBag.ReadProperty("CommentColor", DEFAULT_COMMENTCOLOR)
        .StringColor = PropBag.ReadProperty("StringColor", DEFAULT_STRINGCOLOR)
        .DefaultColor = PropBag.ReadProperty("ForeColor", DEFAULT_FORECOLOR)
    End With
    '-------------------------------
    SetupControl
    
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", rtfCodeBox.BackColor, &H80000005)
    Call PropBag.WriteProperty("Enabled", rtfCodeBox.Enabled, True)
    
    Call PropBag.WriteProperty("Text", rtfCodeBox.Text, "")
    Call PropBag.WriteProperty("Font", rtfCodeBox.Font, Ambient.Font)
    
    
    Call PropBag.WriteProperty("Appearance", m_Appearance, DEFAULT_APPEARANCE)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, DEFAULT_BORDERSTYLE)
    Call PropBag.WriteProperty("AutoVerbMenu", rtfCodeBox.AutoVerbMenu, False)
    Call PropBag.WriteProperty("HideSelection", rtfCodeBox.HideSelection, True)
    Call PropBag.WriteProperty("Locked", rtfCodeBox.Locked, False)
    
    Call PropBag.WriteProperty("LineNumbers", m_LineNumbers, DEFAULT_LINENUMBERS)
    Call PropBag.WriteProperty("LineColor", m_LineColor, DEFAULT_LINECOLOR)
        
    Call PropBag.WriteProperty("LineNumberBackColor", picLineNumbers.BackColor, &H8000000F)
    Call PropBag.WriteProperty("CurrentLineNumberColor", m_CurrentLineNumberColor, DEFAULT_CURRENTLINENUMBERCOLOR)
    Call PropBag.WriteProperty("LineNumberForeColor", m_LineNumberForeColor, DEFAULT_LINENUMBERFORECOLOR)
    
    With m_syhHighlighter
        Call PropBag.WriteProperty("KeywordColor", .KeywordColor, DEFAULT_KEYWORDCOLOR)
        Call PropBag.WriteProperty("CommentColor", .CommentColor, DEFAULT_COMMENTCOLOR)
        Call PropBag.WriteProperty("StringColor", .StringColor, DEFAULT_STRINGCOLOR)
        Call PropBag.WriteProperty("ForeColor", .DefaultColor, DEFAULT_FORECOLOR)
    End With
    Call PropBag.WriteProperty("TabSize", m_TabSize, DEFAULT_TABSIZE)
End Sub

'
'
'
'' PROPERTIES -----------------------------------------------------------------------------
'
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Let BlockCommentStartChar(sStart As String)
Attribute BlockCommentStartChar.VB_MemberFlags = "400"
    m_syhHighlighter.BlockCommentStartChar = sStart
End Property

Public Property Get BlockCommentStartChar() As String
    BlockCommentStartChar = m_syhHighlighter.BlockCommentStartChar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Let BlockCommentEndChar(sEnd As String)
Attribute BlockCommentEndChar.VB_MemberFlags = "400"
    m_syhHighlighter.BlockCommentEndChar = sEnd
End Property

Public Property Get BlockCommentEndChar() As String
    BlockCommentEndChar = m_syhHighlighter.BlockCommentEndChar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Let LineCommentChar(sChar As String)
Attribute LineCommentChar.VB_MemberFlags = "400"
    m_syhHighlighter.LineCommentChar = sChar
End Property

Public Property Get LineCommentChar() As String
    LineCommentChar = m_syhHighlighter.LineCommentChar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Let StringChar(sString As String)
Attribute StringChar.VB_MemberFlags = "400"
    m_syhHighlighter.StringChar = sString
End Property

Public Property Get StringChar() As String
    StringChar = m_syhHighlighter.StringChar
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtfCodeBox,rtfCodeBox,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = rtfCodeBox.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    rtfCodeBox.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtfCodeBox,rtfCodeBox,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "ppgOther"
    Enabled = rtfCodeBox.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    rtfCodeBox.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtfCodeBox,rtfCodeBox,-1,Font
Public Property Get Font() As Font
    Set Font = rtfCodeBox.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set rtfCodeBox.Font = New_Font
    Set m_syhHighlighter.Font = New_Font
    SetupControl
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtfCodeBox,rtfCodeBox,-1,Refresh
Public Sub Refresh()
    rtfCodeBox.Refresh
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtfCodeBox,rtfCodeBox,-1,AutoVerbMenu
Public Property Get AutoVerbMenu() As Boolean
Attribute AutoVerbMenu.VB_ProcData.VB_Invoke_Property = "ppgOther"
    AutoVerbMenu = rtfCodeBox.AutoVerbMenu
End Property

Public Property Let AutoVerbMenu(ByVal New_AutoVerbMenu As Boolean)
    rtfCodeBox.AutoVerbMenu() = New_AutoVerbMenu
    PropertyChanged "AutoVerbMenu"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtfCodeBox,rtfCodeBox,-1,HideSelection
Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_ProcData.VB_Invoke_Property = "ppgOther"
    HideSelection = rtfCodeBox.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    rtfCodeBox.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,8388608
Public Property Get KeywordColor() As OLE_COLOR
    KeywordColor = m_syhHighlighter.KeywordColor
End Property

Public Property Let KeywordColor(ByVal New_KeywordColor As OLE_COLOR)
    m_syhHighlighter.KeywordColor = New_KeywordColor
    PropertyChanged "KeywordColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,32768
Public Property Get CommentColor() As OLE_COLOR
    CommentColor = m_syhHighlighter.CommentColor
End Property

Public Property Let CommentColor(ByVal New_CommentColor As OLE_COLOR)
    m_syhHighlighter.CommentColor = New_CommentColor
    PropertyChanged "CommentColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,128
Public Property Get StringColor() As OLE_COLOR
    StringColor = m_syhHighlighter.StringColor
End Property

Public Property Let StringColor(ByVal New_StringColor As OLE_COLOR)
    m_syhHighlighter.StringColor = New_StringColor
    PropertyChanged "StringColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,3,0,1
Public Property Get Appearance() As AppearanceConstants
    If Ambient.UserMode Then Err.Raise 393
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    If Ambient.UserMode Then Err.Raise 382
    m_Appearance = New_Appearance
    UserControl.Appearance = m_Appearance
    UserControl.BackColor = picLineNumbers.BackColor
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=30,3,0,1
Public Property Get BorderStyle() As RichTextLib.BorderStyleConstants
    If Ambient.UserMode Then Err.Raise 393
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As RichTextLib.BorderStyleConstants)
    If Ambient.UserMode Then Err.Raise 382
    m_BorderStyle = New_BorderStyle
    UserControl.BorderStyle = m_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_syhHighlighter.DefaultColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_syhHighlighter.DefaultColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtfCodeBox,rtfCodeBox,-1,Text
Public Property Get Text() As String
Attribute Text.VB_ProcData.VB_Invoke_Property = "ppgOther"
    Text = rtfCodeBox.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    rtfCodeBox.Text() = New_Text
    If UserControl.Ambient.UserMode Then
        If Not m_syhHighlighter Is Nothing Then
            m_syhHighlighter.Colorize False, True
        End If
    End If
    PropertyChanged "Text"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get LineNumbers() As Boolean
Attribute LineNumbers.VB_ProcData.VB_Invoke_Property = "ppgOther"
    LineNumbers = m_LineNumbers
End Property

Public Property Let LineNumbers(ByVal New_LineNumbers As Boolean)
    m_LineNumbers = New_LineNumbers
    SetupControl
    PropertyChanged "LineNumbers"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get LineColor() As OLE_COLOR
    LineColor = m_LineColor
End Property

Public Property Let LineColor(ByVal New_LineColor As OLE_COLOR)
    m_LineColor = New_LineColor
    linSeperator.BorderColor = New_LineColor
    PropertyChanged "LineColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picLineNumbers,picLineNumbers,-1,BackColor
Public Property Get LineNumberBackColor() As OLE_COLOR
    LineNumberBackColor = picLineNumbers.BackColor
End Property

Public Property Let LineNumberBackColor(ByVal New_LineNumberBackColor As OLE_COLOR)
    picLineNumbers.BackColor() = New_LineNumberBackColor
    UserControl.BackColor = picLineNumbers.BackColor
    SetupControl
    PropertyChanged "LineNumberBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CurrentLineNumberColor() As OLE_COLOR
Attribute CurrentLineNumberColor.VB_Description = "The color of the line number where the caret is currenly located"
    CurrentLineNumberColor = m_CurrentLineNumberColor
End Property

Public Property Let CurrentLineNumberColor(ByVal New_CurrentLineNumberColor As OLE_COLOR)
    m_CurrentLineNumberColor = New_CurrentLineNumberColor
    SetupControl
    PropertyChanged "CurrentLineNumberColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get LineNumberForeColor() As OLE_COLOR
    LineNumberForeColor = m_LineNumberForeColor
End Property

Public Property Let LineNumberForeColor(ByVal New_LineNumberForeColor As OLE_COLOR)
    m_LineNumberForeColor = New_LineNumberForeColor
    SetupControl
    PropertyChanged "LineNumberForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=rtfCodeBox,rtfCodeBox,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_ProcData.VB_Invoke_Property = "ppgOther"
    Locked = rtfCodeBox.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    rtfCodeBox.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,2,
Public Property Let Keywords(sKeywords As String)
Attribute Keywords.VB_MemberFlags = "400"
    m_syhHighlighter.Keywords = sKeywords
    PropertyChanged "Keywords"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,4
Public Property Get TabSize() As Integer
    TabSize = m_TabSize
End Property

Public Property Let TabSize(ByVal New_TabSize As Integer)
    m_TabSize = New_TabSize
    PropertyChanged "TabSize"
End Property

Private Sub rtfCodeBox_Change()
    RaiseEvent Change
End Sub

Private Sub rtfCodeBox_Click()
    RaiseEvent Click
End Sub

Private Sub rtfCodeBox_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub rtfCodeBox_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub rtfCodeBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub rtfCodeBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub rtfCodeBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub


