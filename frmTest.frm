VERSION 5.00
Object = "*\AprjCodeBox.vbp"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Symbiote Software - CodeBox v0.8"
   ClientHeight    =   5190
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "frmTest.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6975
   StartUpPosition =   3  'Windows Default
   Begin prjCodeBox.CodeBox cbxSample 
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoVerbMenu    =   -1  'True
      CurrentLineNumberColor=   255
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   $"frmTest.frx":058A
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   6735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Line Count"
      Height          =   195
      Left            =   3780
      TabIndex        =   5
      Top             =   3840
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Col:"
      Height          =   195
      Left            =   2925
      TabIndex        =   4
      Top             =   3840
      Width           =   270
   End
   Begin VB.Label lblLines 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4620
      TabIndex        =   3
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblCol 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3300
      TabIndex        =   2
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label lblRow 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2460
      TabIndex        =   1
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Row:"
      Height          =   195
      Left            =   1980
      TabIndex        =   0
      Top             =   3840
      Width           =   375
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const m_sVBKeywords As String = "|#Const|#Else|#ElseIf|#End If|#If|Alias|Alias|And|As|Attribute|Base|Begin|Binary|Boolean|Byte|ByVal|Call|Case|CBool|CByte|CCur|CDate|CDbl|CDec|CInt|CLng|Close|Compare|Const|CSng|CStr|Currency|CVar|CVErr|Debug|Decimal|Declare|DefBool|DefByte|DefCur|DefDate|DefDbl|DefDec|DefInt|DefLng|DefObj|DefSng|DefStr|DefVar|Dim|Do|Double|Each|Else|ElseIf|End|Enum|Eqv|Erase|Error|Exit|Event|Explicit|False|For|Friend|Function|Get|Global|GoSub|GoTo|If|Imp|In|Input|Input|Integer|Is|LBound|Lcase|Let|Lib|Like|Line|Lock|Long|Loop|LSet|Name|New|Next|Not|Nothing|Object|On|Open|Option|Optional|Or|Output|Print|Private|Property|Public|Put|Random|Read|ReDim|Replace|Resume|Return|RSet|Seek|Select|Set|Single|Spc|Static|Stop|String|Sub|Tab|Then|Then|To|True|Type|UBound|Ucase|Unlock|Variant|Version|Wend|While|With|Xor|"

Private Sub cbxSample_Change()
    lblLines = cbxSample.LineCount
End Sub

Private Sub cbxSample_SelChange()
    lblRow = cbxSample.CurrentRow
    lblCol = cbxSample.CurrentCol
End Sub

Private Sub Form_Load()
    With cbxSample
        .Keywords = m_sVBKeywords
        .LineCommentChar = "'"
        .BlockCommentStartChar = ""
        .BlockCommentEndChar = ""
        .StringChar = """"
    End With
End Sub

