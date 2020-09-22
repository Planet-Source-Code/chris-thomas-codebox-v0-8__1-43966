VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStringConcatenator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --------------------------------------------------------
' Name:     CodeBox v0.8
' Author:   Chris Thomas
' Company:  Symbiote Software
' Date      March 12, 2003
'
' Purpose:  To quickly append stings by converting them to
'           arrays and using CopyMemory to merge it into
'           one array
' --------------------------------------------------------
'
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal cb&)

' Stores the String
Private m_bytArray() As Byte

' Current length of the string
Private m_lLength As Long

' Clears the string
Public Sub Clear()
    Erase m_bytArray
    m_lLength = 0
End Sub

' Appends the given string.
Public Sub Append(sNewString As String)
    Dim lLengthOfNewString As Long
    Dim bytNewString() As Byte
    
    If sNewString = "" Then Exit Sub
    lLengthOfNewString = Len(sNewString)
    ReDim bytNewString(0 To lLengthOfNewString) As Byte
    bytNewString = StrConv(sNewString, vbFromUnicode)
    
    ReDim Preserve m_bytArray(0 To m_lLength + lLengthOfNewString - 1) As Byte
    CopyMemory m_bytArray(m_lLength), bytNewString(0), lLengthOfNewString
    m_lLength = m_lLength + lLengthOfNewString
End Sub

Public Property Get Length()
    Length = m_lLength
End Property

Public Property Get Value() As String
    Dim sRet As String
    sRet = StrConv(m_bytArray, vbUnicode)
    Value = sRet
End Property

Private Sub Class_Initialize()
    m_lLength = 0
End Sub

Private Sub Class_Terminate()
    Clear
End Sub
