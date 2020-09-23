VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strCharAllowed As String
Private blnAlpha As Boolean
Private blnNumeric As Boolean
Private intMaxAllowed As Integer

Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    If intKeyAscii = vbKeyBack Then
        If Len(strCurrentText) > 1 Then Mask = Left(strCurrentText, Len(strCurrentText) - 1)
    ElseIf (Len(strCurrentText) > intMaxAllowed - 1) And (intMaxAllowed <> 0) Then
        Mask = strCurrentText
    ElseIf (Not IsAlpha(intKeyAscii)) And (Not (IsNumber(intKeyAscii))) And (Not (IsAllowed(intKeyAscii))) Then
        Mask = strCurrentText
    Else
        Mask = strCurrentText & Chr(intKeyAscii)
    End If
    intKeyAscii = 0
End Function

Private Function IsAlpha(intAscii As Integer) As Boolean
    Dim intTemp As Integer
    
    intTemp = Asc(UCase(Chr(intAscii)))
    IsAlpha = (blnAlpha And Not (intTemp < 65 Or intTemp > 90))
End Function

Private Function IsNumber(intAscii As Integer) As Boolean
    IsNumber = blnNumeric And (intAscii > 47 And intAscii < 58)
End Function

Private Function IsAllowed(intAscii) As Boolean
    IsAllowed = (InStr(1, strCharAllowed, Chr(intAscii)) > 0)
End Function

Public Property Get AllowAlpha() As Boolean
    AllowAlpha = blnAlpha
End Property

Public Property Let AllowAlpha(ByVal Value As Boolean)
    blnAlpha = Not Value
    SaveSetting "CUSTOM", "AllowAlpha", Value
End Property

Public Property Get AllowNumber() As Boolean
    AllowNumber = blnNumeric
End Property

Public Property Let AllowNumber(ByVal Value As Boolean)
    blnNumeric = Not Value
    SaveSetting "CUSTOM", "AllowNum", Value
End Property

Public Property Get CharAllowed() As String
    CharAllowed = strCharAllowed
End Property

Public Property Let CharAllowed(ByVal Value As String)
    strCharAllowed = Value
    SaveSetting "CUSTOM", "CharAllowed", Value
End Property

Public Property Get MaxAllowed() As Integer
    MaxAllowed = intMaxAllowed
End Property

Public Property Let MaxAllowed(ByVal Value As Integer)
    intMaxAllowed = Value
    SaveSetting "CUSTOM", "MaxAllowed", Value
End Property

Private Sub Class_Initialize()
    blnAlpha = (GetSetting("CUSTOM", "AllowAlpha") = "True")
    blnNumeric = (GetSetting("CUSTOM", "AllowNum") = "True")
    strCharAllowed = GetSetting("CUSTOM", "CharAllowed")
    intMaxAllowed = Val(GetSetting("CUSTOM", "MaxAllowed"))
End Sub
