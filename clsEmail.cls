VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String
    If intKeyAscii = vbKeyBack Then
        If Len(strCurrentText) > 1 Then
            Mask = Left(strCurrentText, Len(strCurrentText) - 1)
        End If
    ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) And ((Asc(UCase(Chr(intKeyAscii))) < 65) Or (Asc(UCase(Chr(intKeyAscii))) > 90)) And intKeyAscii <> 46 And intKeyAscii <> 64 And intKeyAscii <> 95 Then
        Mask = strCurrentText
    Else
        If (intKeyAscii = 46) And (InStr(1, strCurrentText, "@") = 0) Then
            Mask = strCurrentText
        ElseIf (intKeyAscii = 64) And (InStr(1, strCurrentText, "@") > 0) Then
            Mask = strCurrentText
        ElseIf (intKeyAscii = 46) And (InStr(1, strCurrentText, ".") > 0) Then
            Mask = strCurrentText
        Else
            Mask = strCurrentText & Chr(intKeyAscii)
        End If
    End If
    intKeyAscii = 0
End Function
