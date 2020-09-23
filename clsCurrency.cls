VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private blnSign As Boolean

'properties used to store and retrieve
'our control's settings
Public Property Get UseSign() As Boolean
    UseSign = blnSign
End Property

Public Property Let UseSign(ByVal Value As Boolean)
    blnSign = Value
    SaveSetting "CURRENCY", "UseSign", Value
End Property

'when this object is created, we want to retrieve
'any stored data related previously stored
'related to the selected mask type
Private Sub Class_Initialize()
    If Len(GetSetting("CURRENCY", "UseSign")) = 0 Then
        UseSign = True
    Else
        blnSign = (GetSetting("CURRENCY", "UseSign") = "True")
    End If
End Sub

'here is the heart of our Currency masking funciton
Public Function Mask(ByRef intKeyAscii As Integer, ByVal strCurrentText As String) As String

    'if the user typed the backspace key the handle it
    If intKeyAscii = vbKeyBack Then
        If Len(strCurrentText) > (1 + Abs(blnSign) + InStr(1, strCurrentText, "-")) Then
            Mask = Left(strCurrentText, Len(strCurrentText) - 1)
        End If
    
    'otherwise, if the ascii value passed is not a number,
    'decimal or negitive sign then we dont want it
    ElseIf (intKeyAscii < 48 Or intKeyAscii > 57) And intKeyAscii <> 46 And intKeyAscii <> 45 Then
        Mask = strCurrentText
    
    'the user must have typed a number, decimal, or negative sign
    Else
        'the negative sign was pressed
        If intKeyAscii = 45 Then
        
            'remove - sign since it already exists
            If InStr(1, strCurrentText, "-") > 0 Then
                Mask = Mid(strCurrentText, 2)
                
            'no - sign exists so add one to the front
            Else
                Mask = Chr(intKeyAscii) & strCurrentText
            End If
        
        'the decimal was pressed
        ElseIf intKeyAscii = 46 Then
        
            'here we check to see if the user typed a decimal
            'before entering a leading number.  If so, we add
            'a leading zero (0) for them....we're so kind
            If Len(strCurrentText) = (0 + (Abs(Left(strCurrentText, 1) = "-"))) Then
                If blnSign Then
                    Mask = strCurrentText & "$ 0."
                Else
                    Mask = strCurrentText & "0."
                End If
                
            'as long as the decimal character doesn't already exist,
            'then we go ahead and add it to the end
            ElseIf InStr(1, strCurrentText, ".") = 0 Then
                Mask = strCurrentText & "."
            Else
                Mask = strCurrentText
            End If
            
        'a number character must have been pressed by
        'the user
        Else
        
            'if this is the first number pressed, then add the $ sign
            'if the user has it requested in the properties
            If Len(strCurrentText) = (0 + (Abs(Left(strCurrentText, 1) = "-"))) Then
                If blnSign Then
                    Mask = strCurrentText & "$ " & Chr(intKeyAscii)
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)
                End If
            ElseIf InStr(1, strCurrentText, ".") = 0 Then
                Mask = strCurrentText & Chr(intKeyAscii)
            'if a decimal char already exists, then we only want
            'to allow two numbers to follow.  If two numbers
            'already exist, then we wont accept any more
            Else
                If Len(strCurrentText) - InStr(1, strCurrentText, ".") = 2 Then
                    Mask = strCurrentText
                Else
                    Mask = strCurrentText & Chr(intKeyAscii)
                End If
            End If
        End If
    End If
    intKeyAscii = 0
End Function
