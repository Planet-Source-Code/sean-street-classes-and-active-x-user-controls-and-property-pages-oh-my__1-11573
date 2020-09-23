Attribute VB_Name = "modProcedures"
Option Explicit

'this API retrieves INI data
Declare Function GetPrivateProfileStringbyKeyName& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName$, ByVal lpszKey$, ByVal lpszDefault$, ByVal lpszreturnbuffer$, ByVal cchreturnbuffer&, ByVal lpszFile$)

'this API sets the INI data
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String)

'we use this function to return the path of
'our parent application
Public Declare Function GetCurrentDirectoryA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'this variable stores the INI path for
'our control
Public strCTRLName As String

'the StrVersion variable stores version info
'for the about screen
Public strVersion As String

'this function calls the Get INI API to retrieve INI info
Public Function GetSetting(strPropertyPage As String, strProperty As String) As String
    Dim lCharacters As Long
    Dim strTemp As String
    
    strTemp = String$(128, 0)
    lCharacters = GetPrivateProfileStringbyKeyName(strPropertyPage, strProperty, "", strTemp, 127, strCTRLName)
        
    GetSetting = Left$(strTemp, lCharacters)
End Function

'this is our function to save data to the INI file(s)
Public Sub SaveSetting(strPropertyPage As String, strProperty As String, ByVal strValue As String)
    WritePrivateProfileString strPropertyPage, strProperty, strValue, strCTRLName
End Sub
