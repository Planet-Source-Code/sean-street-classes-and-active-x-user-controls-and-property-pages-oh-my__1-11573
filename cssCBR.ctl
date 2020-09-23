VERSION 5.00
Begin VB.UserControl cssCBR 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   PropertyPages   =   "cssCBR.ctx":0000
   ScaleHeight     =   345
   ScaleWidth      =   1260
   ToolboxBitmap   =   "cssCBR.ctx":0060
   Begin VB.TextBox txtCBR 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "cssCBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim intMaskType As Integer
Public objDate As clsDate
'Public objTime As clsTime  '-----------------------removed from project
Public objPhone As clsPhone
Public objSSN As clsSSN
Public objZip As clsZip
Public objCustom As clsCustom
Public objCurrency As clsCurrency
Public objEmail As clsEmail

Public Enum MaskType    'set possiblilities for mask types
    DateMask = 0
    PhoneMask = 2
    SSNMask = 3
    ZipMask = 4
    CurrencyMask = 5
    EmailMask = 6
    CustomMask = 7
End Enum

'Once a key has been pressed while in the text box, catch the ascii
'value and pass it to the appropriate class object, determined by
'the predefined mask type variable
Private Sub txtCBR_KeyPress(KeyAscii As Integer)

    'if return is pressed, simply apply focus to the
    'next enabled & visible control and null the ascii
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        SendKeys "{TAB}"
        
    'otherwise handle the ascii character accordingly
    Else
    
        'this line erases the text if any was highlighted before
        'applying the ascii validations
        If txtCBR.SelLength > 0 Then txtCBR.Text = Left(txtCBR.Text, txtCBR.SelStart)
        
        Select Case intMaskType
            Case 0  'date
                txtCBR.Text = objDate.Mask(KeyAscii, txtCBR.Text)
            Case 1  'time-------------------------removed from project
                'txtCBR.Text = objTime.Mask(KeyAscii, txtCBR.Text)
            Case 2  'phone
                txtCBR.Text = objPhone.Mask(KeyAscii, txtCBR.Text)
            Case 3  'ssn
                txtCBR.Text = objSSN.Mask(KeyAscii, txtCBR.Text)
            Case 4  'zip
                txtCBR.Text = objZip.Mask(KeyAscii, txtCBR.Text)
            Case 5  'currency
                txtCBR.Text = objCurrency.Mask(KeyAscii, txtCBR.Text)
            Case 6  'email
                txtCBR.Text = objEmail.Mask(KeyAscii, txtCBR.Text)
            Case 7  'custom
                txtCBR.Text = objCustom.Mask(KeyAscii, txtCBR.Text)
        End Select
        
        'set the caret to the last postition in the textbox
        txtCBR.SelStart = Len(txtCBR.Text)
    End If
End Sub


'this function is called if the user changes the name of the control
'it retrives the location of the parent application path, then stores
'the current path and name of the control in a variable called strTemp
'then sets the variable strCRTLName to the new path.  Next, it destroys
'any pre-existing files with the same name before it renames the existing
'file to the appropriate name
Private Sub UserControl_AmbientChanged(PropertyName As String)
    Dim strTemp As String
    If (PropertyName = "DisplayName") And (Len(Dir(strCTRLName)) > 0) Then
        strTemp = strCTRLName
        SetINILoc
        If Len(Dir(strCTRLName)) > 0 Then Kill strCTRLName
        Name strTemp As strCTRLName
    End If
End Sub

'here we store the version data in a variable called strVersion
'then we retrieve the all important initialization data...see the
'GetGeneralSettings function for more info
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    strVersion = App.Major & "." & App.Revision & "." & App.Minor
    GetGeneralSettings
End Sub

'here we resize the text box if the user resizes the control
Private Sub UserControl_Resize()
    txtCBR.Width = UserControl.Width
    txtCBR.Height = UserControl.Height
End Sub

'if the user changes the text property of the control
'we change our property accordingly
Public Property Get Text() As String
    Text = txtCBR.Text
End Property

Public Property Let Text(ByVal Value As String)
    txtCBR.Text = Value
End Property

'properties to manage the MaskType variable
Public Property Get MaskType() As MaskType
Attribute MaskType.VB_ProcData.VB_Invoke_Property = "pagGeneral"
    MaskType = intMaskType
End Property

Public Property Let MaskType(ByVal Value As MaskType)
    intMaskType = Value
    SaveSetting "GENERAL", "MaskType", Value
    SetObject
End Property

'This function retrieves the location of the parent application path
'and stores the control's INI file in the same path so that all data
'related to a project is stored under the same directory
Private Sub SetINILoc()
    Dim strTemp As String * 255
    Dim lngTemp As Long
    Dim lngLen As Long
    lngTemp = 255
    lngLen = GetCurrentDirectoryA(lngTemp, strTemp)
    strCTRLName = Left(strTemp, lngLen) & "\" & UserControl.Ambient.DisplayName & ".INI"
    strCTRLName = Replace(strCTRLName, "(", "")
    strCTRLName = Replace(strCTRLName, ")", "")
End Sub

'once the MaskType has been determined, we set the appropriate
'object to prepare it to handle the ascii data.  The reason we destroy
'all objects first, is to save as much memory as possible.  If the user
'switches from one MaskType to another, we wont accumalate a
'stack of unused objects.....just the one we need
Private Sub SetObject()
    Set objDate = Nothing
    'Set objTime = Nothing
    Set objPhone = Nothing
    Set objSSN = Nothing
    Set objZip = Nothing
    Set objCurrency = Nothing
    Set objEmail = Nothing
    Set objCustom = Nothing
    Select Case intMaskType
        Case 0  'date
            Set objDate = New clsDate
        Case 1  'time   --------------------------removed from project
            'Set objTime = New clsTime
        Case 2  'phone
            Set objPhone = New clsPhone
        Case 3  'SSN
            Set objSSN = New clsSSN
        Case 4  'ZIP
            Set objZip = New clsZip
        Case 5  'currency
            Set objCurrency = New clsCurrency
        Case 6  'email
            Set objEmail = New clsEmail
        Case 7  'custom
            Set objCustom = New clsCustom
    End Select
End Sub

'This function is use to set the initial settings for the control.
'First, we retrieve the location of where the control's INI file
'should live.  If it doesn't exist, then we default our settings
'to a date mask type.  Once our mask type has been
'established, we create an instance of the respective object
Public Sub GetGeneralSettings()
    SetINILoc
    
    If Len(Dir(strCTRLName)) = 0 Then MaskType = 0

    intMaskType = Val(GetSetting("GENERAL", "MaskType"))
    
    SetObject
End Sub

'once the parent application has been destroyed, we destroy
'all of the objects to free up memory
Private Sub UserControl_Terminate()
    Set objDate = Nothing
    'Set objTime = Nothing
    Set objPhone = Nothing
    Set objSSN = Nothing
    Set objZip = Nothing
    Set objCurrency = Nothing
    Set objEmail = Nothing
    Set objCustom = Nothing
End Sub
