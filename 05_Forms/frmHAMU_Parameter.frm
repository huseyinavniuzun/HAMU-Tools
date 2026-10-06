VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_Parameter 
   Caption         =   "HAMU Tools"
   ClientHeight    =   4695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6705
   OleObjectBlob   =   "frmHAMU_Parameter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_Parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public accepted As Boolean
Public ValueText As String
Private mNumeric As Boolean
Public Sub Setup(ByVal explanation As String, ByVal titleText As String, ByVal initial As String, ByVal numericOnly As Boolean)
    Me.Caption = titleText
    lblTitle.Caption = titleText
    Dim intro As String
    intro = HAMU_CommandIntro(gHAMU_CommandId)
    If Len(intro) > 0 Then explanation = intro & vbCrLf & vbCrLf & explanation
    txtDescription.Caption = explanation
    txtDescription.AutoSize = True
    Dim contentHeight As Single
    contentHeight = txtDescription.Height
    If contentHeight < 36 Then contentHeight = 36
    txtDescription.Height = contentHeight
    txtValue.Top = txtDescription.Top + contentHeight + 10
    lblHint.Top = txtValue.Top + txtValue.Height + 10
    cmdOK.Top = lblHint.Top + lblHint.Height + 10
    Me.Height = cmdOK.Top + cmdOK.Height + 48
    txtValue.text = initial
    mNumeric = numericOnly
    accepted = False
    If numericOnly Then lblHint.Caption = HAMU_Text("Sayısal bir değer girin. Ondalık ayırıcı Windows ayarınızı kullanır.")
End Sub
Private Sub cmdOK_Click()
    If mNumeric And Not IsNumeric(txtValue.text) Then
        MsgBox HAMU_Text("Geçerli bir sayı girin."), vbInformation, HAMU_APP_NAME
        Exit Sub
    End If
    ValueText = txtValue.text
    accepted = True
    Me.Hide
End Sub
Private Sub cmdCancel_Click()
    accepted = False
    Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        cmdCancel_Click
    End If
End Sub
Private Sub UserForm_Initialize()
    HAMU_ThemeForm Me
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub



