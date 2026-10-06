VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_Validation 
   Caption         =   "HAMU | Veri Doğrulama"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "frmHAMU_Validation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_Validation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public accepted As Boolean
Public PickSource As Boolean
Public Sub Setup(ByVal target As range, ByVal choices As Collection)
    Dim item As Variant
    lblDescription.Caption = HAMU_Text("Seçili hücrelerde açılır liste oluşturun. Kaynağı listeden seçin veya sayfadan bir liste alanı belirtin.")
    lblTarget.Caption = "Hedef: " & target.address(External:=True)
    cboSource.Clear
    For Each item In choices: cboSource.AddItem CStr(item): Next
    cboSource.ListIndex = 0
End Sub
Private Sub cmdPick_Click()
    PickSource = True: Me.Hide
End Sub
Private Sub cmdOK_Click()
    accepted = True: Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: accepted = False: Me.Hide
End Sub
Private Sub UserForm_Initialize()
    Me.Caption = HAMU_Text("HAMU | Veri Doğrulama")
    HAMU_ThemeForm Me
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub



