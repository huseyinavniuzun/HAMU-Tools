VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_Stripes 
   Caption         =   "HAMU | Renkler"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "frmHAMU_Stripes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_Stripes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public accepted As Boolean
Private Sub UserForm_Initialize()
 Dim item As Variant
 For Each item In Array(HAMU_L("Beyaz", "White"), HAMU_L("Açık Gri", "Light Gray"), HAMU_L("Mavi", "Blue"), HAMU_L("Yeşil", "Green"), HAMU_L("Sarı", "Yellow"), HAMU_L("Turuncu", "Orange"), HAMU_L("Mor", "Purple"), HAMU_L("Pembe", "Pink"), HAMU_L("Turkuaz", "Teal"))
  cboFirst.AddItem item: cboSecond.AddItem item
 Next
 cboFirst.Style = fmStyleDropDownList: cboSecond.Style = fmStyleDropDownList
 cboFirst.ListIndex = 0: cboSecond.ListIndex = 1
 cboDirection.Style = fmStyleDropDownList
 cboDirection.AddItem HAMU_L("Satırlar", "Rows"): cboDirection.AddItem HAMU_L("Sütunlar", "Columns")
 cboDirection.ListIndex = 0
 lblIntro.Caption = HAMU_L("İki dolgu rengi seçin; beyaz ve açık gri varsayılandır.", "Choose two fill colours. White and light gray are the defaults.")
 lblFirst.Caption = HAMU_L("1. renk", "1st colour"): lblSecond.Caption = HAMU_L("2. renk", "2nd colour")
 lblDirection.Caption = HAMU_L("Uygulama", "Apply to")
 cmdApply.Caption = HAMU_L("Uygula", "Apply")
End Sub
Private Sub cboFirst_Change()
 sampleFirst.BackStyle = fmBackStyleOpaque
 sampleFirst.BackColor = HAMU_StripeColor(cboFirst.ListIndex)
End Sub
Private Sub cboSecond_Change()
 sampleSecond.BackStyle = fmBackStyleOpaque
 sampleSecond.BackColor = HAMU_StripeColor(cboSecond.ListIndex)
End Sub
Private Sub cmdApply_Click()
 accepted = True: Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 If CloseMode = 0 Then Cancel = True: accepted = False: Me.Hide
End Sub

