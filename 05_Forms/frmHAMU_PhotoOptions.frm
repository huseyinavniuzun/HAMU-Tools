VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_PhotoOptions 
   Caption         =   "HAMU"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   OleObjectBlob   =   "frmHAMU_PhotoOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_PhotoOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public accepted As Boolean
Public Sub Setup(ByVal placement As Long)
    Me.Caption = HAMU_L("HAMU | Toplu Fotoğraf", "HAMU | Batch Photos")
    lblTitle.Caption = HAMU_L("Fotoğraf seçenekleri", "Photo options")
    lblPlacement.Caption = HAMU_L("Yerleşim", "Placement")
    lblQuality.Caption = HAMU_L("Kalite", "Quality")
    cboPlacement.Clear
    cboPlacement.AddItem HAMU_L("Hücre içine", "In cells")
    cboPlacement.AddItem HAMU_L("Hücre üstüne", "Over cells")
    cboPlacement.ListIndex = 0
    If placement = 2 Then cboPlacement.ListIndex = 1
    cboQuality.Clear
    cboQuality.AddItem HAMU_L("Full HD optimize (önerilen)", "Full HD optimized (recommended)")
    cboQuality.AddItem HAMU_L("Orijinal kalite", "Original quality")
    cboQuality.ListIndex = 0
    cmdOK.Caption = HAMU_L("Ekle", "Insert")
    accepted = False
    UpdateDescription
End Sub
Private Sub cboQuality_Change()
    UpdateDescription
End Sub
Private Sub UpdateDescription()
    If cboQuality.ListIndex = 1 Then
        lblDescription.Caption = HAMU_L("Tam çözünürlük korunur. Büyük fotoğraflar Excel dosyasını büyütebilir ve yavaşlatabilir.", "Full resolution is retained. Large photos can increase workbook size and slow Excel.")
    Else
        lblDescription.Caption = HAMU_L("En fazla 1920 × 1080 piksel (dikeyde 1080 × 1920). Oran ve saydamlık korunur; küçük fotoğraflar büyütülmez.", "Up to 1920 × 1080 pixels (portrait: 1080 × 1920). Aspect ratio and transparency are preserved; small photos are not enlarged.")
    End If
    lblSource.Caption = HAMU_L("Kaynak fotoğraflar değiştirilmez.", "Source photos remain unchanged.")
End Sub
Private Sub cmdOK_Click()
    accepted = True
    Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: accepted = False: Me.Hide
End Sub
Private Sub UserForm_Initialize()
    HAMU_ThemeForm Me
End Sub

