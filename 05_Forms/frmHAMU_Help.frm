VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_Help 
   Caption         =   "HAMU Tools"
   ClientHeight    =   7290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   OleObjectBlob   =   "frmHAMU_Help.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
 HAMU_ThemeForm Me
 Me.Caption = HAMU_Text("HAMU | Kullanım Rehberi")
 lblTitle.Caption = "HAMU Tools V1"
 lblTitle.Font.Size = 16
 lblTitle.Font.Bold = True
 cboTopic.Style = 2
 cboTopic.AddItem HAMU_Text("İlk adımlar")
 cboTopic.AddItem HAMU_Text("Aralık ve çıktı seçimi")
 cboTopic.AddItem HAMU_Text("Metin ve veri temizleme")
 cboTopic.AddItem HAMU_Text("Tablolar ve özetleme")
 cboTopic.AddItem HAMU_Text("Dosyalar ve yedekleme")
 cboTopic.AddItem HAMU_Text("Resimler, barkod ve karekod")
 cboTopic.AddItem HAMU_Text("HAMU hücre fonksiyonları")
 cboTopic.AddItem HAMU_Text("Sorun giderme ve sınırlar")
 cboTopic.AddItem HAMU_Text("Tüm araçlar")
 cboTopic.ListIndex = 0
End Sub
Public Sub Setup(ByVal topic As String)
 Dim keys As Variant, i As Long
 keys = Split("start|selection|clean|tables|files|visual|functions|errors|catalog", "|")
 For i = 0 To UBound(keys)
  If keys(i) = topic Then cboTopic.ListIndex = i: Exit For
 Next i
 ShowTopic
End Sub
Private Sub cboTopic_Change()
 txtSearch.text = ""
 ShowTopic
End Sub
Private Sub ShowTopic()
 Dim keys As Variant, key As String
 If cboTopic.ListIndex < 0 Then Exit Sub
 keys = Split("start|selection|clean|tables|files|visual|functions|errors|catalog", "|")
 key = keys(cboTopic.ListIndex)
 lblHeading.Caption = HAMU_GuideTitle(key)
 lblHeading.Font.Size = 12
 lblHeading.Font.Bold = True
 lblBody.Caption = HAMU_GuideBody(key)
 lblBody.WordWrap = True
 lblBody.AutoSize = True
 fraContent.ScrollBars = 2
 fraContent.ScrollHeight = lblBody.Top + lblBody.Height + 12
 fraContent.ScrollTop = 0
End Sub
Private Sub cmdClose_Click()
 Unload Me
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub

Private Sub txtSearch_Change()
 If Len(Trim$(txtSearch.text)) = 0 Then ShowTopic: Exit Sub
 lblHeading.Caption = HAMU_L("Arama sonuçları", "Search results")
 lblBody.Caption = HAMU_SearchGuide(txtSearch.text)
 lblBody.AutoSize = True
 fraContent.ScrollHeight = lblBody.Top + lblBody.Height + 12
 fraContent.ScrollTop = 0
End Sub


