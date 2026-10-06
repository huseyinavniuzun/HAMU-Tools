VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_Code 
   Caption         =   "HAMU Tools"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6975
   OleObjectBlob   =   "frmHAMU_Code.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_Code"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public accepted As Boolean
Public PickTarget As Boolean
Public PickSource As Boolean
Private mQR As Boolean
Private mLoading As Boolean
Public Property Get IsQR() As Boolean
 IsQR = mQR
End Property
Public Sub Setup(ByVal qr As Boolean, ByVal target As range)
    mQR = qr
    lblDescription.Caption = HAMU_Text("Türü ve içeriği seçin. Görsel aşağıdaki hedef hücreye eklenecek.")
    lblInternet.Caption = HAMU_Text("İnternet gereklidir. İçerik dış üretim servisine gönderilir.")
    txtTarget.text = target.address(External:=True)
    mLoading = True
    cboKind.Clear
    cboKind.AddItem HAMU_L("Karekod", "QR Code")
    cboKind.AddItem HAMU_L("Barkod", "Barcode")
    cboKind.ListIndex = IIf(qr, 0, 1)
    mLoading = False
    LoadModels
    cboContent.Clear
    cboContent.AddItem HAMU_Text("Metin / numara"): cboContent.AddItem HAMU_Text("Web bağlantısı")
    cboContent.AddItem HAMU_Text("E-posta"): cboContent.AddItem HAMU_Text("Telefon")
    cboContent.ListIndex = 0
    accepted = False
End Sub
Private Sub cmdOK_Click()
    If Len(Trim$(txtValue.text)) = 0 And Len(Trim$(txtSource.text)) = 0 Then
        lblStatus.Caption = HAMU_Text("İçerik alanını doldurun."): Exit Sub
    End If
    accepted = True: Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: accepted = False: Me.Hide
End Sub
Private Sub UserForm_Initialize()
    HAMU_ThemeForm Me
End Sub

Private Sub txtTarget_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PickTarget = True: Me.Hide
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub

Private Sub txtSource_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
 PickSource = True: Me.Hide
End Sub
Private Sub cmdSource_Click()
 PickSource = True: Me.Hide
End Sub
Private Sub cmdSourceClear_Click()
 txtSource.text = ""
End Sub
Private Sub txtSource_Change()
 txtTarget.Enabled = (Len(Trim$(txtSource.text)) = 0)
 txtValue.Enabled = txtTarget.Enabled
 If Not txtTarget.Enabled Then
  lblStatus.Caption = HAMU_L("Kodlar verilerin sağındaki ilk boş sütuna eklenir.", "Codes go in the first empty column to the right of the data.")
 Else
  lblStatus.Caption = ""
 End If
End Sub



Private Sub cboKind_Change()
 If mLoading Then Exit Sub
 mQR = (cboKind.ListIndex = 0)
 LoadModels
End Sub
Private Sub LoadModels()
 Dim item As Variant
 cboModel.Clear
 Me.Caption = HAMU_L("HAMU | Karekod ve Barkod", "HAMU | QR and Barcode")
 If mQR Then
  cboModel.AddItem HAMU_L("Standart karekod", "Standard QR")
  cboModel.AddItem HAMU_L("Yuvarlak noktalar", "Round dots")
  cboModel.AddItem HAMU_L("Yuvarlatılmış kareler", "Rounded squares")
 Else
  For Each item In Array("Code 128", "Code 39", "EAN 13", "EAN 8", "Data Matrix", "PDF 417")
   cboModel.AddItem CStr(item)
  Next
 End If
 cboModel.ListIndex = 0
End Sub




