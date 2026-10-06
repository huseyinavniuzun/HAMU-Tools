VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_About 
   Caption         =   "HAMU Tools"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   OleObjectBlob   =   "frmHAMU_About.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
 HAMU_ThemeForm Me
 Me.Caption = HAMU_Text("HAMU | Hakkında")
 imgLogo.Left = 0: imgLogo.Top = 0: imgLogo.width = Me.InsideWidth: imgLogo.Height = 104
 imgLogo.BackColor = RGB(228, 231, 235)
 lblTitle.Caption = "HAMU"
 lblTitle.Left = 20: lblTitle.Top = 116: lblTitle.Font.Size = 19: lblTitle.Font.Bold = True
 lblWebsite.Caption = HAMU_Text("Hüseyin Avni UZUN")
 lblWebsite.Left = 20: lblWebsite.Top = 145: lblWebsite.Font.Size = 11
 lblInfo.Left = 20: lblInfo.Top = 174: lblInfo.width = Me.InsideWidth - 40
 lblInfo.Caption = HAMU_Text("Excel’de veriyi düzenlemek, işleri hızlandırmak ve sonuçları anlaşılır kılmak için geliştirilen HAMU Tools." & vbLf & vbLf & "V1 · Excel üretkenlik ve veri araçları")
 lblInfo.AutoSize = True
 lblWeb.Caption = "huseyinavniuzun.com": lblWeb.ControlTipText = "https://huseyinavniuzun.com"
 lblGitHub.Caption = "github.com/huseyinavniuzun": lblGitHub.ControlTipText = "https://github.com/huseyinavniuzun"
 lblEmail.Caption = HAMU_Text("hamu@huseyinavniuzun.com")
 lblWeb.Top = lblInfo.Top + lblInfo.Height + 14
 lblGitHub.Top = lblWeb.Top + 23: lblEmail.Top = lblGitHub.Top + 23
 lblWeb.ForeColor = RGB(0, 110, 170): lblGitHub.ForeColor = lblWeb.ForeColor: lblEmail.ForeColor = lblWeb.ForeColor
 lblWeb.Font.Underline = True: lblGitHub.Font.Underline = True: lblEmail.Font.Underline = True
 cmdClose.Top = lblEmail.Top + 32: Me.Height = cmdClose.Top + 56
End Sub
Private Sub lblWeb_Click()
 HAMU_OpenProfile "https://huseyinavniuzun.com"
End Sub
Private Sub lblGitHub_Click()
 HAMU_OpenProfile "https://github.com/huseyinavniuzun"
End Sub
Private Sub lblEmail_Click()
 HAMU_OpenProfile "mailto:hamu@huseyinavniuzun.com"
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub ShowHelp()
    lblTitle.Caption = HAMU_Text("HAMU Tools | Yardım")
    lblInfo.Caption = HAMU_Text("1. İşlem yapılacak hücreleri seçin.") & vbCrLf & _
        HAMU_Text("2. HAMU Tools sekmesinden ilgili komutu açın.") & vbCrLf & _
        HAMU_Text("3. Açılan ekranda açıklamayı okuyup seçenekleri belirleyin.") & vbCrLf & vbCrLf & _
        HAMU_Text("Ek aralık gerekiyorsa Excel'in seçim kutusu açılır. Tarih Seç, tarihi aktif hücreye yazar.") & vbCrLf & vbCrLf & _
        HAMU_Text("Barkod ve karekod için internet gereklidir. Kod içeriği üretim servisine gönderilir.") & vbCrLf & vbCrLf & _
        HAMU_Text("VBA işlemlerinden önce önemli çalışma kitaplarını yedekleyin.")
    FitInformation
    Me.show vbModal
End Sub

Public Sub ShowFunctions()
    lblTitle.Caption = HAMU_Text("HAMU | Hücre Fonksiyonları")
    lblInfo.Caption = HAMU_Text("=HAMU_URLCoz(A1) | UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur." & vbLf & vbLf & "=HAMU_RengeGoreTopla(A1:A20;B1) | B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası F9 kullanın." & vbLf & vbLf & "=HAMU_RengeGoreSay(A1:A20;B1) | B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası F9 kullanın." & vbLf & vbLf & "=HAMU_YerelDosyaYolu(A1) | Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur." & vbLf & vbLf & "=HAMU_GizliLink(A1) | Hücrenin ilk köprü adresini getirir." & vbLf & vbLf & "=HAMU_LinkParamDegeri(A1;""id"") | URL içindeki sorgu parametresinin değerini getirir." & vbLf & vbLf & "=HAMU_SayidanMetine(123) | Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur." & vbLf & vbLf & "=HAMU_TablodanVeriGetir(A1;""Tablo1"";""Kod"";""Ad"") | Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir.")
    lblInfo.AutoSize = True
    cmdClose.Top = lblInfo.Top + lblInfo.Height + 12
    Me.Height = cmdClose.Top + cmdClose.Height + 48
    Me.show vbModal
End Sub

Private Sub FitInformation()
    lblInfo.AutoSize = True
    cmdClose.Top = lblInfo.Top + lblInfo.Height + 12
    Me.Height = cmdClose.Top + cmdClose.Height + 36
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub




