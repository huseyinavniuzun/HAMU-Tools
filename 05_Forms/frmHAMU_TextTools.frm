VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_TextTools 
   Caption         =   "HAMU | Metin Araçları"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6390
   OleObjectBlob   =   "frmHAMU_TextTools.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_TextTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public accepted As Boolean
Private mKind As String
Public Sub Setup(ByVal kind As String, ByVal command As String)
    mKind = kind: Me.Caption = HAMU_APP_NAME & " | " & HAMU_CommandLabel(command)
    lblDescription.Caption = HAMU_CommandIntro(command)
    cboMode.Clear
    Select Case kind
        Case "affix": cboMode.AddItem HAMU_Text("Başa ve / veya sona ekle")
        Case "zeros": cboMode.AddItem HAMU_Text("Sabit uzunluğa sıfır ekle"): cboMode.AddItem HAMU_Text("Baştaki sıfırları kaldır")
        Case "unique": cboMode.AddItem HAMU_Text("Hücre içindeki tekrarları kaldır")
        Case "part"
            cboMode.AddItem "Soldan al": cboMode.AddItem HAMU_Text("Sağdan al"): cboMode.AddItem HAMU_Text("İşaretten önce al"): cboMode.AddItem HAMU_Text("İşaretten sonra al"): cboMode.AddItem HAMU_Text("İki işaret arasından al")
    End Select
    cboMode.ListIndex = 0
End Sub
Private Sub cboMode_Change()
    lblFirst.Caption = HAMU_Text("İlk değer"): lblSecond.Caption = HAMU_Text("İkinci değer")
    txtFirst.Enabled = True: txtSecond.visible = False: lblSecond.visible = False
    Select Case mKind
        Case "affix"
            lblFirst.Caption = HAMU_Text("Başa eklenecek metin"): lblSecond.Caption = "Sona eklenecek metin"
            txtSecond.visible = True: lblSecond.visible = True
        Case "zeros"
            lblFirst.Caption = HAMU_Text("Toplam karakter sayısı (örn. 5 " & ChrW(8594) & " 00042)")
            txtFirst.Enabled = (cboMode.ListIndex = 0)
            If Len(txtFirst.text) = 0 Then txtFirst.text = "5"
        Case "unique"
            lblFirst.Caption = HAMU_Text("Öğeleri ayıran işaret (örn. virgül)")
            If Len(txtFirst.text) = 0 Then txtFirst.text = ","
        Case "part"
            If cboMode.ListIndex < 2 Then
                lblFirst.Caption = HAMU_Text("Alınacak karakter sayısı")
            Else
                lblFirst.Caption = HAMU_Text("Aranacak işaret")
            End If
            If cboMode.ListIndex = 4 Then
                lblFirst.Caption = HAMU_Text("Başlangıç işareti"): lblSecond.Caption = HAMU_Text("Bitiş işareti")
                txtSecond.visible = True: lblSecond.visible = True
            End If
    End Select
End Sub
Private Sub cmdOK_Click()
    On Error GoTo invalid
    Dim probe As String
    probe = HAMU_TextTransform(HAMU_Text("Örnek00042"), mKind, cboMode.ListIndex + 1, txtFirst.text, txtSecond.text)
    accepted = True: Me.Hide
    Exit Sub
invalid:
    lblStatus.Caption = HAMU_FriendlyError(Err.number, Err.description)
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: accepted = False: Me.Hide
End Sub
Private Sub UserForm_Initialize()
    HAMU_ThemeForm Me
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub


