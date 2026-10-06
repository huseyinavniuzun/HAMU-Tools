VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_DateRange 
   Caption         =   "HAMU | Tarih Listesi"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   OleObjectBlob   =   "frmHAMU_DateRange.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_DateRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public accepted As Boolean
Public StartDate As Date
Public EndDate As Date
Private Sub cmdOK_Click()
    If Not HAMU_ParseDateInput(txtStart.text, StartDate) Or Not HAMU_ParseDateInput(txtEnd.text, EndDate) Then
        lblStatus.Caption = HAMU_Text("Geçerli tarihleri gg.aa.yyyy biçiminde girin."): Exit Sub
    End If
    If EndDate < StartDate Then lblStatus.Caption = HAMU_Text("Bitiş tarihi başlangıçtan önce olamaz."): Exit Sub
    If DateDiff("d", StartDate, EndDate) > 100000 Then lblStatus.Caption = HAMU_Text("En fazla 100.001 günlük bir aralık seçin."): Exit Sub
    accepted = True: Me.Hide
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: accepted = False: Me.Hide
End Sub
Private Sub UserForm_Initialize()
    HAMU_ThemeForm Me
    txtStart.text = Format$(Date, "dd.mm.yyyy")
    txtEnd.text = Format$(Date + 30, "dd.mm.yyyy")
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub



