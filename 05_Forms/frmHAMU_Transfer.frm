VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_Transfer 
   Caption         =   "HAMU | Şablondan Aktarım"
   ClientHeight    =   6150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10020
   OleObjectBlob   =   "frmHAMU_Transfer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function GetExt(ByVal p As String) As String
    Dim dotPos As Long: dotPos = InStrRev(p, ".")
    If dotPos > 0 Then GetExt = LCase$(Mid$(p, dotPos + 1)) Else GetExt = ""
End Function

Private Sub CommandButton1_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Caption = HAMU_Text("HAMU | Şablondan Aktarım")
    HAMU_ThemeForm Me
    Me.chkLambda.value = True
    Me.chkCell.value = False
    Me.Caption = HAMU_L("HAMU | Aktarma Sihirbazı", "HAMU | Transfer Wizard")

    Me.txtPref.text = ""
    Dim lastPath As String: lastPath = Pref_GetTemplatePath()
    Me.txtPath.text = lastPath
    Me.chkHatirla.value = (Len(lastPath) > 0)
    If GetExt(lastPath) = "xls" Then
        Me.chkLambda.value = False
        Me.chkLambda.Enabled = False
    End If
End Sub

Private Sub btnBrowse_Click()
    Dim p As String: p = DosyaSec(HAMU_Text("Şablon dosyayı seçin"))
    If Len(p) > 0 Then
        Me.txtPath.text = p
        If GetExt(p) = "xls" Then
            Me.chkLambda.value = False
            Me.chkLambda.Enabled = False
            MsgBox HAMU_Text(".xls biçimi LAMBDA işlevlerini desteklemez. Yalnızca stiller aktarılabilir."), vbExclamation
        Else
            Me.chkLambda.Enabled = True
        End If
        If Me.chkHatirla.value Then
            Pref_SetTemplatePath p
        Else
            Pref_SetTemplatePath ""   ' de?ilse temizle
        End If
    End If
End Sub

Private Sub btnNext_Click()
    If Len(Me.txtPath.text) = 0 Then
        MsgBox HAMU_Text("Önce şablon dosyayı seçin."), vbExclamation: Exit Sub
    End If
    If (Me.chkLambda.value Or Me.chkCell.value Or Me.chkSheets.value Or Me.chkNames.value Or Me.chkTables.value) = False Then
        MsgBox HAMU_Text("En az bir seçenek işaretleyin (Lambda veya Hücre Stili)."), vbExclamation: Exit Sub
    End If
    If GetExt(Me.txtPath.text) = "xls" And Me.chkLambda.value Then
        MsgBox HAMU_Text(".xls biçiminde LAMBDA aktarımı desteklenmez."), vbCritical
        Exit Sub
    End If
    If Me.chkHatirla.value Then
        Pref_SetTemplatePath Me.txtPath.text
    Else
        Pref_SetTemplatePath ""
    End If

    Me.Hide
    TM_OnNext Me.txtPath.text, Me.txtPref.text, _
              Me.chkLambda.value, Me.chkSheets.value, Me.chkNames.value, _
              Me.chkTables.value, Me.chkCell.value
    Unload Me
End Sub

Private Sub chkHatirla_Click()
    If Me.chkHatirla.value = False Then
        Pref_SetTemplatePath ""
    Else
        If Len(Me.txtPath.text) > 0 Then
            Pref_SetTemplatePath Me.txtPath.text
        End If
    End If
End Sub


Private Sub btnCancel_Click()
    Unload Me
End Sub


























































Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub












































