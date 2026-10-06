VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_Settings 
   Caption         =   "HAMU Tools"
   ClientHeight    =   6975
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9015.001
   OleObjectBlob   =   "frmHAMU_Settings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
 HAMU_ThemeForm Me
 Me.Caption = HAMU_L("HAMU | Ayarlar", "HAMU | Settings")
 lblTitle.Caption = HAMU_L(HAMU_Text("Çalışma tercihleri"), "Work preferences")
 lblNote.Caption = HAMU_L(HAMU_Text("Tercihler bu Windows kullanıcısı için saklanır. Excel'in genel ayarları değiştirilmez."), "Preferences are saved for this Windows user. Excel's global settings remain unchanged.")
 lblLanguage.Caption = HAMU_L("Dil", "Language")
 cboLanguage.AddItem "Türkçe / Turkish": cboLanguage.AddItem "English"
 cboLanguage.ListIndex = IIf(HAMU_Setting("Language") = "en", 1, 0)
 lblBackup.Caption = HAMU_L(HAMU_Text("Yedek klasörü"), "Backup folder")
 txtBackup.text = HAMU_Setting("BackupFolder")
 txtBackup.ControlTipText = HAMU_DefaultBackupFolder()
 lblExport.Caption = HAMU_L(HAMU_Text("PDF klasörü"), "PDF folder")
 txtExport.text = HAMU_Setting("ExportFolder")
 lblRange.Caption = HAMU_L(HAMU_Text("Aralık seçimi"), "Range selection")
 cboRange.AddItem HAMU_L(HAMU_Text("Uygunsa mevcut seçimi kullan"), "Use current selection when suitable")
 cboRange.AddItem HAMU_L(HAMU_Text("Her zaman aralığı sor"), "Always ask for the range")
 cboRange.ListIndex = CLng(HAMU_Setting("RangePrompt"))
 lblScope.Caption = HAMU_L(HAMU_Text("Klasör tarama"), "Folder scanning")
 cboScope.AddItem HAMU_L("Her seferinde sor", "Ask each time")
 cboScope.AddItem HAMU_L(HAMU_Text("Yalnız mevcut klasör"), "Current folder only")
 cboScope.AddItem HAMU_L(HAMU_Text("Alt klasörlerle birlikte"), "Include subfolders")
 cboScope.ListIndex = CLng(HAMU_Setting("FolderScope"))
 lblImage.Caption = HAMU_L(HAMU_Text("Resim yerleşimi"), "Image placement")
 cboImage.AddItem HAMU_L("Her seferinde sor", "Ask each time")
 cboImage.AddItem HAMU_L(HAMU_Text("Hücre içine"), "In cell")
 cboImage.AddItem HAMU_L(HAMU_Text("Hücre üstüne"), "Over cell")
 cboImage.ListIndex = CLng(HAMU_Setting("ImagePlacement"))
 lblOutput.Caption = HAMU_L("Benzersiz liste", "Unique list")
 cboOutput.AddItem HAMU_L("Hedefi her seferinde sor", "Ask for destination each time")
 cboOutput.AddItem HAMU_L(HAMU_Text("Hedef hücre seç"), "Select destination cell")
 cboOutput.AddItem HAMU_L(HAMU_Text("Yeni sayfa oluştur"), "Create a new sheet")
 cboOutput.ListIndex = CLng(HAMU_Setting("UniqueOutput"))
 lblOverwrite.Caption = HAMU_L(HAMU_Text("Dolu çıktı alanı"), "Occupied output")
 cboOverwrite.AddItem HAMU_L(HAMU_Text("Değiştirmeden önce onay iste"), "Confirm before replacing")
 cboOverwrite.AddItem HAMU_L(HAMU_Text("Mevcut içeriğin üzerine yazma"), "Never overwrite existing content")
 cboOverwrite.ListIndex = CLng(HAMU_Setting("Overwrite"))
 cmdSave.Caption = HAMU_L("Kaydet", "Save")
 cmdClose.Caption = HAMU_L("Kapat", "Close")
 cmdDefaults.Caption = HAMU_L(HAMU_Text("Varsayılanlar"), "Defaults")
 lblFolderHint.Caption = HAMU_L(HAMU_Text("Boş yedek yolu: Belgeler/HAMU/Backup. Boş PDF yolu: işlem sırasında sorulur."), "Empty backup path: Documents/HAMU/Backup. Empty PDF path: ask during export.")
End Sub
Private Sub cmdSave_Click()
 On Error GoTo Failed
 HAMU_SetSetting "BackupFolder", Trim$(txtBackup.text)
 HAMU_SetSetting "ExportFolder", Trim$(txtExport.text)
 HAMU_SetSetting "RangePrompt", CStr(cboRange.ListIndex)
 HAMU_SetSetting "FolderScope", CStr(cboScope.ListIndex)
 HAMU_SetSetting "ImagePlacement", CStr(cboImage.ListIndex)
 HAMU_SetSetting "UniqueOutput", CStr(cboOutput.ListIndex)
 HAMU_SetSetting "Overwrite", CStr(cboOverwrite.ListIndex)
 HAMU_SetSetting "Language", IIf(cboLanguage.ListIndex = 1, "en", "tr")
 HAMU_SaveSettings
 HAMU_RefreshRibbon
 Me.Hide
 Exit Sub
Failed:
 HAMU_ReloadSettings
 HAMU_ShowInfo HAMU_L(HAMU_Text("Ayarlar kaydedilemedi. Klasör yollarını ve yazma iznini kontrol edin."), "Settings could not be saved. Check folder paths and write permission.")
End Sub
Private Sub cmdDefaults_Click()
 txtBackup.text = "": txtExport.text = ""
 cboRange.ListIndex = 0: cboScope.ListIndex = 0: cboImage.ListIndex = 0: cboOutput.ListIndex = 0: cboOverwrite.ListIndex = 0
End Sub
Private Sub cmdClose_Click()
 Unload Me
End Sub
Private Sub cmdBackupBrowse_Click()
 Dim dialog As FileDialog
 Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
 If dialog.show = -1 Then txtBackup.text = dialog.SelectedItems(1)
End Sub
Private Sub cmdExportBrowse_Click()
 Dim dialog As FileDialog
 Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
 If dialog.show = -1 Then txtExport.text = dialog.SelectedItems(1)
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub


