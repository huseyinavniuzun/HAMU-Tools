Attribute VB_Name = "modHAMU_Settings"
Option Explicit
Option Private Module
Private mSettings As Object, mSettingsPath As String
Public Function HAMU_DefaultBackupFolder() As String
 HAMU_DefaultBackupFolder = CreateObject("WScript.Shell").SpecialFolders("MyDocuments") & "\HAMU\Backup"
End Function
Public Sub HAMU_LoadSettings(Optional ByVal settingsPath As String = "")
 If Len(settingsPath) = 0 And Not mSettings Is Nothing Then Exit Sub
 Dim fso As Object, stream As Object, line As String, splitAt As Long, key As String, value As String
 Set mSettings = CreateObject("Scripting.Dictionary")
 mSettings.CompareMode = vbTextCompare
 mSettings.Add "Language", "tr": mSettings.Add "BackupFolder", "": mSettings.Add "ExportFolder", ""
 mSettings.Add "RangePrompt", "0": mSettings.Add "FolderScope", "0": mSettings.Add "ImagePlacement", "0"
 mSettings.Add "UniqueOutput", "0": mSettings.Add "Overwrite", "0"
 If Len(settingsPath) = 0 Then settingsPath = Environ$("APPDATA") & "\HAMU Tools\settings.ini"
 mSettingsPath = settingsPath
 On Error GoTo Done
 Set fso = CreateObject("Scripting.FileSystemObject")
 If Not fso.FileExists(settingsPath) Then Exit Sub
 Set stream = fso.OpenTextFile(settingsPath, 1, False, -1)
 Do While Not stream.AtEndOfStream
  line = stream.ReadLine: splitAt = InStr(line, "=")
  If splitAt > 0 Then
   key = Left$(line, splitAt - 1): value = Mid$(line, splitAt + 1)
   If mSettings.exists(key) Then
    If HAMU_ValidSetting(key, value) Then mSettings(key) = value
 If key = "Language" And value <> "en" Then HAMU_ReleaseLanguageCache
   End If
  End If
 Loop
Done:
 On Error Resume Next
 If Not stream Is Nothing Then stream.Close
End Sub
Public Function HAMU_Setting(ByVal key As String) As String
 If mSettings Is Nothing Then HAMU_LoadSettings
 If mSettings.exists(key) Then HAMU_Setting = CStr(mSettings(key))
End Function
Private Function HAMU_ValidSetting(ByVal key As String, ByVal value As String) As Boolean
 If InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then Exit Function
 Select Case key
 Case "Language": HAMU_ValidSetting = (value = "tr" Or value = "en")
 Case "BackupFolder", "ExportFolder": HAMU_ValidSetting = (Len(value) = 0 Or (Len(value) >= 3 And (Mid$(value, 2, 2) = ":\" Or Left$(value, 2) = "\\")))
 Case "RangePrompt", "Overwrite": HAMU_ValidSetting = (value = "0" Or value = "1")
 Case "FolderScope", "ImagePlacement", "UniqueOutput": HAMU_ValidSetting = (value = "0" Or value = "1" Or value = "2")
 End Select
End Function
Public Sub HAMU_SetSetting(ByVal key As String, ByVal value As String)
 If mSettings Is Nothing Then HAMU_LoadSettings
 If Not mSettings.exists(key) Or Not HAMU_ValidSetting(key, value) Then Err.Raise 5, , "Invalid HAMU setting: " & key
 mSettings(key) = value
 If key = "Language" And value <> "en" Then HAMU_ReleaseLanguageCache
End Sub
Public Sub HAMU_SaveSettings()
 Dim fso As Object, stream As Object, key As Variant, temp As String, backup As String
 If mSettings Is Nothing Then HAMU_LoadSettings
 Set fso = CreateObject("Scripting.FileSystemObject")
 HAMU_CreateFolders fso.GetParentFolderName(mSettingsPath)
 temp = mSettingsPath & ".tmp": backup = mSettingsPath & ".bak"
 Set stream = fso.CreateTextFile(temp, True, True)
 For Each key In mSettings.keys
  stream.WriteLine CStr(key) & "=" & CStr(mSettings(key))
 Next key
 stream.Close
 If fso.FileExists(mSettingsPath) Then fso.CopyFile mSettingsPath, backup, True
 fso.CopyFile temp, mSettingsPath, True: fso.DeleteFile temp
End Sub
Public Function HAMU_BackupFolder() As String
 HAMU_BackupFolder = HAMU_Setting("BackupFolder")
 If Len(HAMU_BackupFolder) = 0 Then HAMU_BackupFolder = HAMU_DefaultBackupFolder()
End Function
Public Function HAMU_ConfirmOverwrite() As Boolean
 If HAMU_Setting("Overwrite") = "1" Then
  HAMU_ShowInfo HAMU_L(HAMU_Text("Hedef dolu. Ayarlarda mevcut içeriğin üzerine yazma kapalı."), "The destination is occupied. Overwriting existing content is disabled in Settings.")
  Exit Function
 End If
 HAMU_ConfirmOverwrite = (MsgBox(HAMU_L(HAMU_Text("Hedef alandaki mevcut içerik değiştirilecek. Devam edilsin mi?"), "Existing destination content will be replaced. Continue?"), vbQuestion + vbYesNo, HAMU_APP_NAME) = vbYes)
End Function
Public Sub HAMU_ShowSettings()
 Dim form As frmHAMU_Settings
 Set form = New frmHAMU_Settings
 form.show vbModal
 Unload form
End Sub

Public Sub HAMU_ReloadSettings()
 Dim path As String
 path = mSettingsPath
 HAMU_LoadSettings path
End Sub

