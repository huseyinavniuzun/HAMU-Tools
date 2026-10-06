Attribute VB_Name = "modHAMU_File"
Option Explicit
Option Private Module
Public HAMU_LastFileNotice As String

Public Sub HAMU_Transfer()
    On Error GoTo HAMU_CommandFailed
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_Transfer") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    frmHAMU_Transfer.show
    Exit Sub
HAMU_CommandFailed:
    HAMU_ShowError "HAMU_Transfer", Err.number, Err.description
End Sub
Public Sub HAMU_Backup()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_Backup") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    On Error GoTo Failed
    Dim path As String, documents As String, shell As Object
    Set shell = CreateObject("WScript.Shell")
    documents = shell.SpecialFolders("MyDocuments")
    path = HAMU_BackupToFolder(ActiveWorkbook, HAMU_BackupFolder())
    HAMU_ShowBackupResult path
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_Backup", Err.number, Err.description
End Sub
Public Function HAMU_BackupToFolder(ByVal wb As Workbook, ByVal folder As String) As String
    Dim fso As Object, base As String, target As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(wb.path) = 0 And wb.HasVBProject Then Err.Raise 5, , HAMU_Text("Makroları korumak için önce dosyayı makro etkin biçimde kaydedin; ardından yedek alın.")
    base = fso.GetBaseName(wb.name)
    HAMU_CreateFolders folder
    target = HAMU_UniqueFilePath(folder, HAMU_TemizDosyaAdi(base) & "_YEDEK_" & Format$(Now, "yyyymmdd_hhnnss"), HAMU_WorkbookExtension(wb))
    wb.SaveCopyAs target
    If Not fso.FileExists(target) Then Err.Raise 5, , HAMU_Text("Yedek oluşturulamadı. Klasörün yazılabilir olduğunu kontrol edin.")
    HAMU_BackupToFolder = target
End Function
Public Sub HAMU_AddDateToFilename()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_AddDateToFilename") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    On Error GoTo Failed
    Dim newBook As Workbook
    If Len(ActiveWorkbook.path) = 0 Then
        HAMU_ShowInfo HAMU_Text("Önce dosyayı kaydedin; ardından adına tarih ekleyin."): Exit Sub
    End If
    Set newBook = HAMU_DateRenameCopy(ActiveWorkbook)
    HAMU_ShowInfo HAMU_Text("Tarihli kopya açıldı:") & vbCrLf & newBook.FullName & vbCrLf & HAMU_LastFileNotice
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_AddDateToFilename", Err.number, Err.description
End Sub
Public Function HAMU_DateRenameCopy(ByVal wb As Workbook) As Workbook
    Dim fso As Object, oldPath As String, target As String, newBook As Workbook
    Dim security As Long, events As Boolean, oldSheets As Long
    Dim number As Long, description As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    HAMU_LastFileNotice = ""
    If Len(wb.path) = 0 Or InStr(wb.FullName, "://") > 0 Then Err.Raise 5, , HAMU_Text("İşlem için yerel klasöre kaydedilmiş bir dosya gerekir.")
    If wb.ReadOnly Then Err.Raise 5, , HAMU_Text("Salt okunur dosya yeniden adlandırılamaz. Yazılabilir bir kopyasını kaydedin.")
    oldPath = wb.FullName: oldSheets = wb.sheets.count
    target = HAMU_UniqueFilePath(wb.path, fso.GetBaseName(wb.name) & "_" & Format$(Now, "yyyymmdd_hhnnss"), "." & fso.GetExtensionName(wb.name))
    security = Application.AutomationSecurity: events = Application.EnableEvents
    On Error GoTo Failed
    wb.SaveCopyAs target
    If Not fso.FileExists(target) Then Err.Raise 5, , HAMU_Text("Yeni kopya kaydedilemedi. Eski dosya korunuyor.")
    If fso.GetFile(target).Size = 0 Then Err.Raise 5, , HAMU_Text("Yeni kopya boş; eski dosya korunuyor.")
    Application.AutomationSecurity = 3: Application.EnableEvents = False
    Set newBook = Workbooks.Open(fileName:=target, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True)
    If newBook.ReadOnly Or newBook.sheets.count <> oldSheets Then Err.Raise 5, , HAMU_Text("Yeni kopya doğrulanamadı. Eski dosya korunuyor.")
    wb.Close SaveChanges:=False
    On Error Resume Next
    fso.DeleteFile oldPath, False
    If Err.number <> 0 Then HAMU_LastFileNotice = HAMU_Text("Yeni kopya hazır; eski dosya silinemedi ve yerinde kaldı.")
    Err.Clear: On Error GoTo Failed
    newBook.Activate
    Set HAMU_DateRenameCopy = newBook
    Application.AutomationSecurity = security: Application.EnableEvents = events
    Exit Function
Failed:
    number = Err.number: description = Err.description
    Application.AutomationSecurity = security: Application.EnableEvents = events
    Err.Raise number, , description
End Function
Public Sub HAMU_FolderIndex()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_FolderIndex") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    On Error GoTo Failed
    Dim fd As FileDialog, sheet As Worksheet, wb As Workbook
    Set wb = ActiveWorkbook
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.title = HAMU_Text("Listelenecek klasörü seçin.")
    If fd.show <> -1 Then Exit Sub
    Dim options As frmHAMU_FolderScope, recursive As Boolean
    If HAMU_Setting("FolderScope") = "0" Then
    Set options = New frmHAMU_FolderScope
    options.Setup fd.SelectedItems(1)
    options.show vbModal
    If Not options.accepted Then Unload options: Exit Sub
    recursive = options.optRecursive.value
    Unload options
    Else
        recursive = (HAMU_Setting("FolderScope") = "2")
    End If
    Dim state As clsHAMU_AppState
    Set state = New clsHAMU_AppState: state.Capture
    Application.ScreenUpdating = False
    Application.StatusBar = HAMU_Text("Dosyalar listeleniyor… Alt klasörlerde işlem uzun sürebilir.")
    Set sheet = HAMU_BuildFolderIndex(wb, fd.SelectedItems(1), recursive)
    state.Restore
    sheet.Activate
    Exit Sub
Failed:
    If Not state Is Nothing Then state.Restore
    HAMU_ShowError "HAMU_FolderIndex", Err.number, Err.description
End Sub
Public Sub HAMU_ExportSheetsAsPDF()
 On Error GoTo Failed
 Dim scope As range, mode As Long, layout As Long, folder As String, path As Variant
 If Not HAMU_HasWorkbook() Then Exit Sub
 mode = HAMU_ChooseIndex(HAMU_L("PDF Kaydet", "Save PDF"), HAMU_L("Mevcut seçiminizi veya etkin sayfayı PDF olarak kaydedin.", "Save the current selection or active sheet as PDF."), Array(HAMU_L("Alan Seç", "Select Range"), HAMU_L("Etkin sayfa", "Active sheet")))
 If mode = 0 Then Exit Sub
 If mode = 1 Then
  Set scope = HAMU_PromptRange(HAMU_L("PDF olarak kaydedilecek alanı seçin.", "Select the range to export to PDF."))
 Else
  Set scope = HAMU_BoundedRange(ActiveSheet.UsedRange, True)
 End If
 If scope Is Nothing Then Exit Sub
 layout = HAMU_ChooseIndex(HAMU_L("PDF Yerleşimi", "PDF Layout"), HAMU_L("Otomatik yön, alanın genişlik/yükseklik oranını kullanır. Genişlik bir sayfaya sığdırılır; uzun tablolar aşağıya devam eder.", "Automatic orientation uses the range width/height ratio. Width fits one page; long tables continue downward."), Array(HAMU_L("Otomatik", "Automatic"), HAMU_L("Dikey", "Portrait"), HAMU_L("Yatay", "Landscape")))
 If layout = 0 Then Exit Sub
 folder = HAMU_Setting("ExportFolder")
 If Len(folder) > 0 Then HAMU_CreateFolders folder: folder = folder & Application.PathSeparator
 path = Application.GetSaveAsFilename(folder & HAMU_TemizDosyaAdi(scope.Worksheet.name) & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf", "PDF (*.pdf),*.pdf")
 If VarType(path) = vbBoolean Then Exit Sub
 If LCase$(right$(CStr(path), 4)) <> ".pdf" Then path = CStr(path) & ".pdf"
 HAMU_ExportPDFRange scope, CStr(path), layout
 HAMU_ShowInfo HAMU_L("Kaydedildi: ", "Saved: ") & CStr(path)
 Exit Sub
Failed:
 HAMU_ShowError HAMU_L("PDF Kaydet", "Save PDF"), Err.number, Err.description
End Sub




Public Sub HAMU_CreateFolders(ByVal path As String)
    Dim fso As Object, parent As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    parent = fso.GetParentFolderName(path)
    If Len(parent) = 0 Then Err.Raise 5, , HAMU_Text("Geçerli bir klasör yolu kullanın.")
    If Not fso.FolderExists(parent) Then HAMU_CreateFolders parent
    fso.CreateFolder path
End Sub
Private Function HAMU_UniqueFilePath(ByVal folder As String, ByVal base As String, ByVal extension As String) As String
    Dim fso As Object, path As String, i As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    path = fso.BuildPath(folder, base & extension): i = 1
    Do While fso.FileExists(path)
        i = i + 1: path = fso.BuildPath(folder, base & "_" & i & extension)
    Loop
    HAMU_UniqueFilePath = path
End Function
Public Function HAMU_FileKind(ByVal extension As String) As String
    Select Case LCase$(extension)
        Case "jpg", "jpeg", "png", "gif", "bmp", "tif", "tiff", "webp", "webm", "heic", "avif": HAMU_FileKind = HAMU_Text("Fotoğraf")
        Case "mp4", "mkv", "avi", "mov", "wmv", "flv", "m4v", "mpeg": HAMU_FileKind = "Video"
        Case "xls", "xlsx", "xlsm", "xlsb", "xltx", "xltm", "csv", "tsv": HAMU_FileKind = HAMU_Text("Excel tablo")
        Case "doc", "docx", "rtf", "odt": HAMU_FileKind = HAMU_Text("Word belge")
        Case "ppt", "pptx", "odp": HAMU_FileKind = HAMU_Text("Sunum")
        Case "pdf": HAMU_FileKind = HAMU_Text("PDF belge")
        Case "zip", "rar", "7z", "tar", "gz": HAMU_FileKind = HAMU_Text("Arşiv")
        Case "mp3", "wav", "flac", "ogg", "m4a": HAMU_FileKind = HAMU_Text("Ses")
        Case Else: HAMU_FileKind = HAMU_Text("Diğer")
    End Select
End Function
Public Function HAMU_ReadableSize(ByVal bytes As Double) As String
    If bytes >= 1073741824# Then
        HAMU_ReadableSize = Format$(bytes / 1073741824#, "0.00") & " GB"
    ElseIf bytes >= 1048576# Then
        HAMU_ReadableSize = Format$(bytes / 1048576#, "0.00") & " MB"
    Else
        HAMU_ReadableSize = Format$(bytes / 1024#, "0.00") & " KB"
    End If
End Function
Public Function HAMU_BuildFolderIndex(ByVal wb As Workbook, ByVal path As String, Optional ByVal includeSubfolders As Boolean = True) As Worksheet
    Dim fso As Object, pending As New Collection, folder As Object, subfolder As Object, file As Object
    Dim sheet As Worksheet, row As Long, skips As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then Err.Raise 5, , HAMU_Text("Klasör bulunamadı.")
    Set sheet = HAMU_GetOrCreateSheet("DOSYALAR_" & Format$(Now, "hhmmss"), wb)
    sheet.range("A1:G1").value = Array(HAMU_Text("Dosya adı"), HAMU_Text("Yerel yol"), HAMU_Text("Uzantı"), HAMU_Text("Tür"), HAMU_Text("Boyut"), HAMU_Text("Oluşturma tarihi"), HAMU_Text("Değiştirme tarihi"))
    row = 2: pending.Add path
    Do While pending.count > 0
        path = CStr(pending(1)): pending.remove 1
        Set folder = Nothing
        On Error Resume Next: Set folder = fso.GetFolder(path): On Error GoTo 0
        If Not folder Is Nothing Then
            On Error GoTo SkipFolder
            For Each file In folder.files
                If row > sheet.rows.count Then Err.Raise 5, , HAMU_Text("Dosya sayısı Excel satır sınırını aşıyor.")
                sheet.Hyperlinks.Add Anchor:=sheet.Cells(row, 1), address:=file.path, TextToDisplay:=file.name
                sheet.Cells(row, 2).Value2 = file.path
                sheet.Cells(row, 3).Value2 = LCase$(fso.GetExtensionName(file.name))
                sheet.Cells(row, 4).Value2 = HAMU_FileKind(fso.GetExtensionName(file.name))
                sheet.Cells(row, 5).Value2 = HAMU_ReadableSize(CDbl(file.Size))
                sheet.Cells(row, 6).value = file.DateCreated
                sheet.Cells(row, 7).value = file.DateLastModified
                row = row + 1
                If row Mod 200 = 0 Then DoEvents
            Next
            If includeSubfolders Then
            For Each subfolder In folder.SubFolders
                If (subfolder.Attributes And 1024) = 0 Then pending.Add subfolder.path
            Next
            End If
        Else
            skips = skips + 1
        End If
ContinueFolder:
        On Error GoTo 0
    Loop
    sheet.columns("F:G").NumberFormat = "dd.mm.yyyy hh:mm"
    sheet.range("A1:G1").Font.Bold = True
    sheet.range("A1:G1").Interior.color = RGB(223, 240, 244)
    sheet.range("A1:G" & row - 1).AutoFilter
    sheet.columns.AutoFit
    If sheet.columns(2).ColumnWidth > 65 Then sheet.columns(2).ColumnWidth = 65
    If skips > 0 Then sheet.Cells(1, 9).value = HAMU_Text("Erişilemeyen klasör sayısı: ") & skips
    Set HAMU_BuildFolderIndex = sheet
    Exit Function
SkipFolder:
    If Err.number = 5 Then Err.Raise Err.number, , Err.description
    skips = skips + 1: Resume ContinueFolder
End Function

Public Sub HAMU_ShowBackupResult(ByVal path As String)
    Dim form As frmHAMU_BackupResult
    Set form = New frmHAMU_BackupResult
    form.Setup path
    form.show vbModal
    Unload form
End Sub

Public Sub HAMU_ExportPDFRange(ByVal source As range, ByVal path As String, Optional ByVal layout As Long = 1)
 Dim Setup As PageSetup, printArea As String, orientation As Long, zoom As Variant, wide As Variant, tall As Variant
 Dim number As Long, description As String, captured As Boolean
 Set Setup = source.Worksheet.PageSetup
 On Error GoTo Failed
 printArea = Setup.printArea: orientation = Setup.orientation
 zoom = Setup.zoom: wide = Setup.FitToPagesWide: tall = Setup.FitToPagesTall
 captured = True
 Setup.printArea = source.address
 If layout = 1 Then
  Setup.orientation = IIf(source.width > source.Height, xlLandscape, xlPortrait)
 Else
  Setup.orientation = IIf(layout = 3, xlLandscape, xlPortrait)
 End If
 Setup.zoom = False: Setup.FitToPagesWide = 1: Setup.FitToPagesTall = False
 source.Worksheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=path, Quality:=xlQualityStandard, IgnorePrintAreas:=False, OpenAfterPublish:=False
Done:
 On Error Resume Next
 If captured Then
  Setup.printArea = printArea: Setup.orientation = orientation
  Setup.zoom = zoom: Setup.FitToPagesWide = wide: Setup.FitToPagesTall = tall
 End If
 On Error GoTo 0
 If number <> 0 Then Err.Raise number, , description
 Exit Sub
Failed:
 number = Err.number: description = Err.description: Resume Done
End Sub

Public Sub HAMU_SaveCopy()
 On Error GoTo Failed
 If Not HAMU_HasWorkbook() Then Exit Sub
 Dim wb As Workbook, path As Variant, extension As String, folder As String
 Set wb = ActiveWorkbook
 extension = HAMU_WorkbookExtension(wb)
 folder = HAMU_Setting("ExportFolder")
 If Len(folder) = 0 Then folder = wb.path
 If Len(folder) > 0 Then folder = folder & Application.PathSeparator
 path = Application.GetSaveAsFilename(folder & HAMU_TemizDosyaAdi(CreateObject("Scripting.FileSystemObject").GetBaseName(wb.name)) & "_KOPYA" & extension, "Excel (*" & extension & "),*" & extension)
 If VarType(path) = vbBoolean Then Exit Sub
 HAMU_SaveWorkbookCopy wb, CStr(path)
 HAMU_ShowInfo HAMU_L("Kopya kaydedildi: ", "Copy saved: ") & CStr(path)
 Exit Sub
Failed:
 HAMU_ShowError HAMU_L("Kopya Kaydet", "Save Copy"), Err.number, Err.description
End Sub
Public Sub HAMU_SaveWorkbookCopy(ByVal wb As Workbook, ByVal path As String)
 Dim fso As Object, extension As String
 Set fso = CreateObject("Scripting.FileSystemObject")
 extension = HAMU_WorkbookExtension(wb)
 If Len(wb.path) = 0 And wb.HasVBProject Then Err.Raise 5, , HAMU_L("Makroları korumak için dosyayı önce makro etkin biçimde kaydedin.", "Save the workbook in a macro-enabled format first to preserve macros.")
 If LCase$(fso.GetExtensionName(path)) <> LCase$(Mid$(extension, 2)) Then Err.Raise 5, , HAMU_L("Dosya uzantısı mevcut biçimle aynı olmalı.", "The file extension must match the current workbook format.")
 If Len(wb.path) > 0 Then
  If StrComp(wb.FullName, path, vbTextCompare) = 0 Then Err.Raise 5, , HAMU_L("Kopya için farklı bir dosya adı seçin.", "Choose a different filename for the copy.")
 End If
 If fso.FileExists(path) Then Err.Raise 5, , HAMU_L("Bu dosya zaten var. Farklı bir ad seçin.", "This file already exists. Choose a different name.")
 wb.SaveCopyAs path
 If Not fso.FileExists(path) Then Err.Raise 5, , HAMU_L("Kopya oluşturulamadı. Kayıt konumunu kontrol edin.", "Could not create the copy. Check the save location.")
End Sub

