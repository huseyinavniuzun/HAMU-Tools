Attribute VB_Name = "modHAMU_PhotoBatch"
Option Explicit
Option Private Module

Public Sub HAMU_ImportCellPhotoBatch(ByVal target As Range, ByVal files As Collection, Optional ByVal optimize As Boolean = False)
    Dim state As New clsHAMU_AppState, block As Collection, first As Long, last As Long, i As Long
    Dim number As Long, description As String, cancelKey As Long, cancelCaptured As Boolean
    On Error GoTo Failed
    state.Capture
    cancelKey = Application.EnableCancelKey: cancelCaptured = True
    If target Is Nothing Then Err.Raise 5, , HAMU_L("Başlangıç hücresini seçin.", "Select the starting cell.")
    If files Is Nothing Then Err.Raise 5
    If files.Count = 0 Then Exit Sub
    If files.Count > target.Worksheet.Rows.Count - target.Row + 1 Then Err.Raise 5, , HAMU_L("Fotoğraflar bu alana sığmıyor.", "The photos do not fit the destination.")
    If target.Worksheet.ProtectContents Then Err.Raise 1004, , HAMU_L("Sayfa korumalı. Önce korumayı kaldırın.", "Unprotect the destination sheet first.")
    Dim merges As Variant
    merges = target.Cells(1, 1).Resize(files.Count, 1).MergeCells
    If IsNull(merges) Then Err.Raise 5, , HAMU_L("Birleştirilmemiş hücreler seçin.", "Select unmerged cells.")
    If merges Then Err.Raise 5, , HAMU_L("Birleştirilmemiş hücreler seçin.", "Select unmerged cells.")
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlErrorHandler
    ' Bound temporary memory without keeping a worker or image cache alive.
    For first = 1 To files.Count Step 50
        last = first + 49: If last > files.Count Then last = files.Count
        Set block = New Collection
        For i = first To last: block.Add CStr(files(i)): Next
        HAMU_CopyCellPhotoBlock target.Offset(first - 1, 0), block, optimize
        Set block = Nothing
    Next
    Application.EnableCancelKey = cancelKey
    state.Restore
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    If cancelCaptured Then Application.EnableCancelKey = cancelKey
    state.Restore
    Err.Raise number, "HAMU_ImportCellPhotoBatch", description
End Sub

Private Sub HAMU_CopyCellPhotoBlock(ByVal target As Range, ByVal files As Collection, ByVal optimize As Boolean)
    Dim fso As Object, manifest As Object, shell As Object, source As Workbook, destination As Range
    Dim folder As String, manifestPath As String, bookPath As String, seedPath As String, scriptPath As String, script As String, exitCode As Long
    Dim number As Long, description As String, i As Long, height As Variant, row As Range
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.BuildPath(fso.GetSpecialFolder(2).Path, "HAMU_Photos_" & fso.GetTempName)
    fso.CreateFolder folder
    manifestPath = fso.BuildPath(folder, "files.txt"): bookPath = fso.BuildPath(folder, "HAMU_" & fso.GetTempName & ".xlsx")
    seedPath = fso.BuildPath(folder, "seed.xlsx")
    HAMU_PhotoSaveSeed seedPath
    Set manifest = fso.CreateTextFile(manifestPath, True, True)
    For i = 1 To files.Count: manifest.WriteLine CStr(files(i)): Next
    manifest.Close: Set manifest = Nothing
    scriptPath = fso.BuildPath(folder, "prepare.ps1")
    script = "$ErrorActionPreference='Stop';$optimize=$" & LCase$(CStr(optimize)) & ";$manifest='" & Replace$(manifestPath, "'", "''") & "';$output='" & Replace$(bookPath, "'", "''") & "';$seed64=[Convert]::ToBase64String([IO.File]::ReadAllBytes('" & Replace$(seedPath, "'", "''") & "'));" & vbLf
    If optimize Then script = script & CStr(ThisWorkbook.Worksheets("HAMU_PhotoResources").Range("A3").Value2) & vbLf
    script = script & HAMU_PhotoPackageScript
    Set shell = CreateObject("WScript.Shell")
    HAMU_PhotoWriteScript scriptPath, script
    exitCode = shell.Run("powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -File """ & scriptPath & """", 0, True)
    If exitCode <> 0 Or Not fso.FileExists(bookPath) Then Err.Raise 5, , HAMU_L("Fotoğraflar hazırlanamadı. Dosyaların okunabildiğini kontrol edin.", "The photos could not be prepared. Check that the files are readable.")
    Set source = Workbooks.Open(FileName:=bookPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    source.Windows(1).Visible = False
    Set destination = target.Cells(1, 1).Resize(files.Count, 1)
    source.Worksheets(1).Range("A1").Resize(files.Count, 1).Copy
    destination.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    source.Close False: Set source = Nothing
    height = destination.EntireRow.RowHeight
    If IsNull(height) Then
        For Each row In destination.Rows
            If row.RowHeight < 100 Then row.RowHeight = 100
        Next
    ElseIf height < 100 Then
        destination.EntireRow.RowHeight = 100
    End If
Clean:
    On Error Resume Next
    Application.CutCopyMode = False
    If Not source Is Nothing Then source.Close False
    If Not manifest Is Nothing Then manifest.Close
    HAMU_PhotoCleanFolder folder
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "HAMU_CopyCellPhotoBlock", description
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    Resume Clean
End Sub

Private Sub HAMU_PhotoWriteScript(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    stream.WriteText text: stream.SaveToFile path, 2: stream.Close
End Sub

Public Sub HAMU_PhotoCleanFolder(ByVal folder As String)
    Dim fso As Object, item As Object, temp As String
    If Len(folder) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    temp = fso.GetSpecialFolder(2).Path
    If StrComp(fso.GetParentFolderName(folder), temp, vbTextCompare) <> 0 Then Exit Sub
    If Left$(fso.GetFileName(folder), 12) <> "HAMU_Photos_" Then Exit Sub
    If Not fso.FolderExists(folder) Then Exit Sub
    For Each item In fso.GetFolder(folder).Files: item.Delete True: Next
    fso.DeleteFolder folder, False
End Sub

Public Function HAMU_PreparePhotoFiles(ByVal files As Collection, ByRef folder As String) As Collection
    Dim fso As Object, manifest As Object, result As New Collection, item As Variant
    Dim inputPath As String, outputPath As String, scriptPath As String, script As String, code As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.BuildPath(fso.GetSpecialFolder(2).Path, "HAMU_Photos_" & fso.GetTempName)
    fso.CreateFolder folder
    inputPath = fso.BuildPath(folder, "files.txt")
    outputPath = fso.BuildPath(folder, "prepared.txt")
    scriptPath = fso.BuildPath(folder, "optimize.ps1")
    Set manifest = fso.CreateTextFile(inputPath, True, True)
    For Each item In files: manifest.WriteLine CStr(item): Next
    manifest.Close
    script = "$ErrorActionPreference='Stop'" & vbLf & CStr(ThisWorkbook.Worksheets("HAMU_PhotoResources").Range("A3").Value2) & vbLf
    script = script & "$paths=[IO.File]::ReadAllLines('" & Replace$(inputPath, "'", "''") & "',[Text.Encoding]::Unicode);"
    script = script & "$result=Convert-HamuPhotos $paths '" & Replace$(folder, "'", "''") & "';"
    script = script & "[IO.File]::WriteAllLines('" & Replace$(outputPath, "'", "''") & "',$result,[Text.Encoding]::Unicode)"
    HAMU_PhotoWriteScript scriptPath, script
    code = CreateObject("WScript.Shell").Run("powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -File """ & scriptPath & """", 0, True)
    If code <> 0 Or Not fso.FileExists(outputPath) Then Err.Raise 5, , HAMU_L("Fotoğraflar optimize edilemedi. Kaynak dosyaları kontrol edin.", "Photos could not be optimized. Check the source files.")
    Set manifest = fso.OpenTextFile(outputPath, 1, False, -1)
    Do Until manifest.AtEndOfStream: result.Add manifest.ReadLine: Loop
    manifest.Close
    If result.Count <> files.Count Then Err.Raise 5
    Set HAMU_PreparePhotoFiles = result
End Function

Private Sub HAMU_PhotoSaveSeed(ByVal path As String)
    Dim doc As Object, node As Object, stream As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0"): Set node = doc.createElement("base64")
    node.DataType = "bin.base64": node.Text = HAMU_PhotoSeed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.Write node.nodeTypedValue: stream.SaveToFile path, 2: stream.Close
End Sub

Private Function HAMU_PhotoSeed() As String
    HAMU_PhotoSeed = CStr(ThisWorkbook.Worksheets("HAMU_PhotoResources").Range("A1").Value2)
End Function

Private Function HAMU_PhotoPackageScript() As String
    HAMU_PhotoPackageScript = CStr(ThisWorkbook.Worksheets("HAMU_PhotoResources").Range("A2").Value2)
End Function
