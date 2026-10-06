Attribute VB_Name = "modHAMU_Visual"
Option Explicit
Option Private Module

Public Sub HAMU_SelectionToPNG()
 On Error GoTo Failed
 Dim scope As range, kind As Long, extension As String, path As Variant
 If Not HAMU_HasWorkbook() Then Exit Sub
 Set scope = HAMU_WorkRange(HAMU_L("Resim olarak kaydedilecek alanı seçin.", "Select the range to save as an image."))
 If scope Is Nothing Then Exit Sub
 kind = HAMU_ChooseIndex(HAMU_L("Resim Kaydet", "Save Image"), HAMU_L("Dosya biçimini seçin.", "Choose the image format."), Array("PNG", "JPG", "GIF"))
 If kind = 0 Then Exit Sub
 extension = Choose(kind, "png", "jpg", "gif")
 path = Application.GetSaveAsFilename("HAMU_" & Format$(Now, "yyyymmdd_hhnnss") & "." & extension, UCase$(extension) & " (*." & extension & "),*." & extension)
 If VarType(path) = vbBoolean Then Exit Sub
 If LCase$(right$(CStr(path), 4)) <> "." & extension Then path = CStr(path) & "." & extension
 HAMU_ExportRangeImage scope, CStr(path), extension
 HAMU_ShowInfo HAMU_L("Kaydedildi: ", "Saved: ") & CStr(path)
 Exit Sub
Failed:
 HAMU_ShowError HAMU_L("Resim Kaydet", "Save Image"), Err.number, Err.description
End Sub

Public Sub HAMU_InsertPhotosFromFolder()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_InsertPhotosFromFolder") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_PromptRange(HAMU_Text("Klasörden Toplu Fotoğraf Ekle: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim ws As Worksheet, startCell As range, folderPath As String, files As Collection, i As Long, tgt As range
    Set ws = hamuScope.Worksheet
    If TypeName(hamuScope) = "Range" Then Set startCell = hamuScope.Cells(1, 1) Else Set startCell = ws.range("A1")
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = HAMU_Text("Resim klasörünü seçin")
        If .show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With
    Dim mode As Long
    mode = CLng(HAMU_Setting("ImagePlacement"))
    If mode = 0 Then mode = HAMU_ChooseIndex(HAMU_Text("Resim Yerleşimi"), HAMU_Text("Hücre içine: Excel hücre resmi. Hücre üstüne: hücre sınırları içinde hareket eden görsel."), Array(HAMU_Text("Hücre içine"), HAMU_Text("Hücre üstüne")))
    If mode = 0 Then Exit Sub
    Set files = HAMU_GetImageFiles(folderPath)
    If files.count = 0 Then HAMU_ShowInfo HAMU_Text("Klasörde desteklenen resim dosyası yok."): Exit Sub
    If files.count > ws.rows.count - startCell.row + 1 Then Err.Raise 5, , HAMU_Text("Resimler bu başlangıç hücresinden sayfaya sığmıyor.")
    For i = 1 To files.count
        Set tgt = startCell.Offset(i - 1, 0)
        ws.rows(tgt.row).RowHeight = 100
        Call HAMU_PlacePhoto(tgt, CStr(files(i)), mode = 1)
        If Not tgt.Offset(0, 1).HasFormula And IsEmpty(tgt.Offset(0, 1).Value2) Then
            tgt.Offset(0, 1).Value2 = Mid$(CStr(files(i)), InStrRev(CStr(files(i)), Application.PathSeparator) + 1)
        End If
    Next i
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_InsertPhotosFromFolder", hamuErrorNumber, hamuErrorText
End Sub

Private Function HAMU_GetImageFiles(ByVal folderPath As String) As Collection
    Dim col As New Collection, patt As Variant, f As String, base As String
    base = folderPath
    If right$(base, 1) <> "\" Then base = base & Application.PathSeparator
    For Each patt In Array("*.jpg", "*.jpeg", "*.png", "*.gif", "*.bmp", "*.tif", "*.tiff", "*.webp")
        f = Dir$(base & patt)
        Do While Len(f) > 0
            col.Add base & f
            f = Dir$
        Loop
    Next patt
    Set HAMU_GetImageFiles = col
End Function


Public Sub HAMU_CreateQRCode()
    On Error GoTo HAMU_CommandFailed
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_CreateQRCode") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    HAMU_ShowCode True
    Exit Sub
HAMU_CommandFailed:
    HAMU_ShowError "HAMU_CreateQRCode", Err.number, Err.description
End Sub

Public Sub HAMU_CreateBarcode()
    On Error GoTo HAMU_CommandFailed
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_CreateBarcode") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    HAMU_ShowCode False
    Exit Sub
HAMU_CommandFailed:
    HAMU_ShowError "HAMU_CreateBarcode", Err.number, Err.description
End Sub



Public Sub HAMU_PlacePhoto(ByVal target As range, ByVal filePath As String, ByVal insideCell As Boolean)
    Dim shape As Object
    If target.MergeCells Then Set target = target.MergeArea
    Set shape = target.Worksheet.Shapes.AddPicture(filePath, msoFalse, msoTrue, target.Left, target.Top, -1, -1)
    With shape
        .LockAspectRatio = msoTrue
        .width = target.width
        If .Height > target.Height Then .Height = target.Height
        .Left = target.Left + (target.width - .width) / 2
        .Top = target.Top + (target.Height - .Height) / 2
        .Placement = xlMoveAndSize
    End With
    If insideCell Then
        target.Worksheet.Activate: target.Cells(1, 1).Select
        On Error GoTo Unsupported
        shape.Select
        CallByName shape, "PlacePictureInCell", VbMethod
    End If
    Exit Sub
Unsupported:
    On Error Resume Next: shape.Delete: On Error GoTo 0
    Err.Raise vbObjectError + 405, , HAMU_Text("Bu Excel sürümü hücre içine resmi desteklemiyor. Hücre üstüne seçeneğini kullanın.")
End Sub

Public Sub HAMU_ExportRangeImage(ByVal source As range, ByVal path As String, ByVal extension As String)
 Dim chart As ChartObject, savedCell As range, state As New clsHAMU_AppState
 Dim number As Long, description As String
 Set source = HAMU_BoundedRange(source)
 If source.CountLarge > 100000 Or source.width > 10000 Or source.Height > 10000 Then Err.Raise 5, , HAMU_L("Resim alanı çok büyük. Daha küçük bir alan seçin.", "The image range is too large. Select a smaller range.")
 extension = LCase$(extension)
 If extension <> "png" And extension <> "jpg" And extension <> "gif" Then Err.Raise 5, , HAMU_L("PNG, JPG veya GIF seçin.", "Select PNG, JPG or GIF.")
 state.Capture
 Set savedCell = ActiveCell
 On Error GoTo Failed
 source.Worksheet.Activate
 source.Select
 Application.ScreenUpdating = True
 DoEvents
 HAMU_CopyRangePicture source
 Set chart = source.Worksheet.ChartObjects.Add(source.Left, source.Top, source.width, source.Height)
 chart.Activate
 DoEvents
 chart.chart.Paste
 DoEvents
 chart.chart.Refresh
 Dim drawUntil As Double
 drawUntil = Timer
 Do While Timer >= drawUntil And Timer - drawUntil < 0.3
  DoEvents
 Loop
 If Not chart.chart.Export(path, IIf(extension = "jpg", "JPEG", UCase$(extension))) Then Err.Raise 5, , HAMU_L("Excel resim dosyasını oluşturamadı.", "Excel could not export the image.")
Done:
 On Error Resume Next
 If Not chart Is Nothing Then chart.Delete
 Application.CutCopyMode = False
 If Not savedCell Is Nothing Then savedCell.Worksheet.Activate: savedCell.Select
 state.Restore
 On Error GoTo 0
 If number <> 0 Then Err.Raise number, , description
 Exit Sub
Failed:
 number = Err.number: description = Err.description: Resume Done
End Sub

Private Sub HAMU_CopyRangePicture(ByVal source As range)
 Dim attempt As Long, pictureError As Long, started As Double
 For attempt = 1 To 3
  source.Worksheet.Activate
  source.Select
  started = Timer
  Do
   DoEvents
  Loop Until Timer - started >= 0.2 Or Timer < started
  On Error Resume Next
  Err.Clear
  source.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
  pictureError = Err.number
  Err.Clear
  On Error GoTo 0
  If pictureError = 0 Then Exit Sub
 Next attempt
 Err.Raise 5, , HAMU_L("Excel resmi hazırlayamadı. Yeniden deneyin.", "Excel could not prepare the image. Please retry.")
End Sub

