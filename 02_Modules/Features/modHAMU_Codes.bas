Attribute VB_Name = "modHAMU_Codes"
Option Explicit
Option Private Module
Public Sub HAMU_ShowCode(ByVal qr As Boolean)
    On Error GoTo Failed
    Dim target As range, source As range, f As frmHAMU_Code, payload As String, url As String
    Dim model As Long, content As Long, value As String, address As String, image As String
    If Not HAMU_HasWorkbook() Then Exit Sub
    Set target = ActiveCell
    Set f = New frmHAMU_Code
    f.Setup qr, target
    Do
        f.PickTarget = False: f.PickSource = False
        f.show vbModal
        If Not f.PickTarget And Not f.PickSource Then Exit Do
        If f.PickSource Then
            Set source = HAMU_PromptRange(HAMU_L("Toplu kod üretimi için tek bir veri sütunu seçin. Kodlar verilerin sağındaki ilk boş sütuna eklenir; başlığı seçmeyin.", "Select one data column for batch generation. Codes go in the first empty column to the right of the data; exclude the header."))
            If Not source Is Nothing Then
                If source.columns.count <> 1 Then
                    HAMU_ShowInfo HAMU_L("Toplu kod için tek sütun seçin.", "Select a single column for batch codes.")
                Else
                    f.txtSource.text = source.address(External:=True)
                End If
            End If
        Else
        Dim picked As range
        Set picked = HAMU_PromptRange(HAMU_Text("Kod görselinin ekleneceği hücreyi seçin."))
        If Not picked Is Nothing Then f.txtTarget.text = picked.Cells(1, 1).address(External:=True)
        End If
    Loop
    If Not f.accepted Then Unload f: Exit Sub
    qr = f.IsQR
    model = f.cboModel.ListIndex: content = f.cboContent.ListIndex
    value = f.txtValue.text: address = f.txtTarget.text
    Dim sourceAddress As String
    sourceAddress = Trim$(f.txtSource.text)
    Unload f
    If Len(sourceAddress) > 0 Then
        Set source = Application.range(sourceAddress)
        HAMU_CreateCodeBatch source, qr, model, content
        Exit Sub
    End If
    Set target = Application.range(address).Cells(1, 1)
    payload = HAMU_CodePayload(value, content)
    url = HAMU_CodeURL(qr, model, payload)
    image = HAMU_DownloadCode(url)
    Dim shape As shape
    Set shape = HAMU_InsertCodeImage(target, image, qr Or model >= 4)
    Kill image
    Exit Sub
Failed:
    Dim errorNumber As Long, errorText As String
    errorNumber = Err.number: errorText = Err.description
    On Error Resume Next
    If Not f Is Nothing Then Unload f
    If Len(image) > 0 Then Kill image
    HAMU_ShowError HAMU_Text("Kod Oluştur"), errorNumber, errorText
End Sub
Public Function HAMU_CodePayload(ByVal value As String, ByVal content As Long) As String
    value = Trim$(value)
    Select Case content
        Case 1
            If InStr(value, "://") = 0 Then value = "https://" & value
        Case 2
            If InStr(value, "@") = 0 Then Err.Raise 5, , HAMU_Text("Geçerli bir e-posta adresi girin.")
            If LCase$(Left$(value, 7)) <> "mailto:" Then value = "mailto:" & value
        Case 3
            If LCase$(Left$(value, 4)) <> "tel:" Then value = "tel:" & value
    End Select
    HAMU_CodePayload = value
End Function
Public Function HAMU_CodeURL(ByVal qr As Boolean, ByVal model As Long, ByVal payload As String) As String
    Dim models As Variant, i As Long, expected As Long
    If qr Then
        HAMU_CodeURL = "https://quickchart.io/qr?size=240&format=png&text=" & WorksheetFunction.EncodeURL(payload)
        If model = 1 Then HAMU_CodeURL = HAMU_CodeURL & "&dotStyle=dots"
        If model = 2 Then HAMU_CodeURL = HAMU_CodeURL & "&dotStyle=rounded"
    Else
        models = Array("Code128", "Code39", "EAN13", "EAN8", "DataMatrix", "PDF417", "QRCode")
        If model < 0 Or model > 6 Then Err.Raise 5, , HAMU_Text("Kod modeli seçin.")
        If model = 2 Or model = 3 Then
            expected = 13: If model = 3 Then expected = 8
            If Len(payload) <> expected Then Err.Raise 5, , "EAN kodu " & expected & HAMU_Text(" rakam içermelidir.")
            For i = 1 To Len(payload)
                If Not Mid$(payload, i, 1) Like "#" Then Err.Raise 5, , HAMU_Text("EAN yalnızca rakam içerebilir.")
            Next
        End If
        HAMU_CodeURL = "https://barcode.tec-it.com/barcode.ashx?data=" & WorksheetFunction.EncodeURL(payload) & "&code=" & models(model) & "&format=png"
    End If
End Function
Private Function HAMU_DownloadCode(ByVal url As String) As String
    Dim http As Object, stream As Object, path As String
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 5000, 10000, 15000
    http.Open "GET", url, False
    http.send
    If http.Status <> 200 Then Err.Raise 5, , HAMU_Text("Kod servisi yanıt vermedi. İnternet bağlantısını ve içerik/model uyumunu kontrol edin.")
    If InStr(1, http.getResponseHeader("Content-Type"), "image/", vbTextCompare) = 0 Then Err.Raise 5, , HAMU_Text("Servis görsel döndürmedi. Model ve içeriği kontrol edin.")
    path = Environ$("TEMP") & "\HAMU_code_" & Format$(Now, "yyyymmddhhnnss") & "_" & CLng(Timer * 100) & ".png"
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.Write http.responseBody
    stream.SaveToFile path, 2: stream.Close
    HAMU_DownloadCode = path
End Function
Public Function HAMU_InsertCodeImage(ByVal target As range, ByVal imagePath As String, ByVal square As Boolean, Optional ByVal fitCell As Boolean = False) As shape
    Dim w As Single, h As Single
    If target.MergeCells Then Set target = target.MergeArea.Cells(1, 1)
    w = 216: h = 72
    If square Then w = 144: h = 144
    If fitCell Then
        w = target.width - 6: h = target.Height - 6
        If square Then
            If w < h Then h = w Else w = h
        End If
    End If
    Dim prior As shape, key As String
    key = "HAMU_Code_" & target.row & "_" & target.column
    On Error Resume Next
    Set prior = target.Worksheet.Shapes(key)
    On Error GoTo 0
    If Not prior Is Nothing Then prior.Delete
    Set HAMU_InsertCodeImage = target.Worksheet.Shapes.AddPicture(imagePath, msoFalse, msoTrue, target.Left, target.Top, w, h)
    HAMU_InsertCodeImage.name = key
    If fitCell Then HAMU_InsertCodeImage.Left = target.Left + 3: HAMU_InsertCodeImage.Top = target.Top + 3
    HAMU_InsertCodeImage.Placement = xlMoveAndSize
    HAMU_InsertCodeImage.AlternativeText = HAMU_Text("HAMU | Kod görseli")
End Function

Public Function HAMU_CodeCellValue(ByVal cell As range) As String
 If IsError(cell.Value2) Or IsEmpty(cell.Value2) Then Exit Function
 If VarType(cell.Value2) = vbString Then
  HAMU_CodeCellValue = Trim$(CStr(cell.Value2))
 Else
  HAMU_CodeCellValue = CStr(cell.text)
  If InStr(HAMU_CodeCellValue, "#") > 0 Then HAMU_CodeCellValue = CStr(cell.Value2)
 End If
End Function
Public Function HAMU_CodeBatchCount(ByVal source As range) As Long
 Set source = HAMU_BoundedRange(source)
 If source.areas.count <> 1 Or source.columns.count <> 1 Then Err.Raise 5, , HAMU_L("Tek bir kesintisiz veri sütunu seçin.", "Select one contiguous data column.")
 If source.column = source.Worksheet.columns.count Then Err.Raise 5, , HAMU_L("Sağda kod eklemek için boş sütun bulunmalı.", "An empty column is required to the right.")
 If source.CountLarge > 500000 Then Err.Raise 5, , HAMU_L("Kaynak çok büyük. Veri alanını daraltın.", "Source is too large. Narrow the data range.")
 Dim cell As range
 For Each cell In source.Cells
  If Len(HAMU_CodeCellValue(cell)) > 0 Then
   If cell.MergeCells Then Err.Raise 5, , HAMU_L("Birleşik hücrelerde toplu kod oluşturulmaz.", "Batch codes do not support merged cells.")

   HAMU_CodeBatchCount = HAMU_CodeBatchCount + 1
   If HAMU_CodeBatchCount > 250 Then Err.Raise 5, , HAMU_L("Bir işlemde en fazla 250 kod üretilebilir. Daha küçük bir alan seçin.", "Generate at most 250 codes per batch. Select a smaller range.")
  End If
 Next
End Function
Public Sub HAMU_CreateCodeBatch(ByVal source As range, ByVal qr As Boolean, ByVal model As Long, ByVal content As Long)
 Dim cell As range, total As Long, completed As Long, failedCount As Long, errorText As String
 Dim state As New clsHAMU_AppState, value As String, image As String, url As String, shape As shape, target As range, priorCancel As Long, outputColumn As Long
 Set source = HAMU_BoundedRange(source)
 total = HAMU_CodeBatchCount(source)
 outputColumn = HAMU_CodeOutputColumn(source)
 If total = 0 Then Err.Raise 5, , HAMU_L("Kod üretilecek veri bulunamadı.", "No data was found for code generation.")
 If source.Worksheet.ProtectContents Then Err.Raise 5, , HAMU_L("Sayfa korumalı. Önce korumayı kaldırın.", "Unprotect the worksheet first.")
 ' Validate every payload/model before the first network request or sheet edit.
 For Each cell In source.Cells
  value = HAMU_CodeCellValue(cell)
  If Len(value) > 0 Then url = HAMU_CodeURL(qr, model, HAMU_CodePayload(value, content))
 Next
 state.Capture
 priorCancel = Application.EnableCancelKey
 Application.EnableCancelKey = xlErrorHandler
 On Error GoTo Failed
 For Each cell In source.Cells
  value = HAMU_CodeCellValue(cell)
  If Len(value) > 0 Then
   Application.StatusBar = "HAMU: " & completed + failedCount + 1 & " / " & total
   Set target = source.Worksheet.Cells(cell.row, outputColumn)
   url = HAMU_CodeURL(qr, model, HAMU_CodePayload(value, content))
   On Error GoTo ItemFailed
   image = HAMU_DownloadCode(url)
   target.EntireRow.RowHeight = IIf(qr Or model >= 4, 108, 72)
   If target.EntireColumn.ColumnWidth < IIf(qr Or model >= 4, 22, 36) Then target.EntireColumn.ColumnWidth = IIf(qr Or model >= 4, 22, 36)
   Set shape = HAMU_InsertCodeImage(target, image, qr Or model >= 4, True)
   shape.AlternativeText = "HAMU | " & value
   Kill image: image = ""
   completed = completed + 1
  End If
NextItem:
  On Error GoTo Failed
  DoEvents
 Next
BatchDone:
 On Error Resume Next
 If Len(image) > 0 Then Kill image
 image = ""
 On Error GoTo 0
 Application.EnableCancelKey = priorCancel
 state.Restore
 HAMU_ShowInfo HAMU_L("Toplu kod üretimi tamamlandı. Oluşturulan: ", "Batch generation complete. Created: ") & completed & "; " & HAMU_L("başarısız: ", "failed: ") & failedCount & IIf(Len(errorText) > 0, vbCrLf & errorText, "")
 Exit Sub
ItemFailed:
 If Err.number = 18 Then errorText = HAMU_L("İşlem iptal edildi.", "Operation cancelled."): Resume BatchDone
 failedCount = failedCount + 1
 If Len(errorText) = 0 Then errorText = cell.address & ": " & Err.description
 On Error Resume Next
 If Len(image) > 0 Then Kill image
 image = ""
 Resume NextItem
Failed:
 Dim number As Long, description As String
 number = Err.number: description = Err.description
 On Error Resume Next
 If Len(image) > 0 Then Kill image
 Application.EnableCancelKey = priorCancel
 state.Restore
 On Error GoTo 0
 Err.Raise number, , description
End Sub


Public Function HAMU_CodeOutputColumn(ByVal source As range) As Long
 Dim ws As Worksheet, found As range, c As Long, region As range, candidate As range, picture As shape, occupied As Boolean
 Set ws = source.Worksheet
 Set region = ws.range(ws.Cells(source.row, 1), ws.Cells(source.row + source.rows.count - 1, ws.columns.count))
 Set found = region.Find(What:="*", After:=region.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
 c = source.column + 1
 If Not found Is Nothing Then If found.column >= c Then c = found.column + 1
 Do While c <= ws.columns.count
  Set candidate = ws.Cells(source.row, c).Resize(source.rows.count, 1)
  occupied = False
  If Application.CountA(candidate) > 0 Then occupied = True
  If IsNull(candidate.MergeCells) Then occupied = True Else If candidate.MergeCells Then occupied = True
  For Each picture In ws.Shapes
   If Not Intersect(ws.range(picture.TopLeftCell, picture.BottomRightCell), candidate) Is Nothing Then occupied = True: Exit For
  Next picture
  If Not occupied Then HAMU_CodeOutputColumn = c: Exit Function
  c = c + 1
 Loop
 Err.Raise 5, , HAMU_L("Kodlar için sağda boş çıktı sütunu bulunamadı.", "No empty output column is available to the right.")
End Function
