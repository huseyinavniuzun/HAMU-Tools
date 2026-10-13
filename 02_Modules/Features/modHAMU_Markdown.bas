Attribute VB_Name = "modHAMU_Markdown"
Option Explicit
Option Private Module
Private Const MD_MAX_CELLS As Long = 50000
Public Sub HAMU_ImportMarkdown()
 On Error GoTo Failed
 If Not HAMU_HasWorkbook() Then Exit Sub
 Dim wb As Workbook, picker As Object, path As String, tables As Collection
 Dim options() As String, i As Long, selected As Long, count As Long, part As Variant
 Set wb = ActiveWorkbook
 Set picker = Application.FileDialog(3)
 With picker
  .Title = HAMU_L("Markdown dosyası seç", "Select a Markdown file")
  .AllowMultiSelect = False: .Filters.Clear: .Filters.Add "Markdown", "*.md;*.markdown"
  If .Show <> -1 Then Exit Sub
  path = .SelectedItems(1)
 End With
 Set picker = Nothing
 Set tables = HAMU_MDParse(HAMU_MDRead(path))
 If tables.Count = 0 Then Err.Raise 5, , HAMU_L("Dosyada başlık ve --- ayırıcı satırı olan bir Markdown tablosu bulunamadı.", "No Markdown table with a header and --- separator was found.")
 If tables.Count = 1 Then
  selected = 1
 Else
  ReDim options(0 To tables.Count)
  options(0) = HAMU_L("Tüm tablolar (ayrı sayfalara)", "All tables (separate sheets)")
  For i = 1 To tables.Count
   part = tables(i)
   options(i) = HAMU_L("Tablo ", "Table ") & i & " — " & UBound(part, 1) - 1 & HAMU_L(" satır / ", " rows / ") & UBound(part, 2) & HAMU_L(" sütun", " columns")
  Next
  selected = HAMU_ChooseIndex(HAMU_L("Markdown Aktar", "Import Markdown"), HAMU_L("Her tablo yeni bir sayfaya alınır. Başlıklar ilk satırdır; değerler telefon ve kodları korumak için metin olarak aktarılır.", "Each table goes into a new sheet. Headers are the first row; values are imported as text to preserve phone numbers and codes."), options)
  If selected = 0 Then Exit Sub
  selected = selected - 1
 End If
 count = HAMU_MDWrite(wb, tables, selected)
 HAMU_ShowInfo count & HAMU_L(" tablo yeni sayfalara aktarıldı.", " table(s) imported into new sheets.")
Clean:
 Set tables = Nothing: Set picker = Nothing: Set wb = Nothing
 HAMU_ReleaseIdleCaches
 Exit Sub
Failed:
 HAMU_ShowError HAMU_L("Markdown Aktar", "Import Markdown"), Err.Number, Err.Description
 Resume Clean
End Sub
Public Function HAMU_MDRead(ByVal path As String) As String
 Dim stream As Object, n As Long, message As String
 On Error GoTo Failed
 If FileLen(path) > 2097152 Then Err.Raise 5, , HAMU_L("Markdown dosyası en fazla 2 MB olabilir.", "The Markdown file must be at most 2 MB.")
 Set stream = CreateObject("ADODB.Stream")
 stream.Type = 2: stream.Charset = "utf-8": stream.Open: stream.LoadFromFile path
 HAMU_MDRead = stream.ReadText
 stream.Close: Set stream = Nothing
 Exit Function
Failed:
 n = Err.Number: message = Err.Description
 On Error Resume Next
 If Not stream Is Nothing Then stream.Close
 On Error GoTo 0
 Err.Raise n, "Markdown", message
End Function
Public Function HAMU_MDParse(ByVal text As String) As Collection
 Dim lines As Variant, i As Long, j As Long, c As Long, row As Variant, head As Variant, sep As Variant
 Dim rows As Collection, result As New Collection, values() As Variant, cells As Long, total As Long
 Dim fence As String, line As String, marker As String, run As Long, columns As Long
 text = Replace$(Replace$(text, vbCrLf, vbLf), vbCr, vbLf)
 If Len(text) > 2097152 Then Err.Raise 5, , HAMU_L("Markdown dosyası en fazla 2 MB olabilir.", "The Markdown file must be at most 2 MB.")
 If Len(text) > 0 Then
  If AscW(Left$(text, 1)) = -257 Then text = Mid$(text, 2)
 End If
 lines = Split(text, vbLf)
 Do While i <= UBound(lines)
  line = Trim$(CStr(lines(i)))
  marker = "": run = 0
  If Len(line) >= 3 Then
   marker = Left$(line, 1)
   If marker = "`" Or marker = "~" Then
    Do While run < Len(line)
     If Mid$(line, run + 1, 1) <> marker Then Exit Do
     run = run + 1
    Loop
   End If
  End If
  If Len(fence) > 0 Then
   If marker = Left$(fence, 1) And run >= Len(fence) Then
    If Len(Trim$(Mid$(line, run + 1))) = 0 Then fence = ""
   End If
   i = i + 1: GoTo NextLine
  End If
  If run >= 3 Then fence = String$(run, marker): i = i + 1: GoTo NextLine
  If i >= UBound(lines) Then Exit Do
  If InStr(line, "|") = 0 Then i = i + 1: GoTo NextLine
  head = MDCells(line): sep = MDCells(Trim$(CStr(lines(i + 1))))
  If Not MDSeparator(sep) Then i = i + 1: GoTo NextLine
  If UBound(head) <> UBound(sep) Then Err.Raise 5, , HAMU_L("Tablo başlığı ve ayırıcı sütun sayıları farklı. Satır: ", "Header and separator column counts differ. Line: ") & i + 1
  columns = UBound(head) + 1
  If columns > 256 Then Err.Raise 5, , HAMU_L("Markdown tablosu en fazla 256 sütun içerebilir.", "Use at most 256 columns per Markdown table.")
  Set rows = New Collection: rows.Add head
  i = i + 2
  Do While i <= UBound(lines)
   line = Trim$(CStr(lines(i)))
   If Len(line) = 0 Or InStr(line, "|") = 0 Or Left$(line, 3) = "```" Or Left$(line, 3) = "~~~" Then Exit Do
   row = MDCells(line)
   If UBound(row) <> UBound(head) Then Err.Raise 5, , HAMU_L("Tablo satırının sütun sayısı başlıkla aynı değil. Satır: ", "The row column count differs from the header. Line: ") & i + 1
   rows.Add row: i = i + 1
   If (rows.Count * columns) + total > MD_MAX_CELLS Then Err.Raise 5, , HAMU_L("Markdown aktarımı en fazla 50.000 hücre içerebilir. Dosyayı daha küçük parçalara ayırın.", "Use at most 50,000 cells per Markdown import. Split larger files.")
  Loop
  cells = rows.Count * columns: total = total + cells
  If total > MD_MAX_CELLS Then Err.Raise 5, , HAMU_L("Markdown aktarımı en fazla 50.000 hücre içerebilir. Dosyayı daha küçük parçalara ayırın.", "Use at most 50,000 cells per Markdown import. Split larger files.")
  ReDim values(1 To rows.Count, 1 To columns)
  For j = 1 To rows.Count
   row = rows(j)
   For c = 1 To columns: values(j, c) = row(c - 1): Next
  Next
  result.Add values
NextLine:
 Loop
 Set HAMU_MDParse = result
End Function
Private Function MDCells(ByVal line As String) As Variant
 Dim parts As New Collection, token As String, i As Long, ch As String, result() As String
 line = Trim$(line)
 If Left$(line, 1) = "|" Then line = Mid$(line, 2)
 For i = 1 To Len(line)
  ch = Mid$(line, i, 1)
  If ch = "\" And i < Len(line) Then
   If Mid$(line, i + 1, 1) = "|" Or Mid$(line, i + 1, 1) = "\" Then
    token = token & Mid$(line, i + 1, 1): i = i + 1: GoTo NextCharacter
   End If
  End If
  If ch = "|" Then
   MDAddCell parts, token: token = ""
  Else
   token = token & ch
  End If
NextCharacter:
 Next
 If ch <> "|" Or Len(token) > 0 Or parts.Count = 0 Then MDAddCell parts, token
 ReDim result(0 To parts.Count - 1)
 For i = 1 To parts.Count: result(i - 1) = parts(i): Next
 MDCells = result
End Function
Private Function MDSeparator(ByVal cells As Variant) As Boolean
 Dim value As Variant, s As String
 For Each value In cells
  s = Trim$(CStr(value))
  If Left$(s, 1) = ":" Then s = Mid$(s, 2)
  If Right$(s, 1) = ":" Then s = Left$(s, Len(s) - 1)
  If Len(s) < 3 Or Len(Replace$(s, "-", "")) > 0 Then Exit Function
 Next
 MDSeparator = True
End Function
Public Function HAMU_MDWrite(ByVal wb As Workbook, ByVal tables As Collection, ByVal selected As Long) As Long
 Dim state As New clsHAMU_AppState, sheet As Worksheet, created As New Collection, i As Long, values As Variant
 Dim data As Range, table As ListObject, n As Long, message As String, cancelKey As Long, captured As Boolean
 On Error GoTo Failed
 state.Capture: captured = True: cancelKey = Application.EnableCancelKey
 If wb Is Nothing Then Err.Raise 5, , HAMU_L("Önce bir çalışma kitabı açın.", "Open a workbook first.")
 If wb.ProtectStructure Then Err.Raise 5, , HAMU_L("Çalışma kitabının yapı korumasını kaldırın.", "Unprotect the workbook structure.")
 If selected < 0 Or selected > tables.Count Then Err.Raise 5, , HAMU_L("Aktarılacak tablo seçimi geçersiz.", "The selected table is invalid.")
 Application.EnableCancelKey = xlErrorHandler
 Application.ScreenUpdating = False: Application.EnableEvents = False
 For i = 1 To tables.Count
  If selected = 0 Or selected = i Then
   values = tables(i)
   Set sheet = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count)): created.Add sheet
   Set data = sheet.Cells(1, 1).Resize(UBound(values, 1), UBound(values, 2))
   data.NumberFormat = "@": data.Value2 = values
   Set table = sheet.ListObjects.Add(xlSrcRange, data, , xlYes)
   table.TableStyle = "TableStyleMedium1"
   data.Columns.AutoFit
   Dim column As Range
   For Each column In data.Columns
    If column.ColumnWidth > 60 Then column.ColumnWidth = 60
   Next
   HAMU_MDWrite = HAMU_MDWrite + 1
  End If
 Next
 Application.EnableCancelKey = cancelKey: state.Restore
 Exit Function
Failed:
 n = Err.Number: message = Err.Description
 On Error Resume Next
 If created.Count > 0 Then Application.DisplayAlerts = False
 For i = created.Count To 1 Step -1: created(i).Delete: Next
 If captured Then Application.EnableCancelKey = cancelKey
 state.Restore
 On Error GoTo 0
 Err.Raise n, "Markdown", message
End Function

Private Sub MDAddCell(ByVal parts As Collection, ByVal text As String)
 If Len(text) > 32767 Then Err.Raise 5, , HAMU_L("Bir hücrenin metni Excel sınırını aşıyor (32.767 karakter).", "A cell exceeds the Excel limit of 32,767 characters.")
 parts.Add Trim$(text)
End Sub
