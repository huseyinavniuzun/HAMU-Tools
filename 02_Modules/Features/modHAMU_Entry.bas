Attribute VB_Name = "modHAMU_Entry"
Option Explicit
Option Private Module
Private mHighlights As Object
Private mHighlightBooksSignature As String
Private mHighlightBooksKnown As Boolean
Public Sub HAMU_DataEntry()
 On Error GoTo Failed
 If Not HAMU_HasWorkbook() Then Exit Sub
 Dim Data As range
 Set Data = HAMU_DataEntryRange(ActiveCell)
 Data.Cells(1, 1).Select
 ' Restore the dispatcher flags before the native modal dialog.
 Application.ScreenUpdating = True: Application.EnableEvents = True: Application.DisplayAlerts = True
 Data.Worksheet.ShowDataForm
 Exit Sub
Failed:
 HAMU_ShowError HAMU_L("Veri Girişi", "Data Entry"), Err.number, Err.description
End Sub
Public Function HAMU_DataEntryRange(ByVal cell As range) As range
 Dim Data As range, table As ListObject, header As range, seen As Object
 If TypeName(cell.parent) <> "Worksheet" Then Err.Raise 5
 On Error Resume Next: Set table = cell.ListObject: On Error GoTo 0
 If table Is Nothing Then Set Data = cell.CurrentRegion Else Set Data = table.range
 Set Data = HAMU_BoundedRange(Data)
 If Data.columns.count > 32 Then Err.Raise 5, , HAMU_L("Excel veri giriş formu en fazla 32 sütun destekler.", "Excel's data form supports at most 32 columns.")
 Set seen = CreateObject("Scripting.Dictionary"): seen.CompareMode = vbTextCompare
 For Each header In Data.rows(1).Cells
  If IsError(header.Value2) Then Err.Raise 5, , HAMU_L("Sütun başlıkları metin olmalı.", "Column headers must be text.")
  If Len(Trim$(CStr(header.Value2))) = 0 Then Err.Raise 5, , HAMU_L("Veri girişinden önce boş sütun başlıklarını doldurun.", "Fill blank column headers before entering data.")
  If seen.exists(CStr(header.Value2)) Then Err.Raise 5, , HAMU_L("Sütun başlıkları benzersiz olmalı.", "Column headers must be unique.")
  seen.Add CStr(header.Value2), True
 Next
 Set HAMU_DataEntryRange = Data
End Function
Public Sub HAMU_HighlightCell(ByVal cell As range, ByVal color As Long)
 Dim key As String
 If mHighlights Is Nothing Then
  Set mHighlights = CreateObject("Scripting.Dictionary")
  mHighlightBooksKnown = False
 End If
 key = cell.address(External:=True)
 If Not mHighlights.exists(key) Then
  If mHighlights.count >= 50000 Then Err.Raise 5, , HAMU_L("Vurgu belleği 50.000 hücre sınırına ulaştı. Önce Vurguyu Kaldır aracını kullanın.", "Highlight history reached 50,000 cells. Use Clear Highlights first.")
  mHighlights.Add key, Array(cell.Interior.color, cell.Interior.pattern, cell.Interior.TintAndShade)
 End If
 cell.Interior.color = color
End Sub
Public Sub HAMU_ClearHighlights()
 On Error GoTo Failed
 Dim source As range
 Set source = HAMU_WorkRange(HAMU_L("Vurgu renklerini kaldıracağınız alanı seçin.", "Select the range whose highlight colors should be removed."))
 If source Is Nothing Then Exit Sub
 HAMU_RemoveHighlightColors source
 Exit Sub
Failed:
 HAMU_ShowError HAMU_L("Vurguyu Kaldır", "Clear Highlights"), Err.number, Err.description
End Sub
Public Sub HAMU_RemoveHighlightColors(ByVal source As range)
 Dim cell As range, key As String, entry As Variant
 For Each cell In source.Cells
  key = cell.address(External:=True)
  If Not mHighlights Is Nothing Then
   If mHighlights.exists(key) Then
    entry = mHighlights(key)
    cell.Interior.color = entry(0): cell.Interior.pattern = entry(1): cell.Interior.TintAndShade = entry(2)
    mHighlights.remove key
    GoTo NextCell
   End If
  End If
  ' Previously saved HAMU highlights can only be recognized by their palette.
  Select Case cell.Interior.color
   Case RGB(255, 242, 0), RGB(255, 199, 44), RGB(255, 230, 153), RGB(255, 199, 206), RGB(226, 239, 218): cell.Interior.pattern = xlNone
  End Select
NextCell:
 Next
 If Not mHighlights Is Nothing Then
  If mHighlights.count = 0 Then Set mHighlights = Nothing
 End If
End Sub


Public Sub HAMU_PruneHighlightCache()
 If mHighlights Is Nothing Then Exit Sub
 Dim openBooks As Object, book As Workbook, key As Variant, first As Long, last As Long, bookName As String, signature As String
 For Each book In Application.Workbooks: signature = signature & book.name & vbNullChar: Next
 If mHighlightBooksKnown And signature = mHighlightBooksSignature Then Exit Sub
 Set openBooks = CreateObject("Scripting.Dictionary"): openBooks.CompareMode = vbTextCompare
 For Each book In Application.Workbooks: openBooks(book.name) = True: Next
 For Each key In mHighlights.keys
  first = InStrRev(CStr(key), "["): last = InStr(first + 1, CStr(key), "]")
  If first > 0 And last > first Then
   bookName = Mid$(CStr(key), first + 1, last - first - 1)
   If Not openBooks.exists(bookName) Then mHighlights.remove key
  End If
 Next
 mHighlightBooksSignature = signature: mHighlightBooksKnown = True
 If mHighlights.count = 0 Then Set mHighlights = Nothing
End Sub

