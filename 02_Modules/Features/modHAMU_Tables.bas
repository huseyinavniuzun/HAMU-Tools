Attribute VB_Name = "modHAMU_Tables"
Option Explicit
Option Private Module

Public Sub HAMU_AutoFit()
    On Error GoTo HAMU_CommandFailed
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_AutoFit") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim r As range
    Set r = HAMU_WorkRange(HAMU_Text("Sığdırılacak veri alanını seçin."))
    If r Is Nothing Then Exit Sub
    r.rows.AutoFit
    Exit Sub
HAMU_CommandFailed:
    HAMU_ShowError "HAMU_AutoFit", Err.number, Err.description
End Sub

Public Sub HAMU_FitRange(ByVal r As range)
    Dim c As range
    r.columns.AutoFit
    For Each c In r.columns
        If c.ColumnWidth > 60 Then c.ColumnWidth = 60
    Next
    r.WrapText = True
    r.rows.AutoFit
End Sub

Public Sub HAMU_AutoTable()
    On Error GoTo HAMU_CommandFailed
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_AutoTable") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim r As range, lo As ListObject
    Set r = HAMU_WorkRange(HAMU_Text("Başlık satırı dahil tablo alanını seçin. Tek hücre için çevresindeki veri kullanılır."))
    If r Is Nothing Then Exit Sub
    If r.CountLarge = 1 Then Set r = r.CurrentRegion
    Set lo = HAMU_CreateTable(r)
    HAMU_ShowInfo HAMU_Text("Tablo hazır: ") & lo.name & vbCrLf & HAMU_Text("Tabloya bitişik satıra veri girerek Excel'in standart genişlemesini kullanabilirsiniz.")
    Exit Sub
HAMU_CommandFailed:
    HAMU_ShowError "HAMU_AutoTable", Err.number, Err.description
End Sub

Public Function HAMU_CreateTable(ByVal r As range) As ListObject
    Dim lo As ListObject, other As ListObject
    If r.areas.count <> 1 Then Err.Raise vbObjectError + 330, , HAMU_Text("Tek bir kesintisiz aralık seçin.")
    If r.Worksheet.ProtectContents Then Err.Raise vbObjectError + 331, , HAMU_Text("Sayfa korumasını kaldırın.")
    If r.rows.count < 2 Then Err.Raise vbObjectError + 332, , HAMU_Text("Başlık ve en az bir veri satırı gerekir.")
    If IsNull(r.MergeCells) Then Err.Raise vbObjectError + 333, , HAMU_Text("Birleşik hücreleri önce çözün.")
    If r.MergeCells Then Err.Raise vbObjectError + 333, , HAMU_Text("Birleşik hücreleri önce çözün.")
    For Each other In r.Worksheet.ListObjects
        If Not Intersect(other.range, r) Is Nothing Then
            If other.range.address = r.address Then
                Set lo = other
            Else
                Err.Raise vbObjectError + 334, , HAMU_Text("Seçim mevcut bir tabloyla kısmen çakışıyor.")
            End If
        End If
    Next
    If lo Is Nothing Then Set lo = r.Worksheet.ListObjects.Add(xlSrcRange, r, , xlYes)
    lo.TableStyle = "TableStyleMedium1"
    lo.ShowTableStyleRowStripes = True
    lo.ShowAutoFilter = True
    With lo.HeaderRowRange
        .Interior.color = RGB(230, 230, 230)
        .Font.color = vbBlack
        .Font.Bold = True
    End With
    HAMU_FitRange lo.range
    Set HAMU_CreateTable = lo
End Function

Public Sub HAMU_ResizeTable()
    On Error GoTo HAMU_CommandFailed
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ResizeTable") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim r As range, lo As ListObject
    Set r = HAMU_WorkRange(HAMU_Text("Genişletilecek tablo içinden bir hücre seçin."))
    If r Is Nothing Then Exit Sub
    Set lo = r.Cells(1, 1).ListObject
    If lo Is Nothing Then
        HAMU_ShowInfo HAMU_Text("Bu hücre bir tablo içinde değil. Önce Otomatik Tablo Oluştur kullanın veya mevcut tablodan bir hücre seçin.")
        Exit Sub
    End If
    HAMU_GrowTable lo
    Exit Sub
HAMU_CommandFailed:
    HAMU_ShowError "HAMU_ResizeTable", Err.number, Err.description
End Sub

Public Sub HAMU_GrowTable(ByVal lo As ListObject)
    Dim lastRow As Long, nextRow As range, target As range, other As ListObject
    If lo.SourceType <> xlSrcRange Then Err.Raise vbObjectError + 336, , HAMU_Text("Bağlantılı tablolar bu araçla genişletilmez.")
    If lo.ShowTotals Then Err.Raise vbObjectError + 337, , HAMU_Text("Genişletmeden önce toplam satırını kapatın.")
    lastRow = lo.range.row + lo.range.rows.count - 1
    Do While lastRow < lo.parent.rows.count
        Set nextRow = lo.parent.Cells(lastRow + 1, lo.range.column).Resize(1, lo.range.columns.count)
        If Application.CountA(nextRow) = 0 Then Exit Do
        lastRow = lastRow + 1
    Loop
    Set target = lo.parent.Cells(lo.range.row, lo.range.column).Resize(lastRow - lo.range.row + 1, lo.range.columns.count)
    For Each other In lo.parent.ListObjects
        If other.name <> lo.name Then
            If Not Intersect(other.range, target) Is Nothing Then Err.Raise vbObjectError + 338, , HAMU_Text("Başka bir tabloyla çakışma var.")
        End If
    Next
    lo.Resize target
    HAMU_FitRange lo.range
End Sub



Public Sub HAMU_StandardFormat()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_StandardFormat") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Standart Sayı ve Tarih Biçimi: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim c As range
    For Each c In hamuScope.Cells
        If Not IsError(c.Value2) Then
            If VarType(c.value) = vbDate Then
                c.NumberFormat = HAMU_DEFAULT_DATE_FMT
            ElseIf IsNumeric(c.Value2) Then
                c.NumberFormat = "0"
            End If
        End If
    Next c
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_StandardFormat", hamuErrorNumber, hamuErrorText
End Sub


Public Sub HAMU_ExpandColumns()
 On Error GoTo Failed
 Dim r As range
 If Not HAMU_HasWorkbook() Then Exit Sub
 Set r = HAMU_WorkRange(HAMU_L("Genişletilecek hücreleri seçin.", "Select cells to fit by width."))
 If r Is Nothing Then Exit Sub
 r.columns.AutoFit
 Exit Sub
Failed:
 HAMU_ShowError "HAMU_ExpandColumns", Err.number, Err.description
End Sub
Public Sub HAMU_TableTotals()
 On Error GoTo Failed
 Dim lo As ListObject, col As ListColumn
 Set lo = HAMU_ActiveTable()
 If lo Is Nothing Then Exit Sub
 lo.ShowTotals = Not lo.ShowTotals
 If lo.ShowTotals Then
  For Each col In lo.ListColumns
   If Not col.DataBodyRange Is Nothing Then
    If Application.count(col.DataBodyRange) > 0 Then col.TotalsCalculation = xlTotalsCalculationSum
   End If
  Next col
 End If
 Exit Sub
Failed:
 HAMU_ShowError "HAMU_TableTotals", Err.number, Err.description
End Sub
Public Function HAMU_ActiveTable() As ListObject
 If Not HAMU_HasWorkbook() Then Exit Function
 If TypeName(ActiveSheet) <> "Worksheet" Then Exit Function
 Set HAMU_ActiveTable = ActiveCell.ListObject
 If HAMU_ActiveTable Is Nothing Then HAMU_ShowInfo HAMU_L("Önce bir tablo içinden hücre seçin.", "Select a cell inside a table first.")
End Function
Public Sub HAMU_CopyVisibleTable()
 On Error GoTo Failed
 Dim lo As ListObject, result As ListObject
 Set lo = HAMU_ActiveTable()
 If lo Is Nothing Then Exit Sub
 Set result = HAMU_VisibleTableCopy(lo)
 Exit Sub
Failed:
 HAMU_ShowError "HAMU_CopyVisibleTable", Err.number, Err.description
End Sub
Public Function HAMU_VisibleTableCopy(ByVal lo As ListObject) As ListObject
 Dim visible As range, area As range, ws As Worksheet, dest As range, r As Long
 If lo.DataBodyRange Is Nothing Then Err.Raise 5, , HAMU_L("Tabloda veri yok.", "The table has no data.")
 If lo.range.CountLarge > 250000 Then Err.Raise 5, , HAMU_L("Tablo çok büyük. Önce veri alanını daraltın.", "The table is too large. Narrow the data range first.")
 On Error Resume Next
 Set visible = lo.DataBodyRange.SpecialCells(xlCellTypeVisible)
 On Error GoTo 0
 If visible Is Nothing Then Err.Raise 5, , HAMU_L("Görünür satır yok.", "No visible rows are available.")
 Set ws = lo.parent.parent.Worksheets.Add(After:=lo.parent)
 lo.HeaderRowRange.copy Destination:=ws.Cells(1, 1)
 r = 2
 Dim row As range
 For Each row In lo.DataBodyRange.rows
  If Not row.EntireRow.Hidden Then
   Set dest = ws.Cells(r, 1).Resize(1, lo.ListColumns.count)
   dest.Value2 = row.Value2
   row.copy: dest.PasteSpecial xlPasteFormats
   r = r + 1
  End If
 Next row
 Application.CutCopyMode = False
 Set HAMU_VisibleTableCopy = HAMU_CreateTable(ws.Cells(1, 1).Resize(r - 1, lo.ListColumns.count))
End Function
