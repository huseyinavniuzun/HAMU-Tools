Attribute VB_Name = "modHAMU_Practical"
Option Explicit
Option Private Module

Public Sub HAMU_AddTextAffixes()
    HAMU_TextToolCommand "affix", "HAMU_AddTextAffixes"
End Sub
Public Sub HAMU_LeadingZeros()
    HAMU_TextToolCommand "zeros", "HAMU_LeadingZeros"
End Sub
Public Sub HAMU_DeduplicateItems()
    HAMU_TextToolCommand "unique", "HAMU_DeduplicateItems"
End Sub
Public Sub HAMU_ExtractTextPart()
    HAMU_TextToolCommand "part", "HAMU_ExtractTextPart"
End Sub
Private Sub HAMU_TextToolCommand(ByVal kind As String, ByVal command As String)
    On Error GoTo Failed
    Dim source As range, target As range, form As frmHAMU_TextTools
    Dim mode As Long, first As String, second As String
    Set source = HAMU_WorkRange()
    If source Is Nothing Then Exit Sub
    If source.columns.count <> 1 Then HAMU_ShowInfo HAMU_Text("Tek bir veri sütunu seçin."): Exit Sub
    Set form = New frmHAMU_TextTools
    form.Setup kind, command
    form.show vbModal
    If Not form.accepted Then Unload form: Exit Sub
    mode = form.cboMode.ListIndex + 1: first = form.txtFirst.text: second = form.txtSecond.text
    Unload form
    Set target = HAMU_PromptRange(HAMU_Text("Sonucun yazılacağı başlangıç hücresini seçin. Kaynak sütun korunur."))
    If target Is Nothing Then Exit Sub
    HAMU_WriteTextResults source, target, kind, mode, first, second
    Exit Sub
Failed:
    HAMU_ShowError command, Err.number, Err.description
End Sub
Public Function HAMU_TextTransform(ByVal value As String, ByVal kind As String, ByVal mode As Long, ByVal first As String, ByVal second As String) As String
    Dim width As Long, pos As Long, finish As Long, pieces As Variant, piece As Variant, dict As Object, result As String
    If Len(value) = 0 Then Exit Function
    Select Case kind
        Case "affix": result = first & value & second
        Case "zeros"
            If mode = 1 Then
                width = CLng(first)
                If width < 1 Or width > 1000 Then Err.Raise 5, , HAMU_Text("Uzunluk 1–1000 arasında olmalıdır.")
                result = value
                If Len(value) < width Then result = String$(width - Len(value), "0") & value
            Else
                result = value
                Do While Len(result) > 1 And Left$(result, 1) = "0": result = Mid$(result, 2): Loop
            End If
        Case "unique"
            If Len(first) = 0 Then Err.Raise 5, , HAMU_Text("Bir ayraç girin. Örneğin virgül veya noktalı virgül.")
            Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = vbTextCompare
            pieces = Split(value, first)
            For Each piece In pieces
                piece = Trim$(CStr(piece))
                If Len(piece) > 0 Then
                    If Not dict.exists(piece) Then
                        dict.Add piece, True
                        If Len(result) > 0 Then result = result & first & " "
                        result = result & piece
                    End If
                End If
            Next
        Case "part"
            Select Case mode
                Case 1, 2
                    width = CLng(first)
                    If width < 1 Or width > 32767 Then Err.Raise 5, , HAMU_Text("Karakter sayısı 1–32767 arasında olmalıdır.")
                    If mode = 1 Then result = Left$(value, width) Else result = right$(value, width)
                Case 3, 4
                    If Len(first) = 0 Then Err.Raise 5, , HAMU_Text("Aranacak işareti girin.")
                    pos = InStr(1, value, first, vbTextCompare)
                    If pos > 0 Then
                        If mode = 3 Then result = Left$(value, pos - 1) Else result = Mid$(value, pos + Len(first))
                    End If
                Case 5
                    If Len(first) = 0 Or Len(second) = 0 Then Err.Raise 5, , HAMU_Text("Başlangıç ve bitiş işaretlerini girin.")
                    pos = InStr(1, value, first, vbTextCompare)
                    If pos > 0 Then
                        pos = pos + Len(first): finish = InStr(pos, value, second, vbTextCompare)
                        If finish > 0 Then result = Mid$(value, pos, finish - pos)
                    End If
            End Select
    End Select
    If Len(result) > 32767 Then Err.Raise 5, , HAMU_Text("Sonuç bir hücrenin metin sınırını aşıyor.")
    HAMU_TextTransform = result
End Function
Public Function HAMU_OutputRange(ByVal target As range, ByVal rows As Long, ByVal columns As Long, Optional ByVal source As range) As range
    Set target = target.Cells(1, 1)
    If rows > target.parent.rows.count - target.row + 1 Or columns > target.parent.columns.count - target.column + 1 Then Err.Raise 5, , HAMU_Text("Sonuç hedef sayfaya sığmıyor. Başka bir başlangıç hücresi seçin.")
    If target.parent.ProtectContents Then Err.Raise 5, , HAMU_Text("Hedef sayfa korumalı.")
    Dim output As range, cell As range
    Set output = target.Resize(rows, columns)
    If Not source Is Nothing Then
        If output.parent Is source.parent Then
            If Not Intersect(output, source) Is Nothing Then Err.Raise 5, , HAMU_Text("Kaynak ve çıktı çakışıyor. Boş bir hedef sütun seçin.")
        End If
    End If
    If IsNull(output.MergeCells) Then Err.Raise 5, , HAMU_Text("Hedefte birleşik hücreler var. Başka bir alan seçin.")
    If output.MergeCells Then Err.Raise 5, , HAMU_Text("Hedefte birleşik hücreler var. Başka bir alan seçin.")
    If Application.CountA(output) > 0 Then
        If Not HAMU_ConfirmOverwrite() Then Exit Function
    End If
    Set HAMU_OutputRange = output
End Function
Public Sub HAMU_WriteTextResults(ByVal source As range, ByVal target As range, ByVal kind As String, ByVal mode As Long, ByVal first As String, ByVal second As String)
    Dim inputData As Variant, output As Variant, row As Long, value As Variant, dest As range
    inputData = source.Value2: ReDim output(1 To source.rows.count, 1 To 1)
    For row = 1 To source.rows.count
        If source.rows.count = 1 Then value = inputData Else value = inputData(row, 1)
        If IsError(value) Then
            output(row, 1) = value
        Else
            output(row, 1) = HAMU_TextTransform(CStr(value), kind, mode, first, second)
        End If
    Next
    Set dest = HAMU_OutputRange(target, source.rows.count, 1, source)
    If dest Is Nothing Then Exit Sub
    dest.NumberFormat = "@": dest.Value2 = output
End Sub

Public Sub HAMU_SelectSpecialCells()
    On Error GoTo Failed
    Dim source As range, chosen As range, choice As Long
    Set source = HAMU_WorkRange(): If source Is Nothing Then Exit Sub
    choice = HAMU_ChooseIndex(HAMU_Text("HAMU | Özel Hücreleri Seç"), HAMU_Text("Seçili alan içinde hangi hücreler seçilsin? Değerler ve biçimler değişmez."), Array(HAMU_Text("Formül hücreleri"), HAMU_Text("Sabit değerler"), HAMU_Text("Hata hücreleri"), HAMU_Text("Boş hücreler"), HAMU_Text("Görünür hücreler")))
    If choice = 0 Then Exit Sub
    Set chosen = HAMU_SpecialCellRange(source, choice)
    If chosen Is Nothing Then HAMU_ShowInfo HAMU_Text("Bu türde hücre bulunamadı."): Exit Sub
    chosen.parent.Activate: chosen.Select
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_SelectSpecialCells", Err.number, Err.description
End Sub
Public Function HAMU_SpecialCellRange(ByVal source As range, ByVal choice As Long) As range
    Dim found As range, other As range
    ' SpecialCells on one cell expands to UsedRange; Intersect restores the requested scope.
    On Error Resume Next
    Select Case choice
        Case 1: Set found = source.SpecialCells(xlCellTypeFormulas)
        Case 2: Set found = source.SpecialCells(xlCellTypeConstants)
        Case 3
            Set found = source.SpecialCells(xlCellTypeFormulas, xlErrors)
            Set other = source.SpecialCells(xlCellTypeConstants, xlErrors)
            If found Is Nothing Then
                Set found = other
            ElseIf Not other Is Nothing Then
                Set found = Union(found, other)
            End If
        Case 4
            If source.CountLarge = 1 Then
                If IsEmpty(source.Value2) And Not source.HasFormula Then Set found = source
            Else
                Set found = source.SpecialCells(xlCellTypeBlanks)
            End If
        Case 5: Set found = source.SpecialCells(xlCellTypeVisible)
    End Select
    On Error GoTo 0
    If Not found Is Nothing Then Set HAMU_SpecialCellRange = Intersect(source, found)
End Function

Public Sub HAMU_FormulaAudit()
    On Error GoTo Failed
    If Not HAMU_HasWorkbook() Then HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın."): Exit Sub
    Dim report As Worksheet
    Set report = HAMU_BuildFormulaAudit(ActiveWorkbook)
    report.Activate
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_FormulaAudit", Err.number, Err.description
End Sub
Public Function HAMU_BuildFormulaAudit(ByVal wb As Workbook) As Worksheet
    Dim sources As New Collection, sheet As Worksheet, formulas As range, cell As range, report As Worksheet, row As Long
    For Each sheet In wb.Worksheets: sources.Add sheet: Next
    Set report = wb.Worksheets.Add(After:=wb.sheets(wb.sheets.count))
    report.name = HAMU_Data_UniqueSheetName(wb, "FORMUL_DENETIMI")
    report.range("A1:F1").value = Array("Sayfa", HAMU_Text("Hücre"), HAMU_Text("Formül"), HAMU_Text("Sonuç"), "Durum", HAMU_Text("Dış kitap başvurusu"))
    report.columns("C:D").NumberFormat = "@": row = 2
    For Each sheet In sources
        Set formulas = Nothing
        On Error Resume Next: Set formulas = sheet.UsedRange.SpecialCells(xlCellTypeFormulas): On Error GoTo 0
        If Not formulas Is Nothing Then
            For Each cell In formulas.Cells
                If row > report.rows.count Then Err.Raise 5, , HAMU_Text("Formül sayısı Excel satır sınırını aşıyor.")
                report.Cells(row, 1).Value2 = sheet.name
                report.Hyperlinks.Add report.Cells(row, 2), "", "'" & Replace(sheet.name, "'", "''") & "'!" & cell.address, , cell.address(False, False)
                report.Cells(row, 3).Value2 = "'" & CStr(cell.formula)
                report.Cells(row, 4).Value2 = cell.text
                If IsError(cell.Value2) Then report.Cells(row, 5).Value2 = "Hata" Else report.Cells(row, 5).Value2 = "Normal"
                If HAMU_IsExternalFormula(CStr(cell.formula)) Then report.Cells(row, 6).Value2 = "Var" Else report.Cells(row, 6).Value2 = "Yok"
                row = row + 1
            Next
        End If
    Next
    HAMU_FormatPracticalReport report, row - 1, 6
    Set HAMU_BuildFormulaAudit = report
End Function
Public Function HAMU_IsExternalFormula(ByVal formula As String) As Boolean
    ' Require a bracketed workbook before a sheet reference, ignoring quoted string literals.
    Dim clean As String, i As Long, quoted As Boolean, token As String, closing As Long, bang As Long
    For i = 1 To Len(formula)
        token = Mid$(formula, i, 1)
        If token = """" Then quoted = Not quoted
        If Not quoted And token <> """" Then clean = clean & token
    Next
    i = InStr(clean, "[")
    Do While i > 0
        closing = InStr(i + 1, clean, "]"): bang = InStr(closing + 1, clean, "!")
        If closing > i And bang > closing Then
            token = Mid$(clean, i + 1, closing - i - 1)
            If InStr(1, token, ".xls", vbTextCompare) > 0 Or InStr(1, token, ".csv", vbTextCompare) > 0 Then HAMU_IsExternalFormula = True: Exit Function
        End If
        i = InStr(i + 1, clean, "[")
    Loop
End Function
Public Sub HAMU_ListDefinedNames()
    On Error GoTo Failed
    If Not HAMU_HasWorkbook() Then HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın."): Exit Sub
    Dim report As Worksheet
    Set report = HAMU_BuildNamesReport(ActiveWorkbook): report.Activate
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_ListDefinedNames", Err.number, Err.description
End Sub
Public Function HAMU_BuildNamesReport(ByVal wb As Workbook) As Worksheet
    Dim report As Worksheet, nm As name, row As Long
    Set report = wb.Worksheets.Add(After:=wb.sheets(wb.sheets.count))
    report.name = HAMU_Data_UniqueSheetName(wb, "TANIMLI_ADLAR")
    report.range("A1:D1").value = Array("Ad", HAMU_Text("Başvuru / ifade"), HAMU_Text("Görünür"), "Durum")
    report.columns("B").NumberFormat = "@": row = 2
    For Each nm In wb.names
        report.Cells(row, 1).Value2 = nm.name
        report.Cells(row, 2).Value2 = "'" & nm.RefersTo
        If nm.visible Then report.Cells(row, 3).Value2 = "Evet" Else report.Cells(row, 3).Value2 = HAMU_Text("Hayır")
        If InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Then report.Cells(row, 4).Value2 = HAMU_Text("Geçersiz başvuru") Else report.Cells(row, 4).Value2 = HAMU_Text("Tanımlı")
        row = row + 1
    Next
    HAMU_FormatPracticalReport report, row - 1, 4
    Set HAMU_BuildNamesReport = report
End Function
Private Sub HAMU_FormatPracticalReport(ByVal report As Worksheet, ByVal rows As Long, ByVal columns As Long)
    report.range(report.Cells(1, 1), report.Cells(1, columns)).Font.Bold = True: report.range(report.Cells(1, 1), report.Cells(1, columns)).Interior.color = RGB(221, 237, 245)
    report.range(report.Cells(1, 1), report.Cells(rows, columns)).AutoFilter
    report.columns.AutoFit
    Dim column As range
    For Each column In report.UsedRange.columns
        If column.ColumnWidth > 65 Then column.ColumnWidth = 65: column.WrapText = True
    Next
End Sub

Public Sub HAMU_DateRangeList()
    On Error GoTo Failed
    Dim form As frmHAMU_DateRange, start As Date, finish As Date, weekdays As Boolean, target As range, values As Variant, count As Long
    If Not HAMU_HasWorkbook() Then HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın."): Exit Sub
    Set form = New frmHAMU_DateRange
    form.show vbModal
    If Not form.accepted Then Unload form: Exit Sub
    start = form.StartDate: finish = form.EndDate: weekdays = form.chkWeekdays.value
    Unload form
    values = HAMU_DateListValues(start, finish, weekdays, count)
    If count = 0 Then HAMU_ShowInfo HAMU_Text("Bu aralıkta hafta içi günü yok."): Exit Sub
    Set target = HAMU_PromptRange(HAMU_Text("Tarih listesinin başlangıç hücresini seçin."))
    If target Is Nothing Then Exit Sub
    Set target = HAMU_OutputRange(target, count, 1)
    If target Is Nothing Then Exit Sub
    target.Value2 = values: target.NumberFormat = "dd.mm.yyyy"
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_DateRangeList", Err.number, Err.description
End Sub
Public Function HAMU_DateListValues(ByVal start As Date, ByVal finish As Date, ByVal weekdays As Boolean, ByRef count As Long) As Variant
    Dim days As Long, index As Long, day As Date, values As Variant
    start = DateValue(start): finish = DateValue(finish)
    days = DateDiff("d", start, finish)
    If days < 0 Then Err.Raise 5, , HAMU_Text("Bitiş tarihi başlangıç tarihinden önce olamaz.")
    If days > 100000 Then Err.Raise 5, , HAMU_Text("Tarih aralığı çok geniş. En fazla 100.001 gün seçin.")
    count = 0
    For index = 0 To days
        day = DateAdd("d", index, start)
        If Not weekdays Or Weekday(day, vbMonday) < 6 Then count = count + 1
    Next
    If count = 0 Then Exit Function
    ReDim values(1 To count, 1 To 1): count = 0
    For index = 0 To days
        day = DateAdd("d", index, start)
        If Not weekdays Or Weekday(day, vbMonday) < 6 Then count = count + 1: values(count, 1) = CDbl(day)
    Next
    HAMU_DateListValues = values
End Function

