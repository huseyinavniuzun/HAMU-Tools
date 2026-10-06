Attribute VB_Name = "modHAMU_Sheet"
Option Explicit
Option Private Module

Public Sub HAMU_SheetIndex()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SheetIndex") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    On Error GoTo ErrH
    Dim wb As Workbook, ws As Worksheet, toc As Worksheet, i As Long, r As Long
    Set wb = ActiveWorkbook
    If HAMU_SheetExists(HAMU_Text("İçindekiler"), wb) Then wb.Worksheets(HAMU_Text("İçindekiler")).Delete
    Set toc = wb.Worksheets.Add(Before:=wb.Worksheets(1))
    toc.name = HAMU_Text("İçindekiler")
    toc.range("A2:C2").value = Array("No", HAMU_Text("Sayfa Adı"), HAMU_Text("Bağlantı"))
    r = 3
    For i = 1 To wb.Worksheets.count
        Set ws = wb.Worksheets(i)
        If ws.visible = xlSheetVisible And ws.name <> toc.name Then
            toc.Cells(r, 1).value = r - 2
            toc.Cells(r, 2).value = ws.name
            toc.Hyperlinks.Add Anchor:=toc.Cells(r, 3), address:="", subAddress:="'" & ws.name & "'!A1", TextToDisplay:="Git"
            r = r + 1
        End If
    Next i
    toc.columns.AutoFit
    If toc.columns(1).ColumnWidth < 6 Then toc.columns(1).ColumnWidth = 6
    If toc.columns(2).ColumnWidth < 20 Then toc.columns(2).ColumnWidth = 20
    If toc.columns(3).ColumnWidth < 12 Then toc.columns(3).ColumnWidth = 12
    With toc.range("A1:C1")
        .Interior.color = RGB(230, 230, 230)
        .Font.color = vbBlack
        .RowHeight = 30
    End With
    toc.range("B1:C1").Merge
    With toc.range("B1")
        .value = HAMU_L("İçindekiler", "Contents")
        .Font.Size = 14: .Font.Bold = True
        .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
    End With
    toc.range("A2:C2").Interior.color = RGB(230, 230, 230)
    toc.range("A2:C2").Font.color = vbBlack
    HAMU_IndexLogo toc
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_SheetIndex", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_SortSheets()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SortSheets") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    On Error GoTo ErrH
    Dim i As Long, j As Long
    For i = 1 To Worksheets.count - 1
        For j = i + 1 To Worksheets.count
            If UCase$(Worksheets(j).name) < UCase$(Worksheets(i).name) Then Worksheets(j).Move Before:=Worksheets(i)
        Next j
    Next i
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_SortSheets", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_ShowHidden()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ShowHidden") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    On Error GoTo ErrH
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.visible = xlSheetVisible
        ws.rows.Hidden = False
        ws.columns.Hidden = False
    Next ws
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_ShowHidden", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_DynamicValidation()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_DynamicValidation") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    On Error GoTo Failed
    Dim target As range, source As range, f As frmHAMU_Validation
    Dim labels As New Collection, expressions As New Collection
    Dim ws As Worksheet, table As ListObject, column As ListColumn, nm As name
    Dim expr As String, chosen As Long
    Set target = HAMU_WorkRange()
    If target Is Nothing Then Exit Sub
    labels.Add HAMU_Text("Sayfadan liste alanı seç"): expressions.Add ""
    For Each ws In target.parent.parent.Worksheets
        For Each table In ws.ListObjects
            For Each column In table.ListColumns
                labels.Add table.name & " / " & column.name
                expressions.Add "=" & table.name & "[" & column.name & "]"
            Next
        Next
    Next
    For Each nm In target.parent.parent.names
        On Error Resume Next
        Set source = Nothing: Set source = nm.RefersToRange
        On Error GoTo Failed
        If Not source Is Nothing Then
            labels.Add nm.name: expressions.Add nm.RefersTo
        End If
    Next
    Set f = New frmHAMU_Validation
    f.Setup target, labels
    Do
        f.PickSource = False
        f.show vbModal
        If f.PickSource Or (f.accepted And f.cboSource.ListIndex = 0) Then
            f.accepted = False
            Set source = HAMU_PromptRange(HAMU_Text("Liste değerlerini içeren tek sütunu veya satırı seçin."))
            If Not source Is Nothing Then
                If source.rows.count > 1 And source.columns.count > 1 Then
                    HAMU_ShowInfo HAMU_Text("Liste kaynağı tek sütun veya tek satır olmalıdır.")
                Else
                    labels.Add HAMU_Text("Seçilen alan: ") & source.address(External:=True)
                    expressions.Add "=" & source.address(External:=True)
                    f.cboSource.AddItem labels(labels.count)
                    f.cboSource.ListIndex = labels.count - 1
                    f.lblSelected.Caption = source.address(External:=True)
                End If
            End If
        Else
            Exit Do
        End If
    Loop
    If Not f.accepted Then Unload f: Exit Sub
    chosen = f.cboSource.ListIndex + 1: expr = expressions(chosen)
    Unload f
    HAMU_ApplyValidation target, expr
    HAMU_ShowInfo HAMU_Text("Açılır liste oluşturuldu.")
    Exit Sub
Failed:
    Dim number As Long, description As String
    number = Err.number: description = Err.description
    On Error Resume Next
    If Not f Is Nothing Then Unload f
    HAMU_ShowError HAMU_Text("Veri Doğrulama"), number, description
End Sub
Public Sub HAMU_ApplyValidation(ByVal target As range, ByVal sourceFormula As String)
    Dim wb As Workbook, nm As name, NameText As String, i As Long
    Set wb = target.Worksheet.parent
    i = 1
    Do
        NameText = "HAMU_Liste_" & i
        Set nm = Nothing
        On Error Resume Next: Set nm = wb.names(NameText): On Error GoTo 0
        If nm Is Nothing Then Exit Do
        i = i + 1
    Loop
    Set nm = wb.names.Add(name:=NameText, RefersTo:=sourceFormula, visible:=False)
    On Error GoTo Failed
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & NameText
        .IgnoreBlank = True: .InCellDropdown = True
        .ShowError = True: .ErrorTitle = HAMU_Text("Geçersiz değer")
        .ErrorMessage = HAMU_Text("Listeden bir değer seçin.")
    End With
    Exit Sub
Failed:
    Dim errorNumber As Long, errorText As String
    errorNumber = Err.number: errorText = Err.description
    nm.Delete
    Err.Raise errorNumber, , errorText
End Sub

Private Function HAMU_GetValidationFormula(ByVal SourceExpression As String) As String

    Dim src As String
    Dim nm As name
    Dim simpleName As String

    src = Trim$(SourceExpression)
    If Len(src) = 0 Then Exit Function
    Set nm = HAMU_FindWorkbookName(src)

    If Not nm Is Nothing Then
        simpleName = HAMU_SimpleName(nm)
        HAMU_GetValidationFormula = "=" & simpleName
        Exit Function
    End If
    If InStr(1, src, "[", vbTextCompare) > 1 And right$(src, 1) = "]" Then
        HAMU_GetValidationFormula = HAMU_TableValidationFormula(src)
        Exit Function
    End If
    If right$(src, 1) = "#" Then
        HAMU_GetValidationFormula = HAMU_SpillValidationFormula(src)
        Exit Function
    End If
    HAMU_GetValidationFormula = HAMU_RangeValidationFormula(src)
End Function

Private Function HAMU_FindWorkbookName(ByVal NameText As String) As name

    Dim nm As name
    Dim localName As String

    For Each nm In ActiveWorkbook.names
        localName = HAMU_SimpleName(nm)

        If StrComp(Trim$(localName), Trim$(NameText), vbTextCompare) = 0 Then
            Set HAMU_FindWorkbookName = nm
            Exit Function
        End If
    Next nm
End Function

Private Function HAMU_SimpleName(ByVal nm As name) As String

    Dim s As String
    Dim p As Long

    s = nm.NameLocal
    p = InStrRev(s, "!")

    If p > 0 Then s = Mid$(s, p + 1)

    HAMU_SimpleName = s
End Function

Private Function HAMU_TableValidationFormula(ByVal TableExpression As String) As String

    On Error GoTo Fail

    Dim p As Long
    Dim tableName As String
    Dim columnName As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lc As ListColumn
    Dim safeName As String
    Dim formulaText As String

    p = InStr(1, TableExpression, "[")
    If p <= 1 Then GoTo Fail

    tableName = Left$(TableExpression, p - 1)
    columnName = Mid$(TableExpression, p + 1, Len(TableExpression) - p - 1)

    For Each ws In ActiveWorkbook.Worksheets
        Set lo = Nothing

        On Error Resume Next
        Set lo = ws.ListObjects(tableName)
        On Error GoTo Fail

        If Not lo Is Nothing Then
            Set lc = Nothing

            On Error Resume Next
            Set lc = lo.ListColumns(columnName)
            On Error GoTo Fail

            If Not lc Is Nothing Then Exit For
        End If
    Next ws

    If lo Is Nothing Then GoTo Fail
    If lc Is Nothing Then GoTo Fail

    safeName = "HAMU_DV_" & _
               HAMU_CleanValidationName(tableName) & "_" & _
               HAMU_CleanValidationName(columnName)

    formulaText = "=" & tableName & "[" & columnName & "]"

    HAMU_CreateOrReplaceValidationName safeName, formulaText
    HAMU_TableValidationFormula = "=" & safeName
    Exit Function

Fail:
    HAMU_TableValidationFormula = ""
End Function

Private Function HAMU_SpillValidationFormula(ByVal SpillExpression As String) As String

    On Error GoTo Fail

    Dim src As String
    Dim safeName As String
    Dim formulaText As String

    src = Trim$(SpillExpression)
    If Left$(src, 1) = "=" Then src = Mid$(src, 2)

    safeName = "HAMU_DV_SPILL_" & Format$(Now, "hhmmss")
    formulaText = "=" & src

    HAMU_CreateOrReplaceValidationName safeName, formulaText
    HAMU_SpillValidationFormula = "=" & safeName
    Exit Function

Fail:
    HAMU_SpillValidationFormula = ""
End Function

Private Function HAMU_RangeValidationFormula(ByVal RangeExpression As String) As String

    On Error GoTo Fail

    Dim rng As range
    Dim safeName As String

    Set rng = HAMU_ResolveNormalRange(RangeExpression)
    If rng Is Nothing Then GoTo Fail

    If rng.Worksheet.name = ActiveSheet.name Then
        HAMU_RangeValidationFormula = "=" & _
            rng.address(RowAbsolute:=True, ColumnAbsolute:=True)
    Else
        safeName = "HAMU_DV_RANGE_" & Format$(Now, "hhmmss")

        HAMU_CreateOrReplaceValidationName _
            safeName, _
            "=" & rng.address( _
                RowAbsolute:=True, _
                ColumnAbsolute:=True, _
                ReferenceStyle:=xlA1, _
                External:=True)

        HAMU_RangeValidationFormula = "=" & safeName
    End If

    Exit Function

Fail:
    HAMU_RangeValidationFormula = ""
End Function

Private Function HAMU_ResolveNormalRange(ByVal RangeExpression As String) As range

    On Error Resume Next

    Dim src As String
    Dim p As Long
    Dim sheetName As String
    Dim addressText As String
    Dim ws As Worksheet

    src = Trim$(RangeExpression)
    If Left$(src, 1) = "=" Then src = Mid$(src, 2)

    p = InStrRev(src, "!")

    If p > 0 Then
        sheetName = Left$(src, p - 1)
        addressText = Mid$(src, p + 1)
        sheetName = Replace$(sheetName, "'", "")

        Set ws = ActiveWorkbook.Worksheets(sheetName)

        If Not ws Is Nothing Then
            Set HAMU_ResolveNormalRange = ws.range(addressText)
        End If
    Else
        Set HAMU_ResolveNormalRange = ActiveSheet.range(src)
    End If

    On Error GoTo 0
End Function

Private Sub HAMU_CreateOrReplaceValidationName( _
    ByVal NameText As String, _
    ByVal RefersToText As String)

    On Error Resume Next
    ActiveWorkbook.names(NameText).Delete
    On Error GoTo 0

    ActiveWorkbook.names.Add _
        name:=NameText, _
        RefersTo:=RefersToText, _
        visible:=False
End Sub

Private Function HAMU_CleanValidationName(ByVal textValue As String) As String

    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    s = textValue

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)

        If (ch >= "A" And ch <= "Z") Or _
           (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Or _
           ch = "_" Then

            out = out & ch
        Else
            out = out & "_"
        End If
    Next i

    HAMU_CleanValidationName = out
End Function

Private Function HAMU_ResolveValidationSource(ByVal expr As String) As range
    Dim nm As name, ws As Worksheet, lo As ListObject, lc As ListColumn, p As Long, tName As String, cName As String
    p = InStr(1, expr, "[")
    If p > 1 And right$(expr, 1) = "]" Then
        tName = Left$(expr, p - 1)
        cName = Mid$(expr, p + 1, Len(expr) - p - 1)
        For Each ws In ActiveWorkbook.Worksheets
            For Each lo In ws.ListObjects
                If StrComp(lo.name, tName, vbTextCompare) = 0 Then
                    For Each lc In lo.ListColumns
                        If StrComp(lc.name, cName, vbTextCompare) = 0 Then Set HAMU_ResolveValidationSource = lc.DataBodyRange: Exit Function
                    Next lc
                End If
            Next lo
        Next ws
    Else
        For Each nm In ActiveWorkbook.names
            If StrComp(nm.name, expr, vbTextCompare) = 0 Then Set HAMU_ResolveValidationSource = nm.RefersToRange: Exit Function
        Next nm
    End If
End Function



Public Sub HAMU_DeleteDropdowns()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_DeleteDropdowns") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    On Error GoTo Failed
    Dim count As Long
    count = HAMU_RemoveDropdownObjects(ActiveSheet)
    HAMU_ShowInfo count & HAMU_Text(" açılır kutu nesnesi silindi. Hücre listeleri korundu.")
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_DeleteDropdowns", Err.number, Err.description
End Sub

Public Function HAMU_RemoveDropdownObjects(ByVal sheet As Worksheet) As Long
    Dim i As Long, shape As shape, remove As Boolean
    For i = sheet.Shapes.count To 1 Step -1
        Set shape = sheet.Shapes(i): remove = False
        If shape.Type = msoFormControl Then
            remove = (shape.FormControlType = xlDropDown)
        ElseIf shape.Type = msoOLEControlObject Then
            remove = (LCase$(shape.OLEFormat.ProgID) = "forms.combobox.1")
        End If
        If remove Then shape.Delete: HAMU_RemoveDropdownObjects = HAMU_RemoveDropdownObjects + 1
    Next
End Function



Public Sub HAMU_IndexLogo(ByVal ws As Worksheet)
 Dim path As String, picture As shape
 path = ThisWorkbook.path & "\..\06_Icons\64\hamu_logo_64.png"
 If Len(Dir$(path)) = 0 Then path = ThisWorkbook.path & "\..\06_Icons\hamu_logo_64.png"
 If Len(Dir$(path)) = 0 Then ws.range("A1").value = "HAMU": Exit Sub
 Set picture = ws.Shapes.AddPicture(path, msoFalse, msoTrue, ws.range("A1").Left + 3, ws.range("A1").Top + 2, 26, 26)
 picture.name = "HAMU_IndexLogo"
 picture.Placement = xlMoveAndSize
End Sub
