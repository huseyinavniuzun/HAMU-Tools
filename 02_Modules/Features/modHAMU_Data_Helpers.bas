Attribute VB_Name = "modHAMU_Data_Helpers"
Option Explicit
Option Private Module

Public Function HAMU_Data_ParseColumnList( _
    ByVal rng As range, _
    ByVal userText As String) As Collection

    Dim result As New Collection
    Dim seen As Object

    Dim parts() As String
    Dim token As String

    Dim i As Long
    Dim idx As Long
    Dim n As Double

    Set seen = CreateObject("Scripting.Dictionary")

    parts = Split(userText, ",")

    For i = LBound(parts) To UBound(parts)

        token = Trim$(parts(i))

        If Len(token) = 0 Then GoTo NextToken

        If IsNumeric(token) Then

            n = CDbl(token)

            If n <> Fix(n) Then

                Err.Raise _
                    vbObjectError + 2101, _
                    "HAMU_Data_ParseColumnList", _
                    "'" & token & HAMU_Text("' geçerli bir sütun numarası değil.")

            End If

            If n < 1 Or n > rng.columns.count Then Err.Raise 5, , HAMU_L("Sütun numarasi secim disinda.", "Column number is outside the selected range.")
            idx = CLng(n)

        Else

            idx = HAMU_HeaderToIndex( _
                    rng.rows(1), _
                    token)

            If idx = 0 Then

                Err.Raise _
                    vbObjectError + 2102, _
                    "HAMU_Data_ParseColumnList", _
                    "'" & token & HAMU_Text("' başlıklı sütun bulunamadı.")

            End If

        End If

        If idx < 1 Or idx > rng.columns.count Then

            Err.Raise _
                vbObjectError + 2103, _
                "HAMU_Data_ParseColumnList", _
                "'" & token & HAMU_Text("' seçim aralığının dışında.")

        End If

        If Not seen.exists(CStr(idx)) Then

            seen.Add CStr(idx), True
            result.Add idx

        End If

NextToken:

    Next i

    If result.count = 0 Then

        Err.Raise _
            vbObjectError + 2104, _
            "HAMU_Data_ParseColumnList", _
            HAMU_Text("Geçerli sütun seçilmedi.")

    End If

    Set HAMU_Data_ParseColumnList = result

End Function

Public Function HAMU_Data_ColumnListsOverlap( _
    ByVal list1 As Collection, _
    ByVal list2 As Collection) As Boolean

    Dim seen As Object
    Dim i As Long

    Set seen = CreateObject("Scripting.Dictionary")

    For i = 1 To list1.count
        seen(CStr(CLng(list1(i)))) = True
    Next i

    For i = 1 To list2.count

        If seen.exists(CStr(CLng(list2(i)))) Then

            HAMU_Data_ColumnListsOverlap = True
            Exit Function

        End If

    Next i

End Function

Public Function HAMU_Data_BuildGroupKey( _
    ByRef srcArr As Variant, _
    ByVal rowIndex As Long, _
    ByVal groupCols As Collection) As String

    Dim i As Long
    Dim idx As Long

    Dim part As String
    Dim result As String
    Dim v As Variant

    For i = 1 To groupCols.count

        idx = CLng(groupCols(i))
        v = srcArr(rowIndex, idx)

        If IsError(v) Then
            part = "#HATA"
        Else
            part = CStr(v)
        End If

        result = result & _
                 CStr(Len(part)) & ":" & _
                 part & "|"

    Next i

    HAMU_Data_BuildGroupKey = result

End Function

Public Function HAMU_HeaderToIndex( _
    ByVal hdr As range, _
    ByVal headerName As String) As Long

    Dim c As range

    For Each c In hdr.Cells

        If StrComp( _
            Trim$(CStr(c.Value2)), _
            Trim$(headerName), _
            vbTextCompare) = 0 Then

            HAMU_HeaderToIndex = _
                c.column - hdr.column + 1

            Exit Function

        End If

    Next c

End Function

Public Function HAMU_Data_UniqueFilePath( _
    ByVal folder As String, _
    ByVal baseName As String, _
    ByVal extension As String) As String

    Dim path As String
    Dim i As Long

    path = folder & Application.PathSeparator & _
           baseName & extension

    If Len(Dir$(path)) = 0 Then

        HAMU_Data_UniqueFilePath = path
        Exit Function

    End If

    i = 1

    Do

        path = folder & Application.PathSeparator & _
               baseName & "_" & Format$(i, "00") & extension

        If Len(Dir$(path)) = 0 Then

            HAMU_Data_UniqueFilePath = path
            Exit Function

        End If

        i = i + 1

    Loop

End Function

Public Function HAMU_Data_UniqueSheetName( _
    ByVal wb As Workbook, _
    ByVal baseName As String) As String

    Dim testName As String
    Dim i As Long

    baseName = Left$(baseName, 31)

    testName = baseName

    If Not HAMU_Data_SheetExists(wb, testName) Then

        HAMU_Data_UniqueSheetName = testName
        Exit Function

    End If

    i = 1

    Do

        testName = Left$(baseName, 27) & _
                   "_" & Format$(i, "00")

        If Not HAMU_Data_SheetExists(wb, testName) Then

            HAMU_Data_UniqueSheetName = testName
            Exit Function

        End If

        i = i + 1

    Loop

End Function

Private Function HAMU_Data_SheetExists( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = wb.Worksheets(sheetName)

    HAMU_Data_SheetExists = Not ws Is Nothing

    On Error GoTo 0

End Function

Public Sub HAMU_Data_AddSplitRow( _
    ByVal dictRows As Object, _
    ByVal dictLabels As Object, _
    ByVal keyOrder As Collection, _
    ByVal groupKey As String, _
    ByVal groupLabel As String, _
    ByVal sourceRow As Long)

    Dim rows As Collection

    If Not dictRows.exists(groupKey) Then

        Set rows = New Collection

        dictRows.Add groupKey, rows
        dictLabels.Add groupKey, groupLabel
        keyOrder.Add groupKey

    Else

        Set rows = dictRows(groupKey)

    End If

    rows.Add sourceRow

End Sub

Public Function HAMU_Data_ResolveColumn( _
    ByVal rng As range, _
    ByVal userValue As String) As Long

    Dim idx As Long
    Dim n As Double

    userValue = Trim$(userValue)

    If IsNumeric(userValue) Then

        n = CDbl(userValue)

        If n <> Fix(n) Then

            Err.Raise _
                vbObjectError + 2201, _
                "HAMU_Data_ResolveColumn", _
                HAMU_Text("Sütun numarası tam sayı olmalıdır.")

        End If

        idx = CLng(n)

    Else

        idx = _
            HAMU_HeaderToIndex( _
                rng.rows(1), _
                userValue)

        If idx = 0 Then

            Err.Raise _
                vbObjectError + 2202, _
                "HAMU_Data_ResolveColumn", _
                "'" & userValue & HAMU_Text("' başlıklı sütun bulunamadı.")

        End If

    End If

    If idx < 1 Or idx > rng.columns.count Then

        Err.Raise _
            vbObjectError + 2203, _
            "HAMU_Data_ResolveColumn", _
            HAMU_Text("Sütun seçim aralığının dışında.")

    End If

    HAMU_Data_ResolveColumn = idx

End Function

Public Sub HAMU_Data_CreateSplitSheet( _
    ByVal wb As Workbook, _
    ByVal sourceRange As range, _
    ByRef srcArr As Variant, _
    ByVal rowList As Collection, _
    ByVal groupLabel As String)

    Dim ws As Worksheet
    Dim resultArr() As Variant

    Dim r As Long
    Dim c As Long
    Dim sourceRow As Long

    Dim rowCount As Long
    Dim colCount As Long
    Dim sheetName As String

    rowCount = rowList.count
    colCount = UBound(srcArr, 2)

    ReDim resultArr(1 To rowCount + 1, _
                    1 To colCount)
    For c = 1 To colCount
        resultArr(1, c) = sourceRange.Cells(1, c).Value2
    Next c
    For r = 1 To rowCount

        sourceRow = CLng(rowList(r))

        For c = 1 To colCount

            resultArr(r + 1, c) = _
                HAMU_Data_SplitOutputValue( _
                    sourceRange.Cells(sourceRow, c), _
                    sourceRange.Cells(1, c).Value2)

        Next c

    Next r

    sheetName = _
        HAMU_Data_CleanSplitName(groupLabel)

    sheetName = _
        HAMU_Data_UniqueSheetName( _
            wb, _
            sheetName)

    Set ws = wb.Worksheets.Add( _
                After:=wb.Worksheets( _
                    wb.Worksheets.count))

    ws.name = sheetName
    For c = 1 To colCount

        If HAMU_Data_IsWeekColumn(sourceRange, c) Then

            ws.columns(c).NumberFormat = "@"

        End If

    Next c

    ws.Cells(1, 1) _
      .Resize(rowCount + 1, colCount) _
      .Value2 = resultArr

    With ws

        With .range( _
                .Cells(1, 1), _
                .Cells(1, colCount))

            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack

        End With

        .range( _
            .Cells(1, 1), _
            .Cells(rowCount + 1, colCount)) _
            .AutoFilter

        .columns.AutoFit

    End With

End Sub

Public Sub HAMU_Data_CreateSplitFile( _
    ByVal sourceRange As range, _
    ByRef srcArr As Variant, _
    ByVal rowList As Collection, _
    ByVal groupLabel As String, _
    ByVal folderPath As String)

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim resultArr() As Variant

    Dim r As Long
    Dim c As Long
    Dim sourceRow As Long

    Dim rowCount As Long
    Dim colCount As Long

    Dim baseName As String
    Dim savePath As String

    rowCount = rowList.count
    colCount = UBound(srcArr, 2)

    ReDim resultArr(1 To rowCount + 1, _
                    1 To colCount)

    For c = 1 To colCount
        resultArr(1, c) = sourceRange.Cells(1, c).Value2
    Next c

    For r = 1 To rowCount

        sourceRow = CLng(rowList(r))

        For c = 1 To colCount

            resultArr(r + 1, c) = _
                HAMU_Data_SplitOutputValue( _
                    sourceRange.Cells(sourceRow, c), _
                    sourceRange.Cells(1, c).Value2)

        Next c

    Next r

    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    For c = 1 To colCount

        If HAMU_Data_IsWeekColumn(sourceRange, c) Then

            ws.columns(c).NumberFormat = "@"

        End If

    Next c

    ws.Cells(1, 1) _
      .Resize(rowCount + 1, colCount) _
      .Value2 = resultArr

    With ws

        With .range( _
                .Cells(1, 1), _
                .Cells(1, colCount))

            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack

        End With

        .range( _
            .Cells(1, 1), _
            .Cells(rowCount + 1, colCount)) _
            .AutoFilter

        .columns.AutoFit

    End With

    baseName = _
        HAMU_Data_CleanSplitName(groupLabel)

    savePath = _
        HAMU_Data_UniqueFilePath( _
            folderPath, _
            baseName, _
            ".xlsx")

    wb.SaveAs _
        fileName:=savePath, _
        fileFormat:=xlOpenXMLWorkbook

    wb.Close SaveChanges:=False

End Sub

Private Function HAMU_Data_CleanSplitName( _
    ByVal value As String) As String

    Dim s As String
    Dim badChars As Variant
    Dim X As Variant

    s = Trim$(value)

    badChars = Array( _
        "\", "/", ":", "*", "?", _
        """", "<", ">", "|", _
        "[", "]")

    For Each X In badChars

        s = Replace$( _
                s, _
                CStr(X), _
                "_")

    Next X

    Do While InStr(s, "__") > 0
        s = Replace$(s, "__", "_")
    Loop

    If Len(s) = 0 Then
        s = "BOS"
    End If

    HAMU_Data_CleanSplitName = _
        Left$(s, 31)

End Function

Public Function HAMU_Data_WeekLabel( _
    ByVal dt As Date) As String

    Dim weekNo As Long
    Dim weekYear As Long
    Dim thursdayDate As Date

    weekNo = _
        DatePart( _
            "ww", _
            dt, _
            vbMonday, _
            vbFirstFourDays)

    thursdayDate = _
        DateAdd( _
            "d", _
            4 - Weekday(dt, vbMonday), _
            dt)

    weekYear = year(thursdayDate)

    HAMU_Data_WeekLabel = _
        CStr(weekYear) & _
        "-H" & _
        Format$(weekNo, "00")

End Function

Public Function HAMU_Data_SplitCellText( _
    ByVal target As range, _
    ByVal headerValue As Variant) As String

    Dim v As Variant
    Dim s As String
    Dim decSep As String
    Dim thouSep As String

    On Error GoTo Fallback

    v = target.Value2

    If IsError(v) Then
        HAMU_Data_SplitCellText = "#HATA"
        Exit Function
    End If

    If IsEmpty(v) Or Len(CStr(v)) = 0 Then
        HAMU_Data_SplitCellText = ""
        Exit Function
    End If

    If HAMU_Data_IsWeekCell(target, CStr(headerValue)) Then
        HAMU_Data_SplitCellText = HAMU_Data_FormatWeekCell(target)
        Exit Function
    End If

    If VarType(v) = vbString Then
        HAMU_Data_SplitCellText = CStr(v)
        Exit Function
    End If

    If IsNumeric(v) Then

        decSep = Application.International(xlDecimalSeparator)
        thouSep = Application.International(xlThousandsSeparator)

        s = Format$(CDbl(v), "0.###############")

        If Len(thouSep) > 0 Then
            If thouSep <> decSep Then
                s = Replace$(s, thouSep, "")
            End If
        End If

        If decSep <> "." Then
            s = Replace$(s, decSep, ".")
        End If

        HAMU_Data_SplitCellText = s
        Exit Function
    End If

    HAMU_Data_SplitCellText = CStr(v)
    Exit Function

Fallback:
    HAMU_Data_SplitCellText = CStr(target.Value2)

End Function

Public Function HAMU_Data_SplitOutputValue( _
    ByVal target As range, _
    ByVal headerValue As Variant) As Variant

    Dim v As Variant

    v = target.Value2

    If IsError(v) Then
        HAMU_Data_SplitOutputValue = v
        Exit Function
    End If

    If HAMU_Data_IsWeekCell(target, CStr(headerValue)) Then
        HAMU_Data_SplitOutputValue = HAMU_Data_FormatWeekCell(target)
        Exit Function
    End If

    If VarType(v) = vbString Then
        HAMU_Data_SplitOutputValue = CStr(v)
    Else
        HAMU_Data_SplitOutputValue = v
    End If

End Function

Private Function HAMU_Data_IsWeekHeader( _
    ByVal headerText As String) As Boolean

    Dim s As String

    s = LCase$(Trim$(headerText))
    s = Replace$(s, " ", "")
    s = Replace$(s, "_", "")
    s = Replace$(s, "-", "")
    s = Replace$(s, ".", "")

    If InStr(1, s, "hafta", vbTextCompare) > 0 Then
        HAMU_Data_IsWeekHeader = True
        Exit Function
    End If

    If InStr(1, s, "week", vbTextCompare) > 0 Then
        HAMU_Data_IsWeekHeader = True
        Exit Function
    End If

    If s = "hf" Or s = "hfno" Or s = "hfkey" Then
        HAMU_Data_IsWeekHeader = True
    End If

End Function

Private Function HAMU_Data_IsWeekCell( _
    ByVal target As range, _
    ByVal headerText As String) As Boolean

    Dim shown As String
    Dim normalized As String
    Dim v As Variant

    If HAMU_Data_IsWeekHeader(headerText) Then
        HAMU_Data_IsWeekCell = True
        Exit Function
    End If

    On Error Resume Next
    shown = Trim$(CStr(target.text))
    On Error GoTo 0

    normalized = Replace$(shown, ",", ".")

    If HAMU_Data_IsWeekText(normalized) Then
        HAMU_Data_IsWeekCell = True
        Exit Function
    End If

    v = target.Value2

    If VarType(v) = vbString Then
        normalized = Replace$(Trim$(CStr(v)), ",", ".")

        If HAMU_Data_IsWeekText(normalized) Then
            HAMU_Data_IsWeekCell = True
        End If
    End If

End Function

Public Function HAMU_Data_IsWeekColumn( _
    ByVal sourceRange As range, _
    ByVal columnIndex As Long) As Boolean

    Dim maxRow As Long
    Dim r As Long
    Dim headerText As String

    headerText = CStr(sourceRange.Cells(1, columnIndex).Value2)

    If HAMU_Data_IsWeekHeader(headerText) Then
        HAMU_Data_IsWeekColumn = True
        Exit Function
    End If

    maxRow = sourceRange.rows.count
    If maxRow > 21 Then maxRow = 21

    For r = 2 To maxRow

        If HAMU_Data_IsWeekCell( _
                sourceRange.Cells(r, columnIndex), _
                headerText) Then

            HAMU_Data_IsWeekColumn = True
            Exit Function

        End If

    Next r

End Function

Private Function HAMU_Data_IsWeekText( _
    ByVal ValueText As String) As Boolean

    Dim s As String
    Dim p As Long
    Dim yearPart As Long
    Dim weekPart As Long

    s = Trim$(ValueText)

    If Len(s) = 0 Then Exit Function
    If Len(s) = 6 And HAMU_Data_IsDigits(s) Then

        yearPart = CLng(Left$(s, 4))
        weekPart = CLng(right$(s, 2))

        If yearPart >= 1900 And yearPart <= 2200 And _
           weekPart >= 1 And weekPart <= 53 Then

            HAMU_Data_IsWeekText = True
        End If

        Exit Function
    End If
    p = InStr(1, s, ".", vbBinaryCompare)

    If p = 5 Then

        If HAMU_Data_IsDigits(Left$(s, 4)) And _
           HAMU_Data_IsDigits(Mid$(s, 6)) Then

            yearPart = CLng(Left$(s, 4))
            weekPart = CLng(Mid$(s, 6))

            If yearPart >= 1900 And yearPart <= 2200 And _
               weekPart >= 1 And weekPart <= 53 Then

                HAMU_Data_IsWeekText = True
            End If

        End If
    End If

End Function

Private Function HAMU_Data_FormatWeekCell( _
    ByVal target As range) As String

    Dim shown As String
    Dim s As String
    Dim v As Variant

    Dim numericValue As Double
    Dim integerValue As Long

    Dim yearPart As Long
    Dim weekPart As Long
    On Error Resume Next
    shown = Trim$(CStr(target.text))
    On Error GoTo 0

    s = Replace$(shown, ",", ".")

    If HAMU_Data_IsWeekText(s) Then
        HAMU_Data_FormatWeekCell = HAMU_Data_NormalizeWeekText(s)
        Exit Function
    End If
    v = target.Value2

    If VarType(v) = vbString Then

        s = Replace$(Trim$(CStr(v)), ",", ".")

        If HAMU_Data_IsWeekText(s) Then
            HAMU_Data_FormatWeekCell = HAMU_Data_NormalizeWeekText(s)
            Exit Function
        End If

    End If

    If IsNumeric(v) Then

        numericValue = CDbl(v)
        If numericValue = Fix(numericValue) And _
           numericValue >= 190001 And _
           numericValue <= 220053 Then

            integerValue = CLng(numericValue)

            yearPart = integerValue \ 100
            weekPart = integerValue Mod 100

            If yearPart >= 1900 And yearPart <= 2200 And _
               weekPart >= 1 And weekPart <= 53 Then

                HAMU_Data_FormatWeekCell = _
                    Format$(yearPart, "0000") & "." & _
                    Format$(weekPart, "00")

                Exit Function
            End If
        End If
        yearPart = Fix(numericValue)

        weekPart = CLng( _
            Round((numericValue - yearPart) * 100, 0))

        If yearPart >= 1900 And yearPart <= 2200 And _
           weekPart >= 1 And weekPart <= 53 Then

            HAMU_Data_FormatWeekCell = _
                Format$(yearPart, "0000") & "." & _
                Format$(weekPart, "00")

            Exit Function
        End If

    End If

    HAMU_Data_FormatWeekCell = Replace$(CStr(target.Value2), ",", ".")

End Function

Private Function HAMU_Data_NormalizeWeekText( _
    ByVal ValueText As String) As String

    Dim s As String
    Dim p As Long
    Dim yearPart As Long
    Dim weekPart As Long

    s = Replace$(Trim$(ValueText), ",", ".")

    If Len(s) = 6 And HAMU_Data_IsDigits(s) Then

        yearPart = CLng(Left$(s, 4))
        weekPart = CLng(right$(s, 2))

    Else

        p = InStr(1, s, ".", vbBinaryCompare)

        yearPart = CLng(Left$(s, p - 1))
        weekPart = CLng(Mid$(s, p + 1))

    End If

    HAMU_Data_NormalizeWeekText = _
        Format$(yearPart, "0000") & "." & _
        Format$(weekPart, "00")

End Function

Private Function HAMU_Data_IsDigits( _
    ByVal ValueText As String) As Boolean

    Dim i As Long
    Dim ch As String

    If Len(ValueText) = 0 Then Exit Function

    For i = 1 To Len(ValueText)

        ch = Mid$(ValueText, i, 1)

        If ch < "0" Or ch > "9" Then
            Exit Function
        End If

    Next i

    HAMU_Data_IsDigits = True

End Function

Public Function HAMU_Data_IsBlankCell( _
    ByVal target As range) As Boolean

    If IsError(target.value) Then Exit Function

    HAMU_Data_IsBlankCell = _
        (Len(Trim$(CStr(target.Value2))) = 0)

End Function

Private Function HAMU_Data_LooksLikeDateFormat( _
    ByVal target As range) As Boolean

    Dim f As String

    On Error Resume Next
    f = LCase$(target.NumberFormat)
    On Error GoTo 0

    If InStr(f, "yy") > 0 Or _
       InStr(f, "yyyy") > 0 Or _
       InStr(f, "dd") > 0 Then

        HAMU_Data_LooksLikeDateFormat = True

    End If

End Function

Public Function HAMU_Data_CellKind( _
    ByVal target As range) As String

    Dim v As Variant

    If IsError(target.value) Then

        HAMU_Data_CellKind = "Hata"
        Exit Function

    End If

    If HAMU_Data_IsBlankCell(target) Then

        HAMU_Data_CellKind = HAMU_Text("Boş")
        Exit Function

    End If

    v = target.Value2

    If VarType(v) = vbBoolean Then

        HAMU_Data_CellKind = HAMU_Text("Mantıksal")

    ElseIf IsNumeric(v) And HAMU_Data_LooksLikeDateFormat(target) Then

        HAMU_Data_CellKind = "Tarih"

    ElseIf IsNumeric(v) Then

        HAMU_Data_CellKind = HAMU_Text("Sayı")

    Else

        HAMU_Data_CellKind = HAMU_Text("Metin")

    End If

End Function

Public Function HAMU_Data_ProfileType( _
    ByVal numCount As Long, _
    ByVal dateCount As Long, _
    ByVal textCount As Long, _
    ByVal boolCount As Long, _
    ByVal errorCount As Long) As String

    Dim totalTyped As Long
    Dim maxCount As Long
    Dim maxType As String

    totalTyped = numCount + dateCount + textCount + boolCount

    If totalTyped = 0 Then

        If errorCount > 0 Then
            HAMU_Data_ProfileType = "Hata"
        Else
            HAMU_Data_ProfileType = HAMU_Text("Boş")
        End If

        Exit Function

    End If

    maxCount = numCount
    maxType = HAMU_Text("Sayı")

    If dateCount > maxCount Then
        maxCount = dateCount
        maxType = "Tarih"
    End If

    If textCount > maxCount Then
        maxCount = textCount
        maxType = HAMU_Text("Metin")
    End If

    If boolCount > maxCount Then
        maxCount = boolCount
        maxType = HAMU_Text("Mantıksal")
    End If

    If maxCount / totalTyped >= 0.9 Then
        HAMU_Data_ProfileType = maxType
    Else
        HAMU_Data_ProfileType = HAMU_Text("Karışık")
    End If

End Function

Public Function HAMU_Data_ProfileKey( _
    ByVal target As range) As String

    If IsError(target.value) Then

        HAMU_Data_ProfileKey = "#HATA"

    ElseIf HAMU_Data_IsWeekCell( _
                target, _
                CStr(target.Worksheet.Cells(1, target.column).Value2)) Then

        HAMU_Data_ProfileKey = HAMU_Data_FormatWeekCell(target)

    Else

        HAMU_Data_ProfileKey = CStr(target.Value2)

    End If

End Function

Public Function HAMU_Data_NormalizeHeaderKey( _
    ByVal headerText As String) As String

    HAMU_Data_NormalizeHeaderKey = _
        LCase$(HAMU_Data_CleanHeaderText(headerText))

End Function

Public Function HAMU_Data_CleanHeaderText( _
    ByVal headerText As String) As String

    Dim s As String

    s = headerText

    s = Replace$(s, Chr$(160), " ")
    s = Replace$(s, vbTab, " ")
    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")

    s = Trim$(s)

    Do While InStr(s, "  ") > 0
        s = Replace$(s, "  ", " ")
    Loop

    HAMU_Data_CleanHeaderText = s

End Function

Public Function HAMU_Data_ToSnakeCase( _
    ByVal headerText As String) As String

    Dim s As String
    Dim i As Long
    Dim ch As String
    Dim outText As String

    s = HAMU_Data_CleanHeaderText(headerText)
    s = Replace$(s, ChrW(304), "i")
    s = Replace$(s, "I", "i")
    s = Replace$(s, ChrW(305), "i")
    s = LCase$(s)

    s = Replace$(s, HAMU_Text("ç"), "c")
    s = Replace$(s, HAMU_Text("ğ"), "g")
    s = Replace$(s, HAMU_Text("ı"), "i")
    s = Replace$(s, HAMU_Text("ö"), "o")
    s = Replace$(s, HAMU_Text("ş"), "s")
    s = Replace$(s, HAMU_Text("ü"), "u")

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)

        If (ch >= "a" And ch <= "z") Or _
           (ch >= "0" And ch <= "9") Then

            outText = outText & ch

        Else

            If Len(outText) > 0 Then

                If right$(outText, 1) <> "_" Then
                    outText = outText & "_"
                End If

            End If

        End If

    Next i

    Do While Left$(outText, 1) = "_"
        outText = Mid$(outText, 2)
    Loop

    Do While Len(outText) > 0 And right$(outText, 1) = "_"
        outText = Left$(outText, Len(outText) - 1)
    Loop

    HAMU_Data_ToSnakeCase = outText

End Function

Public Function HAMU_Data_MakeUniqueHeader( _
    ByVal baseHeader As String, _
    ByVal seen As Object) As String

    Dim candidate As String
    Dim i As Long

    candidate = baseHeader

    If Not seen.exists(candidate) Then

        seen.Add candidate, True
        HAMU_Data_MakeUniqueHeader = candidate
        Exit Function

    End If

    i = 2

    Do

        candidate = baseHeader & "_" & i

        If Not seen.exists(candidate) Then

            seen.Add candidate, True
            HAMU_Data_MakeUniqueHeader = candidate
            Exit Function

        End If

        i = i + 1

    Loop

End Function

Public Function HAMU_Data_AllNonKeyColumns( _
    ByVal columnCount As Long, _
    ByVal keyCols As Collection) As Collection

    Dim result As New Collection
    Dim c As Long

    For c = 1 To columnCount

        If Not HAMU_Data_CollectionContainsLong(keyCols, c) Then
            result.Add c
        End If

    Next c

    Set HAMU_Data_AllNonKeyColumns = result

End Function

Public Function HAMU_Data_CollectionContainsLong( _
    ByVal items As Collection, _
    ByVal value As Long) As Boolean

    Dim i As Long

    For i = 1 To items.count

        If CLng(items(i)) = value Then

            HAMU_Data_CollectionContainsLong = True
            Exit Function

        End If

    Next i

End Function

Public Function HAMU_Data_BuildKeyFromRangeRow( _
    ByVal rng As range, _
    ByVal rowIndex As Long, _
    ByVal cols As Collection) As String

    Dim i As Long
    Dim idx As Long
    Dim partText As String
    Dim result As String

    For i = 1 To cols.count

        idx = CLng(cols(i))

        partText = HAMU_Data_SplitCellText( _
                        rng.Cells(rowIndex, idx), _
                        rng.Cells(1, idx).Value2)

        result = result & _
                 Len(partText) & ":" & _
                 partText & "|"

    Next i

    HAMU_Data_BuildKeyFromRangeRow = result

End Function

Private Function HAMU_Data_KeyExistsInRange( _
    ByVal rng As range, _
    ByVal cols As Collection, _
    ByVal searchKey As String) As Boolean

    Dim r As Long

    For r = 2 To rng.rows.count

        If HAMU_Data_BuildKeyFromRangeRow( _
                rng, _
                r, _
                cols) = searchKey Then

            HAMU_Data_KeyExistsInRange = True
            Exit Function

        End If

    Next r

End Function

Public Sub HAMU_Data_FillJoinOutputRow( _
    ByRef outArr As Variant, _
    ByVal outRow As Long, _
    ByVal leftRng As range, _
    ByVal leftRow As Long, _
    ByVal rightRng As range, _
    ByVal rightRow As Long, _
    ByVal rightCols As Collection)

    Dim c As Long
    Dim outCol As Long
    Dim idx As Long

    outCol = 1

    For c = 1 To leftRng.columns.count

        If leftRow > 0 Then

            outArr(outRow, outCol) = _
                HAMU_Data_SplitOutputValue( _
                    leftRng.Cells(leftRow, c), _
                    leftRng.Cells(1, c).Value2)

        End If

        outCol = outCol + 1

    Next c

    For c = 1 To rightCols.count

        idx = CLng(rightCols(c))

        If rightRow > 0 Then

            outArr(outRow, outCol) = _
                HAMU_Data_SplitOutputValue( _
                    rightRng.Cells(rightRow, idx), _
                    rightRng.Cells(1, idx).Value2)

        End If

        outCol = outCol + 1

    Next c

End Sub

Public Function HAMU_Data_RightHeaderName( _
    ByVal leftRng As range, _
    ByVal rightHeader As String) As String

    Dim c As Long

    For c = 1 To leftRng.columns.count

        If StrComp( _
            Trim$(CStr(leftRng.Cells(1, c).value)), _
            Trim$(rightHeader), _
            vbTextCompare) = 0 Then

            HAMU_Data_RightHeaderName = HAMU_Text("Sağ - ") & rightHeader
            Exit Function

        End If

    Next c

    HAMU_Data_RightHeaderName = rightHeader

End Function

Public Sub HAMU_Data_PrepareJoinFormats( _
    ByVal outSh As Worksheet, _
    ByVal leftRng As range, _
    ByVal rightRng As range, _
    ByVal rightCols As Collection)

    Dim c As Long
    Dim outCol As Long
    Dim idx As Long

    outCol = 1

    For c = 1 To leftRng.columns.count

        If HAMU_Data_IsWeekColumn(leftRng, c) Then
            outSh.columns(outCol).NumberFormat = "@"
        End If

        outCol = outCol + 1

    Next c

    For c = 1 To rightCols.count

        idx = CLng(rightCols(c))

        If HAMU_Data_IsWeekColumn(rightRng, idx) Then
            outSh.columns(outCol).NumberFormat = "@"
        End If

        outCol = outCol + 1

    Next c

End Sub

Public Sub HAMU_Data_PrepareOutputWeekColumns( _
    ByVal outSh As Worksheet, _
    ByVal sourceRange As range, _
    ByVal sourceColumnCount As Long)

    Dim c As Long

    For c = 1 To sourceColumnCount

        If HAMU_Data_IsWeekColumn(sourceRange, c) Then
            outSh.columns(c).NumberFormat = "@"
        End If

    Next c

End Sub

Public Function HAMU_Data_ParseFilterConditions( _
    ByVal rng As range, _
    ByVal expressionText As String) As Collection

    Dim result As New Collection
    Dim pieces() As String
    Dim piece As String

    Dim operators As Variant
    Dim op As Variant
    Dim pos As Long

    Dim leftText As String
    Dim rightText As String
    Dim colIndex As Long

    Dim conditionData As Variant
    Dim i As Long

    operators = Array(">=", "<=", "<>", "~", "=", ">", "<")

    pieces = Split(expressionText, ";")

    For i = LBound(pieces) To UBound(pieces)

        piece = Trim$(pieces(i))

        If Len(piece) = 0 Then GoTo NextCondition

        pos = 0

        For Each op In operators

            pos = InStr(1, piece, CStr(op), vbBinaryCompare)

            If pos > 0 Then Exit For

        Next op

        If pos = 0 Then

            Err.Raise _
                vbObjectError + 2401, _
                "HAMU_Data_ParseFilterConditions", _
                HAMU_Text("Koşul operatörü bulunamadı: ") & piece

        End If

        leftText = Trim$(Left$(piece, pos - 1))
        rightText = Trim$(Mid$(piece, pos + Len(CStr(op))))

        If Len(rightText) >= 2 Then

            If (Left$(rightText, 1) = """" And right$(rightText, 1) = """") Or _
               (Left$(rightText, 1) = "'" And right$(rightText, 1) = "'") Then

                rightText = Mid$(rightText, 2, Len(rightText) - 2)

            End If

        End If

        colIndex = HAMU_Data_ResolveColumn(rng, leftText)

        conditionData = Array(colIndex, CStr(op), rightText)
        result.Add conditionData

NextCondition:

    Next i

    If result.count = 0 Then

        Err.Raise _
            vbObjectError + 2402, _
            "HAMU_Data_ParseFilterConditions", _
            HAMU_Text("Geçerli koşul bulunamadı.")

    End If

    Set HAMU_Data_ParseFilterConditions = result

End Function

Public Function HAMU_Data_RowMatchesConditions( _
    ByVal rng As range, _
    ByVal rowIndex As Long, _
    ByVal conditions As Collection) As Boolean

    Dim i As Long
    Dim cond As Variant

    HAMU_Data_RowMatchesConditions = True

    For i = 1 To conditions.count

        cond = conditions(i)

        If Not HAMU_Data_CompareFilterValue( _
                rng.Cells(rowIndex, CLng(cond(0))), _
                CStr(cond(1)), _
                CStr(cond(2))) Then

            HAMU_Data_RowMatchesConditions = False
            Exit Function

        End If

    Next i

End Function

Private Function HAMU_Data_CompareFilterValue( _
    ByVal target As range, _
    ByVal op As String, _
    ByVal criteria As String) As Boolean

    Dim cellText As String
    Dim cmp As Long

    Dim leftNum As Double
    Dim rightNum As Double

    cellText = HAMU_Data_SplitCellText( _
                    target, _
                    target.Worksheet.Cells(1, target.column).Value2)

    If op = "~" Then

        HAMU_Data_CompareFilterValue = _
            (InStr(1, cellText, criteria, vbTextCompare) > 0)

        Exit Function

    End If

    If IsNumeric(target.Value2) And IsNumeric(criteria) Then

        leftNum = CDbl(target.Value2)
        rightNum = CDbl(criteria)

        Select Case op

            Case "="
                HAMU_Data_CompareFilterValue = (leftNum = rightNum)

            Case "<>"
                HAMU_Data_CompareFilterValue = (leftNum <> rightNum)

            Case ">"
                HAMU_Data_CompareFilterValue = (leftNum > rightNum)

            Case ">="
                HAMU_Data_CompareFilterValue = (leftNum >= rightNum)

            Case "<"
                HAMU_Data_CompareFilterValue = (leftNum < rightNum)

            Case "<="
                HAMU_Data_CompareFilterValue = (leftNum <= rightNum)

        End Select

        Exit Function

    End If

    cmp = StrComp(cellText, criteria, vbTextCompare)

    Select Case op

        Case "="
            HAMU_Data_CompareFilterValue = (cmp = 0)

        Case "<>"
            HAMU_Data_CompareFilterValue = (cmp <> 0)

        Case ">"
            HAMU_Data_CompareFilterValue = (cmp > 0)

        Case ">="
            HAMU_Data_CompareFilterValue = (cmp >= 0)

        Case "<"
            HAMU_Data_CompareFilterValue = (cmp < 0)

        Case "<="
            HAMU_Data_CompareFilterValue = (cmp <= 0)

    End Select

End Function

Public Sub HAMU_Data_WriteRowsToNewSheet( _
    ByVal sourceRange As range, _
    ByVal rowList As Collection, _
    ByVal sheetBaseName As String)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outArr() As Variant

    Dim r As Long
    Dim c As Long
    Dim sourceRow As Long
    Dim colCount As Long

    colCount = sourceRange.columns.count

    Set wb = sourceRange.Worksheet.parent

    Set ws = wb.Worksheets.Add( _
                After:=wb.Worksheets(wb.Worksheets.count))

    ws.name = HAMU_Data_UniqueSheetName( _
                wb, _
                Left$(sheetBaseName, 31))

    For c = 1 To colCount

        ws.Cells(1, c).value = sourceRange.Cells(1, c).value

        If HAMU_Data_IsWeekColumn(sourceRange, c) Then
            ws.columns(c).NumberFormat = "@"
        End If

    Next c

    If rowList.count > 0 Then

        ReDim outArr(1 To rowList.count, 1 To colCount)

        For r = 1 To rowList.count

            sourceRow = CLng(rowList(r))

            For c = 1 To colCount

                outArr(r, c) = _
                    HAMU_Data_SplitOutputValue( _
                        sourceRange.Cells(sourceRow, c), _
                        sourceRange.Cells(1, c).Value2)

            Next c

        Next r

        ws.Cells(2, 1) _
          .Resize(rowList.count, colCount) _
          .Value2 = outArr

    End If

    With ws

        With .range(.Cells(1, 1), .Cells(1, colCount))
            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack
        End With

        If rowList.count > 0 Then
            .range(.Cells(1, 1), .Cells(rowList.count + 1, colCount)).AutoFilter
        End If

        .columns.AutoFit

    End With

End Sub

Public Function HAMU_Data_IsSupportedDataFile( _
    ByVal extension As String) As Boolean

    Select Case LCase$(extension)

        Case "xlsx", "xlsm", "xlsb", "xls", "csv"
            HAMU_Data_IsSupportedDataFile = True

    End Select

End Function

Public Sub HAMU_Data_ShuffleLongArray( _
    ByRef values() As Long)

    Dim i As Long
    Dim j As Long
    Dim tmp As Long

    For i = UBound(values) To LBound(values) + 1 Step -1

        j = Int((i - LBound(values) + 1) * Rnd) + LBound(values)

        tmp = values(i)
        values(i) = values(j)
        values(j) = tmp

    Next i

End Sub

Public Sub HAMU_Data_GenerateMissingCombinations( _
    ByVal valueSets As Collection, _
    ByVal level As Long, _
    ByRef parts() As String, _
    ByVal existing As Object, _
    ByRef outArr As Variant, _
    ByRef outCount As Long)

    Dim values As Collection
    Dim v As Variant

    Dim key As String
    Dim i As Long

    Set values = valueSets(level)

    For Each v In values

        parts(level) = CStr(v)

        If level < valueSets.count Then

            HAMU_Data_GenerateMissingCombinations _
                valueSets, _
                level + 1, _
                parts, _
                existing, _
                outArr, _
                outCount

        Else

            key = ""

            For i = 1 To UBound(parts)

                key = key & _
                      Len(parts(i)) & ":" & _
                      parts(i) & "|"

            Next i

            If Not existing.exists(key) Then

                outCount = outCount + 1

                For i = 1 To UBound(parts)
                    outArr(outCount, i) = parts(i)
                Next i

            End If

        End If

    Next v

End Sub

Public Function HAMU_Data_DominantColumnType( _
    ByVal rng As range, _
    ByVal columnIndex As Long) As String

    Dim counts As Object
    Dim kind As String
    Dim r As Long

    Dim maxCount As Long
    Dim maxType As String
    Dim total As Long

    Dim k As Variant

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare

    For r = 2 To rng.rows.count

        kind = HAMU_Data_CellKind(rng.Cells(r, columnIndex))

        If kind <> HAMU_Text("Boş") And kind <> "Hata" Then

            If counts.exists(kind) Then
                counts(kind) = CLng(counts(kind)) + 1
            Else
                counts.Add kind, 1
            End If

            total = total + 1

        End If

    Next r

    If total = 0 Then

        HAMU_Data_DominantColumnType = HAMU_Text("Boş")
        Exit Function

    End If

    For Each k In counts.keys

        If CLng(counts(k)) > maxCount Then

            maxCount = CLng(counts(k))
            maxType = CStr(k)

        End If

    Next k

    If maxCount / total >= 0.9 Then
        HAMU_Data_DominantColumnType = maxType
    Else
        HAMU_Data_DominantColumnType = HAMU_Text("Karışık")
    End If

End Function

Public Sub HAMU_Data_WriteQualityIssue( _
    ByVal outSh As Worksheet, _
    ByRef outRow As Long, _
    ByVal issueType As String, _
    ByVal sourceRange As range, _
    ByVal sourceRow As Long, _
    ByVal sourceCol As Long, _
    ByVal displayValue As String, _
    ByVal expectedValue As String, _
    ByVal detailText As String)

    outSh.Cells(outRow, 1).value = issueType
    outSh.Cells(outRow, 2).value = _
        sourceRange.row + sourceRow - 1
    outSh.Cells(outRow, 3).value = _
        sourceRange.Cells(1, sourceCol).value
    outSh.Cells(outRow, 4).value = displayValue
    outSh.Cells(outRow, 5).value = expectedValue
    outSh.Cells(outRow, 6).value = detailText

    outRow = outRow + 1

End Sub

Public Function HAMU_Data_KeyDisplay( _
    ByVal rng As range, _
    ByVal rowIndex As Long, _
    ByVal cols As Collection) As String

    Dim i As Long
    Dim idx As Long
    Dim partText As String
    Dim result As String

    For i = 1 To cols.count

        idx = CLng(cols(i))

        partText = HAMU_Data_SplitCellText( _
                        rng.Cells(rowIndex, idx), _
                        rng.Cells(1, idx).Value2)

        If Len(result) > 0 Then
            result = result & " | "
        End If

        result = result & _
                 CStr(rng.Cells(1, idx).value) & _
                 "=" & _
                 partText

    Next i

    HAMU_Data_KeyDisplay = result

End Function


