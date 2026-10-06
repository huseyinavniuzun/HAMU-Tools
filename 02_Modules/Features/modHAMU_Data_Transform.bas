Attribute VB_Name = "modHAMU_Data_Transform"
Option Explicit
Option Private Module

Public Sub HAMU_GroupAndSummarize()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_GroupAndSummarize") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim groupInput As String
    Dim sumInput As String

    Dim groupCols As Collection
    Dim sumCols As Collection

    Dim dictLabels As Object
    Dim dictSums As Object
    Dim keyOrder As Collection

    Dim srcArr As Variant
    Dim outArr() As Variant

    Dim labels As Variant
    Dim sums As Variant

    Dim key As String
    Dim k As Variant

    Dim r As Long
    Dim i As Long
    Dim outRow As Long
    Dim outCol As Long

    Dim idx As Long
    Dim v As Variant

    Dim outSh As Worksheet

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Gruplanacak veri alanını seçin.") & vbCrLf & _
                HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    If rng.areas.count > 1 Then
        MsgBox HAMU_Text("Grupla ve Özetle için tek parça bir aralık seçin."), _
               vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    If rng.rows.count < 2 Then
        MsgBox HAMU_Text("Seçimde en az bir başlık ve bir veri satırı olmalıdır."), _
               vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    groupInput = HAMU_InputText( _
                    HAMU_Text("Gruplama sütunlarını virgülle ayırarak girin.") & vbCrLf & vbCrLf & _
                    HAMU_Text("Sütun numarası kullanabilirsiniz:") & vbCrLf & _
                    "1,2,3" & vbCrLf & vbCrLf & _
                    HAMU_Text("veya başlık adlarını yazabilirsiniz:") & vbCrLf & _
                    HAMU_Text("Mağaza, Şehir, Ürün"), _
                    HAMU_APP_NAME & HAMU_Text(" - Grupla ve Özetle"))

    If Len(Trim$(groupInput)) = 0 Then Exit Sub

    Set groupCols = HAMU_Data_ParseColumnList(rng, groupInput)

    sumInput = HAMU_InputText( _
                    HAMU_Text("Toplanacak sayısal sütunları virgülle ayırarak girin.") & vbCrLf & vbCrLf & _
                    HAMU_Text("Örnek:") & vbCrLf & _
                    "4,5" & vbCrLf & vbCrLf & _
                    HAMU_Text("veya başlık adları:") & vbCrLf & _
                    HAMU_Text("Satış, Ciro"), _
                    HAMU_APP_NAME & HAMU_Text(" - Grupla ve Özetle"))

    If Len(Trim$(sumInput)) = 0 Then Exit Sub

    Set sumCols = HAMU_Data_ParseColumnList(rng, sumInput)
    If HAMU_Data_ColumnListsOverlap(groupCols, sumCols) Then

        MsgBox _
            HAMU_Text("Aynı sütun hem gruplama alanı hem de toplanacak alan olarak kullanılamaz."), _
            vbExclamation, _
            HAMU_APP_NAME

        Exit Sub

    End If

    srcArr = rng.Value2

    Set dictLabels = CreateObject("Scripting.Dictionary")
    Set dictSums = CreateObject("Scripting.Dictionary")
    Set keyOrder = New Collection

    dictLabels.CompareMode = vbBinaryCompare
    dictSums.CompareMode = vbBinaryCompare

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 2 To UBound(srcArr, 1)

        key = HAMU_Data_BuildGroupKey(srcArr, r, groupCols)

        If Not dictLabels.exists(key) Then

            ReDim labels(1 To groupCols.count)
            ReDim sums(1 To sumCols.count)

            For i = 1 To groupCols.count

                idx = CLng(groupCols(i))
                labels(i) = srcArr(r, idx)

            Next i

            For i = 1 To sumCols.count
                sums(i) = 0#
            Next i

            dictLabels.Add key, labels
            dictSums.Add key, sums
            keyOrder.Add key

        Else

            sums = dictSums(key)

        End If

        sums = dictSums(key)

        For i = 1 To sumCols.count

            idx = CLng(sumCols(i))
            v = srcArr(r, idx)

            If Not IsError(v) Then

                If IsNumeric(v) And Len(CStr(v)) > 0 Then
                    sums(i) = CDbl(sums(i)) + CDbl(v)
                End If

            End If

        Next i

        dictSums(key) = sums

    Next r

    Set outSh = HAMU_GetOrCreateSheet( _
                    "OZET_" & Format$(Now, "hhmmss"))
    outCol = 1

    For i = 1 To groupCols.count

        idx = CLng(groupCols(i))
        outSh.Cells(1, outCol).value = rng.Cells(1, idx).value

        outCol = outCol + 1

    Next i
    For i = 1 To sumCols.count

        idx = CLng(sumCols(i))

        outSh.Cells(1, outCol).value = _
            "Toplam - " & CStr(rng.Cells(1, idx).value)

        outCol = outCol + 1

    Next i
    If dictLabels.count > 0 Then

        ReDim outArr(1 To dictLabels.count, _
                     1 To groupCols.count + sumCols.count)

        outRow = 1

        For Each k In keyOrder

            labels = dictLabels(CStr(k))
            sums = dictSums(CStr(k))

            outCol = 1

            For i = 1 To groupCols.count

                outArr(outRow, outCol) = labels(i)
                outCol = outCol + 1

            Next i

            For i = 1 To sumCols.count

                outArr(outRow, outCol) = sums(i)
                outCol = outCol + 1

            Next i

            outRow = outRow + 1

        Next k

    outSh.Cells(2, 1) _
             .Resize(dictLabels.count, _
                     groupCols.count + sumCols.count) _
             .value = outArr

    End If

    With outSh
        With .range( _
                .Cells(1, 1), _
                .Cells(1, groupCols.count + sumCols.count))

            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack

        End With

        If dictLabels.count > 0 Then

            .range( _
                .Cells(1, 1), _
                .Cells(dictLabels.count + 1, _
                       groupCols.count + sumCols.count)) _
                .AutoFilter

        End If

        .columns.AutoFit

    End With

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Gruplama ve özetleme tamamlandı.") & vbCrLf & vbCrLf & _
        "Gruplama sütunu: " & groupCols.count & vbCrLf & _
        "Toplanan sütun: " & sumCols.count & vbCrLf & _
        HAMU_Text("Oluşan grup: ") & Format$(dictLabels.count, "#,##0") & vbCrLf & _
        HAMU_Text("Sonuç sayfası: ") & outSh.name, _
        vbInformation, _
        HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError _
        "HAMU_GroupAndSummarize", _
        Err.number, _
        Err.description

End Sub

Public Sub HAMU_UnpivotColumnsToRows()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_UnpivotColumnsToRows") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim outSh As Worksheet

    Dim answer As Variant
    Dim fixedCols As Long

    Dim rowCount As Long
    Dim colCount As Long

    Dim resultRows As Double
    Dim outRows As Long
    Dim outCols As Long

    Dim srcArr As Variant
    Dim outArr() As Variant

    Dim r As Long
    Dim c As Long
    Dim f As Long
    Dim outRow As Long

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Sütunları satırlara dönüştürmek için veri aralığını seçin.") & vbCrLf & _
                HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    If rng.areas.count > 1 Then
        MsgBox HAMU_Text("UNPIVOT için tek parça bir aralık seçin."), _
               vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    If rng.rows.count < 2 Or rng.columns.count < 2 Then
        MsgBox HAMU_Text("Seçimde en az bir başlık satırı, bir veri satırı ve iki sütun olmalıdır."), _
               vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    answer = HAMU_InputValue( _
                Prompt:= _
                    HAMU_Text("Solda kaç sütun sabit kalacak?") & vbCrLf & vbCrLf & _
                    HAMU_Text("Örnek:") & vbCrLf & _
                    HAMU_Text("Mağaza | Şehir | Ürün | Ocak | Şubat | Mart") & vbCrLf & _
                    HAMU_Text("Bu örnekte sabit sütun sayısı 3'tür."), _
                title:=HAMU_APP_NAME & HAMU_Text(" - Sütunları Satırlara Dönüştür"), _
                Default:=1, _
                ValueType:=1)

    If VarType(answer) = vbBoolean Then
        If answer = False Then Exit Sub
    End If

    If CDbl(answer) <> Fix(CDbl(answer)) Then
        MsgBox HAMU_Text("Sabit sütun sayısı tam sayı olmalıdır."), _
               vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    fixedCols = CLng(answer)

    If fixedCols < 1 Or fixedCols >= rng.columns.count Then
        MsgBox HAMU_Text("Sabit sütun sayısı 1 ile ") & _
               (rng.columns.count - 1) & HAMU_Text(" arasında olmalıdır."), _
               vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    rowCount = rng.rows.count
    colCount = rng.columns.count

    resultRows = CDbl(rowCount - 1) * CDbl(colCount - fixedCols)

    If resultRows > ActiveSheet.rows.count - 1 Then
        MsgBox HAMU_Text("Oluşacak sonuç ") & Format$(resultRows, "#,##0") & _
               HAMU_Text(" satırdır ve Excel'in sayfa satır sınırını aşar."), _
               vbCritical, HAMU_APP_NAME
        Exit Sub
    End If

    outRows = CLng(resultRows)
    outCols = fixedCols + 2

    srcArr = rng.Value2

    ReDim outArr(1 To outRows, 1 To outCols)

    outRow = 1

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 2 To rowCount

        For c = fixedCols + 1 To colCount
            For f = 1 To fixedCols

                If rng.Cells(r, f).MergeCells Then
                    outArr(outRow, f) = _
                        rng.Cells(r, f).MergeArea.Cells(1, 1).value
                Else
                    outArr(outRow, f) = srcArr(r, f)
                End If

            Next f
            If rng.Cells(1, c).MergeCells Then
                outArr(outRow, fixedCols + 1) = _
                    HAMU_UnpivotHeader(rng.Cells(1, c).MergeArea.Cells(1, 1))
            Else
                outArr(outRow, fixedCols + 1) = HAMU_UnpivotHeader(rng.Cells(1, c))
            End If
            outArr(outRow, fixedCols + 2) = srcArr(r, c)

            outRow = outRow + 1

        Next c

    Next r

    Set outSh = HAMU_GetOrCreateSheet( _
                    "UNPIVOT_" & Format$(Now, "hhmmss"))
    For f = 1 To fixedCols
        outSh.Cells(1, f).value = rng.Cells(1, f).value
    Next f

    outSh.Cells(1, fixedCols + 1).value = "Alan"
    outSh.Cells(1, fixedCols + 2).value = HAMU_Text("Değer")

    outSh.columns(fixedCols + 1).NumberFormat = "@"
    outSh.Cells(2, 1) _
         .Resize(outRows, outCols).value = outArr

    With outSh
        With .range(.Cells(1, 1), .Cells(1, outCols))

            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack

        End With

        .range(.Cells(1, 1), .Cells(outRows + 1, outCols)).AutoFilter
        .columns.AutoFit

    End With

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox HAMU_Text("UNPIVOT tamamlandı.") & vbCrLf & vbCrLf & _
           "Sabit sütun: " & fixedCols & vbCrLf & _
           HAMU_Text("Satıra çevrilen sütun: ") & (colCount - fixedCols) & vbCrLf & _
           HAMU_Text("Oluşturulan veri satırı: ") & Format$(outRows, "#,##0") & vbCrLf & _
           HAMU_Text("Sonuç sayfası: ") & outSh.name, _
           vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError "HAMU_UnpivotColumnsToRows", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_SmartJoin()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SmartJoin") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim leftRng As range
    Dim rightRng As range
    Dim outSh As Worksheet

    Dim leftKeys As Collection
    Dim rightKeys As Collection
    Dim rightCols As Collection

    Dim leftInput As String
    Dim rightInput As String
    Dim bringInput As String

    Dim answer As Variant
    Dim joinType As Long

    Dim rightMap As Object
    Dim leftKeySet As Object
    Dim matchedRight As Object
    Dim keyOrder As Collection
    Dim rowsForKey As Collection

    Dim leftArr As Variant
    Dim rightArr As Variant
    Dim outArr() As Variant

    Dim r As Long
    Dim rr As Long
    Dim c As Long
    Dim outRow As Long
    Dim outCol As Long
    Dim outputRows As Long

    Dim key As String
    Dim k As Variant

    Set leftRng = HAMU_PromptRange( _
                    HAMU_Text("Eşleştirme için SOL tabloyu seçin.") & vbCrLf & _
                    HAMU_Text("İlk satır başlık olmalıdır."))

    If leftRng Is Nothing Then Exit Sub

    Set rightRng = HAMU_PromptRange( _
                    HAMU_Text("Eşleştirme için SAĞ tabloyu seçin.") & vbCrLf & _
                    HAMU_Text("İlk satır başlık olmalıdır."))

    If rightRng Is Nothing Then Exit Sub

    If leftRng.areas.count > 1 Or rightRng.areas.count > 1 Then

        MsgBox HAMU_Text("Her iki veri kümesi de tek parça bir aralık olmalıdır."), _
               vbExclamation, HAMU_APP_NAME

        Exit Sub

    End If

    leftInput = HAMU_InputText( _
                    HAMU_Text("SOL tablodaki anahtar sütunları virgülle girin.") & vbCrLf & _
                    HAMU_Text("Örnek: store_id") & vbCrLf & _
                    "veya: Tarih, store_id, product_code", _
                    HAMU_APP_NAME & HAMU_Text(" - Anahtara Göre Tablo Eşleştir"))

    If Len(Trim$(leftInput)) = 0 Then Exit Sub

    rightInput = HAMU_InputText( _
                    HAMU_Text("SAĞ tablodaki karşılık gelen anahtar sütunları virgülle girin.") & _
                    vbCrLf & _
                    HAMU_Text("Anahtar sayısı sol tabloyla aynı olmalıdır."), _
                    HAMU_APP_NAME & HAMU_Text(" - Anahtara Göre Tablo Eşleştir"))

    If Len(Trim$(rightInput)) = 0 Then Exit Sub

    Set leftKeys = HAMU_Data_ParseColumnList(leftRng, leftInput)
    Set rightKeys = HAMU_Data_ParseColumnList(rightRng, rightInput)

    If leftKeys.count <> rightKeys.count Then

        MsgBox _
            HAMU_Text("Sol ve sağ anahtar sayıları aynı olmalıdır."), _
            vbExclamation, HAMU_APP_NAME

        Exit Sub

    End If

    bringInput = HAMU_InputText( _
                    HAMU_Text("SAĞ tablodan getirilecek sütunları girin.") & vbCrLf & _
                    HAMU_Text("Örnek: store_name, city") & vbCrLf & vbCrLf & _
                    HAMU_Text("Boş bırakırsanız anahtarlar dışındaki tüm sağ sütunlar alınır."), _
                    HAMU_APP_NAME & HAMU_Text(" - Anahtara Göre Tablo Eşleştir"))

    If gHAMU_InputCancelled Then Exit Sub
    If Len(Trim$(bringInput)) = 0 Then

        Set rightCols = HAMU_Data_AllNonKeyColumns( _
                            rightRng.columns.count, _
                            rightKeys)

    Else

        Set rightCols = HAMU_Data_ParseColumnList( _
                            rightRng, _
                            bringInput)

    End If

    joinType = HAMU_ChooseIndex(HAMU_CommandLabel("HAMU_SmartJoin"), HAMU_Text("Hangi kayıtlar sonuçta yer alsın? Sol tablo ilk seçtiğiniz, sağ tablo ikinci seçtiğiniz veridir."), Array(HAMU_Text("İlk tablonun tüm kayıtları (LEFT)"), HAMU_Text("Yalnız iki tabloda eşleşenler (INNER)"), HAMU_Text("İki tablonun tüm kayıtları (FULL)")))
    If joinType = 0 Then Exit Sub

    leftArr = leftRng.Value2
    rightArr = rightRng.Value2

    Set rightMap = CreateObject("Scripting.Dictionary")
    Set leftKeySet = CreateObject("Scripting.Dictionary")
    Set matchedRight = CreateObject("Scripting.Dictionary")
    Set keyOrder = New Collection

    rightMap.CompareMode = vbBinaryCompare
    leftKeySet.CompareMode = vbBinaryCompare

    For r = 2 To rightRng.rows.count

        key = HAMU_Data_BuildKeyFromRangeRow( _
                    rightRng, _
                    r, _
                    rightKeys)

        If Not rightMap.exists(key) Then

            Set rowsForKey = New Collection
            rightMap.Add key, rowsForKey
            keyOrder.Add key

        Else

            Set rowsForKey = rightMap(key)

        End If

        rowsForKey.Add r

    Next r
    For r = 2 To leftRng.rows.count

        key = HAMU_Data_BuildKeyFromRangeRow( _
                    leftRng, _
                    r, _
                    leftKeys)

        leftKeySet(key) = True

        If rightMap.exists(key) Then

            Set rowsForKey = rightMap(key)
            outputRows = outputRows + rowsForKey.count

        ElseIf joinType = 1 Or joinType = 3 Then

            outputRows = outputRows + 1

        End If

    Next r

    If joinType = 3 Then

        For r = 2 To rightRng.rows.count

            key = HAMU_Data_BuildKeyFromRangeRow( _
                        rightRng, _
                        r, _
                        rightKeys)

            If Not leftKeySet.exists(key) Then
                outputRows = outputRows + 1
            End If

        Next r

    End If

    If outputRows > ActiveSheet.rows.count - 1 Then

        MsgBox _
            HAMU_Text("JOIN sonucu Excel satır sınırını aşıyor: ") & _
            Format$(outputRows, "#,##0"), _
            vbCritical, HAMU_APP_NAME

        Exit Sub

    End If

    ReDim outArr(1 To outputRows, _
                 1 To leftRng.columns.count + rightCols.count)

    outRow = 1

    For r = 2 To leftRng.rows.count

        key = HAMU_Data_BuildKeyFromRangeRow( _
                    leftRng, _
                    r, _
                    leftKeys)

        If rightMap.exists(key) Then

            Set rowsForKey = rightMap(key)

            For Each k In rowsForKey

                rr = CLng(k)

                HAMU_Data_FillJoinOutputRow _
                    outArr, _
                    outRow, _
                    leftRng, _
                    r, _
                    rightRng, _
                    rr, _
                    rightCols

                matchedRight(CStr(rr)) = True
                outRow = outRow + 1

            Next k

        ElseIf joinType = 1 Or joinType = 3 Then

            HAMU_Data_FillJoinOutputRow _
                outArr, _
                outRow, _
                leftRng, _
                r, _
                rightRng, _
                0, _
                rightCols

            outRow = outRow + 1

        End If

    Next r

    If joinType = 3 Then

        For r = 2 To rightRng.rows.count

            If Not matchedRight.exists(CStr(r)) Then

                HAMU_Data_FillJoinOutputRow _
                    outArr, _
                    outRow, _
                    leftRng, _
                    0, _
                    rightRng, _
                    r, _
                    rightCols

                outRow = outRow + 1

            End If

        Next r

    End If

    Set outSh = HAMU_GetOrCreateSheet( _
                    "JOIN_" & Format$(Now, "hhmmss"))

    outCol = 1

    For c = 1 To leftRng.columns.count

        outSh.Cells(1, outCol).value = leftRng.Cells(1, c).value
        outCol = outCol + 1

    Next c

    For c = 1 To rightCols.count

        rr = CLng(rightCols(c))

        outSh.Cells(1, outCol).value = _
            HAMU_Data_RightHeaderName( _
                leftRng, _
                CStr(rightRng.Cells(1, rr).value))

        outCol = outCol + 1

    Next c

    If outputRows > 0 Then

        HAMU_Data_PrepareJoinFormats _
            outSh, _
            leftRng, _
            rightRng, _
            rightCols

        outSh.Cells(2, 1) _
             .Resize(outputRows, _
                     leftRng.columns.count + rightCols.count) _
             .Value2 = outArr

    End If

    With outSh

        With .range( _
                .Cells(1, 1), _
                .Cells(1, leftRng.columns.count + rightCols.count))

            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack

        End With

        If outputRows > 0 Then

            .range( _
                .Cells(1, 1), _
                .Cells(outputRows + 1, _
                       leftRng.columns.count + rightCols.count)) _
                .AutoFilter

        End If

        .columns.AutoFit

    End With

    MsgBox _
        HAMU_Text("Anahtara Göre Tablo Eşleştir tamamlandı.") & vbCrLf & _
        HAMU_Text("Oluşturulan satır: ") & Format$(outputRows, "#,##0") & vbCrLf & _
        HAMU_Text("Sonuç sayfası: ") & outSh.name, _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_SmartJoin", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_FillDownBlanks()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_FillDownBlanks") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim r As Long
    Dim c As Long
    Dim changed As Long
    Dim lastValue As Variant
    Dim hasLast As Boolean

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Boş hücreleri yukarıdaki değerle doldurmak istediğiniz alanı seçin."))

    If rng Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For c = 1 To rng.columns.count

        hasLast = False

        For r = 1 To rng.rows.count

            If HAMU_Data_IsBlankCell(rng.Cells(r, c)) Then

                If hasLast Then

                    rng.Cells(r, c).value = lastValue
                    changed = changed + 1

                End If

            Else

                lastValue = rng.Cells(r, c).value
                hasLast = True

            End If

        Next r

    Next c

CleanExit:

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        changed & HAMU_Text(" boş hücre yukarıdaki değerle dolduruldu."), _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError "HAMU_FillDownBlanks", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_MergeDuplicates()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_MergeDuplicates") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim keyInput As String
    Dim sumInput As String

    Dim keyCols As Collection
    Dim sumCols As Collection

    Dim dictRows As Object
    Dim dictCounts As Object
    Dim keyOrder As Collection

    Dim rowData As Variant
    Dim outArr() As Variant

    Dim key As String
    Dim k As Variant

    Dim r As Long
    Dim c As Long
    Dim outRow As Long

    Dim currentValue As Variant

    Dim outSh As Worksheet

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Tekrar eden kayıtları birleştirmek istediğiniz veri alanını seçin.") & _
                vbCrLf & HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    keyInput = HAMU_InputText( _
                    HAMU_Text("Tekrarı belirleyecek anahtar sütunları virgülle girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Tarih, store_id, product_code"), _
                    HAMU_APP_NAME & HAMU_Text(" - Tekrarları Birleştir"))

    If Len(Trim$(keyInput)) = 0 Then Exit Sub

    Set keyCols = HAMU_Data_ParseColumnList(rng, keyInput)

    sumInput = HAMU_InputText( _
                    HAMU_Text("Toplanacak sayısal sütunları girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Satış, Stok") & vbCrLf & vbCrLf & _
                    HAMU_Text("Hiçbir sütun toplanmayacaksa boş bırakın."), _
                    HAMU_APP_NAME & HAMU_Text(" - Tekrarları Birleştir"))

    If Len(Trim$(sumInput)) > 0 Then

        Set sumCols = HAMU_Data_ParseColumnList(rng, sumInput)

        If HAMU_Data_ColumnListsOverlap(keyCols, sumCols) Then

            MsgBox _
                HAMU_Text("Anahtar sütunlar aynı zamanda toplama sütunu olamaz."), _
                vbExclamation, HAMU_APP_NAME

            Exit Sub

        End If

    Else

        Set sumCols = New Collection

    End If

    Set dictRows = CreateObject("Scripting.Dictionary")
    Set dictCounts = CreateObject("Scripting.Dictionary")
    Set keyOrder = New Collection

    dictRows.CompareMode = vbBinaryCompare

    For r = 2 To rng.rows.count

        key = HAMU_Data_BuildKeyFromRangeRow( _
                    rng, _
                    r, _
                    keyCols)

        If Not dictRows.exists(key) Then

            ReDim rowData(1 To rng.columns.count)

            For c = 1 To rng.columns.count

                rowData(c) = _
                    HAMU_Data_SplitOutputValue( _
                        rng.Cells(r, c), _
                        rng.Cells(1, c).Value2)

            Next c

            dictRows.Add key, rowData
            dictCounts.Add key, 1
            keyOrder.Add key

        Else

            rowData = dictRows(key)

            For c = 1 To rng.columns.count

                currentValue = rng.Cells(r, c).Value2

                If HAMU_Data_CollectionContainsLong(sumCols, c) Then

                    If IsNumeric(currentValue) Then

                        If IsNumeric(rowData(c)) Then
                            rowData(c) = CDbl(rowData(c)) + CDbl(currentValue)
                        Else
                            rowData(c) = CDbl(currentValue)
                        End If

                    End If

                ElseIf Not HAMU_Data_CollectionContainsLong(keyCols, c) Then

                    If Len(Trim$(CStr(rowData(c)))) = 0 And _
                       Len(Trim$(CStr(currentValue))) > 0 Then

                        rowData(c) = _
                            HAMU_Data_SplitOutputValue( _
                                rng.Cells(r, c), _
                                rng.Cells(1, c).Value2)

                    End If

                End If

            Next c

            dictRows(key) = rowData
            dictCounts(key) = CLng(dictCounts(key)) + 1

        End If

    Next r

    ReDim outArr(1 To dictRows.count, _
                 1 To rng.columns.count + 1)

    outRow = 1

    For Each k In keyOrder

        rowData = dictRows(CStr(k))

        For c = 1 To rng.columns.count
            outArr(outRow, c) = rowData(c)
        Next c

        outArr(outRow, rng.columns.count + 1) = dictCounts(CStr(k))
        outRow = outRow + 1

    Next k

    Set outSh = HAMU_GetOrCreateSheet( _
                    "TEKRAR_BIRLESIK_" & Format$(Now, "hhmmss"))

    For c = 1 To rng.columns.count
        outSh.Cells(1, c).value = rng.Cells(1, c).value
    Next c

    outSh.Cells(1, rng.columns.count + 1).value = HAMU_Text("HAMU_KayıtSayısı")

    HAMU_Data_PrepareOutputWeekColumns _
        outSh, _
        rng, _
        rng.columns.count

    If dictRows.count > 0 Then

        outSh.Cells(2, 1) _
             .Resize(dictRows.count, rng.columns.count + 1) _
             .Value2 = outArr

    End If

    With outSh

        With .range(.Cells(1, 1), .Cells(1, rng.columns.count + 1))
            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack
        End With

        .range(.Cells(1, 1), _
               .Cells(dictRows.count + 1, rng.columns.count + 1)).AutoFilter

        .columns.AutoFit

    End With

    MsgBox _
        HAMU_Text("Tekrarlar birleştirildi.") & vbCrLf & _
        HAMU_Text("Kaynak satır: ") & Format$(rng.rows.count - 1, "#,##0") & vbCrLf & _
        HAMU_Text("Sonuç satır: ") & Format$(dictRows.count, "#,##0"), _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_MergeDuplicates", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_StandardizeHeaders()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_StandardizeHeaders") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim answer As Variant
    Dim mode As Long

    Dim seen As Object
    Dim c As Long

    Dim oldText As String
    Dim newText As String

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Başlıklarını standardize etmek istediğiniz veri alanını seçin.") & _
                vbCrLf & HAMU_Text("İlk satır başlık olarak kullanılacaktır."))

    If rng Is Nothing Then Exit Sub

    mode = HAMU_ChooseIndex(HAMU_L("Başlıkları Düzelt", "Normalize Headers"), HAMU_L("Teknik biçim Türkçe harfleri ASCII karşılığına dönüştürür; temiz başlık yalnız fazla boşlukları kaldırır.", "Technical format transliterates Turkish letters to ASCII; clean headers only remove excess whitespace."), Array("snake_case", HAMU_L("Temiz başlık", "Clean header")))
    If mode = 0 Then Exit Sub

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For c = 1 To rng.columns.count

        oldText = CStr(rng.Cells(1, c).value)

        If mode = 1 Then
            newText = HAMU_Data_ToSnakeCase(oldText)
        Else
            newText = HAMU_Data_CleanHeaderText(oldText)
        End If

        If Len(newText) = 0 Then
            newText = "sütun_" & c
        End If

        newText = HAMU_Data_MakeUniqueHeader(newText, seen)

        rng.Cells(1, c).value = newText

    Next c

    MsgBox _
        rng.columns.count & HAMU_Text(" başlık standardize edildi."), _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_StandardizeHeaders", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_ConvertColumnTypes()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ConvertColumnTypes") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim cols As Collection

    Dim inputText As String
    Dim answer As Variant
    Dim mode As Long

    Dim i As Long
    Dim r As Long
    Dim idx As Long

    Dim v As Variant
    Dim converted As Long
    Dim Failed As Long

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Veri tipi dönüştürülecek veri alanını seçin.") & vbCrLf & _
                HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    inputText = HAMU_InputText( _
                    HAMU_Text("Dönüştürülecek sütunları virgülle girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Satış, Stok") & vbCrLf & _
                    "veya: 4,5", _
                    HAMU_APP_NAME & HAMU_Text(" - Veri Tipini Düzelt"))

    If Len(Trim$(inputText)) = 0 Then Exit Sub

    Set cols = HAMU_Data_ParseColumnList(rng, inputText)

    answer = HAMU_InputValue( _
                Prompt:= _
                    HAMU_Text("Hedef veri tipini seçin:") & vbCrLf & vbCrLf & _
                    "1 - Metin" & vbCrLf & _
                    HAMU_Text("2 - Sayı") & vbCrLf & _
                    "3 - Tarih" & vbCrLf & _
                    HAMU_Text("4 - Tam sayı") & vbCrLf & _
                    HAMU_Text("5 - Yüzde"), _
                title:=HAMU_APP_NAME & HAMU_Text(" - Veri Tipini Düzelt"), _
                Default:=2, _
                ValueType:=1)

    If VarType(answer) = vbBoolean Then
        If answer = False Then Exit Sub
    End If

    mode = CLng(answer)

    If mode < 1 Or mode > 5 Then

        MsgBox HAMU_Text("1 ile 5 arasında bir veri tipi seçin."), _
               vbExclamation, HAMU_APP_NAME

        Exit Sub

    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 1 To cols.count

        idx = CLng(cols(i))

        If mode = 1 Then
            rng.columns(idx).NumberFormat = "@"
        ElseIf mode = 3 Then
            rng.columns(idx).NumberFormat = "dd.mm.yyyy"
        ElseIf mode = 5 Then
            rng.columns(idx).NumberFormat = "0.00%"
        Else
            rng.columns(idx).NumberFormat = "General"
        End If

        For r = 2 To rng.rows.count

            If Not HAMU_Data_IsBlankCell(rng.Cells(r, idx)) Then

                v = rng.Cells(r, idx).Value2

                Select Case mode

                    Case 1

                        rng.Cells(r, idx).value = _
                            HAMU_Data_SplitCellText( _
                                rng.Cells(r, idx), _
                                rng.Cells(1, idx).Value2)

                        converted = converted + 1

                    Case 2

                        If IsNumeric(v) Then
                            rng.Cells(r, idx).Value2 = CDbl(v)
                            converted = converted + 1
                        Else
                            Failed = Failed + 1
                        End If

                    Case 3

                        If IsDate(rng.Cells(r, idx).value) Then
                            rng.Cells(r, idx).value = CDate(rng.Cells(r, idx).value)
                            converted = converted + 1
                        Else
                            Failed = Failed + 1
                        End If

                    Case 4

                        If IsNumeric(v) Then
                            rng.Cells(r, idx).Value2 = CLng(CDbl(v))
                            converted = converted + 1
                        Else
                            Failed = Failed + 1
                        End If

                    Case 5

                        If IsNumeric(v) Then

                            If Abs(CDbl(v)) > 1 Then
                                rng.Cells(r, idx).Value2 = CDbl(v) / 100#
                            Else
                                rng.Cells(r, idx).Value2 = CDbl(v)
                            End If

                            converted = converted + 1

                        Else

                            Failed = Failed + 1

                        End If

                End Select

            End If

        Next r

    Next i

CleanExit:

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Veri tipi dönüşümü tamamlandı.") & vbCrLf & _
        HAMU_Text("Dönüştürülen: ") & converted & vbCrLf & _
        HAMU_Text("Dönüştürülemeyen: ") & Failed, _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError "HAMU_ConvertColumnTypes", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_PivotRowsToColumns()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_PivotRowsToColumns") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range

    Dim keyInput As String
    Dim pivotInput As String
    Dim valueInput As String

    Dim keyCols As Collection
    Dim pivotCol As Long
    Dim valueCol As Long

    Dim answer As Variant
    Dim aggMode As Long

    Dim groupMap As Object
    Dim groupLabels As Object
    Dim groupOrder As Collection

    Dim pivotMap As Object
    Dim pivotOrder As Collection

    Dim labelData As Variant
    Dim outArr() As Variant

    Dim key As String
    Dim pivotLabel As String
    Dim k As Variant

    Dim r As Long
    Dim i As Long
    Dim outRow As Long
    Dim outCol As Long
    Dim idx As Long
    Dim gIndex As Long
    Dim pIndex As Long

    Dim valueData As Variant
    Dim outSh As Worksheet

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Çapraz tablo oluşturulacak veri alanını seçin.") & vbCrLf & _
                HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    keyInput = HAMU_InputText( _
                    HAMU_Text("Satırlarda sabit kalacak anahtar sütunları girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Mağaza, Ürün"), _
                    HAMU_APP_NAME & HAMU_Text(" - Çapraz Tablo"))

    If Len(Trim$(keyInput)) = 0 Then Exit Sub

    Set keyCols = HAMU_Data_ParseColumnList(rng, keyInput)

    pivotInput = HAMU_InputText( _
                    HAMU_Text("Yeni sütun başlıklarını oluşturacak alanı girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Ay"), _
                    HAMU_APP_NAME & HAMU_Text(" - Çapraz Tablo"))

    If Len(Trim$(pivotInput)) = 0 Then Exit Sub

    pivotCol = HAMU_Data_ResolveColumn(rng, pivotInput)

    valueInput = HAMU_InputText( _
                    HAMU_Text("Hücrelere yazılacak değer sütununu girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Satış"), _
                    HAMU_APP_NAME & HAMU_Text(" - Çapraz Tablo"))

    If Len(Trim$(valueInput)) = 0 Then Exit Sub

    valueCol = HAMU_Data_ResolveColumn(rng, valueInput)

    answer = HAMU_InputValue( _
                Prompt:= _
                    HAMU_Text("Aynı hücreye birden fazla kayıt düşerse:") & vbCrLf & vbCrLf & _
                    "1 - Topla" & vbCrLf & _
                    HAMU_Text("2 - İlk değeri kullan") & vbCrLf & _
                    HAMU_Text("3 - Kayıt sayısını yaz"), _
                title:=HAMU_APP_NAME & HAMU_Text(" - Çapraz Tablo"), _
                Default:=1, _
                ValueType:=1)

    If VarType(answer) = vbBoolean Then
        If answer = False Then Exit Sub
    End If

    aggMode = CLng(answer)

    If aggMode < 1 Or aggMode > 3 Then
        MsgBox HAMU_Text("1 ile 3 arasında seçim yapın."), vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    Set groupMap = CreateObject("Scripting.Dictionary")
    Set groupLabels = CreateObject("Scripting.Dictionary")
    Set groupOrder = New Collection

    Set pivotMap = CreateObject("Scripting.Dictionary")
    Set pivotOrder = New Collection

    groupMap.CompareMode = vbBinaryCompare
    pivotMap.CompareMode = vbTextCompare
    For r = 2 To rng.rows.count

        key = HAMU_Data_BuildKeyFromRangeRow(rng, r, keyCols)

        If Not groupMap.exists(key) Then

            groupMap.Add key, groupMap.count + 1
            groupOrder.Add key

            ReDim labelData(1 To keyCols.count)

            For i = 1 To keyCols.count

                idx = CLng(keyCols(i))

                labelData(i) = _
                    HAMU_Data_SplitOutputValue( _
                        rng.Cells(r, idx), _
                        rng.Cells(1, idx).Value2)

            Next i

            groupLabels.Add key, labelData

        End If

        pivotLabel = HAMU_Data_SplitCellText( _
                        rng.Cells(r, pivotCol), _
                        rng.Cells(1, pivotCol).Value2)

        If Not pivotMap.exists(pivotLabel) Then

            pivotMap.Add pivotLabel, pivotMap.count + 1
            pivotOrder.Add pivotLabel

        End If

    Next r

    ReDim outArr( _
        1 To groupMap.count, _
        1 To keyCols.count + pivotMap.count)
    For Each k In groupOrder

        gIndex = CLng(groupMap(CStr(k)))
        labelData = groupLabels(CStr(k))

        For i = 1 To keyCols.count
            outArr(gIndex, i) = labelData(i)
        Next i

    Next k
    For r = 2 To rng.rows.count

        key = HAMU_Data_BuildKeyFromRangeRow(rng, r, keyCols)
        gIndex = CLng(groupMap(key))

        pivotLabel = HAMU_Data_SplitCellText( _
                        rng.Cells(r, pivotCol), _
                        rng.Cells(1, pivotCol).Value2)

        pIndex = CLng(pivotMap(pivotLabel))
        outCol = keyCols.count + pIndex

        valueData = rng.Cells(r, valueCol).Value2

        Select Case aggMode

            Case 1

                If IsNumeric(valueData) Then

                    If IsNumeric(outArr(gIndex, outCol)) Then
                        outArr(gIndex, outCol) = _
                            CDbl(outArr(gIndex, outCol)) + CDbl(valueData)
                    Else
                        outArr(gIndex, outCol) = CDbl(valueData)
                    End If

                End If

            Case 2

                If Len(CStr(outArr(gIndex, outCol))) = 0 Then

                    outArr(gIndex, outCol) = _
                        HAMU_Data_SplitOutputValue( _
                            rng.Cells(r, valueCol), _
                            rng.Cells(1, valueCol).Value2)

                End If

            Case 3

                If IsNumeric(outArr(gIndex, outCol)) Then
                    outArr(gIndex, outCol) = CLng(outArr(gIndex, outCol)) + 1
                Else
                    outArr(gIndex, outCol) = 1
                End If

        End Select

    Next r

    Set outSh = HAMU_GetOrCreateSheet( _
                    "CAPRAZ_" & Format$(Now, "hhmmss"))

    outCol = 1

    For i = 1 To keyCols.count

        idx = CLng(keyCols(i))
        outSh.Cells(1, outCol).value = rng.Cells(1, idx).value

        If HAMU_Data_IsWeekColumn(rng, idx) Then
            outSh.columns(outCol).NumberFormat = "@"
        End If

        outCol = outCol + 1

    Next i

    For Each k In pivotOrder

        outSh.Cells(1, outCol).value = CStr(k)
        outCol = outCol + 1

    Next k

    If groupMap.count > 0 Then

        outSh.Cells(2, 1) _
             .Resize(groupMap.count, keyCols.count + pivotMap.count) _
             .Value2 = outArr

    End If

    With outSh

        With .range(.Cells(1, 1), _
                    .Cells(1, keyCols.count + pivotMap.count))

            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack

        End With

        .range(.Cells(1, 1), _
               .Cells(groupMap.count + 1, _
                      keyCols.count + pivotMap.count)).AutoFilter

        .columns.AutoFit

    End With

    MsgBox _
        HAMU_Text("Çapraz tablo oluşturuldu.") & vbCrLf & _
        HAMU_Text("Satır grubu: ") & groupMap.count & vbCrLf & _
        "Yeni veri sütunu: " & pivotMap.count, _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_PivotRowsToColumns", hamuErrorNumber, hamuErrorText

End Sub

Public Function HAMU_UnpivotHeader(ByVal cell As range) As String
    If IsError(cell.Value2) Then
        HAMU_UnpivotHeader = cell.text
    ElseIf IsNumeric(cell.Value2) And LCase$(cell.NumberFormat) = "general" Then
        HAMU_UnpivotHeader = Trim$(Str$(cell.Value2))
    ElseIf VarType(cell.Value2) = vbString Then
        HAMU_UnpivotHeader = cell.Value2
    Else
        HAMU_UnpivotHeader = cell.text
    End If
End Function


