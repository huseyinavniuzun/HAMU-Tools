Attribute VB_Name = "modHAMU_Data_Quality"
Option Explicit
Option Private Module

Public Sub HAMU_DataProfile()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_DataProfile") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim outSh As Worksheet

    Dim c As Long
    Dim r As Long
    Dim outRow As Long

    Dim dataRows As Long
    Dim filledCount As Long
    Dim blankCount As Long
    Dim errorCount As Long

    Dim numCount As Long
    Dim dateCount As Long
    Dim textCount As Long
    Dim boolCount As Long

    Dim uniqueCount As Long
    Dim repeatCount As Long

    Dim minNum As Double
    Dim maxNum As Double
    Dim sumNum As Double
    Dim numericSeen As Boolean

    Dim minDate As Double
    Dim maxDate As Double
    Dim dateSeen As Boolean

    Dim dict As Object
    Dim cell As range

    Dim kind As String
    Dim profileType As String
    Dim key As String
    Dim noteText As String

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Profilini çıkarmak istediğiniz veri alanını seçin.") & vbCrLf & _
                HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    If rng.areas.count > 1 Or rng.rows.count < 2 Then

        MsgBox _
            HAMU_Text("Tek parça bir veri alanı seçin. İlk satır başlık olmalıdır."), _
            vbExclamation, HAMU_APP_NAME

        Exit Sub

    End If

    dataRows = rng.rows.count - 1

    Set outSh = HAMU_GetOrCreateSheet( _
                    "VERI_PROFIL_" & Format$(Now, "hhmmss"))

    outSh.range("A1:L1").value = Array( _
        "Sütun", _
        HAMU_Text("Tahmini Tür"), _
        HAMU_Text("Veri Satırı"), _
        "Dolu", _
        HAMU_Text("Boş"), _
        "Benzersiz", _
        "Tekrar", _
        "Hata", _
        "Minimum", _
        "Maksimum", _
        "Ortalama", _
        "Not")

    outRow = 2

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For c = 1 To rng.columns.count

        filledCount = 0
        blankCount = 0
        errorCount = 0

        numCount = 0
        dateCount = 0
        textCount = 0
        boolCount = 0

        sumNum = 0
        numericSeen = False
        dateSeen = False

        Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = vbTextCompare

        For r = 2 To rng.rows.count

            Set cell = rng.Cells(r, c)

            If HAMU_Data_IsBlankCell(cell) Then

                blankCount = blankCount + 1

            Else

                filledCount = filledCount + 1

                If IsError(cell.value) Then

                    errorCount = errorCount + 1

                Else

                    kind = HAMU_Data_CellKind(cell)

                    Select Case kind

                        Case "Tarih"

                            dateCount = dateCount + 1

                            If IsDate(cell.value) Then

                                If Not dateSeen Then
                                    minDate = CDbl(CDate(cell.value))
                                    maxDate = minDate
                                    dateSeen = True
                                Else
                                    If CDbl(CDate(cell.value)) < minDate Then _
                                        minDate = CDbl(CDate(cell.value))
                                    If CDbl(CDate(cell.value)) > maxDate Then _
                                        maxDate = CDbl(CDate(cell.value))
                                End If

                            End If

                        Case HAMU_Text("Sayı")

                            numCount = numCount + 1

                            If IsNumeric(cell.Value2) Then

                                If Not numericSeen Then
                                    minNum = CDbl(cell.Value2)
                                    maxNum = minNum
                                    numericSeen = True
                                Else
                                    If CDbl(cell.Value2) < minNum Then _
                                        minNum = CDbl(cell.Value2)
                                    If CDbl(cell.Value2) > maxNum Then _
                                        maxNum = CDbl(cell.Value2)
                                End If

                                sumNum = sumNum + CDbl(cell.Value2)

                            End If

                        Case HAMU_Text("Mantıksal")
                            boolCount = boolCount + 1

                        Case Else
                            textCount = textCount + 1

                    End Select

                    key = kind & "|" & HAMU_Data_ProfileKey(cell)

                    If Not dict.exists(key) Then
                        dict.Add key, True
                    End If

                End If

            End If

        Next r

        uniqueCount = dict.count
        repeatCount = filledCount - errorCount - uniqueCount

        If repeatCount < 0 Then repeatCount = 0

        profileType = HAMU_Data_ProfileType( _
                        numCount, _
                        dateCount, _
                        textCount, _
                        boolCount, _
                        errorCount)

        noteText = ""

        If blankCount > 0 Then
            noteText = blankCount & HAMU_Text(" boş hücre")
        End If

        If errorCount > 0 Then

            If Len(noteText) > 0 Then noteText = noteText & "; "

            noteText = noteText & errorCount & " hata"

        End If

        If repeatCount > 0 Then

            If Len(noteText) > 0 Then noteText = noteText & "; "

            noteText = noteText & repeatCount & " tekrar"

        End If

        outSh.Cells(outRow, 1).value = rng.Cells(1, c).value
        outSh.Cells(outRow, 2).value = profileType
        outSh.Cells(outRow, 3).value = dataRows
        outSh.Cells(outRow, 4).value = filledCount
        outSh.Cells(outRow, 5).value = blankCount
        outSh.Cells(outRow, 6).value = uniqueCount
        outSh.Cells(outRow, 7).value = repeatCount
        outSh.Cells(outRow, 8).value = errorCount

        If profileType = "Tarih" And dateSeen Then

            outSh.Cells(outRow, 9).value = CDate(minDate)
            outSh.Cells(outRow, 10).value = CDate(maxDate)

            outSh.Cells(outRow, 9).NumberFormat = "dd.mm.yyyy"
            outSh.Cells(outRow, 10).NumberFormat = "dd.mm.yyyy"

        ElseIf numericSeen Then

            outSh.Cells(outRow, 9).value = minNum
            outSh.Cells(outRow, 10).value = maxNum

            If numCount > 0 Then
                outSh.Cells(outRow, 11).value = sumNum / numCount
            End If

        End If

        outSh.Cells(outRow, 12).value = noteText

        outRow = outRow + 1

    Next c

    With outSh

        With .range(.Cells(1, 1), .Cells(1, 12))
            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack
        End With

        .range(.Cells(1, 1), .Cells(outRow - 1, 12)).AutoFilter
        .columns.AutoFit

    End With

CleanExit:

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Veri profili oluşturuldu.") & vbCrLf & _
        HAMU_Text("İncelenen sütun: ") & rng.columns.count & vbCrLf & _
        HAMU_Text("Sonuç sayfası: ") & outSh.name, _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError "HAMU_DataProfile", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_FindMissingCombinations()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_FindMissingCombinations") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim cols As Collection
    Dim inputText As String

    Dim valueSets As Collection
    Dim d As Object
    Dim values As Collection
    Dim existing As Object

    Dim parts() As String

    Dim candidateCount As Double
    Dim outCount As Long

    Dim tempArr() As Variant
    Dim finalArr() As Variant

    Dim r As Long
    Dim i As Long
    Dim idx As Long
    Dim c As Long

    Dim partText As String
    Dim key As String
    Dim k As Variant

    Dim answer As VbMsgBoxResult
    Dim outSh As Worksheet

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Eksik kombinasyonları kontrol edeceğiniz veri alanını seçin.") & _
                vbCrLf & HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    inputText = HAMU_InputText( _
                    HAMU_Text("Kombinasyon oluşturacak sütunları virgülle girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Mağaza, Ürün, Hafta"), _
                    HAMU_APP_NAME & " - Eksik Kombinasyonlar")

    If Len(Trim$(inputText)) = 0 Then Exit Sub

    Set cols = HAMU_Data_ParseColumnList(rng, inputText)

    If cols.count < 2 Then

        MsgBox _
            HAMU_Text("Eksik kombinasyon için en az iki sütun seçin."), _
            vbExclamation, HAMU_APP_NAME

        Exit Sub

    End If

    Set valueSets = New Collection
    Set existing = CreateObject("Scripting.Dictionary")

    existing.CompareMode = vbBinaryCompare

    For i = 1 To cols.count

        Set d = CreateObject("Scripting.Dictionary")
        d.CompareMode = vbTextCompare

        idx = CLng(cols(i))

        For r = 2 To rng.rows.count

            partText = HAMU_Data_SplitCellText( _
                            rng.Cells(r, idx), _
                            rng.Cells(1, idx).Value2)

            If Len(Trim$(partText)) > 0 Then

                If Not d.exists(partText) Then
                    d.Add partText, True
                End If

            End If

        Next r

        Set values = New Collection

        For Each k In d.keys
            values.Add CStr(k)
        Next k

        valueSets.Add values

    Next i

    For r = 2 To rng.rows.count

        key = ""

        For i = 1 To cols.count

            idx = CLng(cols(i))

            partText = HAMU_Data_SplitCellText( _
                            rng.Cells(r, idx), _
                            rng.Cells(1, idx).Value2)

            key = key & Len(partText) & ":" & partText & "|"

        Next i

        existing(key) = True

    Next r

    candidateCount = 1

    For i = 1 To valueSets.count

        Set values = valueSets(i)
        candidateCount = candidateCount * values.count

    Next i

    If candidateCount > 500000 Then

        MsgBox _
            HAMU_Text("Olası kombinasyon sayısı çok yüksek: ") & _
            Format$(candidateCount, "#,##0") & vbCrLf & _
            HAMU_Text("Güvenlik için işlem durduruldu. Daha az sütun/değer seçin."), _
            vbCritical, HAMU_APP_NAME

        Exit Sub

    ElseIf candidateCount > 100000 Then

        answer = MsgBox( _
                    HAMU_Text("Yaklaşık ") & Format$(candidateCount, "#,##0") & _
                    HAMU_L(" kombinasyon kontrol edilecek.", " combinations will be checked.") & vbCrLf & _
                    HAMU_L("Devam edilsin mi?", "Continue?"), _
                    vbQuestion + vbYesNo, _
                    HAMU_APP_NAME)

        If answer <> vbYes Then Exit Sub

    End If

    ReDim parts(1 To cols.count)
    ReDim tempArr(1 To CLng(candidateCount), 1 To cols.count)

    HAMU_Data_GenerateMissingCombinations _
        valueSets, _
        1, _
        parts, _
        existing, _
        tempArr, _
        outCount

    Set outSh = HAMU_GetOrCreateSheet( _
                    "EKSIK_KOMBIN_" & Format$(Now, "hhmmss"))

    For i = 1 To cols.count

        idx = CLng(cols(i))
        outSh.Cells(1, i).value = rng.Cells(1, idx).value

    Next i

    If outCount > 0 Then

        ReDim finalArr(1 To outCount, 1 To cols.count)

        For r = 1 To outCount
            For c = 1 To cols.count
                finalArr(r, c) = tempArr(r, c)
            Next c
        Next r

        outSh.Cells(2, 1) _
             .Resize(outCount, cols.count) _
             .Value2 = finalArr

    End If

    With outSh

        With .range(.Cells(1, 1), .Cells(1, cols.count))
            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack
        End With

        If outCount > 0 Then
            .range(.Cells(1, 1), .Cells(outCount + 1, cols.count)).AutoFilter
        End If

        .columns.AutoFit

    End With

    MsgBox _
        HAMU_Text("Eksik kombinasyon kontrolü tamamlandı.") & vbCrLf & _
        HAMU_Text("Olası kombinasyon: ") & Format$(candidateCount, "#,##0") & vbCrLf & _
        "Eksik kombinasyon: " & Format$(outCount, "#,##0"), _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_FindMissingCombinations", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_DataQualityReport()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_DataQualityReport") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim requiredInput As String
    Dim keyInput As String

    Dim requiredCols As Collection
    Dim keyCols As Collection

    Dim duplicateCounts As Object

    Dim dominantTypes() As String

    Dim outSh As Worksheet

    Dim r As Long
    Dim c As Long
    Dim i As Long
    Dim outRow As Long
    Dim idx As Long

    Dim key As String
    Dim actualType As String
    Dim detailText As String

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Kalite kontrolü yapılacak veri alanını seçin.") & vbCrLf & _
                HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    requiredInput = HAMU_InputText( _
                    HAMU_Text("Zorunlu olması gereken sütunları virgülle girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Tarih, store_id, product_code") & vbCrLf & vbCrLf & _
                    HAMU_Text("Zorunlu alan kontrolü istemiyorsanız boş bırakın."), _
                    HAMU_APP_NAME & " - Veri Kalitesi")

    If gHAMU_InputCancelled Then Exit Sub
    If Len(Trim$(requiredInput)) > 0 Then
        Set requiredCols = HAMU_Data_ParseColumnList(rng, requiredInput)
    Else
        Set requiredCols = New Collection
    End If

    keyInput = HAMU_InputText( _
                    HAMU_Text("Tekrar kontrolü için anahtar sütunları girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Tarih, store_id, product_code") & vbCrLf & vbCrLf & _
                    HAMU_Text("Tekrar kontrolü istemiyorsanız boş bırakın."), _
                    HAMU_APP_NAME & " - Veri Kalitesi")

    If gHAMU_InputCancelled Then Exit Sub
    If Len(Trim$(keyInput)) > 0 Then
        Set keyCols = HAMU_Data_ParseColumnList(rng, keyInput)
    Else
        Set keyCols = New Collection
    End If

    Set duplicateCounts = CreateObject("Scripting.Dictionary")
    duplicateCounts.CompareMode = vbBinaryCompare

    If keyCols.count > 0 Then

        For r = 2 To rng.rows.count

            key = HAMU_Data_BuildKeyFromRangeRow(rng, r, keyCols)

            If duplicateCounts.exists(key) Then
                duplicateCounts(key) = CLng(duplicateCounts(key)) + 1
            Else
                duplicateCounts.Add key, 1
            End If

        Next r

    End If

    ReDim dominantTypes(1 To rng.columns.count)

    For c = 1 To rng.columns.count
        dominantTypes(c) = HAMU_Data_DominantColumnType(rng, c)
    Next c

    Set outSh = HAMU_GetOrCreateSheet( _
                    "KALITE_RAPOR_" & Format$(Now, "hhmmss"))

    outSh.range("A1:F1").value = Array( _
        HAMU_Text("Sorun Türü"), _
        HAMU_Text("Kaynak Satır"), _
        "Sütun", _
        HAMU_Text("Değer"), _
        "Beklenen / Durum", _
        HAMU_Text("Açıklama"))

    outRow = 2

    Application.ScreenUpdating = False

    For r = 2 To rng.rows.count
        For i = 1 To requiredCols.count

            idx = CLng(requiredCols(i))

            If HAMU_Data_IsBlankCell(rng.Cells(r, idx)) Then

                HAMU_Data_WriteQualityIssue _
                    outSh, _
                    outRow, _
                    HAMU_Text("Zorunlu Alan Boş"), _
                    rng, _
                    r, _
                    idx, _
                    "", _
                    HAMU_Text("Dolu değer"), _
                    HAMU_Text("Zorunlu sütun boş bırakılmış.")

            End If

        Next i
        For c = 1 To rng.columns.count

            If IsError(rng.Cells(r, c).value) Then

                HAMU_Data_WriteQualityIssue _
                    outSh, _
                    outRow, _
                    HAMU_Text("Excel Hatası"), _
                    rng, _
                    r, _
                    c, _
                    rng.Cells(r, c).text, _
                    HAMU_Text("Hatasız değer"), _
                    HAMU_Text("Hücre Excel hata değeri içeriyor.")

            ElseIf Not HAMU_Data_IsBlankCell(rng.Cells(r, c)) Then

                actualType = HAMU_Data_CellKind(rng.Cells(r, c))

                If dominantTypes(c) <> HAMU_Text("Karışık") And _
                   dominantTypes(c) <> HAMU_Text("Boş") And _
                   actualType <> dominantTypes(c) Then

                    detailText = _
                        HAMU_Text("Sütunun baskın tipi ") & dominantTypes(c) & _
                        HAMU_Text(", bu hücrenin tipi ") & actualType & "."

                    HAMU_Data_WriteQualityIssue _
                        outSh, _
                        outRow, _
                        HAMU_Text("Veri Tipi Sapması"), _
                        rng, _
                        r, _
                        c, _
                        rng.Cells(r, c).text, _
                        dominantTypes(c), _
                        detailText

                End If

            End If

        Next c
        If keyCols.count > 0 Then

            key = HAMU_Data_BuildKeyFromRangeRow(rng, r, keyCols)

            If CLng(duplicateCounts(key)) > 1 Then

                HAMU_Data_WriteQualityIssue _
                    outSh, _
                    outRow, _
                    "Tekrar Eden Anahtar", _
                    rng, _
                    r, _
                    CLng(keyCols(1)), _
                    HAMU_Data_KeyDisplay(rng, r, keyCols), _
                    "Tekil anahtar", _
                    "Bu anahtar kombinasyonu " & _
                    duplicateCounts(key) & " kez bulunuyor."

            End If

        End If

    Next r

    With outSh

        With .range(.Cells(1, 1), .Cells(1, 6))
            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack
        End With

        If outRow > 2 Then
            .range(.Cells(1, 1), .Cells(outRow - 1, 6)).AutoFilter
        End If

        .columns.AutoFit

    End With

    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Veri kalite raporu tamamlandı.") & vbCrLf & _
        "Bulunan sorun: " & Format$(outRow - 2, "#,##0") & vbCrLf & _
        HAMU_Text("Sonuç sayfası: ") & outSh.name, _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError "HAMU_DataQualityReport", hamuErrorNumber, hamuErrorText

End Sub


