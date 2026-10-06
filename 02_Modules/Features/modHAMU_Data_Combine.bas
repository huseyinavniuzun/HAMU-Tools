Attribute VB_Name = "modHAMU_Data_Combine"
Option Explicit
Option Private Module

Public Sub HAMU_CombineCSVs()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_CombineCSVs") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim fd As FileDialog
    Dim folder As String
    Dim f As String

    Dim wb As Workbook
    Dim sourceOwned As Boolean
    Dim ws As Worksheet
    Dim outSh As Worksheet
    Dim rng As range

    Dim firstFile As Boolean
    Dim outRow As Long
    Dim fileCount As Long

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .title = HAMU_Text("Birleştirilecek CSV dosyalarının bulunduğu klasörü seçin")
        If .show <> -1 Then Exit Sub
        folder = .SelectedItems(1)
    End With

    f = Dir$(folder & Application.PathSeparator & "*.csv")

    If Len(f) = 0 Then
        MsgBox HAMU_Text("Seçilen klasörde CSV dosyası bulunamadı."), _
               vbInformation, HAMU_APP_NAME
        Exit Sub
    End If

    Set outSh = HAMU_GetOrCreateSheet( _
                    "CSV_BIRLESIK_" & Format$(Now, "hhmmss"))

    outRow = 1
    firstFile = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Do While Len(f) > 0

        Set wb = HAMU_OpenSourceWorkbook(folder & Application.PathSeparator & f, sourceOwned)

        Set ws = wb.Worksheets(1)
        Set rng = HAMU_BoundedRange(ws.UsedRange, True)

        If WorksheetFunction.CountA(rng) > 0 Then

            If firstFile Then

                rng.copy Destination:=outSh.Cells(outRow, 1)

                outRow = outRow + rng.rows.count
                firstFile = False

            ElseIf rng.rows.count > 1 Then

                rng.Offset(1, 0) _
                   .Resize(rng.rows.count - 1, rng.columns.count) _
                   .copy Destination:=outSh.Cells(outRow, 1)

                outRow = outRow + rng.rows.count - 1

            End If

            fileCount = fileCount + 1

        End If

        If sourceOwned Then wb.Close SaveChanges:=False
        Set wb = Nothing

        f = Dir$

    Loop

    outSh.columns.AutoFit

CleanExit:

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    If fileCount > 0 Then
        MsgBox fileCount & HAMU_Text(" CSV dosyası birleştirildi.") & vbCrLf & _
               HAMU_Text("Sonuç sayfası: ") & outSh.name, _
               vbInformation, HAMU_APP_NAME
    End If

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    On Error Resume Next

    If Not wb Is Nothing Then
        If sourceOwned Then wb.Close SaveChanges:=False
    End If

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    On Error GoTo 0

    HAMU_ShowError "HAMU_CombineCSVs", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_ConsolidateSheets()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ConsolidateSheets") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim outSh As Worksheet
    Dim ws As Worksheet
    Dim rng As range

    Dim outRow As Long
    Dim sourceCol As Long
    Dim dataRows As Long
    Dim firstSheet As Boolean
    Dim sheetCount As Long

    Set outSh = HAMU_GetOrCreateSheet( _
                    "SEKMELER_" & Format$(Now, "hhmmss"))

    outRow = 1
    firstSheet = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each ws In ActiveWorkbook.Worksheets

        If ws.name <> outSh.name Then

            If WorksheetFunction.CountA(HAMU_BoundedRange(ws.UsedRange, True)) > 0 Then

                Set rng = HAMU_BoundedRange(ws.UsedRange, True)

                If firstSheet Then

                    sourceCol = rng.columns.count + 1

                    rng.copy Destination:=outSh.Cells(1, 1)

                    outSh.Cells(1, sourceCol).value = "KaynakSayfa"

                    dataRows = rng.rows.count - 1

                    If dataRows > 0 Then
                        outSh.Cells(2, sourceCol) _
                             .Resize(dataRows, 1).value = ws.name
                    End If

                    outRow = rng.rows.count + 1
                    firstSheet = False
                    sheetCount = sheetCount + 1

                Else

                    If rng.columns.count <> sourceCol - 1 Then
                        Err.Raise vbObjectError + 2001, _
                                  "HAMU_ConsolidateSheets", _
                                  "'" & ws.name & HAMU_Text("' sayfasının sütun sayısı diğer ") & _
                                  HAMU_Text("sayfalarla aynı değil.")
                    End If

                    If rng.rows.count > 1 Then

                        dataRows = rng.rows.count - 1

                        rng.Offset(1, 0) _
                           .Resize(dataRows, rng.columns.count) _
                           .copy Destination:=outSh.Cells(outRow, 1)

                        outSh.Cells(outRow, sourceCol) _
                             .Resize(dataRows, 1).value = ws.name

                        outRow = outRow + dataRows

                    End If

                    sheetCount = sheetCount + 1

                End If

            End If

        End If

    Next ws

    If firstSheet Then
        MsgBox HAMU_Text("Birleştirilecek veri bulunan çalışma sayfası bulunamadı."), _
               vbInformation, HAMU_APP_NAME
        GoTo CleanExit
    End If

    outSh.columns.AutoFit

    MsgBox sheetCount & HAMU_Text(" çalışma sayfası birleştirildi.") & vbCrLf & _
           HAMU_Text("Sonuç sayfası: ") & outSh.name, _
           vbInformation, HAMU_APP_NAME

CleanExit:

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError "HAMU_ConsolidateSheets", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_AppendSelectedTables()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_AppendSelectedTables") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim msg As String
    Dim nm As Variant
    Dim names As Variant

    Dim outSh As Worksheet
    Dim outRow As Long

    Dim lo As ListObject
    Dim ws As Worksheet

    Dim expectedCols As Long
    Dim foundCount As Long
    Dim missingList As String

    msg = HAMU_InputText( _
            HAMU_Text("Birleştirilecek tablo adlarını virgülle ayırarak girin.") & vbCrLf & _
            HAMU_Text("Örnek: Tablo1, Tablo2, Tablo3"), _
            HAMU_APP_NAME)

    If Len(Trim$(msg)) = 0 Then Exit Sub

    names = Split(msg, ",")

    Set outSh = HAMU_GetOrCreateSheet( _
                    "TABLO_BIRLESIK_" & Format$(Now, "hhmmss"))

    outRow = 1

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each nm In names

        Set lo = Nothing

        For Each ws In ActiveWorkbook.Worksheets

            On Error Resume Next
            Set lo = ws.ListObjects(Trim$(CStr(nm)))
            On Error GoTo ErrH

            If Not lo Is Nothing Then Exit For

        Next ws

        If lo Is Nothing Then

            If Len(missingList) > 0 Then missingList = missingList & ", "
            missingList = missingList & Trim$(CStr(nm))

        Else

            If expectedCols = 0 Then
                expectedCols = lo.range.columns.count
            ElseIf lo.range.columns.count <> expectedCols Then
                Err.Raise vbObjectError + 2002, _
                          "HAMU_AppendSelectedTables", _
                          "'" & lo.name & HAMU_Text("' tablosunun sütun sayısı diğer tablolarla aynı değil.")
            End If

            If outRow = 1 Then

                lo.range.copy Destination:=outSh.Cells(outRow, 1)
                outRow = lo.range.rows.count + 1

            ElseIf Not lo.DataBodyRange Is Nothing Then

                lo.DataBodyRange.copy Destination:=outSh.Cells(outRow, 1)
                outRow = outRow + lo.DataBodyRange.rows.count

            End If

            foundCount = foundCount + 1

        End If

    Next nm

    outSh.columns.AutoFit

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    If foundCount = 0 Then

        MsgBox HAMU_Text("Girilen tablo adlarından hiçbiri bulunamadı."), _
               vbExclamation, HAMU_APP_NAME

    Else

        msg = foundCount & HAMU_Text(" tablo birleştirildi.") & vbCrLf & _
              HAMU_Text("Sonuç sayfası: ") & outSh.name

        If Len(missingList) > 0 Then
            msg = msg & vbCrLf & vbCrLf & _
                  "Bulunamayan tablolar: " & missingList
        End If

        MsgBox msg, vbInformation, HAMU_APP_NAME

    End If

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError "HAMU_AppendSelectedTables", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_SmartAppendByHeaders()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SmartAppendByHeaders") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim outSh As Worksheet
    Dim sheets As Collection

    Dim inputText As String
    Dim names As Variant
    Dim nm As Variant

    Dim headerMap As Object
    Dim weekMap As Object
    Dim headerNames As Collection
    Dim localMap As Object

    Dim rng As range
    Dim arr As Variant
    Dim outArr() As Variant

    Dim totalRows As Long
    Dim outCols As Long
    Dim outRow As Long

    Dim r As Long
    Dim c As Long
    Dim destCol As Long

    Dim headerText As String
    Dim norm As String
    Dim sourceHeader As String

    Set wb = ActiveWorkbook
    Set sheets = New Collection

    inputText = HAMU_InputText( _
                    HAMU_Text("Birleştirilecek sayfaları virgülle girin.") & vbCrLf & _
                    HAMU_Text("Örnek: Ocak, Şubat, Mart") & vbCrLf & vbCrLf & _
                    HAMU_Text("Tüm sayfalar için * yazabilirsiniz."), _
                    HAMU_APP_NAME & HAMU_Text(" - Başlığa Göre Akıllı Birleştir"), _
                    "*")

    If Len(Trim$(inputText)) = 0 Then Exit Sub

    If Trim$(inputText) = "*" Then

        For Each ws In wb.Worksheets

            If WorksheetFunction.CountA(HAMU_BoundedRange(ws.UsedRange, True)) > 0 Then
                sheets.Add ws
            End If

        Next ws

    Else

        names = Split(inputText, ",")

        For Each nm In names

            Set ws = Nothing

            On Error Resume Next
            Set ws = wb.Worksheets(Trim$(CStr(nm)))
            On Error GoTo ErrH

            If ws Is Nothing Then

                Err.Raise _
                    vbObjectError + 2301, _
                    "HAMU_SmartAppendByHeaders", _
                    "'" & Trim$(CStr(nm)) & HAMU_Text("' sayfası bulunamadı.")

            End If

            sheets.Add ws

        Next nm

    End If

    If sheets.count = 0 Then

        MsgBox HAMU_Text("Birleştirilecek sayfa bulunamadı."), _
               vbInformation, HAMU_APP_NAME

        Exit Sub

    End If

    Set headerMap = CreateObject("Scripting.Dictionary")
    Set weekMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = vbTextCompare
    weekMap.CompareMode = vbTextCompare

    Set headerNames = New Collection
    For Each ws In sheets

        Set rng = HAMU_BoundedRange(ws.UsedRange, True)

        If rng.rows.count > 1 Then
            totalRows = totalRows + rng.rows.count - 1
        End If

        Set localMap = CreateObject("Scripting.Dictionary")
        localMap.CompareMode = vbTextCompare

        For c = 1 To rng.columns.count

            headerText = Trim$(CStr(rng.Cells(1, c).value))

            If Len(headerText) = 0 Then
                headerText = "Sütun_" & c
            End If

            norm = HAMU_Data_NormalizeHeaderKey(headerText)

            If localMap.exists(norm) Then

                Err.Raise _
                    vbObjectError + 2302, _
                    "HAMU_SmartAppendByHeaders", _
                    "'" & ws.name & HAMU_Text("' sayfasında aynı başlık birden fazla kez var: ") & _
                    headerText

            End If

            localMap.Add norm, c

            If Not headerMap.exists(norm) Then

                headerMap.Add norm, headerNames.count + 1
                headerNames.Add headerText

            End If

            If HAMU_Data_IsWeekColumn(rng, c) Then
                weekMap(norm) = True
            End If

        Next c

    Next ws

    sourceHeader = "KaynakSayfa"

    If headerMap.exists(HAMU_Data_NormalizeHeaderKey(sourceHeader)) Then
        sourceHeader = "HAMU_KaynakSayfa"
    End If

    outCols = headerNames.count + 1

    If totalRows > ActiveSheet.rows.count - 1 Then

        MsgBox _
            HAMU_Text("Birleştirme sonucu Excel satır sınırını aşıyor."), _
            vbCritical, HAMU_APP_NAME

        Exit Sub

    End If

    ReDim outArr(1 To totalRows, 1 To outCols)

    outRow = 1
    For Each ws In sheets

        Set rng = HAMU_BoundedRange(ws.UsedRange, True)
        arr = rng.Value2

        Set localMap = CreateObject("Scripting.Dictionary")
        localMap.CompareMode = vbTextCompare

        For c = 1 To rng.columns.count

            headerText = Trim$(CStr(rng.Cells(1, c).value))

            If Len(headerText) = 0 Then
                headerText = "Sütun_" & c
            End If

            localMap(HAMU_Data_NormalizeHeaderKey(headerText)) = c

        Next c

        For r = 2 To rng.rows.count

            For c = 1 To headerNames.count

                norm = HAMU_Data_NormalizeHeaderKey(CStr(headerNames(c)))

                If localMap.exists(norm) Then

                    destCol = CLng(localMap(norm))

                    outArr(outRow, c) = _
                        HAMU_Data_SplitOutputValue( _
                            rng.Cells(r, destCol), _
                            rng.Cells(1, destCol).Value2)

                End If

            Next c

            outArr(outRow, outCols) = ws.name
            outRow = outRow + 1

        Next r

    Next ws

    Set outSh = HAMU_GetOrCreateSheet( _
                    "AKILLI_BIRLESIK_" & Format$(Now, "hhmmss"))

    For c = 1 To headerNames.count

        outSh.Cells(1, c).value = headerNames(c)

        norm = HAMU_Data_NormalizeHeaderKey(CStr(headerNames(c)))

        If weekMap.exists(norm) Then
            outSh.columns(c).NumberFormat = "@"
        End If

    Next c

    outSh.Cells(1, outCols).value = sourceHeader

    If totalRows > 0 Then
        outSh.Cells(2, 1).Resize(totalRows, outCols).Value2 = outArr
    End If

    With outSh

        With .range(.Cells(1, 1), .Cells(1, outCols))
            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack
        End With

        If totalRows > 0 Then
            .range(.Cells(1, 1), .Cells(totalRows + 1, outCols)).AutoFilter
        End If

        .columns.AutoFit

    End With

    MsgBox _
        sheets.count & HAMU_Text(" sayfa başlıklarına göre birleştirildi.") & vbCrLf & _
        HAMU_Text("Toplam satır: ") & Format$(totalRows, "#,##0") & vbCrLf & _
        HAMU_Text("Sonuç sayfası: ") & outSh.name, _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_SmartAppendByHeaders", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_CollectFilesData()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_CollectFilesData") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim folderPath As String
    Dim sheetName As String
    Dim fileName As String
    Dim fullPath As String
    Dim ext As String

    Dim sourceOwned As Boolean
    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wbOut As Workbook
    Dim outSh As Worksheet

    Dim rng As range
    Dim arr As Variant
    Dim colArr() As Variant

    Dim headerMap As Object

    Dim outRow As Long
    Dim nextCol As Long
    Dim dataRows As Long

    Dim r As Long
    Dim c As Long
    Dim destCol As Long
    Dim fileCount As Long

    Dim headerText As String
    Dim norm As String

    Set wbOut = ActiveWorkbook

    With Application.FileDialog(msoFileDialogFolderPicker)

        .title = HAMU_Text("Veri toplanacak Excel / CSV dosyalarının klasörünü seçin")

        If .show <> -1 Then Exit Sub

        folderPath = .SelectedItems(1)

    End With

    sheetName = HAMU_InputText( _
                    HAMU_Text("Okunacak sayfa adını yazın.") & vbCrLf & _
                    HAMU_Text("Boş bırakırsanız her dosyanın ilk sayfası kullanılır."), _
                    HAMU_APP_NAME & " - Dosyalardan Veri Topla")

    If gHAMU_InputCancelled Then Exit Sub

    Set outSh = HAMU_GetOrCreateSheet( _
                    "DOSYA_TOPLA_" & Format$(Now, "hhmmss"))

    Set headerMap = CreateObject("Scripting.Dictionary")
    headerMap.CompareMode = vbTextCompare

    outSh.Cells(1, 1).value = "KaynakDosya"

    nextCol = 2
    outRow = 2

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    fileName = Dir$(folderPath & Application.PathSeparator & "*.*")

    Do While Len(fileName) > 0

        fullPath = folderPath & Application.PathSeparator & fileName
        ext = LCase$(Mid$(fileName, InStrRev(fileName, ".") + 1))

        If HAMU_Data_IsSupportedDataFile(ext) Then

            If StrComp(fullPath, wbOut.FullName, vbTextCompare) <> 0 Then

                Set wbSource = HAMU_OpenSourceWorkbook(fullPath, sourceOwned)

                Set wsSource = Nothing

                If Len(Trim$(sheetName)) > 0 Then

                    On Error Resume Next
                    Set wsSource = wbSource.Worksheets(sheetName)
                    On Error GoTo ErrH

                    If wsSource Is Nothing Then

                        If sourceOwned Then wbSource.Close SaveChanges:=False
                        Set wbSource = Nothing

                        GoTo NextFile

                    End If

                Else

                    Set wsSource = wbSource.Worksheets(1)

                End If

                Set rng = HAMU_BoundedRange(wsSource.UsedRange, True)

                If WorksheetFunction.CountA(rng) > 0 And _
                   rng.rows.count >= 1 Then

                    arr = rng.Value2

                    For c = 1 To rng.columns.count

                        headerText = Trim$(CStr(rng.Cells(1, c).value))

                        If Len(headerText) = 0 Then
                            headerText = "Sütun_" & c
                        End If

                        norm = HAMU_Data_NormalizeHeaderKey(headerText)

                        If Not headerMap.exists(norm) Then

                            headerMap.Add norm, nextCol
                            outSh.Cells(1, nextCol).value = headerText

                            If HAMU_Data_IsWeekColumn(rng, c) Then
                                outSh.columns(nextCol).NumberFormat = "@"
                            End If

                            nextCol = nextCol + 1

                        Else

                            If HAMU_Data_IsWeekColumn(rng, c) Then
                                outSh.columns(CLng(headerMap(norm))).NumberFormat = "@"
                            End If

                        End If

                    Next c

                    If rng.rows.count > 1 Then

                        dataRows = rng.rows.count - 1

                        outSh.Cells(outRow, 1) _
                             .Resize(dataRows, 1).value = fileName

                        For c = 1 To rng.columns.count

                            headerText = Trim$(CStr(rng.Cells(1, c).value))

                            If Len(headerText) = 0 Then
                                headerText = "Sütun_" & c
                            End If

                            norm = HAMU_Data_NormalizeHeaderKey(headerText)
                            destCol = CLng(headerMap(norm))

                            ReDim colArr(1 To dataRows, 1 To 1)

                            For r = 2 To rng.rows.count

                                colArr(r - 1, 1) = _
                                    HAMU_Data_SplitOutputValue( _
                                        rng.Cells(r, c), _
                                        rng.Cells(1, c).Value2)

                            Next r

                            outSh.Cells(outRow, destCol) _
                                 .Resize(dataRows, 1).Value2 = colArr

                        Next c

                        outRow = outRow + dataRows

                    End If

                    fileCount = fileCount + 1

                End If

                If sourceOwned Then wbSource.Close SaveChanges:=False
                Set wbSource = Nothing

            End If

        End If

NextFile:

        fileName = Dir$

    Loop

    With outSh

        With .range(.Cells(1, 1), .Cells(1, nextCol - 1))
            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack
        End With

        If outRow > 2 Then
            .range(.Cells(1, 1), .Cells(outRow - 1, nextCol - 1)).AutoFilter
        End If

        .columns.AutoFit

    End With

CleanExit:

    On Error Resume Next

    If Not wbSource Is Nothing Then
        If sourceOwned Then wbSource.Close SaveChanges:=False
    End If

    Application.DisplayAlerts = True
    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    On Error GoTo 0

    MsgBox _
        fileCount & HAMU_Text(" dosyadan veri toplandı.") & vbCrLf & _
        HAMU_Text("Sonuç sayfası: ") & outSh.name, _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    On Error Resume Next

    If Not wbSource Is Nothing Then
        If sourceOwned Then wbSource.Close SaveChanges:=False
    End If

    Application.DisplayAlerts = True
    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    On Error GoTo 0

    HAMU_ShowError "HAMU_CollectFilesData", hamuErrorNumber, hamuErrorText

End Sub


