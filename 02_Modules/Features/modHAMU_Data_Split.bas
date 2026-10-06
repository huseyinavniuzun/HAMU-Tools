Attribute VB_Name = "modHAMU_Data_Split"
Option Explicit
Option Private Module

Public Sub HAMU_SplitDataIntoSheets()
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    HAMU_SplitDataWizard
End Sub

Public Sub HAMU_SplitByValueToFiles()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SplitByValueToFiles") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim ws As Worksheet
    Dim ur As range
    Dim hdr As range

    Dim keyCol As String
    Dim idx As Long

    Dim fldr As String
    Dim dict As Object
    Dim arr As Variant

    Dim r As Long
    Dim k As Variant

    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim outRow As Long

    Dim baseName As String
    Dim savePath As String
    Dim fileCount As Long

    Set ws = ActiveSheet
    Set ur = HAMU_PromptRange(HAMU_Text("Başlık satırı dahil bölünecek veri aralığını seçin."))
    If ur Is Nothing Then Exit Sub
    Set ws = ur.Worksheet
    Set hdr = ur.rows(1)

    keyCol = HAMU_InputText( _
                HAMU_Text("Bölünecek sütunun başlığını veya seçim içindeki sütun numarasını girin:"), _
                HAMU_APP_NAME)

    If Len(Trim$(keyCol)) = 0 Then Exit Sub

    If IsNumeric(keyCol) Then
        idx = CLng(keyCol)
    Else
        idx = HAMU_HeaderToIndex(hdr, keyCol)
    End If

    If idx < 1 Or idx > ur.columns.count Then
        Err.Raise vbObjectError + 2003, _
                  "HAMU_SplitByValueToFiles", _
                  HAMU_Text("Bölme sütunu bulunamadı.")
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)

        .title = HAMU_Text("Dosyaların kaydedileceği klasörü seçin")

        If .show <> -1 Then Exit Sub

        fldr = .SelectedItems(1)

    End With

    arr = ur.Value2

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For r = 2 To UBound(arr, 1)

        If Len(Trim$(CStr(arr(r, idx)))) > 0 Then
            If Not dict.exists(CStr(arr(r, idx))) Then
                dict.Add CStr(arr(r, idx)), True
            End If
        End If

    Next r

    If dict.count = 0 Then
        MsgBox HAMU_Text("Seçilen sütunda dosyaya bölünecek değer bulunamadı."), _
               vbInformation, HAMU_APP_NAME
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    For Each k In dict.keys

        Set wbNew = Workbooks.Add(xlWBATWorksheet)
        Set wsNew = wbNew.Worksheets(1)

        ur.rows(1).copy Destination:=wsNew.range("A1")

        outRow = 2

        For r = 2 To UBound(arr, 1)

            If StrComp( _
                CStr(arr(r, idx)), _
                CStr(k), _
                vbTextCompare) = 0 Then

                ur.rows(r).copy Destination:=wsNew.Cells(outRow, 1)

                outRow = outRow + 1

            End If

        Next r

        wsNew.columns.AutoFit

        baseName = HAMU_TemizDosyaAdi(CStr(k))

        If Len(baseName) = 0 Then baseName = "Bos"

        savePath = HAMU_Data_UniqueFilePath( _
                        fldr, _
                        baseName, _
                        ".xlsx")

        wbNew.SaveAs _
            fileName:=savePath, _
            fileFormat:=xlOpenXMLWorkbook

        wbNew.Close SaveChanges:=False
        Set wbNew = Nothing

        fileCount = fileCount + 1

    Next k

    Application.DisplayAlerts = True
    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox fileCount & HAMU_Text(" dosya oluşturuldu."), _
           vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    On Error Resume Next

    If Not wbNew Is Nothing Then
        wbNew.Close SaveChanges:=False
    End If

    Application.DisplayAlerts = True
    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    On Error GoTo 0

    HAMU_ShowError "HAMU_SplitByValueToFiles", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_SplitURLParameters()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SplitURLParameters") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_PromptRange(HAMU_Text("URL Parametrelerini Sütunlara Ayır: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub

    On Error GoTo ErrH

    Dim src As range
    Dim c As range
    Dim sh As Worksheet

    Dim allKeys As Object
    Dim keyIndex As Object

    Dim urlText As String
    Dim queryPart As String
    Dim pairText As String

    Dim pairs() As String

    Dim qPos As Long
    Dim hashPos As Long
    Dim eqPos As Long

    Dim i As Long
    Dim r As Long
    Dim col As Long

    Dim keyName As String
    Dim ValueText As String
    Dim ky As Variant

    If TypeName(hamuScope) <> "Range" Then
        MsgBox HAMU_Text("Önce URL listesini seçin."), _
               vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    Set src = hamuScope.columns(1)

    Set allKeys = CreateObject("Scripting.Dictionary")
    allKeys.CompareMode = vbTextCompare

    For Each c In src.Cells

        urlText = CStr(c.Value2)

        qPos = InStr(1, urlText, "?", vbBinaryCompare)

        If qPos > 0 Then

            queryPart = Mid$(urlText, qPos + 1)

            hashPos = InStr(1, queryPart, "#", vbBinaryCompare)

            If hashPos > 0 Then
                queryPart = Left$(queryPart, hashPos - 1)
            End If

            If Len(queryPart) > 0 Then

                pairs = Split(queryPart, "&")

                For i = LBound(pairs) To UBound(pairs)

                    eqPos = InStr(1, pairs(i), "=", vbBinaryCompare)

                    If eqPos > 1 Then

                        keyName = LCase$(Left$(pairs(i), eqPos - 1))

                        If Not allKeys.exists(keyName) Then
                            allKeys.Add keyName, True
                        End If

                    End If

                Next i

            End If

        End If

    Next c

    If allKeys.count = 0 Then
        MsgBox HAMU_Text("Seçilen hücrelerde URL sorgu parametresi bulunamadı."), _
               vbInformation, HAMU_APP_NAME
        Exit Sub
    End If

    Set sh = HAMU_GetOrCreateSheet( _
                "URLPAR_" & Format$(Now, "hhmmss"))

    Set keyIndex = CreateObject("Scripting.Dictionary")
    keyIndex.CompareMode = vbTextCompare

    sh.Cells(1, 1).value = HAMU_Text("Kaynak")

    col = 2

    For Each ky In allKeys.keys

        sh.Cells(1, col).value = ky
        keyIndex.Add CStr(ky), col

        col = col + 1

    Next ky

    r = 2

    For Each c In src.Cells

        urlText = CStr(c.Value2)

        sh.Cells(r, 1).value = urlText

        qPos = InStr(1, urlText, "?", vbBinaryCompare)

        If qPos > 0 Then

            queryPart = Mid$(urlText, qPos + 1)

            hashPos = InStr(1, queryPart, "#", vbBinaryCompare)

            If hashPos > 0 Then
                queryPart = Left$(queryPart, hashPos - 1)
            End If

            If Len(queryPart) > 0 Then

                pairs = Split(queryPart, "&")

                For i = LBound(pairs) To UBound(pairs)

                    pairText = pairs(i)
                    eqPos = InStr(1, pairText, "=", vbBinaryCompare)

                    If eqPos > 1 Then

                        keyName = LCase$(Left$(pairText, eqPos - 1))
                        ValueText = Mid$(pairText, eqPos + 1)

                        If keyIndex.exists(keyName) Then
                            sh.Cells(r, CLng(keyIndex(keyName))).value = ValueText
                        End If

                    End If

                Next i

            End If

        End If

        r = r + 1

    Next c

    sh.rows(1).Font.Bold = True
    sh.columns.AutoFit

    MsgBox allKeys.count & HAMU_Text(" farklı URL parametresi ayrıştırıldı.") & vbCrLf & _
           HAMU_Text("Sonuç sayfası: ") & sh.name, _
           vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_SplitURLParameters", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_PivotExport()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_PivotExport") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim pt As PivotTable
    Dim pf As PivotField
    Dim fldName As String

    Dim it As PivotItem
    Dim sh As Worksheet

    Dim wb As Workbook
    Dim safeName As String
    Dim oldPage As String

    Set wb = ActiveWorkbook

    On Error Resume Next
    Dim pivotScope As range
    Set pivotScope = HAMU_PromptRange(HAMU_Text("Dışa aktarılacak PivotTable içinden hücre seçin."))
    If pivotScope Is Nothing Then Exit Sub
    Set pt = pivotScope.Cells(1, 1).PivotTable
    On Error GoTo ErrH

    If pt Is Nothing Then
        Err.Raise vbObjectError + 2004, _
                  "HAMU_PivotExport", _
                  HAMU_Text("Lütfen bir PivotTable hücresi seçin.")
    End If

    fldName = HAMU_InputText( _
                HAMU_Text("Rapor filtresi alan adını girin:"), _
                HAMU_APP_NAME)

    If Len(Trim$(fldName)) = 0 Then Exit Sub

    Set pf = pt.PivotFields(fldName)

    If pf.orientation <> xlPageField Then
        Err.Raise vbObjectError + 2005, _
                  "HAMU_PivotExport", _
                  HAMU_Text("Seçilen alan bir rapor filtresi değil.")
    End If

    On Error Resume Next
    oldPage = pf.CurrentPage.name
    pf.EnableMultiplePageItems = False
    On Error GoTo ErrH

    Application.ScreenUpdating = False

    For Each it In pf.PivotItems

        pf.CurrentPage = it.name

        safeName = Left$( _
                    "PX_" & HAMU_TemizDosyaAdi(CStr(it.name)), _
                    31)

        safeName = HAMU_Data_UniqueSheetName(wb, safeName)

        Set sh = wb.Worksheets.Add( _
                    After:=wb.Worksheets(wb.Worksheets.count))

        sh.name = safeName

        pt.TableRange2.copy

        sh.range("A1").PasteSpecial xlPasteAll

        sh.columns.AutoFit

    Next it

    If Len(oldPage) > 0 Then
        On Error Resume Next
        pf.CurrentPage = oldPage
        On Error GoTo ErrH
    End If

    Application.CutCopyMode = False
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox HAMU_Text("Pivot filtreleri ayrı sayfalara aktarıldı."), _
           vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.CutCopyMode = False
    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError "HAMU_PivotExport", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_SplitDataWizard()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SplitDataWizard") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim wbSource As Workbook

    Dim srcArr As Variant
    Dim dictRows As Object
    Dim dictLabels As Object
    Dim keyOrder As Collection
    Dim rowList As Collection

    Dim methodAnswer As Variant
    Dim outputAnswer As Variant
    Dim answer As Variant

    Dim splitMethod As Long
    Dim outputMode As Long

    Dim dataRowCount As Long
    Dim chunkSize As Long
    Dim partCount As Long

    Dim columnInput As String
    Dim dateColumnInput As String

    Dim splitCols As Collection
    Dim splitCol As Long
    Dim dateCol As Long
    Dim dateMode As Long

    Dim folderPath As String

    Dim r As Long
    Dim i As Long
    Dim partNo As Long

    Dim key As String
    Dim label As String

    Dim v As Variant
    Dim dt As Date

    Dim k As Variant
    Dim createdCount As Long

    Set rng = HAMU_PromptRange( _
        HAMU_Text("Bölmek istediğiniz veri alanını seçin.") & vbCrLf & _
        HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    If rng.areas.count > 1 Then

        MsgBox _
            HAMU_Text("Veriyi Böl aracı için tek parça bir alan seçin."), _
            vbExclamation, _
            HAMU_APP_NAME

        Exit Sub

    End If

    If rng.rows.count < 2 Then

        MsgBox _
            HAMU_Text("Seçimde en az bir başlık ve bir veri satırı bulunmalıdır."), _
            vbExclamation, _
            HAMU_APP_NAME

        Exit Sub

    End If

    Set wbSource = rng.Worksheet.parent

    srcArr = rng.Value2
    dataRowCount = UBound(srcArr, 1) - 1

    methodAnswer = HAMU_InputValue( _
        Prompt:= _
            HAMU_Text("Veri hangi yönteme göre bölünsün?") & vbCrLf & vbCrLf & _
            HAMU_Text("1 - Satır sayısına göre") & vbCrLf & _
            HAMU_Text("2 - Eşit parçalara göre") & vbCrLf & _
            HAMU_Text("3 - Tek sütun değerine göre") & vbCrLf & _
            HAMU_Text("4 - Birden fazla sütuna göre") & vbCrLf & _
            HAMU_Text("5 - Tarihe göre"), _
        title:=HAMU_APP_NAME & HAMU_Text(" - Veriyi Böl"), _
        Default:=1, _
        ValueType:=1)

    If VarType(methodAnswer) = vbBoolean Then
        If methodAnswer = False Then Exit Sub
    End If

    splitMethod = CLng(methodAnswer)

    If splitMethod < 1 Or splitMethod > 5 Then

        MsgBox _
            HAMU_Text("1 ile 5 arasında bir yöntem seçin."), _
            vbExclamation, _
            HAMU_APP_NAME

        Exit Sub

    End If

    outputAnswer = HAMU_InputValue( _
        Prompt:= _
            HAMU_Text("Sonuç nasıl oluşturulsun?") & vbCrLf & vbCrLf & _
            HAMU_Text("1 - Aynı dosyada yeni sayfalar") & vbCrLf & _
            HAMU_Text("2 - Ayrı Excel dosyaları"), _
        title:=HAMU_APP_NAME & HAMU_Text(" - Çıktı"), _
        Default:=1, _
        ValueType:=1)

    If VarType(outputAnswer) = vbBoolean Then
        If outputAnswer = False Then Exit Sub
    End If

    outputMode = CLng(outputAnswer)

    If outputMode <> 1 And outputMode <> 2 Then

        MsgBox _
            HAMU_Text("1 veya 2 seçin."), _
            vbExclamation, _
            HAMU_APP_NAME

        Exit Sub

    End If

    If outputMode = 2 Then

        With Application.FileDialog(msoFileDialogFolderPicker)

            .title = HAMU_Text("Oluşturulacak dosyaların kaydedileceği klasörü seçin")

            If .show <> -1 Then Exit Sub

            folderPath = .SelectedItems(1)

        End With

    End If

    Set dictRows = CreateObject("Scripting.Dictionary")
    Set dictLabels = CreateObject("Scripting.Dictionary")
    Set keyOrder = New Collection

    dictRows.CompareMode = vbTextCompare
    dictLabels.CompareMode = vbTextCompare

    If splitMethod = 1 Then

        answer = HAMU_InputValue( _
            HAMU_Text("Her parçada kaç veri satırı olsun?"), _
            HAMU_APP_NAME, _
            5000, _
            ValueType:=1)

        If VarType(answer) = vbBoolean Then
            If answer = False Then Exit Sub
        End If

        If CDbl(answer) < 1 Or _
           CDbl(answer) <> Fix(CDbl(answer)) Then

            MsgBox _
                HAMU_Text("Satır sayısı pozitif bir tam sayı olmalıdır."), _
                vbExclamation, _
                HAMU_APP_NAME

            Exit Sub

        End If

        chunkSize = CLng(answer)

        For r = 2 To UBound(srcArr, 1)

            partNo = ((r - 2) \ chunkSize) + 1

            label = "PARCA_" & Format$(partNo, "000")
            key = label

            HAMU_Data_AddSplitRow _
                dictRows, _
                dictLabels, _
                keyOrder, _
                key, _
                label, _
                r

        Next r

    End If

    If splitMethod = 2 Then

        answer = HAMU_InputValue( _
            HAMU_Text("Veri kaç eşit parçaya ayrılsın?"), _
            HAMU_APP_NAME, _
            4, _
            ValueType:=1)

        If VarType(answer) = vbBoolean Then
            If answer = False Then Exit Sub
        End If

        If CDbl(answer) < 1 Or _
           CDbl(answer) <> Fix(CDbl(answer)) Then

            MsgBox _
                HAMU_Text("Parça sayısı pozitif bir tam sayı olmalıdır."), _
                vbExclamation, _
                HAMU_APP_NAME

            Exit Sub

        End If

        partCount = CLng(answer)

        If partCount > dataRowCount Then
            partCount = dataRowCount
        End If

        chunkSize = _
            Application.WorksheetFunction.RoundUp( _
                dataRowCount / partCount, _
                0)

        For r = 2 To UBound(srcArr, 1)

            partNo = ((r - 2) \ chunkSize) + 1

            label = "PARCA_" & Format$(partNo, "000")
            key = label

            HAMU_Data_AddSplitRow _
                dictRows, _
                dictLabels, _
                keyOrder, _
                key, _
                label, _
                r

        Next r

    End If

    If splitMethod = 3 Then

        columnInput = HAMU_InputText( _
            HAMU_Text("Bölme yapılacak sütunun numarasını veya başlığını girin.") & _
            vbCrLf & vbCrLf & _
            HAMU_Text("Örnek:") & vbCrLf & _
            "3" & vbCrLf & _
            "veya" & vbCrLf & _
            HAMU_Text("Şehir"), _
            HAMU_APP_NAME)

        If Len(Trim$(columnInput)) = 0 Then Exit Sub

        splitCol = _
            HAMU_Data_ResolveColumn( _
                rng, _
                columnInput)

        For r = 2 To UBound(srcArr, 1)

            label = HAMU_Data_SplitCellText( _
                        rng.Cells(r, splitCol), _
                        rng.Cells(1, splitCol).Value2)

            If Len(Trim$(label)) = 0 Then
                label = "BOS"
            End If

            key = CStr(Len(label)) & ":" & label

            HAMU_Data_AddSplitRow _
                dictRows, _
                dictLabels, _
                keyOrder, _
                key, _
                label, _
                r

        Next r

    End If

    If splitMethod = 4 Then

        columnInput = HAMU_InputText( _
            HAMU_Text("Bölme yapılacak sütunları virgülle ayırarak girin.") & _
            vbCrLf & vbCrLf & _
            HAMU_Text("Örnek:") & vbCrLf & _
            "1,2,3" & vbCrLf & _
            "veya" & vbCrLf & _
            HAMU_Text("Account, Şehir, Ürün"), _
            HAMU_APP_NAME)

        If Len(Trim$(columnInput)) = 0 Then Exit Sub

        Set splitCols = _
            HAMU_Data_ParseColumnList( _
                rng, _
                columnInput)

        For r = 2 To UBound(srcArr, 1)

            key = ""
            label = ""

            For i = 1 To splitCols.count

                Dim partText As String

                splitCol = CLng(splitCols(i))

                partText = HAMU_Data_SplitCellText( _
                                rng.Cells(r, splitCol), _
                                rng.Cells(1, splitCol).Value2)

                If Len(Trim$(partText)) = 0 Then
                    partText = "BOS"
                End If

                key = key & _
                      Len(partText) & ":" & _
                      partText & "|"

                If Len(label) > 0 Then
                    label = label & "_"
                End If

                label = label & partText

            Next i

            HAMU_Data_AddSplitRow _
                dictRows, _
                dictLabels, _
                keyOrder, _
                key, _
                label, _
                r

        Next r

    End If

    If splitMethod = 5 Then

        dateColumnInput = HAMU_InputText( _
            HAMU_Text("Tarih sütununun numarasını veya başlığını girin.") & _
            vbCrLf & vbCrLf & _
            HAMU_Text("Örnek:") & vbCrLf & _
            "1" & vbCrLf & _
            "veya" & vbCrLf & _
            "Tarih", _
            HAMU_APP_NAME)

        If Len(Trim$(dateColumnInput)) = 0 Then Exit Sub

        dateCol = _
            HAMU_Data_ResolveColumn( _
                rng, _
                dateColumnInput)

        answer = HAMU_InputValue( _
            Prompt:= _
                HAMU_Text("Tarih hangi seviyeye göre bölünsün?") & vbCrLf & vbCrLf & _
                HAMU_Text("1 - Yıl") & vbCrLf & _
                HAMU_Text("2 - Çeyrek") & vbCrLf & _
                "3 - Ay" & vbCrLf & _
                "4 - Hafta", _
            title:=HAMU_APP_NAME, _
            Default:=3, _
            ValueType:=1)

        If VarType(answer) = vbBoolean Then
            If answer = False Then Exit Sub
        End If

        dateMode = CLng(answer)

        If dateMode < 1 Or dateMode > 4 Then

            MsgBox _
                HAMU_Text("1 ile 4 arasında bir tarih yöntemi seçin."), _
                vbExclamation, _
                HAMU_APP_NAME

            Exit Sub

        End If

        For r = 2 To UBound(srcArr, 1)

            v = srcArr(r, dateCol)

            If IsDate(v) Then

                dt = CDate(v)

                Select Case dateMode

                    Case 1

                        label = _
                            Format$(year(dt), "0000")

                    Case 2

                        label = _
                            Format$(year(dt), "0000") & _
                            "_Q" & _
                            DatePart("q", dt)

                    Case 3

                        label = _
                            Format$(dt, "yyyy-mm")

                    Case 4

                        label = _
                            HAMU_Data_WeekLabel(dt)

                End Select

            Else

                label = "TARIH_YOK"

            End If

            key = label

            HAMU_Data_AddSplitRow _
                dictRows, _
                dictLabels, _
                keyOrder, _
                key, _
                label, _
                r

        Next r

    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    For Each k In keyOrder

        Set rowList = dictRows(CStr(k))
        label = CStr(dictLabels(CStr(k)))

        If outputMode = 1 Then

            HAMU_Data_CreateSplitSheet _
                wbSource, _
                rng, _
                srcArr, _
                rowList, _
                label

        Else

            HAMU_Data_CreateSplitFile _
                rng, _
                srcArr, _
                rowList, _
                label, _
                folderPath

        End If

        createdCount = createdCount + 1

    Next k

CleanExit:

    Application.DisplayAlerts = True
    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    If createdCount > 0 Then

        MsgBox _
            HAMU_Text("Veri bölme işlemi tamamlandı.") & vbCrLf & vbCrLf & _
            HAMU_Text("Oluşturulan parça: ") & createdCount & vbCrLf & _
            IIf( _
                outputMode = 1, _
                HAMU_Text("Çıktı: Yeni çalışma sayfaları"), _
                HAMU_Text("Çıktı: Ayrı Excel dosyaları")), _
            vbInformation, _
            HAMU_APP_NAME

    End If

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    Application.DisplayAlerts = True
    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    HAMU_ShowError _
        "HAMU_SplitDataWizard", _
        Err.number, _
        Err.description

End Sub


