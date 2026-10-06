Attribute VB_Name = "modHAMU_Data_Select"
Option Explicit
Option Private Module

Public Sub HAMU_CompareTwoLists()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_CompareTwoLists") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng1 As range
    Dim rng2 As range

    Dim outSh As Worksheet
    Dim dict As Object

    Dim c As range
    Dim r As Long

    Dim v As String
    Dim k As Variant

    Set rng1 = HAMU_PromptRange( _
                HAMU_Text("1. listeyi seçin (tek sütun)."))

    If rng1 Is Nothing Then Exit Sub

    Set rng2 = HAMU_PromptRange( _
                HAMU_Text("2. listeyi seçin (tek sütun)."))

    If rng2 Is Nothing Then Exit Sub

    If rng1.columns.count <> 1 Or rng2.columns.count <> 1 Then
        MsgBox HAMU_Text("Her iki seçim de tek sütun olmalıdır."), _
               vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For Each c In rng1.Cells

        v = Trim$(CStr(c.Value2))

        If Len(v) > 0 Then
            dict(v) = "Sadece Liste 1"
        End If

    Next c

    For Each c In rng2.Cells

        v = Trim$(CStr(c.Value2))

        If Len(v) > 0 Then

            If dict.exists(v) Then
                dict(v) = HAMU_Text("İki Listede")
            Else
                dict.Add v, "Sadece Liste 2"
            End If

        End If

    Next c

    Set outSh = HAMU_GetOrCreateSheet( _
                    "LISTE_KARSILASTIR_" & Format$(Now, "hhmmss"))

    outSh.range("A1:B1").value = Array(HAMU_Text("Değer"), "Durum")

    r = 2

    For Each k In dict.keys

        outSh.Cells(r, 1).value = k
        outSh.Cells(r, 2).value = dict(k)

        r = r + 1

    Next k

    outSh.rows(1).Font.Bold = True
    outSh.columns.AutoFit

    MsgBox dict.count & HAMU_Text(" benzersiz değer karşılaştırıldı.") & vbCrLf & _
           HAMU_Text("Sonuç sayfası: ") & outSh.name, _
           vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_CompareTwoLists", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_ExtractUniqueList()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ExtractUniqueList") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim source As range, target As range, choice As Long, count As Long
    Set source = HAMU_PromptRange(HAMU_Text("Benzersiz değerlerin alınacağı alanı seçin."))
    If source Is Nothing Then Exit Sub
    Dim values As Variant
    values = HAMU_UniqueValues(source, count)
    If count = 0 Then HAMU_ShowInfo HAMU_Text("Seçimde listelenecek değer yok."): Exit Sub
    choice = CLng(HAMU_Setting("UniqueOutput"))
    If choice = 0 Then choice = HAMU_ChooseIndex(HAMU_Text("HAMU | Benzersiz Liste"), HAMU_Text("Liste nereye yazılsın? Mevcut alanda bir başlangıç hücresi seçebilir veya yeni bir sayfa oluşturabilirsiniz. Başlıkla birlikte ") & count + 1 & HAMU_Text(" satır gerekir."), Array(HAMU_Text("Seçtiğim alana / sayfaya yaz"), HAMU_Text("Yeni sayfa oluştur")))
    If choice = 0 Then Exit Sub
    If choice = 1 Then
        Set target = HAMU_PromptRange(HAMU_Text("Listenin başlangıç hücresini seçin. Başka bir sayfaya geçebilirsiniz."))
        If target Is Nothing Then Exit Sub
    Else
        Dim sheet As Worksheet
        Set sheet = source.parent.parent.Worksheets.Add(After:=source.parent.parent.sheets(source.parent.parent.sheets.count))
        sheet.name = HAMU_Data_UniqueSheetName(source.parent.parent, "BENZERSIZ")
        Set target = sheet.range("A1")
    End If
    HAMU_WriteUniqueList values, count, target
    target.parent.Activate: target.Cells(1, 1).Select

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_ExtractUniqueList", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_SelectColumnsToNewTable()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SelectColumnsToNewTable") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim cols As Collection
    Dim inputText As String

    Dim outSh As Worksheet
    Dim outArr() As Variant

    Dim r As Long
    Dim i As Long
    Dim idx As Long

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Sütun seçmek istediğiniz veri alanını seçin.") & vbCrLf & _
                HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub
    Set rng = HAMU_BoundedRange(rng)
    If rng.rows.count < 2 Then
        HAMU_ShowInfo HAMU_L("Başlık ve en az bir veri satırı seçin.", "Select a header and at least one data row.")
        Exit Sub
    End If

    inputText = HAMU_InputText(HAMU_L("Sütun adlarını veya seçtiğiniz alanın içindeki sütun numaralarını virgülle ayırarak yazın. Örnek: Tarih, Ürün veya 1,2,3. Numaralar seçimin soldan sağa sırasıdır; tüm sayfanın sütun numarası değildir.", "Enter column headers or column numbers within the selection, separated by commas. Example: Date, Product or 1,2,3. Numbers refer to the selected area, from left" & _
        " to right."), HAMU_APP_NAME)

    If Len(Trim$(inputText)) = 0 Then Exit Sub

    Set cols = HAMU_Data_ParseColumnList(rng, inputText)

    ReDim outArr(1 To rng.rows.count - 1, _
                 1 To cols.count)

    For r = 2 To rng.rows.count

        For i = 1 To cols.count

            idx = CLng(cols(i))

            outArr(r - 1, i) = _
                HAMU_Data_SplitOutputValue( _
                    rng.Cells(r, idx), _
                    rng.Cells(1, idx).Value2)

        Next i

    Next r

    Set outSh = HAMU_GetOrCreateSheet( _
                    "SUTUN_KOPYALA_" & Format$(Now, "hhmmss"))

    For i = 1 To cols.count

        idx = CLng(cols(i))

        outSh.Cells(1, i).value = rng.Cells(1, idx).value

        If HAMU_Data_IsWeekColumn(rng, idx) Then
            outSh.columns(i).NumberFormat = "@"
        End If

    Next i

    If rng.rows.count > 1 Then

        outSh.Cells(2, 1) _
             .Resize(rng.rows.count - 1, cols.count) _
             .Value2 = outArr

    End If

    With outSh

        With .range(.Cells(1, 1), .Cells(1, cols.count))
            .Font.Bold = True
            .Interior.color = RGB(230, 230, 230)
            .Font.color = vbBlack
        End With

        .range(.Cells(1, 1), _
               .Cells(rng.rows.count, cols.count)).AutoFilter

        .columns.AutoFit

    End With

    Dim resultTable As ListObject
    Set resultTable = outSh.ListObjects.Add(xlSrcRange, outSh.range("A1").Resize(rng.rows.count, cols.count), , xlYes)
    resultTable.TableStyle = "TableStyleMedium1"
    HAMU_ShowInfo _
        cols.count & HAMU_Text(" sütun yeni tabloya aktarıldı.") & vbCrLf & _
        HAMU_Text("Sonuç sayfası: ") & outSh.name

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_SelectColumnsToNewTable", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_FilterRowsToNewTable()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_FilterRowsToNewTable") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim conditions As Collection
    Dim rowsOut As Collection

    Dim inputText As String
    Dim r As Long

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Filtrelemek istediğiniz veri alanını seçin.") & vbCrLf & _
                HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    inputText = HAMU_InputText( _
                    HAMU_Text("Koşulları noktalı virgülle ayırın.") & vbCrLf & vbCrLf & _
                    HAMU_Text("Örnek:") & vbCrLf & _
                    HAMU_Text("Şehir=Bursa;Satış>0;Ürün~TV") & vbCrLf & vbCrLf & _
                    HAMU_Text("Operatörler: =  <>  >  >=  <  <=  ~"), _
                    HAMU_APP_NAME & HAMU_Text(" - Satır Filtrele"))

    If Len(Trim$(inputText)) = 0 Then Exit Sub

    Set conditions = HAMU_Data_ParseFilterConditions( _
                        rng, _
                        inputText)

    Set rowsOut = New Collection

    For r = 2 To rng.rows.count

        If HAMU_Data_RowMatchesConditions( _
                rng, _
                r, _
                conditions) Then

            rowsOut.Add r

        End If

    Next r

    HAMU_Data_WriteRowsToNewSheet _
        rng, _
        rowsOut, _
        "FILTRE_" & Format$(Now, "hhmmss")

    MsgBox _
        Format$(rowsOut.count, "#,##0") & _
        HAMU_Text(" satır koşullara uydu."), _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_FilterRowsToNewTable", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_SampleData()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SampleData") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim rowsOut As Collection

    Dim answer As Variant
    Dim mode As Long
    Dim sampleSize As Long

    Dim groupInput As String
    Dim groupCols As Collection

    Dim dict As Object
    Dim keyOrder As Collection
    Dim rowsForKey As Collection

    Dim indices() As Long
    Dim r As Long
    Dim i As Long
    Dim takeCount As Long

    Dim key As String
    Dim k As Variant

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Örnek alınacak veri alanını seçin.") & vbCrLf & _
                HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    answer = HAMU_InputValue( _
                Prompt:= _
                    HAMU_Text("Örnekleme yöntemini seçin:") & vbCrLf & vbCrLf & _
                    HAMU_Text("1 - Tüm veriden rastgele N satır") & vbCrLf & _
                    HAMU_Text("2 - Her gruptan rastgele N satır"), _
                title:=HAMU_APP_NAME & HAMU_Text(" - Veri Örnekle"), _
                Default:=1, _
                ValueType:=1)

    If VarType(answer) = vbBoolean Then
        If answer = False Then Exit Sub
    End If

    mode = CLng(answer)

    If mode <> 1 And mode <> 2 Then
        MsgBox HAMU_Text("1 veya 2 seçin."), vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    answer = HAMU_InputValue( _
                Prompt:=IIf(mode = 1, _
                            HAMU_Text("Kaç rastgele satır alınsın?"), _
                            HAMU_Text("Her gruptan en fazla kaç satır alınsın?")), _
                title:=HAMU_APP_NAME & HAMU_Text(" - Veri Örnekle"), _
                Default:=100, _
                ValueType:=1)

    If VarType(answer) = vbBoolean Then
        If answer = False Then Exit Sub
    End If

    sampleSize = CLng(answer)

    If sampleSize < 1 Then
        MsgBox HAMU_Text("Örnek sayısı en az 1 olmalıdır."), vbExclamation, HAMU_APP_NAME
        Exit Sub
    End If

    Set rowsOut = New Collection

    Randomize

    If mode = 1 Then

        ReDim indices(1 To rng.rows.count - 1)

        For i = 1 To UBound(indices)
            indices(i) = i + 1
        Next i

        HAMU_Data_ShuffleLongArray indices

        takeCount = Application.Min(sampleSize, UBound(indices))

        For i = 1 To takeCount
            rowsOut.Add indices(i)
        Next i

    Else

        groupInput = HAMU_InputText( _
                        HAMU_Text("Gruplama sütunlarını virgülle girin.") & vbCrLf & _
                        HAMU_Text("Örnek: Account, Şehir"), _
                        HAMU_APP_NAME & HAMU_Text(" - Veri Örnekle"))

        If Len(Trim$(groupInput)) = 0 Then Exit Sub

        Set groupCols = HAMU_Data_ParseColumnList(rng, groupInput)

        Set dict = CreateObject("Scripting.Dictionary")
        Set keyOrder = New Collection

        dict.CompareMode = vbBinaryCompare

        For r = 2 To rng.rows.count

            key = HAMU_Data_BuildKeyFromRangeRow(rng, r, groupCols)

            If Not dict.exists(key) Then

                Set rowsForKey = New Collection

                dict.Add key, rowsForKey
                keyOrder.Add key

            Else

                Set rowsForKey = dict(key)

            End If

            rowsForKey.Add r

        Next r

        For Each k In keyOrder

            Set rowsForKey = dict(CStr(k))

            ReDim indices(1 To rowsForKey.count)

            For i = 1 To rowsForKey.count
                indices(i) = CLng(rowsForKey(i))
            Next i

            HAMU_Data_ShuffleLongArray indices

            takeCount = Application.Min(sampleSize, UBound(indices))

            For i = 1 To takeCount
                rowsOut.Add indices(i)
            Next i

        Next k

    End If

    HAMU_Data_WriteRowsToNewSheet _
        rng, _
        rowsOut, _
        "ORNEK_" & Format$(Now, "hhmmss")

    MsgBox _
        Format$(rowsOut.count, "#,##0") & _
        HAMU_Text(" örnek satır oluşturuldu."), _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_SampleData", hamuErrorNumber, hamuErrorText

End Sub

Public Sub HAMU_ShuffleRows()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ShuffleRows") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrH

    Dim rng As range
    Dim rowsOut As Collection
    Dim indices() As Long
    Dim i As Long

    Set rng = HAMU_PromptRange( _
                HAMU_Text("Satırlarını rastgele karıştırmak istediğiniz veri alanını seçin.") & _
                vbCrLf & HAMU_Text("İlk satır başlık olmalıdır."))

    If rng Is Nothing Then Exit Sub

    If rng.rows.count < 2 Then Exit Sub

    ReDim indices(1 To rng.rows.count - 1)

    For i = 1 To UBound(indices)
        indices(i) = i + 1
    Next i

    Randomize
    HAMU_Data_ShuffleLongArray indices

    Set rowsOut = New Collection

    For i = 1 To UBound(indices)
        rowsOut.Add indices(i)
    Next i

    HAMU_Data_WriteRowsToNewSheet _
        rng, _
        rowsOut, _
        "KARISIK_" & Format$(Now, "hhmmss")

    MsgBox _
        HAMU_Text("Satırlar rastgele karıştırıldı."), _
        vbInformation, HAMU_APP_NAME

    Exit Sub

ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description

    HAMU_ShowError "HAMU_ShuffleRows", hamuErrorNumber, hamuErrorText

End Sub

Public Function HAMU_UniqueValues(ByVal source As range, ByRef count As Long) As Variant
    Dim dict As Object, cell As range, key As Variant, values As Variant, index As Long
    Set dict = CreateObject("Scripting.Dictionary"): dict.CompareMode = vbTextCompare
    For Each cell In source.Cells
        If Not IsError(cell.Value2) Then
            key = Trim$(CStr(cell.Value2))
            If Len(key) > 0 Then
                If Not dict.exists(key) Then dict.Add key, cell.Value2
            End If
        End If
    Next
    count = dict.count
    If count = 0 Then Exit Function
    ReDim values(1 To count + 1, 1 To 1)
    values(1, 1) = HAMU_Text("Değer")
    For Each key In dict.keys
        index = index + 1: values(index + 1, 1) = dict(key)
    Next
    HAMU_UniqueValues = values
End Function
Public Sub HAMU_WriteUniqueList(ByVal values As Variant, ByVal count As Long, ByVal target As range)
    Set target = target.Cells(1, 1)
    If count + 1 > target.parent.rows.count - target.row + 1 Then Err.Raise 5, , HAMU_Text("Liste hedef sayfanın satır sınırına sığmıyor.")
    Dim output As range
    Set output = target.Resize(count + 1, 1)
    If target.parent.ProtectContents Then Err.Raise 5, , HAMU_Text("Hedef sayfa korumalı; başka bir alan seçin.")
    Dim cell As range
    For Each cell In output.Cells
    If cell.MergeCells Then Err.Raise 5, , HAMU_Text("Hedefte birleşik hücreler var; boş bir sütun seçin.")
    Next
    If Application.CountA(output) > 0 Then
        If Not HAMU_ConfirmOverwrite() Then Exit Sub
    End If
    output.Value2 = values: target.Font.Bold = True
    output.columns.AutoFit
End Sub


