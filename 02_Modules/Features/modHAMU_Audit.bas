Attribute VB_Name = "modHAMU_Audit"
Option Explicit
Option Private Module

Public Sub HAMU_CompareSheets()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_CompareSheets") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    On Error GoTo ErrH
    Dim s1 As String, s2 As String, keyName As String
    Dim ws1 As Worksheet, ws2 As Worksheet, r1 As range, r2 As range, h1 As range, h2 As range
    Dim idx1 As Long, idx2 As Long, d1 As Object, d2 As Object, a1, a2, i As Long, k As Variant, rep As Worksheet, row As Long
    s1 = HAMU_InputText(HAMU_Text("1. sayfa adı:"), HAMU_APP_NAME)
    If Len(s1) = 0 Then Exit Sub
    s2 = HAMU_InputText(HAMU_Text("2. sayfa adı:"), HAMU_APP_NAME)
    If Len(s2) = 0 Then Exit Sub
    keyName = HAMU_InputText(HAMU_Text("Anahtar sütun başlığı veya numarası:"), HAMU_APP_NAME, "1")
    Set ws1 = ActiveWorkbook.Worksheets(s1)
    Set ws2 = ActiveWorkbook.Worksheets(s2)
    Set r1 = HAMU_BoundedRange(ws1.UsedRange, True): Set r2 = HAMU_BoundedRange(ws2.UsedRange, True)
    Set h1 = r1.rows(1): Set h2 = r2.rows(1)
    If IsNumeric(keyName) Then idx1 = CLng(keyName): idx2 = idx1 Else idx1 = HAMU_HeaderToIndex(h1, keyName): idx2 = HAMU_HeaderToIndex(h2, keyName)
    Set d1 = CreateObject("Scripting.Dictionary"): Set d2 = CreateObject("Scripting.Dictionary")
    a1 = r1.Value2: a2 = r2.Value2
    For i = 2 To UBound(a1, 1)
    k = CStr(a1(i, idx1))
    If Len(k) > 0 Then d1(k) = Join(Application.index(a1, i, 0), "|")
    Next i
    For i = 2 To UBound(a2, 1)
    k = CStr(a2(i, idx2))
    If Len(k) > 0 Then d2(k) = Join(Application.index(a2, i, 0), "|")
    Next i
    Set rep = HAMU_GetOrCreateSheet("DELTA_RAPOR")
    rep.range("A1:C1").value = Array("Durum", "Anahtar", "Detay")
    row = 2
    For Each k In d1.keys
        If Not d2.exists(k) Then
            rep.Cells(row, 1).value = HAMU_Text("SİLİNEN"): rep.Cells(row, 2).value = k: rep.Cells(row, 3).value = d1(k): row = row + 1
        ElseIf d1(k) <> d2(k) Then
            rep.Cells(row, 1).value = HAMU_Text("GÜNCELLENEN"): rep.Cells(row, 2).value = k: rep.Cells(row, 3).value = HAMU_Text("Satır farklı"): row = row + 1
        End If
    Next k
    For Each k In d2.keys
        If Not d1.exists(k) Then rep.Cells(row, 1).value = "EKLENEN": rep.Cells(row, 2).value = k: rep.Cells(row, 3).value = d2(k): row = row + 1
    Next k
    rep.columns.AutoFit
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_CompareSheets", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_MarkProblemCells()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_MarkProblemCells") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Sorunlu Hücreleri İşaretle: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim c As range, n As Long, f As Long
    For Each c In hamuScope.Cells
        If Not c.HasFormula And VarType(c.Value2) = vbString And IsNumeric(c.Value2) Then
            HAMU_HighlightCell c, RGB(255, 242, 0)
            n = n + 1
        ElseIf Not c.HasFormula And VarType(c.Value2) = vbString And Left$(CStr(c.Value2), 1) = "=" Then
            HAMU_HighlightCell c, RGB(255, 199, 44)
            f = f + 1
        End If
    Next c
    HAMU_ShowInfo HAMU_Text("Sayı gibi metin: ") & n & HAMU_Text(" | Formül gibi metin: ") & f
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_MarkProblemCells", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_HighlightByCriteria()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_HighlightByCriteria") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Hücre Türüne Göre Vurgula: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim mode As Variant, c As range
    mode = HAMU_InputValue(HAMU_Text("1=Sayı 2=Metin 3=Hata 4=Boş"), HAMU_APP_NAME, 1, ValueType:=1)
    If mode = False Then Exit Sub
    hamuScope.Interior.pattern = xlNone
    For Each c In hamuScope.Cells
        Select Case CLng(mode)
            Case 1: If Not IsError(c.Value2) Then If IsNumeric(c.Value2) Then HAMU_HighlightCell c, RGB(255, 199, 44)
            Case 2: If VarType(c.Value2) = vbString Then HAMU_HighlightCell c, RGB(255, 199, 44)
            Case 3: If IsError(c.Value2) Then HAMU_HighlightCell c, RGB(255, 199, 44)
            Case 4: If Len(CStr(c.Value2)) = 0 Then HAMU_HighlightCell c, RGB(255, 199, 44)
        End Select
    Next c
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_HighlightByCriteria", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_HighlightDuplicates()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_HighlightDuplicates") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Tekrar Edenleri Vurgula: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim dict As Object, c As range, key As String
    Set dict = CreateObject("Scripting.Dictionary")
    For Each c In hamuScope.Cells
        key = CStr(c.value)
        If Len(key) > 0 Then dict(key) = dict(key) + 1
    Next c
    For Each c In hamuScope.Cells
        key = CStr(c.value)
        If Len(key) > 0 And dict(key) > 1 Then HAMU_HighlightCell c, RGB(255, 230, 153)
    Next c
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_HighlightDuplicates", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_AddDuplicateCount()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_AddDuplicateCount") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    On Error GoTo Failed
    Dim inputRange As range, outputRange As range, values() As Variant
    Set inputRange = HAMU_PromptRange(HAMU_Text("Tekrarları sayılacak giriş listesini seçin. Tek sütun seçin."))
    If inputRange Is Nothing Then Exit Sub
    If inputRange.columns.count <> 1 Then
        HAMU_ShowInfo HAMU_Text("Giriş aralığı tek sütun olmalıdır."): Exit Sub
    End If
    If inputRange.rows.count = inputRange.Worksheet.rows.count Then Set inputRange = Intersect(inputRange, inputRange.Worksheet.UsedRange)
    If inputRange Is Nothing Then Exit Sub
    Set outputRange = HAMU_PromptRange(HAMU_Text("Tekrar sayılarının yazılacağı çıkış hücresini veya aynı uzunlukta tek sütunu seçin."))
    If outputRange Is Nothing Then Exit Sub
    HAMU_WriteDuplicateCounts inputRange, outputRange
    Exit Sub
Failed:
    HAMU_ShowError HAMU_Text("Tekrar Sayısını Yaz"), Err.number, Err.description
End Sub
Public Sub HAMU_WriteDuplicateCounts(ByVal inputRange As range, ByVal outputRange As range)
    Dim values() As Variant
    If inputRange.columns.count <> 1 Or outputRange.columns.count <> 1 Then Err.Raise 5, , HAMU_Text("Giriş ve çıkış tek sütun olmalıdır.")
    If outputRange.rows.count <> 1 And outputRange.rows.count <> inputRange.rows.count Then Err.Raise 5, , HAMU_Text("Çıkış aralığı tek hücre veya girişle aynı satır sayısında olmalıdır.")
    Set outputRange = outputRange.Cells(1, 1).Resize(inputRange.rows.count, 1)
    If inputRange.Worksheet Is outputRange.Worksheet Then
        If Not Intersect(inputRange, outputRange) Is Nothing Then Err.Raise 5, , HAMU_Text("Çıkış giriş listesiyle çakışamaz.")
    End If
    If Application.CountA(outputRange) > 0 Then
        If Not HAMU_ConfirmOverwrite() Then Exit Sub
    End If
    HAMU_DuplicateCounts inputRange, values
    outputRange.Value2 = values
End Sub
Public Sub HAMU_ListLinks()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ListLinks") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    On Error GoTo ErrH
    Dim links As Variant, i As Long, sh As Worksheet
    links = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    If IsEmpty(links) Then HAMU_ShowInfo HAMU_Text("Dış bağlantı yok."): Exit Sub
    Set sh = HAMU_GetOrCreateSheet("LINKS_" & Format(Now, "hhmmss"))
    sh.Cells(1, 1).value = HAMU_Text("Kaynak")
    For i = LBound(links) To UBound(links)
        sh.Cells(i + 1, 1).value = links(i)
    Next i
    sh.columns.AutoFit
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_ListLinks", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_BreakLinks()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_BreakLinks") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    On Error GoTo ErrH
    Dim links As Variant, i As Long
    links = ActiveWorkbook.LinkSources(Type:=xlLinkTypeExcelLinks)
    If IsEmpty(links) Then HAMU_ShowInfo HAMU_Text("Kırılacak dış bağlantı yok."): Exit Sub
    If MsgBox(HAMU_Text("Tüm dış bağlantılar kırılacak. Emin misiniz?"), vbOKCancel + vbExclamation, HAMU_APP_NAME) <> vbOK Then Exit Sub
    For i = LBound(links) To UBound(links)
        ActiveWorkbook.BreakLink name:=links(i), Type:=xlLinkTypeExcelLinks
    Next i
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_BreakLinks", hamuErrorNumber, hamuErrorText
End Sub




Public Sub HAMU_DuplicateCounts(ByVal r As range, ByRef result() As Variant)
    Dim dict As Object, values As Variant, i As Long, key As String
    Set dict = CreateObject("Scripting.Dictionary")
    If r.rows.count = 1 Then
        ReDim values(1 To 1, 1 To 1): values(1, 1) = r.Value2
    Else
        values = r.Value2
    End If
    ReDim result(1 To UBound(values, 1), 1 To 1)
    For i = 1 To UBound(values, 1)
        If Not IsError(values(i, 1)) Then
            key = CStr(values(i, 1))
            If Len(key) > 0 Then dict(key) = dict(key) + 1
        End If
    Next
    For i = 1 To UBound(values, 1)
        If Not IsError(values(i, 1)) Then
            key = CStr(values(i, 1))
            If Len(key) > 0 Then result(i, 1) = dict(key) Else result(i, 1) = ""
        Else
            result(i, 1) = ""
        End If
    Next
End Sub

