Attribute VB_Name = "modHAMU_Date"
Option Explicit
Option Private Module

Public Sub HAMU_DatePicker()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_DatePicker") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim target As range, selected As Date, initial As Date
    On Error GoTo Failed
    If Not HAMU_HasWorkbook() Or TypeName(Selection) <> "Range" Then Exit Sub
    Set target = ActiveCell
    initial = Date
    If Not IsError(target.Value2) Then
        If IsDate(target.value) Then initial = CDate(target.value)
    End If
    selected = frmHAMU_DateSimple.GetDate(initial)
    If selected = 0 Then Exit Sub
    target.value = selected
    target.NumberFormat = HAMU_DEFAULT_DATE_FMT
    Exit Sub
Failed:
    HAMU_ShowError HAMU_Text("Tarih Seç"), Err.number, Err.description
End Sub
Public Sub HAMU_FixDates()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_FixDates") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Tarihleri Düzelt: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub

    On Error GoTo ErrHandler

    Dim rng As range
    Dim c As range
    Dim dt As Date

    Dim duzelen As Long
    Dim zatenTarih As Long
    Dim gecersiz As Long
    Dim bos As Long
    Dim formul As Long

    If TypeName(Selection) <> "Range" Then

        MsgBox _
            HAMU_Text("Önce tarihlerin bulunduğu alanı seçin."), _
            vbExclamation, _
            "HAMU Tools"

        Exit Sub

    End If

    Set rng = hamuScope

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each c In rng.Cells

        If IsError(c.value) Then

            gecersiz = gecersiz + 1
            GoTo NextCell

        End If

        If Len(Trim$(CStr(c.Value2))) = 0 Then

            bos = bos + 1
            GoTo NextCell

        End If

        If c.HasFormula Then

            If HAMU_TryParseDate(c.value, dt) Then
                c.NumberFormat = "dd.mm.yyyy"
            End If

            formul = formul + 1
            GoTo NextCell

        End If

        If HAMU_TryParseDate(c.Value2, dt) Then
            If IsNumeric(c.Value2) Then

                If CDbl(c.Value2) > 0 And _
                   CDbl(c.Value2) < 2958466# And _
                   Len(CStr(Fix(CDbl(c.Value2)))) < 8 Then

                    zatenTarih = zatenTarih + 1

                Else
                    duzelen = duzelen + 1
                End If

            Else
                duzelen = duzelen + 1
            End If

            c.value = dt
            c.NumberFormat = "dd.mm.yyyy"

        Else

            gecersiz = gecersiz + 1

        End If

NextCell:

    Next c

CleanExit:

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Tarih düzeltme tamamlandı.") & vbCrLf & vbCrLf & _
        HAMU_Text("Düzeltilen: ") & Format$(duzelen, "#,##0") & vbCrLf & _
        "Zaten tarih: " & Format$(zatenTarih, "#,##0") & vbCrLf & _
        HAMU_Text("Formül: ") & Format$(formul, "#,##0") & vbCrLf & _
        HAMU_Text("Boş: ") & Format$(bos, "#,##0") & vbCrLf & _
        HAMU_Text("Tanımlanamayan: ") & Format$(gecersiz, "#,##0"), _
        vbInformation, _
        "HAMU Tools"

    Exit Sub

ErrHandler:

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Tarih düzeltme hatası:") & vbCrLf & _
        HAMU_FriendlyError(Err.number, Err.description), _
        vbExclamation, _
        "HAMU Tools"

End Sub

Public Sub HAMU_CheckDates()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_CheckDates") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Tarih Kontrolü: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub

    On Error GoTo ErrHandler

    Dim rng As range
    Dim c As range
    Dim dt As Date

    Dim gecerli As Long
    Dim gecersiz As Long
    Dim bos As Long

    If TypeName(Selection) <> "Range" Then

        MsgBox _
            HAMU_Text("Önce kontrol edilecek alanı seçin."), _
            vbExclamation, _
            "HAMU Tools"

        Exit Sub

    End If

    Set rng = hamuScope

    Application.ScreenUpdating = False

    HAMU_CheckDateCells rng, gecerli, gecersiz, bos

    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Tarih kontrolü tamamlandı.") & vbCrLf & vbCrLf & _
        HAMU_Text("Geçerli: ") & Format$(gecerli, "#,##0") & vbCrLf & _
        HAMU_Text("Geçersiz: ") & Format$(gecersiz, "#,##0") & vbCrLf & _
        HAMU_Text("Boş: ") & Format$(bos, "#,##0"), _
        vbInformation, _
        "HAMU Tools"

    Exit Sub

ErrHandler:

    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Tarih kontrolü hatası:") & vbCrLf & _
        HAMU_FriendlyError(Err.number, Err.description), _
        vbExclamation, _
        "HAMU Tools"

End Sub

Public Sub HAMU_DateBuilder()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_DateBuilder") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    On Error GoTo ErrHandler

    Dim rngYear As range
    Dim rngMonth As range
    Dim rngDay As range
    Dim target As range

    Dim rowCount As Long
    Dim i As Long

    Dim Y As Long
    Dim m As Long
    Dim d As Long

    Dim dt As Date
    Dim monthValue As Variant

    Dim created As Long
    Dim invalid As Long

    Set rngYear = HAMU_SelectRange( _
        HAMU_Text("YIL değerlerinin bulunduğu alanı seçin."))

    If rngYear Is Nothing Then Exit Sub

    Set rngMonth = HAMU_SelectRange( _
        HAMU_Text("AY değerlerinin bulunduğu alanı seçin."))

    If rngMonth Is Nothing Then Exit Sub

    Set rngDay = HAMU_SelectRange( _
        HAMU_Text("GÜN değerlerinin bulunduğu alanı seçin."))

    If rngDay Is Nothing Then Exit Sub

    Set target = HAMU_SelectRange( _
        HAMU_Text("Oluşturulan tarihlerin başlayacağı TEK hücreyi seçin."))

    If target Is Nothing Then Exit Sub

    Set target = target.Cells(1, 1)

    If rngYear.Cells.count <> rngMonth.Cells.count Or _
       rngYear.Cells.count <> rngDay.Cells.count Then

        MsgBox _
            HAMU_Text("Yıl, Ay ve Gün alanlarının hücre sayıları eşit olmalıdır."), _
            vbExclamation, _
            "HAMU Tools"

        Exit Sub

    End If

    rowCount = rngYear.Cells.count

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 1 To rowCount

        Y = 0
        m = 0
        d = 0

        If IsNumeric(rngYear.Cells(i).value) Then

            Y = CLng(rngYear.Cells(i).value)
            Y = HAMU_NormalizeYear(Y)

        End If

        monthValue = rngMonth.Cells(i).value

        If IsNumeric(monthValue) Then

            m = CLng(monthValue)

        Else

            m = HAMU_TurkishMonthNumber( _
                    CStr(monthValue))

        End If

        If IsNumeric(rngDay.Cells(i).value) Then

            d = CLng(rngDay.Cells(i).value)

        End If

        If HAMU_MakeValidDate(Y, m, d, dt) Then

            target.Offset(i - 1, 0).value = dt
            target.Offset(i - 1, 0).NumberFormat = "dd.mm.yyyy"

            created = created + 1

        Else

            target.Offset(i - 1, 0).value = ""
            HAMU_HighlightCell target.Offset(i - 1, 0), RGB(255, 199, 206)

            invalid = invalid + 1

        End If

    Next i

CleanExit:

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Tarih oluşturma tamamlandı.") & vbCrLf & vbCrLf & _
        HAMU_Text("Oluşturulan: ") & Format$(created, "#,##0") & vbCrLf & _
        HAMU_Text("Geçersiz: ") & Format$(invalid, "#,##0"), _
        vbInformation, _
        "HAMU Tools"

    Exit Sub

ErrHandler:

    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox _
        HAMU_Text("Tarih oluşturucu hatası:") & vbCrLf & _
        HAMU_FriendlyError(Err.number, Err.description), _
        vbExclamation, _
        "HAMU Tools"

End Sub

Private Function HAMU_TryParseDate( _
    ByVal InputValue As Variant, _
    ByRef ResultDate As Date) As Boolean

    On Error GoTo Fail

    Dim s As String
    Dim parts() As String

    Dim d As Long
    Dim m As Long
    Dim Y As Long

    Dim n As Double

    If IsNumeric(InputValue) And _
       VarType(InputValue) <> vbString Then

        n = CDbl(InputValue)

        If n >= 10000000# And _
           n <= 99999999# Then

            s = Format$(Fix(n), "00000000")

            If HAMU_ParseEightDigitDate( _
                s, _
                ResultDate) Then

                HAMU_TryParseDate = True
                Exit Function

            End If

        End If

        If n > 0 And _
           n < 2958466# Then

            ResultDate = CDate(Fix(n))

            HAMU_TryParseDate = True
            Exit Function

        End If

    End If

    s = CStr(InputValue)
    s = Replace$(s, Chr$(160), " ")

    s = Trim$(s)

    If Len(s) = 0 Then GoTo Fail

    If Left$(s, 1) = "'" Then

        s = Mid$(s, 2)
        s = Trim$(s)

    End If

    If InStr(1, s, "T", vbTextCompare) > 0 Then

        s = Split(s, "T")(0)

    End If

    If HAMU_ParseTurkishNamedDate( _
        s, _
        ResultDate) Then

        HAMU_TryParseDate = True
        Exit Function

    End If

    If InStr(s, " ") > 0 Then

        s = Split(s, " ")(0)

    End If

    If HAMU_IsDigitsOnly(s) Then

        Select Case Len(s)

            Case 8

                If HAMU_ParseEightDigitDate( _
                    s, _
                    ResultDate) Then

                    HAMU_TryParseDate = True
                    Exit Function

                End If

            Case 6

                d = CLng(Left$(s, 2))
                m = CLng(Mid$(s, 3, 2))
                Y = CLng(right$(s, 2))

                Y = HAMU_NormalizeYear(Y)

                If HAMU_MakeValidDate( _
                    Y, m, d, ResultDate) Then

                    HAMU_TryParseDate = True
                    Exit Function

                End If

            Case 4, 5

                n = CDbl(s)

                If n > 0 And _
                   n < 2958466# Then

                    ResultDate = CDate(Fix(n))

                    HAMU_TryParseDate = True
                    Exit Function

                End If

        End Select

    End If

    s = Replace$(s, "\", ".")
    s = Replace$(s, "/", ".")
    s = Replace$(s, "-", ".")
    s = Replace$(s, "_", ".")

    Do While InStr(s, "..") > 0

        s = Replace$(s, "..", ".")

    Loop

    parts = Split(s, ".")

    If UBound(parts) <> 2 Then GoTo Fail

    If Len(Trim$(parts(0))) = 4 Then

        Y = CLng(val(parts(0)))
        m = CLng(val(parts(1)))
        d = CLng(val(parts(2)))

    Else

        d = CLng(val(parts(0)))
        m = CLng(val(parts(1)))
        Y = CLng(val(parts(2)))

        Y = HAMU_NormalizeYear(Y)

    End If

    If HAMU_MakeValidDate( _
        Y, m, d, ResultDate) Then

        HAMU_TryParseDate = True
        Exit Function

    End If

Fail:

    HAMU_TryParseDate = False

End Function

Private Function HAMU_ParseEightDigitDate( _
    ByVal s As String, _
    ByRef ResultDate As Date) As Boolean

    Dim d As Long
    Dim m As Long
    Dim Y As Long

    If Len(s) <> 8 Then Exit Function

    Y = CLng(Left$(s, 4))

    If Y >= 1900 And _
       Y <= 9999 Then

        m = CLng(Mid$(s, 5, 2))
        d = CLng(right$(s, 2))

        If HAMU_MakeValidDate( _
            Y, m, d, ResultDate) Then

            HAMU_ParseEightDigitDate = True
            Exit Function

        End If

    End If

    d = CLng(Left$(s, 2))
    m = CLng(Mid$(s, 3, 2))
    Y = CLng(right$(s, 4))

    If HAMU_MakeValidDate( _
        Y, m, d, ResultDate) Then

        HAMU_ParseEightDigitDate = True

    End If

End Function

Private Function HAMU_MakeValidDate( _
    ByVal Y As Long, _
    ByVal m As Long, _
    ByVal d As Long, _
    ByRef ResultDate As Date) As Boolean

    On Error GoTo Fail

    Dim temp As Date

    If Y < 1900 Or Y > 9999 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function

    temp = DateSerial(Y, m, d)

    If year(temp) <> Y Then Exit Function
    If month(temp) <> m Then Exit Function
    If day(temp) <> d Then Exit Function

    ResultDate = temp

    HAMU_MakeValidDate = True

    Exit Function

Fail:

    HAMU_MakeValidDate = False

End Function

Private Function HAMU_NormalizeYear( _
    ByVal Y As Long) As Long

    If Y < 100 Then

        If Y <= 49 Then

            HAMU_NormalizeYear = _
                2000 + Y

        Else

            HAMU_NormalizeYear = _
                1900 + Y

        End If

    Else

        HAMU_NormalizeYear = Y

    End If

End Function

Private Function HAMU_IsDigitsOnly( _
    ByVal s As String) As Boolean

    Dim i As Long
    Dim ch As String

    If Len(s) = 0 Then Exit Function

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)

        If ch < "0" Or _
           ch > "9" Then

            Exit Function

        End If

    Next i

    HAMU_IsDigitsOnly = True

End Function

Private Function HAMU_ParseTurkishNamedDate( _
    ByVal s As String, _
    ByRef ResultDate As Date) As Boolean

    Dim temp As String
    Dim parts() As String

    Dim d As Long
    Dim m As Long
    Dim Y As Long

    temp = Trim$(s)
    Do While InStr(temp, "  ") > 0

        temp = Replace$(temp, "  ", " ")

    Loop

    parts = Split(temp, " ")

    If UBound(parts) <> 2 Then Exit Function

    d = CLng(val(parts(0)))

    m = HAMU_TurkishMonthNumber( _
        CStr(parts(1)))

    Y = CLng(val(parts(2)))

    Y = HAMU_NormalizeYear(Y)

    If m = 0 Then Exit Function

    HAMU_ParseTurkishNamedDate = _
        HAMU_MakeValidDate( _
            Y, _
            m, _
            d, _
            ResultDate)

End Function

Private Function HAMU_TurkishMonthNumber( _
    ByVal MonthText As String) As Long

    Dim s As String

    s = LCase$(Trim$(MonthText))

    Select Case s

        Case "ocak"

            HAMU_TurkishMonthNumber = 1

        Case HAMU_Text("şubat"), "subat"

            HAMU_TurkishMonthNumber = 2

        Case "mart"

            HAMU_TurkishMonthNumber = 3

        Case "nisan"

            HAMU_TurkishMonthNumber = 4

        Case HAMU_Text("mayıs"), "mayis"

            HAMU_TurkishMonthNumber = 5

        Case "haziran"

            HAMU_TurkishMonthNumber = 6

        Case "temmuz"

            HAMU_TurkishMonthNumber = 7

        Case HAMU_Text("ağustos"), "agustos"

            HAMU_TurkishMonthNumber = 8

        Case HAMU_Text("eylül"), "eylul"

            HAMU_TurkishMonthNumber = 9

        Case "ekim"

            HAMU_TurkishMonthNumber = 10

        Case HAMU_Text("kasım"), "kasim"

            HAMU_TurkishMonthNumber = 11

        Case HAMU_Text("aralık"), "aralik"

            HAMU_TurkishMonthNumber = 12

    End Select

End Function

Private Function HAMU_SelectRange(ByVal promptText As String) As range
    Set HAMU_SelectRange = HAMU_PromptRange(promptText)
End Function


Public Function HAMU_DateDistance(ByVal value As Date) As String
    Dim days As Long
    days = DateDiff("d", Date, value)
    If days = 0 Then
        HAMU_DateDistance = HAMU_Text("Bugün")
    ElseIf days > 0 Then
        HAMU_DateDistance = days & HAMU_Text(" gün sonra")
    Else
        HAMU_DateDistance = Abs(days) & HAMU_Text(" gün önce")
    End If
End Function

Public Function HAMU_ParseDateInput(ByVal value As String, ByRef parsed As Date) As Boolean
    HAMU_ParseDateInput = HAMU_TryParseDate(value, parsed)
End Function



Public Sub HAMU_CheckDateCells(ByVal rng As range, ByRef gecerli As Long, ByRef gecersiz As Long, ByRef bos As Long)
    Dim c As range, dt As Date
    gecerli = 0: gecersiz = 0: bos = 0
    For Each c In rng.Cells
        If IsError(c.value) Then

            HAMU_HighlightCell c, RGB(255, 199, 206)
            gecersiz = gecersiz + 1

            GoTo NextCell

        End If
        If Len(Trim$(CStr(c.Value2))) = 0 Then

            bos = bos + 1
            GoTo NextCell

        End If
        If HAMU_TryParseDate(c.value, dt) Then

            HAMU_HighlightCell c, RGB(226, 239, 218)
            gecerli = gecerli + 1

        Else

            HAMU_HighlightCell c, RGB(255, 199, 206)
            gecersiz = gecersiz + 1

        End If

NextCell:

    Next c
End Sub
