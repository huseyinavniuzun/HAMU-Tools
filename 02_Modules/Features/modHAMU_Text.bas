Attribute VB_Name = "modHAMU_Text"
Option Explicit
Option Private Module

Public Sub HAMU_CleanData()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_CleanData") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Metin ve Boşluk Temizle: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub

    On Error GoTo ErrHandler

    Dim area As range
    Dim c As range
    Dim s As String
    Dim cleaned As String
    Dim re As Object
    Dim changed As Long
    Dim formulaSkipped As Long
    Dim textChecked As Long

    If TypeName(hamuScope) <> "Range" Then
        MsgBox HAMU_Text("Önce temizlenecek metin alanını seçin."), _
               vbExclamation, "HAMU Tools"
        Exit Sub
    End If

    Set re = CreateObject("VBScript.RegExp")

    With re
        .Global = True
        .pattern = " +"
    End With

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For Each area In hamuScope.areas
        For Each c In area.Cells

            If c.HasFormula Then
                formulaSkipped = formulaSkipped + 1
                GoTo NextCell
            End If

            If Not c.HasFormula And VarType(c.Value2) = vbString Then
                textChecked = textChecked + 1

                s = CStr(c.Value2)
                cleaned = s

                cleaned = Replace$(cleaned, Chr$(160), " ")
                cleaned = Replace$(cleaned, vbTab, " ")
                cleaned = Replace$(cleaned, vbCr, " ")
                cleaned = Replace$(cleaned, vbLf, " ")
                cleaned = Replace$(cleaned, ChrW$(8203), "")
                cleaned = Replace$(cleaned, ChrW$(65279), "")
                cleaned = re.Replace(cleaned, " ")
                cleaned = Trim$(cleaned)

                If cleaned <> s Then
                    c.Value2 = cleaned
                    changed = changed + 1
                End If
            End If

NextCell:
        Next c
    Next area

CleanExit:
    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox HAMU_Text("Metin temizleme tamamlandı.") & vbCrLf & vbCrLf & _
           "Kontrol edilen metin: " & textChecked & vbCrLf & _
           HAMU_Text("Düzeltilen hücre: ") & changed & vbCrLf & _
           HAMU_Text("Korunan formül: ") & formulaSkipped, _
           vbInformation, "HAMU Tools"
    Exit Sub

ErrHandler:
    Application.EnableEvents = hamuState.EventsBefore
    Application.ScreenUpdating = hamuState.ScreenBefore

    MsgBox HAMU_Text("Metin temizleme hatası:") & vbCrLf & _
           HAMU_FriendlyError(Err.number, Err.description), _
           vbExclamation, "HAMU Tools"
End Sub

Public Sub HAMU_TextCase()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_TextCase") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Büyük / Küçük Harf Dönüştür: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim mode As Variant, c As range, s As String
    mode = HAMU_ChooseIndex(HAMU_Text("Harf Dönüştür"), HAMU_Text("Seçili metinlere uygulanacak dönüşümü seçin."), Array(HAMU_Text("BÜYÜK HARF"), HAMU_Text("küçük harf"), HAMU_Text("Baş Harfler")))
    If mode = False Then Exit Sub
    For Each c In hamuScope.Cells
        If Not c.HasFormula And VarType(c.Value2) = vbString Then
            s = CStr(c.Value2)
            Select Case CLng(mode)
                Case 1: c.value = HAMU_TurkishCase(s, vbUpperCase)
                Case 2: c.value = HAMU_TurkishCase(s, vbLowerCase)
                Case Else: c.value = HAMU_TurkishCase(s, vbProperCase)
            End Select
        End If
    Next c
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_TextCase", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_ExtractDigits()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ExtractDigits") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Metinden Sayıları Ayıkla: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim c As range, s As String, i As Long, ch As String, out As String
    For Each c In hamuScope.Cells
        If Not IsError(c.Value2) Then
            s = CStr(c.Value2): out = ""
            For i = 1 To Len(s)
                ch = Mid$(s, i, 1)
                If ch >= "0" And ch <= "9" Then out = out & ch
            Next i
            c.Offset(0, 1).value = out
        End If
    Next c
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_ExtractDigits", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_ConvertQuotedNumbers()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ConvertQuotedNumbers") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Metin Olarak Saklanan Sayıları Düzelt: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim c As range
    For Each c In hamuScope.Cells
        If Not c.HasFormula And VarType(c.Value2) = vbString Then
            If Left$(c.text, 1) = "'" Or IsNumeric(c.Value2) Then
                If IsNumeric(c.Value2) Then c.value = CDbl(c.Value2)
            End If
        End If
    Next c
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_ConvertQuotedNumbers", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_ConvertFormulasToValues()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_ConvertFormulasToValues") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Formülleri Değere Dönüştür: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim rng As range
    Set rng = hamuScope
    rng.value = rng.value
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_ConvertFormulasToValues", hamuErrorNumber, hamuErrorText
End Sub

Public Sub HAMU_MaskData()
    On Error GoTo HAMU_CommandFailed
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_MaskData") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim r As range, c As range, s As String, leftN As Variant, rightN As Variant
    Set r = HAMU_WorkRange(HAMU_Text("Maskelenecek metin aralığını seçin."))
    If r Is Nothing Then Exit Sub
    leftN = HAMU_InputValue(HAMU_Text("Baştan korunacak karakter sayısı:"), Default:=2, ValueType:=1)
    If VarType(leftN) = vbBoolean Then Exit Sub
    rightN = HAMU_InputValue(HAMU_Text("Sondan korunacak karakter sayısı:"), Default:=2, ValueType:=1)
    If VarType(rightN) = vbBoolean Then Exit Sub
    If leftN < 0 Or rightN < 0 Or leftN <> Fix(leftN) Or rightN <> Fix(rightN) Then Err.Raise vbObjectError + 360, , HAMU_Text("Sıfır veya pozitif tam sayı girin.")
    For Each c In r.Cells
        If Not c.HasFormula And Not IsError(c.Value2) Then
            s = CStr(c.Value2)
            If Len(s) > leftN + rightN Then c.Value2 = Left$(s, leftN) & String$(Len(s) - leftN - rightN, "*") & right$(s, rightN)
        End If
    Next
    Exit Sub
HAMU_CommandFailed:
    HAMU_ShowError "HAMU_MaskData", Err.number, Err.description
End Sub
Public Sub HAMU_RowStriping()
 On Error GoTo Failed
 Dim r As range, f As frmHAMU_Stripes, first As Long, second As Long, columns As Boolean
 Set r = HAMU_WorkRange(HAMU_L("Renklendirilecek alanı seçin.", "Select the range to stripe."))
 If r Is Nothing Then Exit Sub
 Set f = New frmHAMU_Stripes
 f.show vbModal
 If Not f.accepted Then Unload f: Exit Sub
 first = HAMU_StripeColor(f.cboFirst.ListIndex)
 second = HAMU_StripeColor(f.cboSecond.ListIndex)
 columns = (f.cboDirection.ListIndex = 1)
 Unload f
 HAMU_ApplyStripes r, first, second, columns
 Exit Sub
Failed:
 If Not f Is Nothing Then Unload f
 HAMU_ShowError "HAMU_RowStriping", Err.number, Err.description
End Sub
Public Function HAMU_StripeColor(ByVal index As Long) As Long
 Select Case index
 Case 0: HAMU_StripeColor = vbWhite
 Case 1: HAMU_StripeColor = RGB(242, 242, 242)
 Case 2: HAMU_StripeColor = RGB(219, 234, 254)
 Case 3: HAMU_StripeColor = RGB(220, 252, 231)
 Case 4: HAMU_StripeColor = RGB(254, 249, 195)
 Case 5: HAMU_StripeColor = RGB(255, 237, 213)
 Case 6: HAMU_StripeColor = RGB(243, 232, 255)
 Case 7: HAMU_StripeColor = RGB(252, 231, 243)
 Case 8: HAMU_StripeColor = RGB(224, 247, 250)
 End Select
End Function
Public Sub HAMU_ApplyStripes(ByVal r As range, ByVal first As Long, ByVal second As Long, ByVal byColumn As Boolean)
 Dim i As Long, count As Long, part As range
 count = r.rows.count
 If byColumn Then count = r.columns.count
 For i = 1 To count
  If byColumn Then Set part = r.columns(i) Else Set part = r.rows(i)
  If i Mod 2 = 1 Then part.Interior.color = first Else part.Interior.color = second
 Next i
End Sub


Public Sub HAMU_UnmergeAndFill()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_UnmergeAndFill") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Birleşimleri Çöz ve Doldur: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error GoTo ErrH
    Dim c As range, m As range, seen As Object, addr As String, v
    Set seen = CreateObject("Scripting.Dictionary")
    For Each c In hamuScope.Cells
        If c.MergeCells Then
            Set m = c.MergeArea
            addr = m.address(False, False)
            If Not seen.exists(addr) Then
                v = m.Cells(1, 1).value
                m.UnMerge
                m.value = v
                seen.Add addr, True
            End If
        End If
    Next c
    Exit Sub
ErrH:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    HAMU_ShowError "HAMU_UnmergeAndFill", hamuErrorNumber, hamuErrorText
End Sub


Public Sub HAMU_DeleteBlankCells()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_DeleteBlankCells") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture
    Dim hamuScope As range
    Set hamuScope = HAMU_WorkRange(HAMU_Text("Boş Hücreleri Sil ve Yukarı Kaydır: işlem yapılacak hücreleri seçin."))
    If hamuScope Is Nothing Then Exit Sub
    On Error Resume Next
    hamuScope.SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    On Error GoTo 0
End Sub

Public Sub HAMU_FillBlanksZero()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_FillBlanksZero") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    On Error GoTo Failed
    Dim range As range, count As Double
    Set range = HAMU_WorkRange()
    If range Is Nothing Then Exit Sub
    count = HAMU_FillEmptyCells(range)
    If count = 0 Then HAMU_ShowInfo HAMU_Text("Seçimde boş hücre yok. Formüller ve dolu hücreler korundu.")
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_FillBlanksZero", Err.number, Err.description
End Sub
Public Function HAMU_FillEmptyCells(ByVal range As range) As Double
    Dim values As Variant, row As Long, col As Long, first As Long, length As Long
    Dim pending As range, part As range, areas As Long, rows As Long
    If range.CountLarge = 1 Then
        If IsEmpty(range.Value2) And Not range.HasFormula Then
            range.Value2 = 0: HAMU_FillEmptyCells = 1
        End If
        Exit Function
    End If
    values = range.formula: rows = range.rows.count
    For col = 1 To range.columns.count
        first = 0
        For row = 1 To rows + 1
            If row <= rows Then
                Dim blank As Boolean
                blank = IsEmpty(values(row, col))
                If VarType(values(row, col)) = vbString Then blank = (Len(values(row, col)) = 0)
                If blank Then
                    If first = 0 Then first = row
                ElseIf first > 0 Then
                    GoSub AddPart
                End If
            ElseIf first > 0 Then
                GoSub AddPart
            End If
        Next
    Next
    If Not pending Is Nothing Then pending.Value2 = 0
    Exit Function
AddPart:
    length = row - first
    Set part = range.Cells(first, col).Resize(length, 1)
    HAMU_FillEmptyCells = HAMU_FillEmptyCells + length
    If pending Is Nothing Then Set pending = part Else Set pending = Union(pending, part)
    areas = areas + 1: first = 0
    If areas = 32 Then pending.Value2 = 0: Set pending = Nothing: areas = 0
    Return
End Function



Public Function HAMU_TurkishCase(ByVal text As String, ByVal conversion As VbStrConv) As String
    Dim i As Long, ch As String, atWordStart As Boolean
    atWordStart = True
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If conversion = vbUpperCase Then
            ch = HAMU_TRUpperLetter(ch)
        Else
            ch = HAMU_TRLowerLetter(ch)
            If conversion = vbProperCase Then
                If HAMU_TRUpperLetter(ch) <> HAMU_TRLowerLetter(ch) Then
                    If atWordStart Then ch = HAMU_TRUpperLetter(ch)
                    atWordStart = False
                ElseIf ch <> "'" And ch <> "’" Then
                    atWordStart = (ch < "0" Or ch > "9")
                End If
            End If
        End If
        HAMU_TurkishCase = HAMU_TurkishCase & ch
    Next
End Function
Private Function HAMU_TRLowerLetter(ByVal ch As String) As String
    Select Case ch
        Case "I": HAMU_TRLowerLetter = "ı"
        Case "İ": HAMU_TRLowerLetter = "i"
        Case Else: HAMU_TRLowerLetter = LCase$(ch)
    End Select
End Function
Private Function HAMU_TRUpperLetter(ByVal ch As String) As String
    Select Case ch
        Case "i": HAMU_TRUpperLetter = "İ"
        Case "ı": HAMU_TRUpperLetter = "I"
        Case Else: HAMU_TRUpperLetter = UCase$(ch)
    End Select
End Function
