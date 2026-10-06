Attribute VB_Name = "modHAMU_Text_Regex"
Option Explicit
Option Private Module

Private Const HAMU_REGEX_APP As String = "HAMU Tools"
Private Const HAMU_REGEX_FORM As String = "frmHAMU_Regex"
Private Const HAMU_REGEX_SEPARATOR As String = " | "

Public Sub HAMU_PatternFind()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_PatternFind") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    Dim frm As Object

    On Error GoTo FormNotInstalled

    For Each frm In VBA.UserForms
        If TypeName(frm) = HAMU_REGEX_FORM Then
            If frm.IsRunning Then Exit Sub
            frm.Setup "FIND"
            frm.show vbModeless
            Exit Sub
        End If
    Next
    Set frm = VBA.UserForms.Add(HAMU_REGEX_FORM)

    CallByName frm, "Setup", VbMethod, "FIND"
    frm.show vbModeless

    Exit Sub

FormNotInstalled:

    MsgBox _
        HAMU_Text("Regex formu henüz kurulmamış.") & vbCrLf & vbCrLf & _
        HAMU_Text("05_Forms klasöründeki frmHAMU_Regex.frm dosyasını içe aktarın."), _
        vbInformation, _
        HAMU_REGEX_APP

End Sub

Public Sub HAMU_RegexReplace()
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_RegexReplace") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    Dim hamuState As clsHAMU_AppState
    Set hamuState = New clsHAMU_AppState
    hamuState.Capture

    Dim frm As Object

    On Error GoTo FormNotInstalled

    For Each frm In VBA.UserForms
        If TypeName(frm) = HAMU_REGEX_FORM Then
            If frm.IsRunning Then Exit Sub
            frm.Setup "REPLACE"
            frm.show vbModeless
            Exit Sub
        End If
    Next
    Set frm = VBA.UserForms.Add(HAMU_REGEX_FORM)

    CallByName frm, "Setup", VbMethod, "REPLACE"
    frm.show vbModeless

    Exit Sub

FormNotInstalled:

    MsgBox _
        HAMU_Text("Regex formu henüz kurulmamış.") & vbCrLf & vbCrLf & _
        HAMU_Text("05_Forms klasöründeki frmHAMU_Regex.frm dosyasını içe aktarın."), _
        vbInformation, _
        HAMU_REGEX_APP

End Sub

Public Function HAMU_Regex_GetCatalog( _
    ByVal modeName As String) As Variant

    If UCase$(Trim$(modeName)) = "REPLACE" Then
        HAMU_Regex_GetCatalog = HAMU_RegexReplaceCatalog()
    Else
        HAMU_Regex_GetCatalog = HAMU_RegexFindCatalog()
    End If

End Function

Public Function HAMU_Regex_GetPreset( _
    ByVal modeName As String, _
    ByVal categoryName As String, _
    ByVal presetName As String) As Variant

    Dim cat As Variant
    Dim i As Long

    cat = HAMU_Regex_GetCatalog(modeName)

    For i = LBound(cat) To UBound(cat)

        If StrComp(CStr(cat(i)(0)), categoryName, vbTextCompare) = 0 And _
           StrComp(CStr(cat(i)(1)), presetName, vbTextCompare) = 0 Then

            HAMU_Regex_GetPreset = cat(i)
            Exit Function

        End If

    Next i

End Function

Public Function HAMU_Regex_RunFindUI( _
 ByVal sourceRange As range, ByVal selectedTarget As range, ByVal outputMode As Long, _
 ByVal resultMode As Long, ByVal pattern As String, ByVal ignoreCase As Boolean, _
 ByVal firstRowHeader As Boolean, ByVal presetName As String, Optional ByVal owner As Object = Nothing) As String
 On Error GoTo Failed
 Dim state As New clsHAMU_AppState, regex As Object, matches As Object, target As range
 Dim values As Variant, outArr() As Variant, i As Long, startIndex As Long, matchCells As Long, textValue As String
 Dim errorNumber As Long, errorText As String
 state.Capture
 Set sourceRange = HAMU_Regex_SourceRange(sourceRange)
 If sourceRange.columns.count <> 1 Then Err.Raise 5, , HAMU_L("Kaynak tek sütun olmalıdır.", "The source must be one column.")
 If outputMode < 0 Or outputMode > 2 Then Err.Raise 5, , HAMU_L("Çıktı türünü seçin.", "Choose an output type.")
 If resultMode < 1 Or resultMode > 4 Then Err.Raise 5, , HAMU_L("Sonuç türünü seçin.", "Choose a result type.")
 HAMU_Regex_ValidatePattern pattern, ignoreCase, False
 Set regex = HAMU_Regex_Create(pattern, ignoreCase, resultMode = 2 Or resultMode = 3, False)
 values = HAMU_Regex_ReadValues(sourceRange)
 ReDim outArr(1 To sourceRange.rows.count, 1 To 1)
 startIndex = 1
 If firstRowHeader Then startIndex = 2: outArr(1, 1) = HAMU_L("Normal İfade - ", "Pattern - ") & presetName
 For i = startIndex To sourceRange.rows.count
  If Not IsError(values(i, 1)) Then
   textValue = CStr(values(i, 1))
   If resultMode = 4 Then
    If regex.Test(textValue) Then
     outArr(i, 1) = HAMU_L("Var", "Yes"): matchCells = matchCells + 1
    Else
     outArr(i, 1) = HAMU_L("Yok", "No")
    End If
   Else
    Set matches = regex.Execute(textValue)
    If matches.count > 0 Then matchCells = matchCells + 1
    Select Case resultMode
    Case 1
     outArr(i, 1) = ""
     If matches.count > 0 Then outArr(i, 1) = matches(0).value
    Case 2: outArr(i, 1) = HAMU_Regex_JoinMatches(matches)
    Case 3: outArr(i, 1) = matches.count
    End Select
   End If
  End If
  If i Mod 256 = 0 Then HAMU_Regex_CheckCancel owner
 Next
 HAMU_Regex_CheckCancel owner
 Set target = HAMU_Regex_ResolveFindTarget(sourceRange, selectedTarget, outputMode)
 Set target = HAMU_OutputRange(target, sourceRange.rows.count, 1, sourceRange)
 If target Is Nothing Then GoTo Done
 Application.EnableEvents = False: Application.ScreenUpdating = False
 If resultMode <> 3 Then target.NumberFormat = "@" Else target.NumberFormat = "General"
 target.Value2 = outArr
 If firstRowHeader Then target.rows(1).Font.Bold = True
 HAMU_Regex_RunFindUI = HAMU_L("Tamamlandı: ", "Complete: ") & matchCells & HAMU_L(" hücrede eşleşme bulundu.", " cells matched.")
Done:
 state.Restore
 Exit Function
Failed:
 errorNumber = Err.number: errorText = Err.description
 state.Restore
 Err.Raise errorNumber, , errorText
End Function

Public Function HAMU_Regex_RunReplaceUI( _
 ByVal sourceRange As range, ByVal selectedTarget As range, ByVal outputMode As Long, _
 ByVal pattern As String, ByVal replacement As String, ByVal ignoreCase As Boolean, _
 ByVal globalReplace As Boolean, ByVal firstRowHeader As Boolean, Optional ByVal owner As Object = Nothing) As String
 On Error GoTo Failed
 Dim state As New clsHAMU_AppState, regex As Object, target As range
 Dim values As Variant, changedMask() As Boolean, formulaMask() As Boolean, formulas As range, cell As range
 Dim r As Long, c As Long, startRow As Long, changed As Long, skippedFormula As Long, oldText As String, newText As String
 Dim errorNumber As Long, errorText As String, processed As Long
 state.Capture
 Set sourceRange = HAMU_Regex_SourceRange(sourceRange)
 If outputMode < 0 Or outputMode > 2 Then Err.Raise 5, , HAMU_L("Çıktı türünü seçin.", "Choose an output type.")
 HAMU_Regex_ValidatePattern pattern, ignoreCase, False
 Set regex = HAMU_Regex_Create(pattern, ignoreCase, globalReplace, False)
 values = HAMU_Regex_ReadValues(sourceRange)
 ReDim changedMask(1 To sourceRange.rows.count, 1 To sourceRange.columns.count)
 ReDim formulaMask(1 To sourceRange.rows.count, 1 To sourceRange.columns.count)
 If sourceRange.CountLarge = 1 Then
  formulaMask(1, 1) = sourceRange.HasFormula
 Else
  On Error Resume Next
  Set formulas = sourceRange.SpecialCells(xlCellTypeFormulas)
  On Error GoTo Failed
  If Not formulas Is Nothing Then
   Set formulas = Application.Intersect(formulas, sourceRange)
   If Not formulas Is Nothing Then
    For Each cell In formulas.Cells
     formulaMask(cell.row - sourceRange.row + 1, cell.column - sourceRange.column + 1) = True
    Next
   End If
  End If
 End If
 startRow = 1: If firstRowHeader Then startRow = 2
 For r = startRow To sourceRange.rows.count
  For c = 1 To sourceRange.columns.count
   If formulaMask(r, c) Then
    skippedFormula = skippedFormula + 1
   ElseIf Not IsError(values(r, c)) Then
    oldText = CStr(values(r, c))
    newText = regex.Replace(oldText, replacement)
    If StrComp(oldText, newText, vbBinaryCompare) <> 0 Then
     changedMask(r, c) = True: values(r, c) = newText: changed = changed + 1
    End If
   End If
   processed = processed + 1
   If processed Mod 256 = 0 Then HAMU_Regex_CheckCancel owner
  Next
 Next
 HAMU_Regex_CheckCancel owner
 If outputMode = 2 Then
  Set target = sourceRange
 Else
  Set target = HAMU_Regex_ResolveReplaceTarget(sourceRange, selectedTarget, outputMode)
  Set target = HAMU_OutputRange(target, sourceRange.rows.count, sourceRange.columns.count, sourceRange)
  If target Is Nothing Then GoTo Done
 End If
 Application.EnableEvents = False: Application.ScreenUpdating = False
 If outputMode <> 2 Then sourceRange.copy Destination:=target
 HAMU_Regex_WriteChanges target, values, changedMask
 HAMU_Regex_RunReplaceUI = HAMU_L("Tamamlandı: ", "Complete: ") & changed & HAMU_L(" hücre değiştirildi.", " cells changed.")
 If skippedFormula > 0 Then HAMU_Regex_RunReplaceUI = HAMU_Regex_RunReplaceUI & " " & skippedFormula & HAMU_L(" formül korundu.", " formulas preserved.")
Done:
 state.Restore
 Exit Function
Failed:
 errorNumber = Err.number: errorText = Err.description
 state.Restore
 Err.Raise errorNumber, , errorText
End Function

Private Function HAMU_Regex_ResolveFindTarget( _
    ByVal sourceRange As range, _
    ByVal selectedTarget As range, _
    ByVal outputMode As Long) As range

    Dim ws As Worksheet
    Dim colNo As Long
    Dim checkRange As range

    Set ws = sourceRange.Worksheet

    Select Case outputMode

        Case 0

            If selectedTarget Is Nothing Then Exit Function

            Set HAMU_Regex_ResolveFindTarget = selectedTarget.Cells(1, 1)

        Case 1

            colNo = sourceRange.column + sourceRange.columns.count

            Do While colNo <= ws.columns.count

                Set checkRange = ws.range( _
                                    ws.Cells(sourceRange.row, colNo), _
                                    ws.Cells( _
                                        sourceRange.row + sourceRange.rows.count - 1, _
                                        colNo))

                If WorksheetFunction.CountA(checkRange) = 0 Then

                    Set HAMU_Regex_ResolveFindTarget = _
                        ws.Cells(sourceRange.row, colNo)

                    Exit Function

                End If

                colNo = colNo + 1

            Loop

        Case 2

            Set ws = sourceRange.Worksheet.parent.Worksheets.Add( _
                        After:=sourceRange.Worksheet.parent.Worksheets( _
                            sourceRange.Worksheet.parent.Worksheets.count))

            ws.name = HAMU_Regex_UniqueSheetName( _
                        ws.parent, _
                        "REGEX_BUL")

            Set HAMU_Regex_ResolveFindTarget = ws.range("A1")

    End Select

End Function

Private Function HAMU_Regex_ResolveReplaceTarget( _
    ByVal sourceRange As range, _
    ByVal selectedTarget As range, _
    ByVal outputMode As Long) As range

    Dim ws As Worksheet

    Select Case outputMode

        Case 0

            If selectedTarget Is Nothing Then Exit Function

            Set HAMU_Regex_ResolveReplaceTarget = selectedTarget.Cells(1, 1)

        Case 1

            Set ws = sourceRange.Worksheet.parent.Worksheets.Add( _
                        After:=sourceRange.Worksheet.parent.Worksheets( _
                            sourceRange.Worksheet.parent.Worksheets.count))

            ws.name = HAMU_Regex_UniqueSheetName( _
                        ws.parent, _
                        "REGEX_DEGISTIR")

            Set HAMU_Regex_ResolveReplaceTarget = ws.range("A1")

    End Select

End Function

Private Function HAMU_Regex_Create( _
    ByVal pattern As String, _
    ByVal ignoreCase As Boolean, _
    ByVal globalMatches As Boolean, _
    ByVal multiLine As Boolean) As Object

    Dim regex As Object

    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .pattern = pattern
        .ignoreCase = ignoreCase
        .Global = globalMatches
        .multiLine = multiLine
    End With

    Set HAMU_Regex_Create = regex

End Function

Private Sub HAMU_Regex_ValidatePattern( _
    ByVal pattern As String, _
    ByVal ignoreCase As Boolean, _
    ByVal multiLine As Boolean)

    Dim regex As Object

    On Error GoTo InvalidPattern

    Set regex = HAMU_Regex_Create( _
                    pattern, _
                    ignoreCase, _
                    True, _
                    multiLine)

    regex.Test "HAMU"

    Exit Sub

InvalidPattern:

    Err.Raise _
        vbObjectError + 4204, , _
        HAMU_Text("Geçersiz normal ifade:") & vbCrLf & _
        pattern & vbCrLf & vbCrLf & _
        Err.description

End Sub

Private Function HAMU_Regex_JoinMatches( _
    ByVal matches As Object) As String

    Dim i As Long
    Dim result As String

    For i = 0 To matches.count - 1

        If Len(result) > 0 Then
            result = result & HAMU_REGEX_SEPARATOR
        End If

        result = result & matches(i).value

    Next i

    HAMU_Regex_JoinMatches = result

End Function

Private Function HAMU_Regex_UniqueSheetName( _
    ByVal wb As Workbook, _
    ByVal baseName As String) As String

    Dim nameTry As String
    Dim i As Long

    nameTry = Left$(baseName, 31)

    If Not HAMU_Regex_SheetExists(wb, nameTry) Then

        HAMU_Regex_UniqueSheetName = nameTry
        Exit Function

    End If

    i = 1

    Do

        nameTry = Left$(baseName, 27) & "_" & Format$(i, "00")

        If Not HAMU_Regex_SheetExists(wb, nameTry) Then

            HAMU_Regex_UniqueSheetName = nameTry
            Exit Function

        End If

        i = i + 1

    Loop

End Function

Private Function HAMU_Regex_SheetExists( _
    ByVal wb As Workbook, _
    ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = wb.Worksheets(sheetName)

    HAMU_Regex_SheetExists = Not ws Is Nothing

    On Error GoTo 0

End Function

Private Function HAMU_RegexFindCatalog() As Variant

    Dim items() As Variant
    ReDim items(0 To 41)

    items(0) = Array(HAMU_Text("İletişim"), HAMU_Text("E-posta"), "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", "True", "False", "Standart e-posta adresleri.")
    items(1) = Array(HAMU_Text("İletişim"), HAMU_Text("Türkiye Cep Telefonu"), "(\+?90[\s.-]?)?(0[\s.-]?)?5\d{2}[\s.-]?\d{3}[\s.-]?\d{2}[\s.-]?\d{2}", "True", "False", HAMU_Text("+90, 0, boşluk ve tireli cep telefonu."))
    items(2) = Array(HAMU_Text("İletişim"), "Genel Telefon", "(\+?\d{1,3}[\s.-]?)?(\(?\d{2,4}\)?[\s.-]?)?\d{3}[\s.-]?\d{2,4}[\s.-]?\d{2,4}", "True", "False", HAMU_Text("Genel telefon biçimleri."))
    items(3) = Array(HAMU_Text("İletişim"), "Dahili Telefon", "(ext\.?|dahili|x)\s*\d{1,6}", "True", "False", HAMU_Text("Dahili numarası."))
    items(4) = Array(HAMU_Text("Web ve Ağ"), "URL / Web Adresi", "(https?://|www\.)[^\s<>']+", "True", "False", "Web adresleri.")
    items(5) = Array(HAMU_Text("Web ve Ağ"), HAMU_Text("Alan Adı / Domain"), "([A-Z0-9-]+\.)+[A-Z]{2,}", "True", "False", "example.com benzeri domain.")
    items(6) = Array(HAMU_Text("Web ve Ağ"), "IPv4", "\b((25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(25[0-5]|2[0-4]\d|1?\d?\d)\b", "False", "False", "IPv4 adresi.")
    items(7) = Array(HAMU_Text("Web ve Ağ"), "MAC Adresi", "\b[0-9A-F]{2}([:-][0-9A-F]{2}){5}\b", "True", "False", "MAC adresi.")
    items(8) = Array("Tarih ve Zaman", "GG.AA.YYYY / GG-AA-YYYY / GG/AA/YYYY", "\b(0?[1-9]|[12]\d|3[01])[./-](0?[1-9]|1[0-2])[./-]((19|20)?\d{2})\b", "False", "False", HAMU_Text("Gün-ay-yıl tarihleri."))
    items(9) = Array("Tarih ve Zaman", "YYYY-AA-GG", "\b(19|20)\d{2}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])\b", "False", "False", "ISO tarih.")
    items(10) = Array("Tarih ve Zaman", "Saat HH:MM / HH:MM:SS", "\b([01]?\d|2[0-3]):[0-5]\d(:[0-5]\d)?\b", "False", "False", HAMU_Text("24 saat biçimi."))
    items(11) = Array("Tarih ve Zaman", "Tarih ve Saat", "\b(19|20)\d{2}[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])[ T]([01]?\d|2[0-3]):[0-5]\d(:[0-5]\d)?\b", "False", "False", "Tarih ve saat.")
    items(12) = Array("Tarih ve Zaman", HAMU_Text("YYYYMM Ay Anahtarı"), "\b(19|20)\d{2}(0[1-9]|1[0-2])\b", "False", "False", HAMU_Text("202610 gibi ay anahtarı."))
    items(13) = Array("Tarih ve Zaman", HAMU_Text("YYYY.WW / YYYYWW Hafta Anahtarı"), "\b(19|20)\d{2}([.,](0[1-9]|[1-4]\d|5[0-3])|(0[1-9]|[1-4]\d|5[0-3]))\b", "False", "False", "2026.01, 2026,01 veya 202601.")
    items(14) = Array("Finans ve Banka", HAMU_Text("Türkiye IBAN"), "\bTR\d{2}(\s?\d{4}){5}\s?\d{2}\b", "True", "False", HAMU_Text("TR IBAN biçimi."))
    items(15) = Array("Finans ve Banka", "SWIFT / BIC", "\b[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?\b", "False", "False", "8 veya 11 karakterlik SWIFT/BIC.")
    items(16) = Array("Finans ve Banka", HAMU_Text("Kredi / Banka Kartı No"), "\b(\d[ -]?){13,19}\b", "False", "False", HAMU_Text("Kart numarası biçimi; Luhn doğrulamaz."))
    items(17) = Array("Finans ve Banka", HAMU_Text("Türk Lirası Tutarı"), "[-+]?\d{1,3}(\.\d{3})*(,\d+)?(\s*(TL|TRY))?", "True", "False", HAMU_Text("TL / TRY tutarları."))
    items(18) = Array("Finans ve Banka", HAMU_Text("Yüzde"), "[-+]?\d+([.,]\d+)?\s*%", "False", "False", HAMU_Text("Yüzde değerleri."))
    items(19) = Array("Kimlik ve Kod", HAMU_Text("TCKN (Biçim)"), "\b[1-9]\d{10}\b", "False", "False", HAMU_Text("11 rakam; checksum doğrulamaz."))
    items(20) = Array("Kimlik ve Kod", HAMU_Text("VKN (Biçim)"), "\b\d{10}\b", "False", "False", HAMU_Text("10 rakam; checksum doğrulamaz."))
    items(21) = Array("Kimlik ve Kod", "Kimlik No (Genel)", "\b[A-Z0-9]{6,20}\b", "True", "False", "Genel kimlik/kod.")
    items(22) = Array("Kimlik ve Kod", "Pasaport No (Genel)", "\b[A-Z][A-Z0-9]{5,11}\b", "True", "False", HAMU_Text("Genel pasaport biçimi."))
    items(23) = Array("Kimlik ve Kod", HAMU_Text("Türkiye Plaka"), HAMU_Text("\b(0[1-9]|[1-7]\d|8[01])\s?[A-ZÇĞİÖŞÜ]{1,3}\s?\d{2,4}\b"), "True", "False", HAMU_Text("Türkiye plaka."))
    items(24) = Array("Kimlik ve Kod", HAMU_Text("Türkiye Posta Kodu"), "\b\d{5}\b", "False", "False", "5 haneli posta kodu.")
    items(25) = Array("Kimlik ve Kod", "EAN-13", "\b\d{13}\b", "False", "False", HAMU_Text("EAN-13 biçimi."))
    items(26) = Array("Kimlik ve Kod", "GUID / UUID", "\b[0-9A-F]{8}-[0-9A-F]{4}-[1-5][0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\b", "True", "False", "GUID / UUID.")
    items(27) = Array(HAMU_Text("Sayı"), HAMU_Text("Tam Sayı"), "\b\d+\b", "False", "False", HAMU_Text("Pozitif tam sayı."))
    items(28) = Array(HAMU_Text("Sayı"), HAMU_Text("İşaretli Tam Sayı"), "[-+]?\d+", "False", "False", HAMU_Text("Pozitif veya negatif tam sayı."))
    items(29) = Array(HAMU_Text("Sayı"), HAMU_Text("Ondalıklı Sayı - Virgül"), "[-+]?\d+,\d+", "False", "False", HAMU_Text("Virgüllü ondalık."))
    items(30) = Array(HAMU_Text("Sayı"), HAMU_Text("Ondalıklı Sayı - Nokta"), "[-+]?\d+\.\d+", "False", "False", HAMU_Text("Noktalı ondalık."))
    items(31) = Array(HAMU_Text("Sayı"), HAMU_Text("Ondalıklı Sayı - Virgül veya Nokta"), "[-+]?\d+([.,]\d+)?", "False", "False", HAMU_Text("Genel sayı."))
    items(32) = Array(HAMU_Text("Sayı"), HAMU_Text("Bilimsel Gösterim"), "[-+]?\d+([.,]\d+)?[Ee][-+]?\d+", "True", "False", "1.23E+10 benzeri.")
    items(33) = Array("Metin ve Sosyal", HAMU_Text("Türkçe Kelime"), HAMU_Text("[A-ZÇĞİÖŞÜa-zçğıöşü]+"), "False", "False", HAMU_Text("Türkçe kelime."))
    items(34) = Array("Metin ve Sosyal", "Hashtag", HAMU_Text("#[A-Z0-9_ÇĞİÖŞÜçğıöşü]+"), "True", "False", "Hashtag.")
    items(35) = Array("Metin ve Sosyal", HAMU_Text("Mention / @Kullanıcı"), HAMU_Text("@[A-Z0-9_ÇĞİÖŞÜçğıöşü.-]+"), "True", "False", "Mention.")
    items(36) = Array("Metin ve Sosyal", HAMU_Text("Parantez İçindeki Metin"), "\([^)]*\)", "False", "False", HAMU_Text("Parantezli bölüm."))
    items(37) = Array("Teknik / Dosya / HTML", "HTML / XML Etiketi", "<[^>]+>", "True", "False", "HTML/XML etiketi.")
    items(38) = Array("Teknik / Dosya / HTML", "HEX Renk", "#([0-9A-F]{3}|[0-9A-F]{6})\b", "True", "False", "HEX renk.")
    items(39) = Array("Teknik / Dosya / HTML", HAMU_Text("Dosya Uzantısı"), "\.[A-Z0-9]{1,8}\b", "True", "False", HAMU_Text("Dosya uzantısı."))
    items(40) = Array("Teknik / Dosya / HTML", "Windows Dosya Yolu", "[A-Z]:\\([^\\/:*?""<>|\r\n]+\\)*[^\\/:*?""<>|\r\n]*", "True", "False", "Windows dosya yolu.")
    items(41) = Array("Teknik / Dosya / HTML", HAMU_Text("Sürüm Numarası"), "\bv?\d+\.\d+(\.\d+){0,2}\b", "True", "False", HAMU_Text("Sürüm numarası."))

    HAMU_RegexFindCatalog = items

End Function

Private Function HAMU_RegexReplaceCatalog() As Variant

    Dim items() As Variant
    ReDim items(0 To 28)

    items(0) = Array(HAMU_Text("Boşluk ve Görünmeyen Karakterler"), HAMU_Text("Birden Fazla Boşluğu Tek Boşluk Yap"), "[ ]{2,}", " ", "True", "True", "False", HAMU_Text("Arka arkaya gelen boşlukları tek boşluk yapar."))
    items(1) = Array(HAMU_Text("Boşluk ve Görünmeyen Karakterler"), HAMU_Text("Baştaki / Sondaki Boşlukları Sil"), "^\s+|\s+$", "", "True", "True", "False", HAMU_Text("Baş ve sondaki boşlukları siler."))
    items(2) = Array(HAMU_Text("Boşluk ve Görünmeyen Karakterler"), HAMU_Text("Sekmeleri Boşluğa Çevir"), "\t+", " ", "True", "True", "False", HAMU_Text("Tab karakterlerini boşluğa çevirir."))
    items(3) = Array(HAMU_Text("Boşluk ve Görünmeyen Karakterler"), HAMU_Text("Satır Sonlarını Boşluğa Çevir"), "[\r\n]+", " ", "True", "True", "False", HAMU_Text("Satır sonlarını boşluk yapar."))
    items(4) = Array(HAMU_Text("Boşluk ve Görünmeyen Karakterler"), HAMU_Text("NBSP Karakterini Normal Boşluğa Çevir"), HAMU_Text(" +"), " ", "True", "True", "False", HAMU_Text("Kesilmez boşlukları normal boşluk yapar."))
    items(5) = Array(HAMU_Text("Boşluk ve Görünmeyen Karakterler"), HAMU_Text("Tüm Boşlukları Sil"), "\s+", "", "True", "True", "False", HAMU_Text("Tüm boşlukları kaldırır."))
    items(6) = Array("Tarih ve Hafta", "GG/AA/YYYY veya GG-AA-YYYY -> GG.AA.YYYY", "(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})", "$1.$2.$3", "False", "True", "False", HAMU_Text("Tarih ayraçlarını noktaya çevirir."))
    items(7) = Array("Tarih ve Hafta", "YYYY/AA/GG veya YYYY.AA.GG -> YYYY-AA-GG", "(\d{4})[/.](\d{1,2})[/.](\d{1,2})", "$1-$2-$3", "False", "True", "False", HAMU_Text("ISO tarih ayraçlarını tire yapar."))
    items(8) = Array("Tarih ve Hafta", "YYYYWW -> YYYY.WW", "\b((19|20)\d{2})((0[1-9]|[1-4]\d|5[0-3]))\b", "$1.$3", "False", "True", "False", "202601 -> 2026.01")
    items(9) = Array("Tarih ve Hafta", "YYYY,WW -> YYYY.WW", "\b((19|20)\d{2}),(0[1-9]|[1-4]\d|5[0-3])\b", "$1.$3", "False", "True", "False", "2026,01 -> 2026.01")
    items(10) = Array("Tarih ve Hafta", HAMU_Text("Saatte Noktayı İki Noktaya Çevir"), "\b([01]?\d|2[0-3])\.([0-5]\d)\b", "$1:$2", "False", "True", "False", "14.30 -> 14:30")
    items(11) = Array("Telefon / IBAN / Kart", HAMU_Text("Telefonu 0XXX XXX XX XX Biçimine Getir"), "^(\+?90[\s.-]?)?(0[\s.-]?)?(5\d{2})[\s.-]?(\d{3})[\s.-]?(\d{2})[\s.-]?(\d{2})$", "0$3 $4 $5 $6", "False", "True", "False", HAMU_Text("Türkiye cep telefonunu standardize eder."))
    items(12) = Array("Telefon / IBAN / Kart", HAMU_Text("IBAN Boşluklarını Kaldır"), "\s+", "", "True", "True", "False", HAMU_Text("IBAN boşluklarını kaldırır."))
    items(13) = Array("Telefon / IBAN / Kart", HAMU_Text("Boşluksuz TR IBAN'ı Grupla"), "^(TR\d{2})(\d{4})(\d{4})(\d{4})(\d{4})(\d{4})(\d{2})$", "$1 $2 $3 $4 $5 $6 $7", "True", "True", "False", HAMU_Text("TR IBAN'ı gruplar."))
    items(14) = Array("Telefon / IBAN / Kart", HAMU_Text("16 Haneli Kart Numarasını 4'lü Grupla"), "^(\d{4})(\d{4})(\d{4})(\d{4})$", "$1 $2 $3 $4", "False", "True", "False", HAMU_Text("16 haneli kartı gruplar."))
    items(15) = Array("Telefon / IBAN / Kart", HAMU_Text("Kodlardan Tire ve Boşlukları Sil"), "[\s-]+", "", "True", "True", "False", HAMU_Text("Kodlardaki boşluk ve tireleri siler."))
    items(16) = Array(HAMU_Text("Sayı ve Ondalık"), HAMU_Text("Ondalık Virgülü Noktaya Çevir"), "(\d),(\d)", "$1.$2", "False", "True", "False", HAMU_Text("Ondalık virgülü noktaya çevirir."))
    items(17) = Array(HAMU_Text("Sayı ve Ondalık"), HAMU_Text("Ondalık Noktayı Virgüle Çevir"), "(\d)\.(\d)", "$1,$2", "False", "True", "False", HAMU_Text("Ondalık noktayı virgüle çevirir."))
    items(18) = Array(HAMU_Text("Sayı ve Ondalık"), HAMU_Text("Sayı Dışındaki Karakterleri Sil"), "[^\d]+", "", "False", "True", "False", HAMU_Text("Rakam dışındaki karakterleri siler."))
    items(19) = Array(HAMU_Text("Sayı ve Ondalık"), HAMU_Text("Yüzde İşaretini Sil"), "\s*%\s*", "", "False", "True", "False", HAMU_Text("Yüzde işaretini siler."))
    items(20) = Array("Metin Temizleme", HAMU_Text("Parantez İçini ve Parantezleri Sil"), "\([^)]*\)", "", "True", "True", "False", HAMU_Text("Parantezli bölümü siler."))
    items(21) = Array("Metin Temizleme", HAMU_Text("Köşeli Parantez İçini Sil"), "\[[^\]]*\]", "", "True", "True", "False", HAMU_Text("Köşeli parantezli bölümü siler."))
    items(22) = Array("Metin Temizleme", HAMU_Text("Noktalama İşaretlerini Sil"), "[.,;:!?'`()\[\]{}]+", "", "True", "True", "False", HAMU_Text("Noktalama işaretlerini siler."))
    items(23) = Array("Metin Temizleme", HAMU_Text("Alt Çizgi ve Tireleri Boşluğa Çevir"), "[_-]+", " ", "True", "True", "False", HAMU_Text("Alt çizgi ve tireleri boşluğa çevirir."))
    items(24) = Array("Web / HTML", "HTML / XML Etiketlerini Sil", "<[^>]+>", "", "True", "True", "False", "HTML/XML etiketlerini siler.")
    items(25) = Array("Web / HTML", "URL'leri Sil", "(https?://|www\.)[^\s<>']+", "", "True", "True", "False", "Web adreslerini siler.")
    items(26) = Array("Web / HTML", "E-posta Adreslerini Sil", "[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", "", "True", "True", "False", HAMU_Text("E-postaları siler."))
    items(27) = Array("Web / HTML", "Hashtag'leri Sil", HAMU_Text("#[A-Z0-9_ÇĞİÖŞÜçğıöşü]+"), "", "True", "True", "False", "Hashtagleri siler.")
    items(28) = Array("Web / HTML", HAMU_Text("Mention'ları Sil"), HAMU_Text("@[A-Z0-9_ÇĞİÖŞÜçğıöşü.-]+"), "", "True", "True", "False", HAMU_Text("Mentionları siler."))

    HAMU_RegexReplaceCatalog = items

End Function






Private Function HAMU_Regex_SourceRange(ByVal source As range) As range
 If source Is Nothing Then Err.Raise 5, , HAMU_L("Kaynak alanı seçin.", "Select the source range.")
 If source.areas.count <> 1 Then Err.Raise 5, , HAMU_L("Tek bir kesintisiz alan seçin.", "Select one contiguous range.")
 Set source = HAMU_BoundedRange(source)
 If source.CountLarge > 50000 Then Err.Raise 5, , HAMU_L("Normal İfade için en fazla 50.000 hücre seçin. İşlemi parçalara ayırın.", "Use at most 50,000 cells per pattern operation. Split larger jobs.")
 If IsNull(source.MergeCells) Then Err.Raise 5, , HAMU_L("Birleştirilmiş hücreleri ayırın.", "Unmerge the selected cells.")
 If source.MergeCells Then Err.Raise 5, , HAMU_L("Birleştirilmiş hücreleri ayırın.", "Unmerge the selected cells.")
 Set HAMU_Regex_SourceRange = source
End Function
Private Function HAMU_Regex_ReadValues(ByVal source As range) As Variant
 Dim values As Variant
 If source.CountLarge = 1 Then
  ReDim values(1 To 1, 1 To 1): values(1, 1) = source.Value2
 Else
  values = source.Value2
 End If
 HAMU_Regex_ReadValues = values
End Function
Private Sub HAMU_Regex_CheckCancel(ByVal owner As Object)
 DoEvents
 If Not owner Is Nothing Then
  If owner.CancelRequested Then Err.Raise 18, , HAMU_L("İşlem iptal edildi. Sonuç yazılmadı.", "Cancelled. Results were not written.")
 End If
End Sub
Private Sub HAMU_Regex_WriteChanges(ByVal target As range, ByRef values As Variant, ByRef mask() As Boolean)
 Dim r As Long, c As Long, first As Long, last As Long, i As Long, chunk() As Variant, part As range
 For c = 1 To UBound(mask, 2)
  r = 1
  Do While r <= UBound(mask, 1)
   If mask(r, c) Then
    first = r
    Do While r <= UBound(mask, 1)
     If Not mask(r, c) Then Exit Do
     r = r + 1
    Loop
    last = r - 1
    ReDim chunk(1 To last - first + 1, 1 To 1)
    For i = first To last: chunk(i - first + 1, 1) = values(i, c): Next
    Set part = target.Cells(first, c).Resize(last - first + 1, 1)
    part.NumberFormat = "@": part.Value2 = chunk
   Else
    r = r + 1
   End If
  Loop
 Next
End Sub
