Attribute VB_Name = "modHAMU_UDF"
Option Explicit

Public Function HAMU_URLCoz(ByVal strURL As String) As String
Attribute HAMU_URLCoz.VB_Description = "UTF-8 URL karakterlerini çözer. Geçersiz yüzde dizisini korur."
Attribute HAMU_URLCoz.VB_ProcData.VB_Invoke_Func = " \n20"
    HAMU_URLCoz = HAMU_URLDecode(strURL)
End Function

Public Function HAMU_RengeGoreTopla(rng As range, colorCell As range) As Variant
Attribute HAMU_RengeGoreTopla.VB_Description = "B1 ile aynı dolgu rengine sahip sayıları toplar. Biçim değişimi sonrası Ctrl+Alt+F9."
Attribute HAMU_RengeGoreTopla.VB_ProcData.VB_Invoke_Func = " \n20"
    If rng.CountLarge > 250000 Then HAMU_RengeGoreTopla = CVErr(xlErrNum): Exit Function
    Dim c As range, total As Double, sampleColor As Long
    sampleColor = colorCell.Cells(1, 1).Interior.color
    For Each c In rng
        If c.Interior.color = sampleColor Then
            If Not IsError(c.Value2) Then
                If IsNumeric(c.Value2) Then total = total + CDbl(c.Value2)
            End If
        End If
    Next c
    HAMU_RengeGoreTopla = total
End Function

Public Function HAMU_RengeGoreSay(rng As range, colorCell As range) As Variant
Attribute HAMU_RengeGoreSay.VB_Description = "B1 ile aynı dolgu rengine sahip hücreleri sayar. Biçim değişimi sonrası Ctrl+Alt+F9."
Attribute HAMU_RengeGoreSay.VB_ProcData.VB_Invoke_Func = " \n20"
    If rng.CountLarge > 250000 Then HAMU_RengeGoreSay = CVErr(xlErrNum): Exit Function
    Dim c As range, n As Long, sampleColor As Long
    sampleColor = colorCell.Cells(1, 1).Interior.color
    For Each c In rng
        If c.Interior.color = sampleColor Then n = n + 1
    Next c
    HAMU_RengeGoreSay = n
End Function

Public Function HAMU_YerelDosyaYolu(ByVal webPath As String) As String
Attribute HAMU_YerelDosyaYolu.VB_Description = "Kişisel OneDrive URL yolunu yerel OneDrive ortam değişkeniyle eşler; diğer URL türlerini korur."
Attribute HAMU_YerelDosyaYolu.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim prefix As String, base As String, slashPos As Long
    prefix = "https://d.docs.live.net/"
    If LCase$(Left$(webPath, Len(prefix))) <> prefix Then
        HAMU_YerelDosyaYolu = webPath
        Exit Function
    End If
    slashPos = InStr(Len(prefix) + 1, webPath, "/")
    base = Environ$("OneDrive")
    If Len(base) = 0 Or slashPos = 0 Then
        HAMU_YerelDosyaYolu = webPath
    Else
        HAMU_YerelDosyaYolu = base & Replace$(HAMU_URLDecode(Mid$(webPath, slashPos)), "/", Application.PathSeparator)
    End If
End Function
Public Function HAMU_GizliLink(hucre As range) As String
Attribute HAMU_GizliLink.VB_Description = "Hücrenin ilk köprü adresini getirir."
Attribute HAMU_GizliLink.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error Resume Next
    HAMU_GizliLink = hucre.Hyperlinks(1).address
    On Error GoTo 0
End Function

Public Function HAMU_LinkParamDegeri(ByVal hucre As range, ByVal parametreAdi As Variant) As Variant
Attribute HAMU_LinkParamDegeri.VB_Description = "URL içindeki sorgu parametresinin değerini getirir."
Attribute HAMU_LinkParamDegeri.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim url As String, query As String, pos As Long, parts As Variant, piece As Variant
    Dim dict As Object, equalPos As Long, key As String, values As Variant
    Dim i As Long, r As Long, c As Long, isTwo As Boolean, ignored As Long
    On Error GoTo Failed
    url = CStr(hucre.Cells(1, 1).Value2)
    pos = InStr(url, "?")
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    If pos > 0 Then
        query = Mid$(url, pos + 1)
        pos = InStr(query, "#")
        If pos > 0 Then query = Left$(query, pos - 1)
        parts = Split(query, "&")
        For Each piece In parts
            equalPos = InStr(CStr(piece), "=")
            If equalPos > 0 Then dict(HAMU_URLDecode(Left$(CStr(piece), equalPos - 1))) = HAMU_URLDecode(Mid$(CStr(piece), equalPos + 1))
        Next
    End If
    If IsObject(parametreAdi) Then parametreAdi = parametreAdi.Value2
    If Not IsArray(parametreAdi) Then
        key = CStr(parametreAdi)
        If dict.exists(key) Then HAMU_LinkParamDegeri = dict(key) Else HAMU_LinkParamDegeri = ""
        Exit Function
    End If
    On Error Resume Next
    ignored = UBound(parametreAdi, 2)
    isTwo = (Err.number = 0)
    Err.Clear
    On Error GoTo Failed
    If isTwo Then
        ReDim values(LBound(parametreAdi, 1) To UBound(parametreAdi, 1), LBound(parametreAdi, 2) To UBound(parametreAdi, 2))
        For r = LBound(parametreAdi, 1) To UBound(parametreAdi, 1)
            For c = LBound(parametreAdi, 2) To UBound(parametreAdi, 2)
                key = CStr(parametreAdi(r, c))
                If dict.exists(key) Then values(r, c) = dict(key) Else values(r, c) = ""
            Next
        Next
    Else
        ReDim values(LBound(parametreAdi) To UBound(parametreAdi))
        For i = LBound(parametreAdi) To UBound(parametreAdi)
            key = CStr(parametreAdi(i))
            If dict.exists(key) Then values(i) = dict(key) Else values(i) = ""
        Next
    End If
    HAMU_LinkParamDegeri = values
    Exit Function
Failed:
    HAMU_LinkParamDegeri = CVErr(xlErrValue)
End Function
Public Function HAMU_SayidanMetine(ByVal sayi As Double) As String
Attribute HAMU_SayidanMetine.VB_Description = "Türkçe: yüz yirmi üç. 15 basamaktan küçük mutlak değerler; ondalık kısmı tek tek okur."
Attribute HAMU_SayidanMetine.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim whole As Double, groups As Variant, index As Long, part As Long, result As String
    Dim fraction As String, digitNames As Variant, digit As Long, sep As String
    If Abs(sayi) >= 1E+15 Then
        HAMU_SayidanMetine = HAMU_Text("Sayı desteklenen aralığı aşıyor.")
        Exit Function
    End If
    whole = Fix(Abs(sayi))
    groups = Array("", "bin", "milyon", "milyar", "trilyon")
    Do While whole > 0
        part = whole - Fix(whole / 1000) * 1000
        If part > 0 Then
            If index = 1 And part = 1 Then
                result = "bin " & result
            Else
                result = HAMU_ThreeDigits(part) & " " & groups(index) & " " & result
            End If
        End If
        whole = Fix(whole / 1000)
        index = index + 1
    Loop
    If Len(Trim$(result)) = 0 Then result = HAMU_Text("sıfır")
    If sayi < 0 Then result = "eksi " & result
    If Abs(sayi) <> Fix(Abs(sayi)) Then
        sep = Application.International(xlDecimalSeparator)
        fraction = Format$(Abs(sayi), "0.###############")
        fraction = Mid$(fraction, InStr(fraction, sep) + Len(sep))
        digitNames = Array(HAMU_Text("sıfır"), "bir", "iki", HAMU_Text("üç"), HAMU_Text("dört"), HAMU_Text("beş"), HAMU_Text("altı"), "yedi", "sekiz", "dokuz")
        result = result & HAMU_Text(" virgül")
        For digit = 1 To Len(fraction)
            result = result & " " & digitNames(CLng(Mid$(fraction, digit, 1)))
        Next
    End If
    HAMU_SayidanMetine = Application.Trim(result)
End Function
Public Function HAMU_TablodanVeriGetir(ByVal aranan As Variant, ByVal tabloAdi As String, ByVal arananSütun As String, ByVal donenSütun As String, Optional ByVal varsayilan As Variant = "") As Variant
Attribute HAMU_TablodanVeriGetir.VB_Description = "Formülün bulunduğu kitapta tablodan ilk eşleşen değeri getirir."
Attribute HAMU_TablodanVeriGetir.VB_ProcData.VB_Invoke_Func = " \n20"
    On Error GoTo Son
    Dim ws As Worksheet, lo As ListObject, idxFind As Long, idxRet As Long, i As Long
    Dim callerBook As Workbook, keys As Variant
    On Error Resume Next
    Set callerBook = Application.Caller.parent.parent
    On Error GoTo Son
    If callerBook Is Nothing Then Set callerBook = ActiveWorkbook
    For Each ws In callerBook.Worksheets
        On Error Resume Next
        Set lo = ws.ListObjects(tabloAdi)
        On Error GoTo Son
        If Not lo Is Nothing Then Exit For
    Next ws
    If lo Is Nothing Then GoTo Son
    If lo.DataBodyRange Is Nothing Then GoTo Son
    idxFind = lo.ListColumns(arananSütun).index
    idxRet = lo.ListColumns(donenSütun).index
    keys = lo.ListColumns(idxFind).DataBodyRange.value
    If lo.ListRows.count = 1 Then
        If keys = aranan Then HAMU_TablodanVeriGetir = lo.DataBodyRange.Cells(1, idxRet).value: Exit Function
    Else
        For i = 1 To UBound(keys, 1)
            If Not IsError(keys(i, 1)) Then
                If keys(i, 1) = aranan Then HAMU_TablodanVeriGetir = lo.DataBodyRange.Cells(i, idxRet).value: Exit Function
            End If
        Next
    End If
Son:
    HAMU_TablodanVeriGetir = varsayilan
End Function

Private Function HAMU_ThreeDigits(ByVal n As Long) As String
    Dim ones As Variant, tens As Variant, result As String, hundreds As Long
    ones = Array("", "bir", "iki", HAMU_Text("üç"), HAMU_Text("dört"), HAMU_Text("beş"), HAMU_Text("altı"), "yedi", "sekiz", "dokuz")
    tens = Array("", "on", "yirmi", "otuz", HAMU_Text("kırk"), "elli", HAMU_Text("altmış"), HAMU_Text("yetmiş"), "seksen", "doksan")
    hundreds = n \ 100
    If hundreds > 1 Then result = ones(hundreds) & " "
    If hundreds > 0 Then result = result & HAMU_Text("yüz ")
    result = result & tens((n Mod 100) \ 10) & " " & ones(n Mod 10)
    HAMU_ThreeDigits = Trim$(result)
End Function



