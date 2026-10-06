Attribute VB_Name = "modHAMU_Utils"
Option Explicit
Option Private Module
Private mState As clsHAMU_AppState
Private mSafeDepth As Long

Public Sub HAMU_ShowInfo(ByVal msg As String, Optional ByVal ttl As String = HAMU_APP_NAME)
    MsgBox msg, vbInformation, ttl
End Sub

Public Sub HAMU_ShowError(ByVal procName As String, ByVal errNo As Long, ByVal errDesc As String)
    Dim title As String, message As String
    title = HAMU_CommandLabel(procName)
    If Len(title) = 0 Then title = HAMU_CommandLabel(gHAMU_CommandId)
    If Len(title) = 0 Then title = HAMU_APP_NAME
    message = HAMU_FriendlyError(errNo, errDesc)
    MsgBox message, vbExclamation, HAMU_APP_NAME & " | " & title
    On Error Resume Next
    Dim fso As Object, stream As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.OpenTextFile(Environ$("TEMP") & "\HAMU_Tools_Hata.log", 8, True, -1)
    stream.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & procName & " | " & errNo & " | " & errDesc
    stream.Close
End Sub
Public Function HAMU_FriendlyError(ByVal number As Long, ByVal description As String) As String
    Select Case number
        Case 13: HAMU_FriendlyError = HAMU_Text("Seçilen veri bu işlemle uyumlu değil. Hücreleri ve seçenekleri kontrol edin.")
        Case 53, 76: HAMU_FriendlyError = HAMU_Text("Dosya veya klasör bulunamadı. Yolun doğru olduğunu kontrol edin.")
        Case 70, 75: HAMU_FriendlyError = HAMU_Text("Dosyaya erişilemiyor. Dosya açık, salt okunur veya klasör yazmaya kapalı olabilir.")
        Case 18: HAMU_FriendlyError = HAMU_Text("İşlem iptal edildi.")
        Case 1004
            If TypeName(ActiveSheet) = "Worksheet" Then
                If ActiveSheet.ProtectContents Then
                    HAMU_FriendlyError = HAMU_Text("Sayfa korumalı. İşlem için sayfa korumasını kaldırın."): Exit Function
                End If
            End If
            HAMU_FriendlyError = HAMU_Text("Excel işlemi tamamlayamadı. Seçilen alanı, hücre türlerini ve dosya erişimini kontrol edin.")
        Case Else
            If Len(description) > 0 And InStr(description, "defined") = 0 Then
                HAMU_FriendlyError = Left$(description, 240)
            Else
                HAMU_FriendlyError = HAMU_Text("İşlem tamamlanamadı. Seçimi ve işlem seçeneklerini kontrol edin.")
            End If
    End Select
End Function

Public Function HAMU_TemizDosyaAdi(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    s = Trim$(s)
    For i = LBound(bad) To UBound(bad)
        s = Replace$(s, bad(i), "_")
    Next i
    s = Replace$(s, " ", "_")
    s = Replace$(s, HAMU_Text("İ"), "I"): s = Replace$(s, HAMU_Text("ı"), "i")
    s = Replace$(s, HAMU_Text("Ğ"), "G"): s = Replace$(s, HAMU_Text("ğ"), "g")
    s = Replace$(s, HAMU_Text("Ü"), "U"): s = Replace$(s, HAMU_Text("ü"), "u")
    s = Replace$(s, HAMU_Text("Ş"), "S"): s = Replace$(s, HAMU_Text("ş"), "s")
    s = Replace$(s, HAMU_Text("Ö"), "O"): s = Replace$(s, HAMU_Text("ö"), "o")
    s = Replace$(s, HAMU_Text("Ç"), "C"): s = Replace$(s, HAMU_Text("ç"), "c")
    If Len(s) = 0 Then s = "Bos"
    HAMU_TemizDosyaAdi = s
End Function

Public Function HAMU_URLDecode(ByVal strURL As String) As String
    Dim i As Long, j As Long, count As Long, hex As String, bytes() As Byte
    Dim stream As Object, result As String
    strURL = Replace$(strURL, "+", " ")
    i = 1
    Do While i <= Len(strURL)
        If Mid$(strURL, i, 1) = "%" And i + 2 <= Len(strURL) Then
            j = i: count = 0
            ReDim bytes(0 To Len(strURL))
            Do While j + 2 <= Len(strURL)
                If Mid$(strURL, j, 1) <> "%" Then Exit Do
                hex = Mid$(strURL, j + 1, 2)
                If Not hex Like "[0-9A-Fa-f][0-9A-Fa-f]" Then Exit Do
                bytes(count) = CByte("&H" & hex)
                count = count + 1: j = j + 3
            Loop
            If count > 0 Then
                ReDim Preserve bytes(0 To count - 1)
                Set stream = CreateObject("ADODB.Stream")
                stream.Type = 1: stream.Open
                stream.Write bytes
                stream.Position = 0: stream.Type = 2: stream.Charset = "utf-8"
                result = result & stream.ReadText
                stream.Close
                i = j
            Else
                result = result & "%": i = i + 1
            End If
        Else
            result = result & Mid$(strURL, i, 1): i = i + 1
        End If
    Loop
    HAMU_URLDecode = result
End Function
Private Function HAMU_GetSelectedRange() As range
    Set HAMU_GetSelectedRange = HAMU_PromptRange(HAMU_Text("İşlem aralığını seçin."))
End Function
Private Function HAMU_GetWorkbookFolder(ByVal wb As Workbook) As String
    If Len(wb.path) > 0 Then
        HAMU_GetWorkbookFolder = wb.path
    Else
        HAMU_GetWorkbookFolder = Environ$("USERPROFILE")
    End If
End Function

Public Sub HAMU_EnsureFolder(ByVal folderPath As String)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Function HAMU_LastUsedRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    HAMU_LastUsedRow = ws.Cells(ws.rows.count, col).End(xlUp).row
End Function

Public Function HAMU_SheetExists(ByVal nm As String, Optional ByVal wb As Workbook) As Boolean
    Dim sh As Worksheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    On Error Resume Next
    Set sh = wb.Worksheets(nm)
    HAMU_SheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function

Public Function HAMU_GetOrCreateSheet(ByVal nm As String, Optional ByVal wb As Workbook) As Worksheet
    Dim candidate As String, i As Long
    If wb Is Nothing Then Set wb = ActiveWorkbook
    nm = Left$(HAMU_CleanSheetName(nm), 31)
    candidate = nm: i = 1
    Do While HAMU_SheetExists(candidate, wb)
        i = i + 1
        candidate = Left$(nm, 26) & "_" & CStr(i)
    Loop
    Set HAMU_GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    HAMU_GetOrCreateSheet.name = candidate
End Function
Public Sub HAMU_SafeBegin()
    If Not HAMU_HasWorkbook() Then Exit Sub
    If mSafeDepth = 0 Then
        Set mState = New clsHAMU_AppState
        mState.Capture
    End If
    mSafeDepth = mSafeDepth + 1
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub
Public Sub HAMU_SafeEnd()
    If mSafeDepth = 0 Then Exit Sub
    mSafeDepth = mSafeDepth - 1
    If mSafeDepth = 0 Then
        mState.Restore
        Set mState = Nothing
    End If
End Sub
Public Function HAMU_PromptRange(ByVal promptText As String, Optional ByVal allowMulti As Boolean = False) As range
    Dim initial As String, picked As range
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun.")
        Exit Function
    End If
    If TypeName(Selection) = "Range" Then initial = Selection.address(External:=True)
    On Error Resume Next
    Set picked = Application.InputBox(Left$(promptText, 240), HAMU_APP_NAME & HAMU_Text(" | Aralık Seç"), initial, Type:=8)
    On Error GoTo 0
    If picked Is Nothing Then Exit Function
    If Not allowMulti And picked.areas.count > 1 Then
        HAMU_ShowInfo HAMU_Text("Tek bir kesintisiz aralık seçin.")
        Exit Function
    End If
    If picked.areas.count = 1 Then Set picked = HAMU_BoundedRange(picked)
    If picked Is Nothing Then Exit Function
    If picked.CountLarge > 500000 Then
        HAMU_ShowInfo HAMU_L("Seçim 500.000 hücreyi aşıyor. Daha küçük bir veri alanı seçin; tüm sütunlar dolu son hücreye göre daraltılır.", "Selection exceeds 500,000 cells. Select a smaller data area; whole columns are trimmed to the last populated cell.")
        Exit Function
    End If
    Set HAMU_PromptRange = picked
End Function
Private Function HAMU_UniqueName(Optional ByVal prefix As String = "HAMU_") As String
    HAMU_UniqueName = prefix & Format(Now, "yyyymmdd_hhnnss")
End Function

Private Function HAMU_CleanSheetName(ByVal value As String) As String
    Dim bad As Variant, item As Variant
    bad = Array("/", "\", "?", "*", "[", "]", ":", "'")
    For Each item In bad
        value = Replace$(value, CStr(item), "_")
    Next
    If Len(Trim$(value)) = 0 Then value = "HAMU"
    HAMU_CleanSheetName = value
End Function

Public Function HAMU_WorkbookExtension(ByVal wb As Workbook) As String
    Dim extension As String
    If Len(wb.path) > 0 Then
        extension = CreateObject("Scripting.FileSystemObject").GetExtensionName(wb.name)
        If Len(extension) > 0 Then HAMU_WorkbookExtension = "." & extension: Exit Function
    End If
    Select Case wb.fileFormat
        Case xlOpenXMLWorkbookMacroEnabled: HAMU_WorkbookExtension = ".xlsm"
        Case xlExcel12: HAMU_WorkbookExtension = ".xlsb"
        Case xlExcel8: HAMU_WorkbookExtension = ".xls"
        Case xlCSV, 62: HAMU_WorkbookExtension = ".csv"
        Case Else: HAMU_WorkbookExtension = ".xlsx"
    End Select
End Function



Public Function HAMU_WorkRange(Optional ByVal explanation As String = "Islem araligini secin.") As range
    Dim r As range, used As range
    If Not HAMU_HasWorkbook() Then Exit Function
    If HAMU_Setting("RangePrompt") = "1" Then
        Set r = HAMU_PromptRange(explanation)
    Else
        If TypeName(Selection) = "Range" Then Set r = Selection
        If r Is Nothing Then Set r = HAMU_PromptRange(explanation)
    End If
    If r Is Nothing Then Exit Function
    If r.areas.count <> 1 Then
        HAMU_ShowInfo HAMU_Text("Tek bir kesintisiz aralık seçin.")
        Exit Function
    End If
    Set r = HAMU_BoundedRange(r)
    If r Is Nothing Then Exit Function
    If r.CountLarge > 500000 Then
        HAMU_ShowInfo HAMU_L("Seçim çok büyük. En fazla 500.000 hücrelik bir veri alanı seçin.", "Selection is too large. Select at most 500,000 cells.")
        Exit Function
    End If
    Set HAMU_WorkRange = r
End Function

Public Function HAMU_OpenSourceWorkbook(ByVal path As String, ByRef owned As Boolean) As Workbook
    Dim wb As Workbook, security As Long, events As Boolean
    Dim number As Long, description As String
    owned = False
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            Set HAMU_OpenSourceWorkbook = wb: Exit Function
        End If
    Next
    security = Application.AutomationSecurity: events = Application.EnableEvents
    On Error GoTo Failed
    Application.AutomationSecurity = 3: Application.EnableEvents = False
    Set HAMU_OpenSourceWorkbook = Workbooks.Open(fileName:=path, UpdateLinks:=0, ReadOnly:=True, Local:=True, IgnoreReadOnlyRecommended:=True)
    owned = True
    Application.AutomationSecurity = security: Application.EnableEvents = events
    Exit Function
Failed:
    number = Err.number: description = Err.description
    Application.AutomationSecurity = security: Application.EnableEvents = events
    Err.Raise number, , description
End Function

Public Function HAMU_HasWorkbook() As Boolean
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then Exit Function
    HAMU_HasWorkbook = Not wb.IsAddin
End Function

Public Function HAMU_CanRunCommand(ByVal commandTag As String) As Boolean
    Select Case commandTag
        Case "HAMU_Help", "HAMU_About", "HAMU_UDFHelp"
            HAMU_CanRunCommand = True
        Case Else
            If Left$(commandTag, 4) = "UDF_" Then
                HAMU_CanRunCommand = True
            Else
                HAMU_CanRunCommand = HAMU_HasWorkbook()
            End If
    End Select
End Function

Public Function HAMU_BoundedRange(ByVal source As range, Optional ByVal trimUsedExtent As Boolean = False) As range
 Dim scope As range, lastRow As range, lastColumn As range
 If source Is Nothing Then Exit Function
 Set HAMU_BoundedRange = source
 If source.areas.count <> 1 Then Exit Function
 If Not trimUsedExtent And source.rows.count <> source.Worksheet.rows.count And source.columns.count <> source.Worksheet.columns.count Then Exit Function
 Set scope = Intersect(source, source.Worksheet.UsedRange)
 If scope Is Nothing Then Set HAMU_BoundedRange = source.Cells(1, 1): Exit Function
 Set lastRow = scope.Find(What:="*", After:=scope.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
 If lastRow Is Nothing Then Set HAMU_BoundedRange = source.Cells(1, 1): Exit Function
 Set lastColumn = scope.Find(What:="*", After:=scope.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, MatchByte:=False, SearchFormat:=False)
 Dim bottom As Long, right As Long
 bottom = source.row + source.rows.count - 1: right = source.column + source.columns.count - 1
 If trimUsedExtent Or source.rows.count = source.Worksheet.rows.count Then bottom = lastRow.row
 If trimUsedExtent Or source.columns.count = source.Worksheet.columns.count Then right = lastColumn.column
 Set HAMU_BoundedRange = source.Worksheet.range(source.Cells(1, 1), source.Worksheet.Cells(bottom, right))
 If HAMU_BoundedRange.CountLarge > 500000 Then Err.Raise 5, , HAMU_L("Veri alanı çok büyük. En fazla 500.000 hücre içeren daha küçük bir alan seçin.", "The data range is too large. Select at most 500,000 cells.")
End Function

