Attribute VB_Name = "modHAMU_Transfer"
Option Explicit
Option Private Module

Private mTemplate As Workbook, mTarget As Workbook
Private mOwned As Boolean, mItems As Variant, mPrefix As String
Public Function Pref_GetTemplatePath() As String
    Pref_GetTemplatePath = GetSetting("HAMU Tools", "Transfer", "Template", "")
End Function
Public Sub Pref_SetTemplatePath(ByVal value As String)
    SaveSetting "HAMU Tools", "Transfer", "Template", value
End Sub
Public Function DosyaSec(ByVal titleText As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = titleText
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add HAMU_Text("Excel çalışma kitapları"), "*.xlsx;*.xlsm;*.xlsb;*.xls;*.xlam"
        If .show = -1 Then DosyaSec = .SelectedItems(1)
    End With
End Function
Public Sub TM_OnNext(ByVal path As String, ByVal prefix As String, ByVal useLambda As Boolean, ByVal useSheets As Boolean, ByVal useNames As Boolean, ByVal useTables As Boolean, ByVal useStyles As Boolean)
    Dim n As name, st As Style, rows As Collection, item As Variant, sheet As Worksheet, table As ListObject, scope As range
    Dim i As Long, exists() As Boolean, priorSecurity As Long, wb As Workbook
    On Error GoTo Failed
    Set mTarget = ActiveWorkbook
    mPrefix = prefix
    Set mTemplate = HAMU_OpenSourceWorkbook(path, mOwned)
    If mTemplate Is mTarget Then Err.Raise vbObjectError + 351, , HAMU_Text("Kaynak şablon ile hedef kitap aynı olamaz.")
    Set rows = New Collection
    If useLambda Then
        For Each n In mTemplate.names
            If InStr(1, n.RefersTo, "LAMBDA(", vbTextCompare) > 0 And InStr(n.name, "!") = 0 Then rows.Add Array("LAMBDA", n.name)
        Next
    End If
    If useStyles Then
        For Each st In mTemplate.Styles
            If Not st.BuiltIn Then rows.Add Array("STYLE", st.name)
        Next
    End If
    If useNames Then
        For Each n In mTemplate.names
            Set scope = Nothing
            On Error Resume Next: Set scope = n.RefersToRange: On Error GoTo Failed
            If n.visible And Not scope Is Nothing And InStr(n.name, "!") = 0 Then rows.Add Array("NAME", n.name)
        Next
    End If
    For Each sheet In mTemplate.Worksheets
        If useSheets And sheet.visible = xlSheetVisible Then rows.Add Array("SHEET", sheet.name)
        If useTables Then
            For Each table In sheet.ListObjects
                rows.Add Array("TABLE", sheet.name & "!" & table.name)
            Next
        End If
    Next
    If rows.count = 0 Then
        HAMU_ShowInfo HAMU_Text("Şablonda aktarılabilecek kitap kapsamlı LAMBDA veya özel stil bulunamadı.")
        IP_OnClose
        Exit Sub
    End If
    ReDim mItems(1 To rows.count, 1 To 2)
    ReDim exists(1 To rows.count)
    For i = 1 To rows.count
        item = rows(i)
        mItems(i, 1) = item(0): mItems(i, 2) = item(1)
        exists(i) = HAMU_TransferExists(CStr(item(0)), HAMU_TransferItemName(CStr(item(0)), prefix, CStr(item(1))))
    Next
    mTarget.Activate
    frmHAMU_ItemPicker.LoadData path, mItems, exists
    frmHAMU_ItemPicker.show
    Exit Sub
Failed:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    If priorSecurity <> 0 Then Application.AutomationSecurity = priorSecurity
    IP_OnClose
    HAMU_ShowError "HAMU_Transfer", hamuErrorNumber, hamuErrorText
End Sub
Private Function HAMU_TransferExists(ByVal kind As String, ByVal nm As String) As Boolean
    Dim obj As Object
    On Error Resume Next
    Select Case kind
    Case "LAMBDA", "NAME": Set obj = mTarget.names(nm)
    Case "STYLE": Set obj = mTarget.Styles(nm)
    Case "SHEET", "TABLE": Set obj = mTarget.Worksheets(nm)
    End Select
    HAMU_TransferExists = Not obj Is Nothing
End Function
Public Sub IP_OnTransfer(ByVal policy As String, ByVal suffix As String, ByRef selected() As Long)
    Dim i As Long, index As Long, nm As String, kind As String, n As Long
    Dim source As Style, target As Style, number As Long, base As String
    On Error GoTo Failed
    For i = LBound(selected) To UBound(selected)
        index = selected(i): kind = CStr(mItems(index, 1))
        nm = HAMU_TransferItemName(kind, mPrefix, CStr(mItems(index, 2)))
        If HAMU_TransferExists(kind, nm) Then
            If policy = HAMU_Text("Atla") Then GoTo NextItem
            If policy = "Yeniden" Then
                base = nm: number = 1
                If Len(suffix) = 0 Then suffix = "_2"
                nm = IIf(kind = "SHEET" Or kind = "TABLE", Left$(base, 24), base) & suffix
                Do While HAMU_TransferExists(kind, nm)
                    number = number + 1: nm = IIf(kind = "SHEET" Or kind = "TABLE", Left$(base, 20), base) & suffix & "_" & CStr(number)
                Loop
            Else
                Select Case kind
                Case "LAMBDA", "NAME": ' Names.Add replaces the definition after the new value is ready.
                Case "STYLE": ' Update the existing style in place.
                Case Else
                    Err.Raise 5, , HAMU_L("Mevcut sayfa silinmez. Atla veya Yeniden Adlandır seçin.", "Existing sheets are not deleted. Choose Skip or Rename.")
                End Select
            End If
        End If
        If kind = "NAME" Or kind = "SHEET" Or kind = "TABLE" Then
            HAMU_TransferDataItem kind, CStr(mItems(index, 2)), nm
        ElseIf kind = "LAMBDA" Then
            mTarget.names.Add name:=nm, RefersTo:=mTemplate.names(CStr(mItems(index, 2))).RefersTo
        Else
            Set source = mTemplate.Styles(CStr(mItems(index, 2)))
            Set target = Nothing
            On Error Resume Next: Set target = mTarget.Styles(nm): On Error GoTo Failed
            If target Is Nothing Then Set target = mTarget.Styles.Add(nm)
            target.IncludeNumber = source.IncludeNumber
            target.IncludeFont = source.IncludeFont
            target.IncludeAlignment = source.IncludeAlignment
            target.IncludeBorder = source.IncludeBorder
            target.IncludePatterns = source.IncludePatterns
            target.IncludeProtection = source.IncludeProtection
            target.NumberFormat = source.NumberFormat
            target.Font.name = source.Font.name
            target.Font.Size = source.Font.Size
            target.Font.Bold = source.Font.Bold
            target.Font.Italic = source.Font.Italic
            target.Font.color = source.Font.color
            target.Interior.color = source.Interior.color
            target.Interior.pattern = source.Interior.pattern
            target.HorizontalAlignment = source.HorizontalAlignment
            target.VerticalAlignment = source.VerticalAlignment
            target.WrapText = source.WrapText
            target.Locked = source.Locked
            target.FormulaHidden = source.FormulaHidden
            Dim edge As Variant
            For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight)
                target.Borders(edge).LineStyle = source.Borders(edge).LineStyle
                target.Borders(edge).color = source.Borders(edge).color
                target.Borders(edge).Weight = source.Borders(edge).Weight
            Next
        End If
        n = n + 1
NextItem:
    Next
    HAMU_ShowInfo CStr(n) & HAMU_Text(" öğe aktarıldı. Yeniden adlandırılan LAMBDA bağımlılıklarını hedef kitapta kontrol edin.")
    IP_OnClose
    Exit Sub
Failed:
    Dim hamuErrorNumber As Long, hamuErrorText As String
    hamuErrorNumber = Err.number: hamuErrorText = Err.description
    IP_OnClose
    HAMU_ShowError "HAMU_Transfer", hamuErrorNumber, hamuErrorText
End Sub
Public Sub IP_OnClose()
    On Error Resume Next
    If mOwned And Not mTemplate Is Nothing Then mTemplate.Close False
    Set mTemplate = Nothing
    Set mTarget = Nothing
    mOwned = False
    mItems = Empty: mPrefix = ""
End Sub



Public Function HAMU_TransferItemName(ByVal kind As String, ByVal prefix As String, ByVal item As String) As String
 If kind = "TABLE" Then item = Mid$(item, InStr(item, "!") + 1)
 HAMU_TransferItemName = prefix & item
 If kind = "SHEET" Or kind = "TABLE" Then
  Dim bad As Variant
  For Each bad In Array("/", "\", "?", "*", "[", "]", ":")
   HAMU_TransferItemName = Replace$(HAMU_TransferItemName, CStr(bad), "_")
  Next
  HAMU_TransferItemName = Left$(HAMU_TransferItemName, 31)
 End If
End Function
Private Sub HAMU_TransferDataItem(ByVal kind As String, ByVal item As String, ByVal name As String)
 Dim source As range, target As Worksheet, table As ListObject, key As String, created As Boolean
 On Error GoTo Failed
 Select Case kind
 Case "NAME": Set source = mTemplate.names(item).RefersToRange
 Case "SHEET": Set source = HAMU_BoundedRange(mTemplate.Worksheets(item).Cells)
 Case "TABLE"
  Set table = mTemplate.Worksheets(Left$(item, InStr(item, "!") - 1)).ListObjects(Mid$(item, InStr(item, "!") + 1))
  Set source = table.HeaderRowRange.Resize(table.ListRows.count + 1, table.ListColumns.count)
 End Select
 If source.areas.count <> 1 Or source.CountLarge > 500000 Then Err.Raise 5, , HAMU_L("Aktarılan veri tek alan ve en fazla 500.000 hücre olmalı.", "Transferred data must be one area of at most 500,000 cells.")
 If kind = "NAME" Then
  Set target = HAMU_GetOrCreateSheet("HAMU_" & Left$(name, 24), mTarget)
 Else
  Set target = mTarget.Worksheets.Add(After:=mTarget.Worksheets(mTarget.Worksheets.count))
  target.name = name
 End If
 created = True
 source.copy
 target.range("A1").PasteSpecial xlPasteValuesAndNumberFormats
 target.range("A1").PasteSpecial xlPasteFormats
 target.range("A1").PasteSpecial xlPasteColumnWidths
 Application.CutCopyMode = False
 If kind = "NAME" Then mTarget.names.Add name:=name, RefersTo:="=" & target.range("A1").Resize(source.rows.count, source.columns.count).address(External:=True)
 If kind = "TABLE" Then
  Set table = target.ListObjects.Add(xlSrcRange, target.range("A1").Resize(source.rows.count, source.columns.count), , xlYes)
  table.TableStyle = "TableStyleMedium1"
 End If
 Exit Sub
Failed:
 Dim number As Long, description As String, alerts As Boolean
 number = Err.number: description = Err.description
 On Error Resume Next
 alerts = Application.DisplayAlerts: Application.DisplayAlerts = False
 If created Then target.Delete
 Application.DisplayAlerts = alerts: Application.CutCopyMode = False
 On Error GoTo 0
 Err.Raise number, , description
End Sub


