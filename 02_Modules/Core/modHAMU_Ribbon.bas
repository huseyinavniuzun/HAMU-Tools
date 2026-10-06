Attribute VB_Name = "modHAMU_Ribbon"
Option Explicit

Public gRibbon As IRibbonUI
Private mCommandBusy As Boolean
Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    Set gRibbon = ribbon
    HAMU_LoadSettings
End Sub
Public Sub HAMU_OnAction(ByVal control As IRibbonControl)
    On Error GoTo Failed
    HAMU_RunCommand control.Tag
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_OnAction", Err.number, Err.description
End Sub



Public Sub HAMU_RunCommand(ByVal commandTag As String)
    Dim state As clsHAMU_AppState
    If mCommandBusy Then Exit Sub
    On Error GoTo Failed
    If Not HAMU_CanRunCommand(commandTag) Then
        HAMU_ShowInfo HAMU_Text("Bu işlem için önce bir çalışma kitabı açın veya Ctrl+N ile yeni bir kitap oluşturun.")
        Exit Sub
    End If
    If HAMU_NeedsWorksheet(commandTag) And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komutu kullanmak için bir çalışma sayfasına geçin."): Exit Sub
    End If
    HAMU_PruneHighlightCache
    mCommandBusy = True
    gHAMU_CommandId = commandTag
    Set state = New clsHAMU_AppState
    state.Capture
    If Left$(commandTag, 4) = "UDF_" Then HAMU_RegisterFunctions
    Select Case commandTag
        Case "HAMU_DataEntry": HAMU_DataEntry
        Case "HAMU_ClearHighlights": HAMU_ClearHighlights
        Case "HAMU_FillBlanksZero": HAMU_FillBlanksZero
        Case "HAMU_CountByColor": HAMU_CountByColor
        Case "HAMU_SumByColor": HAMU_SumByColor
        Case "HAMU_Transfer": HAMU_Transfer
        Case "HAMU_SaveCopy": HAMU_SaveCopy
        Case "HAMU_Backup": HAMU_Backup
        Case "HAMU_AddDateToFilename": HAMU_AddDateToFilename
        Case "HAMU_ExportSheetsAsPDF": HAMU_ExportSheetsAsPDF
        Case "HAMU_FolderIndex": HAMU_FolderIndex
        Case "HAMU_CombineCSVs": HAMU_CombineCSVs
        Case "HAMU_ConsolidateSheets": HAMU_ConsolidateSheets
        Case "HAMU_AppendSelectedTables": HAMU_AppendSelectedTables
        Case "HAMU_SmartAppendByHeaders": HAMU_SmartAppendByHeaders
        Case "HAMU_CollectFilesData": HAMU_CollectFilesData
        Case "HAMU_SmartJoin": HAMU_SmartJoin
        Case "HAMU_GroupAndSummarize": HAMU_GroupAndSummarize
        Case "HAMU_UnpivotColumnsToRows": HAMU_UnpivotColumnsToRows
        Case "HAMU_SplitDataWizard": HAMU_SplitDataWizard
        Case "HAMU_SplitByValueToFiles": HAMU_SplitByValueToFiles
        Case "HAMU_SplitURLParameters": HAMU_SplitURLParameters
        Case "HAMU_PivotExport": HAMU_PivotExport
        Case "HAMU_PivotRowsToColumns": HAMU_PivotRowsToColumns
        Case "HAMU_FillDownBlanks": HAMU_FillDownBlanks
        Case "HAMU_MergeDuplicates": HAMU_MergeDuplicates
        Case "HAMU_StandardizeHeaders": HAMU_StandardizeHeaders
        Case "HAMU_ConvertColumnTypes": HAMU_ConvertColumnTypes
        Case "HAMU_SelectColumnsToNewTable": HAMU_SelectColumnsToNewTable
        Case "HAMU_FilterRowsToNewTable": HAMU_FilterRowsToNewTable
        Case "HAMU_SampleData": HAMU_SampleData
        Case "HAMU_ShuffleRows": HAMU_ShuffleRows
        Case "HAMU_CompareTwoLists": HAMU_CompareTwoLists
        Case "HAMU_ExtractUniqueList": HAMU_ExtractUniqueList
        Case "HAMU_DataProfile": HAMU_DataProfile
        Case "HAMU_DataQualityReport": HAMU_DataQualityReport
        Case "HAMU_FindMissingCombinations": HAMU_FindMissingCombinations
        Case "HAMU_CleanData": HAMU_CleanData
        Case "HAMU_TextCase": HAMU_TextCase
        Case "HAMU_PatternFind": HAMU_PatternFind
        Case "HAMU_RegexReplace": HAMU_RegexReplace
        Case "HAMU_ExtractDigits": HAMU_ExtractDigits
        Case "HAMU_ConvertQuotedNumbers": HAMU_ConvertQuotedNumbers
        Case "HAMU_MaskData": HAMU_MaskData
        Case "HAMU_ConvertFormulasToValues": HAMU_ConvertFormulasToValues
        Case "HAMU_UnmergeAndFill": HAMU_UnmergeAndFill
        Case "HAMU_StandardFormat": HAMU_StandardFormat
        Case "HAMU_DeleteBlankCells": HAMU_DeleteBlankCells
        Case "HAMU_DeleteDropdowns": HAMU_DeleteDropdowns
        Case "HAMU_RowStriping": HAMU_RowStriping
        Case "HAMU_DatePicker": HAMU_DatePicker
        Case "HAMU_FixDates": HAMU_FixDates
        Case "HAMU_CheckDates": HAMU_CheckDates
        Case "HAMU_DateBuilder": HAMU_DateBuilder
        Case "HAMU_SheetIndex": HAMU_SheetIndex
        Case "HAMU_SortSheets": HAMU_SortSheets
        Case "HAMU_ShowHidden": HAMU_ShowHidden
        Case "HAMU_DynamicValidation": HAMU_DynamicValidation
        Case "HAMU_SelectionToPNG": HAMU_SelectionToPNG
        Case "HAMU_InsertPhotosFromFolder": HAMU_InsertPhotosFromFolder
        Case "HAMU_CreateBarcode": HAMU_CreateBarcode
        Case "HAMU_CreateQRCode": HAMU_CreateQRCode
        Case "HAMU_CompareSheets": HAMU_CompareSheets
        Case "HAMU_MarkProblemCells": HAMU_MarkProblemCells
        Case "HAMU_HighlightByCriteria": HAMU_HighlightByCriteria
        Case "HAMU_HighlightDuplicates": HAMU_HighlightDuplicates
        Case "HAMU_AddDuplicateCount": HAMU_AddDuplicateCount
        Case "HAMU_ListLinks": HAMU_ListLinks
        Case "HAMU_BreakLinks": HAMU_BreakLinks
        Case "HAMU_Help": HAMU_Help
        Case "HAMU_About": HAMU_About
        Case "HAMU_ExpandColumns": HAMU_ExpandColumns
        Case "HAMU_TableTotals": HAMU_TableTotals
        Case "HAMU_CopyVisibleTable": HAMU_CopyVisibleTable
        Case "HAMU_AutoFit": HAMU_AutoFit
        Case "HAMU_AutoTable": HAMU_AutoTable
        Case "HAMU_ResizeTable": HAMU_ResizeTable
        Case "HAMU_AddTextAffixes": HAMU_AddTextAffixes
        Case "HAMU_LeadingZeros": HAMU_LeadingZeros
        Case "HAMU_DeduplicateItems": HAMU_DeduplicateItems
        Case "HAMU_ExtractTextPart": HAMU_ExtractTextPart
        Case "HAMU_DateRangeList": HAMU_DateRangeList
        Case "HAMU_FormulaAudit": HAMU_FormulaAudit
        Case "HAMU_SelectSpecialCells": HAMU_SelectSpecialCells
        Case "HAMU_ListDefinedNames": HAMU_ListDefinedNames
        Case "HAMU_UDFHelp": HAMU_UDFHelp

        Case "UDF_URLDecode": HAMU_ShowInfo HAMU_CommandIntro("UDF_URLDecode")
        Case "UDF_RenklileriTopla": HAMU_ShowInfo HAMU_CommandIntro("UDF_RenklileriTopla")
        Case "UDF_RengeGoreSay": HAMU_ShowInfo HAMU_CommandIntro("UDF_RengeGoreSay")
        Case "UDF_YerelDosyaYolu": HAMU_ShowInfo HAMU_CommandIntro("UDF_YerelDosyaYolu")
        Case "UDF_GizliLink": HAMU_ShowInfo HAMU_CommandIntro("UDF_GizliLink")
        Case "UDF_LinkParamDegeri": HAMU_ShowInfo HAMU_CommandIntro("UDF_LinkParamDegeri")
        Case "UDF_SayidanMetine": HAMU_ShowInfo HAMU_CommandIntro("UDF_SayidanMetine")
        Case "UDF_TablodanVeriGetir": HAMU_ShowInfo HAMU_CommandIntro("UDF_TablodanVeriGetir")
        Case "UDF_LambdaMultiLookup": HAMU_ShowInfo HAMU_CommandIntro("UDF_LambdaMultiLookup")
        Case "UDF_LambdaMultiSum": HAMU_ShowInfo HAMU_CommandIntro("UDF_LambdaMultiSum")
        Case "UDF_LambdaUniqueMerge": HAMU_ShowInfo HAMU_CommandIntro("UDF_LambdaUniqueMerge")
        Case Else: Err.Raise vbObjectError + 340, , HAMU_Text("Tanımsız HAMU komutu: ") & commandTag
    End Select
    state.Restore
    mCommandBusy = False
    gHAMU_CommandId = ""
    Exit Sub
Failed:
    Dim errorNumber As Long, errorText As String
    errorNumber = Err.number: errorText = Err.description
    On Error Resume Next
    If Not state Is Nothing Then state.Restore
    On Error GoTo 0
    mCommandBusy = False
    gHAMU_CommandId = ""
    HAMU_ShowError commandTag, errorNumber, errorText
End Sub
Public Sub HAMU_About()
    frmHAMU_About.show
End Sub
Public Sub HAMU_Help()
    HAMU_OpenGuide "start"
End Sub
Public Sub HAMU_UDFHelp()
 HAMU_RegisterFunctions
 HAMU_OpenGuide "functions"
End Sub

Public Sub HAMU_OnGuide(control As IRibbonControl)
 HAMU_OpenGuide control.Tag
End Sub

Public Sub HAMU_OnSettings(control As IRibbonControl)
 HAMU_ShowSettings
End Sub
Public Sub HAMU_GetLanguageIndex(control As IRibbonControl, ByRef returnedVal)
 returnedVal = IIf(HAMU_Setting("Language") = "en", 1, 0)
End Sub
Public Sub HAMU_OnLanguage(control As IRibbonControl, ByVal selectedID As String, ByVal selectedIndex As Integer)
 On Error GoTo Failed
 HAMU_SetSetting "Language", IIf(selectedIndex = 1, "en", "tr")
 HAMU_SaveSettings
 HAMU_RefreshRibbon
 Exit Sub
Failed:
 HAMU_ReloadSettings
 HAMU_RefreshRibbon
 HAMU_ShowInfo HAMU_L("Ayarlar kaydedilemedi.", "Settings could not be saved.")
End Sub
Public Sub HAMU_GetSettingsSummary(control As IRibbonControl, ByRef returnedVal)
 returnedVal = HAMU_L(HAMU_Text("Yedek: "), "Backup: ") & HAMU_BackupFolder()
End Sub
Public Sub HAMU_RefreshRibbon()
 On Error Resume Next
 If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

Public Sub HAMU_GetLabel(control As IRibbonControl, ByRef returnedVal)
 returnedVal = HAMU_RibbonText(control.id, control.Tag, "label")
End Sub

Public Sub HAMU_GetDescription(control As IRibbonControl, ByRef returnedVal)
 returnedVal = HAMU_RibbonText(control.id, control.Tag, "description")
End Sub

Public Sub HAMU_GetHelperText(control As IRibbonControl, ByRef returnedVal)
 returnedVal = HAMU_RibbonText(control.id, control.Tag, "helperText")
End Sub

Public Sub HAMU_GetScreentip(control As IRibbonControl, ByRef returnedVal)
 returnedVal = HAMU_RibbonText(control.id, control.Tag, "screentip")
End Sub

Public Sub HAMU_GetSupertip(control As IRibbonControl, ByRef returnedVal)
 returnedVal = HAMU_RibbonText(control.id, control.Tag, "supertip")
End Sub

Public Sub HAMU_GetAltText(control As IRibbonControl, ByRef returnedVal)
 returnedVal = HAMU_RibbonText(control.id, control.Tag, "altText")
End Sub

Public Sub HAMU_OnLambdaTemplate(control As IRibbonControl)
 On Error GoTo Failed
 Dim path As String, wb As Workbook, owned As Boolean
 path = CreateObject("Scripting.FileSystemObject").GetAbsolutePathName(ThisWorkbook.path & "\..\09_Templates\HAMU_LAMBDA_Sablonu.xlsx")
 Set wb = HAMU_OpenSourceWorkbook(path, owned)
 wb.Activate
 Exit Sub
Failed:
 HAMU_ShowError "LAMBDA", Err.number, Err.description
End Sub


Private Function HAMU_RibbonText(ByVal controlId As String, ByVal controlTag As String, ByVal field As String) As String
    If Left$(controlId, 3) = "en_" Then controlId = Mid$(controlId, 4)
    ' Menus retain their lookup key when Office creates a Quick Access Toolbar copy.
    If Left$(controlTag, 3) = "UI:" Then controlId = Mid$(controlTag, 4)
    HAMU_RibbonText = HAMU_UIText(controlId, field)
    If Len(HAMU_RibbonText) > 0 Or Len(controlTag) = 0 Or Left$(controlTag, 3) = "UI:" Then Exit Function
    HAMU_RibbonText = HAMU_UIText("r_" & controlTag, field)
    If Len(HAMU_RibbonText) > 0 Then Exit Function
    Select Case field
        Case "label", "screentip": HAMU_RibbonText = HAMU_CommandLabel(controlTag)
        Case "supertip", "description": HAMU_RibbonText = HAMU_CommandIntro(controlTag)
    End Select
End Function

Public Sub HAMU_IsTurkish(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
 returnedVal = (HAMU_Setting("Language") <> "en")
End Sub
Public Sub HAMU_IsEnglish(ByVal control As IRibbonControl, ByRef returnedVal As Variant)
 returnedVal = (HAMU_Setting("Language") = "en")
End Sub

Public Sub HAMU_BeginToolCommand(ByVal id As String, ByRef accepted As Boolean)
 accepted = False
 If mCommandBusy Then Exit Sub
 HAMU_PruneHighlightCache
 mCommandBusy = True: gHAMU_CommandId = id
 accepted = True
End Sub
Public Sub HAMU_EndToolCommand()
 mCommandBusy = False: gHAMU_CommandId = ""
End Sub
