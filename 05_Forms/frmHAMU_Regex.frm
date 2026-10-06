VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_Regex 
   Caption         =   "HAMU - Normal İfade"
   ClientHeight    =   7410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11385
   OleObjectBlob   =   "frmHAMU_Regex.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_Regex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMode As String
Private mCatalog As Variant
Private mSource As range
Private mTarget As range
Private mLoading As Boolean
Private mRunning As Boolean, mCancelRequested As Boolean

Public Sub Setup(ByVal modeName As String)

    mLoading = True
    mMode = UCase$(Trim$(modeName))

    If TypeName(Selection) = "Range" Then
        Set mSource = Selection
    Else
        Set mSource = Nothing
    End If

    Set mTarget = Nothing

    txtSource.text = HAMU_FormRangeText(mSource)
    txtTarget.text = ""

    If mMode = "REPLACE" Then
        Me.Caption = HAMU_Text("HAMU - Normal İfade ile Değiştir")
        lblModeTitle.Caption = HAMU_Text("Normal İfade ile Değiştir")

        lblReplacement.visible = True
        txtReplacement.visible = True
        chkGlobal.visible = True

        lblResultMode.visible = False
        cboResultMode.visible = False

        cboOutput.Clear
        cboOutput.AddItem HAMU_Text("Hedef hücre seç")
        cboOutput.AddItem "Yeni sayfaya kopyala"
        cboOutput.AddItem HAMU_Text("Kaynak hücreleri yerinde değiştir")
        cboOutput.ListIndex = 0

        cmdRun.Caption = HAMU_Text("Değiştir")

    Else
        mMode = "FIND"

        Me.Caption = HAMU_Text("HAMU - Normal İfade ile Bul")
        lblModeTitle.Caption = HAMU_Text("Normal İfade ile Bul")

        lblReplacement.visible = False
        txtReplacement.visible = False
        chkGlobal.visible = False

        lblResultMode.visible = True
        cboResultMode.visible = True

        cboResultMode.Clear
        cboResultMode.AddItem HAMU_Text("İlk eşleşme")
        cboResultMode.AddItem HAMU_Text("Tüm eşleşmeler")
        cboResultMode.AddItem HAMU_Text("Eşleşme sayısı")
        cboResultMode.AddItem "Var / Yok"
        cboResultMode.ListIndex = 0

        cboOutput.Clear
        cboOutput.AddItem HAMU_Text("Hedef hücre seç")
        cboOutput.AddItem HAMU_Text("Sağdaki ilk boş sütuna yaz")
        cboOutput.AddItem "Yeni sayfaya yaz"
        cboOutput.ListIndex = 0

        cmdRun.Caption = "Bul"
    End If

    chkHeader.value = True

    mCatalog = HAMU_Regex_GetCatalog(mMode)

    HAMU_LoadCategories

    mLoading = False

    HAMU_UpdateTargetUI

    If cboCategory.ListCount > 0 Then
        cboCategory.ListIndex = 0
    End If

End Sub

Private Sub HAMU_LoadCategories()

    Dim seen As Object
    Dim i As Long
    Dim categoryName As String

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    cboCategory.Clear

    For i = LBound(mCatalog) To UBound(mCatalog)

        categoryName = CStr(mCatalog(i)(0))

        If Not seen.exists(categoryName) Then
            seen.Add categoryName, True
            cboCategory.AddItem categoryName
        End If

    Next i

    cboCategory.AddItem HAMU_Text("Özel Regex")

End Sub

Private Sub cboCategory_Change()

    If mLoading Then Exit Sub

    HAMU_LoadTemplates

End Sub

Private Sub HAMU_LoadTemplates()

    Dim i As Long
    Dim categoryName As String

    mLoading = True

    cboTemplate.Clear

    categoryName = cboCategory.value

    If StrComp(categoryName, HAMU_Text("Özel Regex"), vbTextCompare) = 0 Then

        cboTemplate.AddItem "Kendi ifadem"
        cboTemplate.AddItem "E-posta (yakalama gruplu)"
        cboTemplate.AddItem "URL (protokol / alan / yol)"
        cboTemplate.AddItem "Tarih (yil / ay / gun)"
        cboTemplate.AddItem "Kod ve numara"
        cboTemplate.AddItem "Parantez ici"
        cboTemplate.ListIndex = 0

        mLoading = False
        HAMU_LoadSelectedPreset
        Exit Sub

    End If

    For i = LBound(mCatalog) To UBound(mCatalog)

        If StrComp( _
            CStr(mCatalog(i)(0)), _
            categoryName, _
            vbTextCompare) = 0 Then

            cboTemplate.AddItem CStr(mCatalog(i)(1))

        End If

    Next i

    If cboTemplate.ListCount > 0 Then
        cboTemplate.ListIndex = 0
    End If

    mLoading = False

    HAMU_LoadSelectedPreset

End Sub

Private Sub cboTemplate_Change()

    If mLoading Then Exit Sub

    HAMU_LoadSelectedPreset

End Sub

Private Sub HAMU_LoadSelectedPreset()

    Dim item As Variant

    If Len(cboCategory.value) = 0 Then Exit Sub
    If Len(cboTemplate.value) = 0 Then Exit Sub

    If StrComp(cboCategory.value, HAMU_Text("Özel Regex"), vbTextCompare) = 0 Then

        txtDescription.text = _
            HAMU_Text("Kendi normal ifadenizi yazabilirsiniz. Yakalama grupları için ") & _
            "$1, $2 ... kullanabilirsiniz."

        Select Case cboTemplate.ListIndex
            Case 1: txtPattern.text = "([A-Z0-9._%+-]+)@([A-Z0-9.-]+\.[A-Z]{2,})"
            Case 2: txtPattern.text = "(https?)://([^/\s]+)(/[^\s]*)?"
            Case 3: txtPattern.text = "(\d{4})[./-](\d{1,2})[./-](\d{1,2})"
            Case 4: txtPattern.text = "([A-Z]+)[-_ ]*(\d+)"
            Case 5: txtPattern.text = "\(([^)]*)\)"
            Case Else: txtPattern.text = ""
        End Select

        If mMode = "REPLACE" Then
            txtReplacement.text = "$1"
            chkGlobal.value = True
        End If

        chkIgnoreCase.value = True
        Exit Sub

    End If

    item = HAMU_Regex_GetPreset( _
                mMode, _
                cboCategory.value, _
                cboTemplate.value)

    If IsEmpty(item) Then Exit Sub

    txtPattern.text = CStr(item(2))

    If mMode = "REPLACE" Then

        txtReplacement.text = CStr(item(3))
        chkIgnoreCase.value = CBool(item(4))
        chkGlobal.value = CBool(item(5))
        txtDescription.text = CStr(item(7))

    Else

        chkIgnoreCase.value = CBool(item(3))
        txtDescription.text = CStr(item(5))

    End If

End Sub

Private Sub cboOutput_Change()

    If mLoading Then Exit Sub

    HAMU_UpdateTargetUI

End Sub

Private Sub HAMU_UpdateTargetUI()

    Dim needsTarget As Boolean

    needsTarget = (cboOutput.ListIndex = 0)

    txtTarget.Enabled = needsTarget
    cmdTargetSelect.Enabled = needsTarget

    If Not needsTarget Then

        Set mTarget = Nothing
        txtTarget.text = ""

    End If

End Sub

Private Sub cmdSourceSelect_Click()

    Dim rg As range

    Me.Hide

    On Error Resume Next

    Set rg = HAMU_PromptRange(HAMU_Text("Kaynak veri aralığını seçin."))

    On Error GoTo 0

    If rg Is Nothing Then
        Me.show vbModeless
        Exit Sub
    End If

    Set mSource = rg
    txtSource.text = HAMU_FormRangeText(mSource)
    Me.show vbModeless
End Sub

Private Sub cmdTargetSelect_Click()

    Dim rg As range

    Me.Hide

    On Error Resume Next

    Set rg = HAMU_PromptRange(HAMU_Text("Sonucun başlayacağı ilk hücreyi seçin."))

    On Error GoTo 0

    If rg Is Nothing Then
        Me.show vbModeless
        Exit Sub
    End If

    Set mTarget = rg.Cells(1, 1)
    txtTarget.text = HAMU_FormRangeText(mTarget)
    Me.show vbModeless
End Sub

Private Sub cmdRun_Click()
    If mRunning Then Exit Sub

    Dim statusText As String
    Dim outputIndex As Long
    Dim resultMode As Long

    If mSource Is Nothing Then

        MsgBox _
            HAMU_Text("Önce kaynak veri aralığını seçin."), _
            vbExclamation, _
            "HAMU Tools"

        Exit Sub

    End If

    If Len(Trim$(txtPattern.text)) = 0 Then

        MsgBox _
            HAMU_Text("Normal ifade boş olamaz."), _
            vbExclamation, _
            "HAMU Tools"

        Exit Sub

    End If

    outputIndex = cboOutput.ListIndex

    If outputIndex < 0 Then

        MsgBox _
            HAMU_Text("Çıktı yöntemini seçin."), _
            vbExclamation, _
            "HAMU Tools"

        Exit Sub

    End If

    If outputIndex = 0 And mTarget Is Nothing Then

        MsgBox _
            HAMU_Text("Çıktı için bir hedef hücre seçin."), _
            vbExclamation, _
            "HAMU Tools"

        Exit Sub

    End If

    On Error GoTo RunError

    Dim toolAccepted As Boolean
    HAMU_BeginToolCommand "HAMU_Regex", toolAccepted
    If Not toolAccepted Then Exit Sub
    mRunning = True: mCancelRequested = False
    cmdRun.Enabled = False: cmdSourceSelect.Enabled = False: cmdTargetSelect.Enabled = False
    lblStatus.Caption = HAMU_Text("İşlem yapılıyor...")
    DoEvents

    If mMode = "REPLACE" Then

        statusText = HAMU_Regex_RunReplaceUI( _
                        mSource, _
                        mTarget, _
                        outputIndex, _
                        txtPattern.text, _
                        txtReplacement.text, _
                        CBool(chkIgnoreCase.value), _
                        CBool(chkGlobal.value), _
                        CBool(chkHeader.value), Me)

    Else

        If mSource.columns.count <> 1 Then

            MsgBox _
                HAMU_Text("Normal İfade ile Bul için kaynak seçim tek sütun olmalıdır."), _
                vbExclamation, _
                "HAMU Tools"

            lblStatus.Caption = ""
            FinishRun
            Exit Sub

        End If

        resultMode = cboResultMode.ListIndex + 1

        statusText = HAMU_Regex_RunFindUI( _
                        mSource, _
                        mTarget, _
                        outputIndex, _
                        resultMode, _
                        txtPattern.text, _
                        CBool(chkIgnoreCase.value), _
                        CBool(chkHeader.value), _
                        cboTemplate.value, Me)

    End If

    lblStatus.Caption = statusText
    FinishRun
    Exit Sub

RunError:
    Dim errorNumber As Long, errorText As String
    errorNumber = Err.number: errorText = Err.description
    FinishRun
    lblStatus.Caption = errorText
    If errorNumber <> 18 Then MsgBox errorText, vbExclamation, "HAMU Tools"

End Sub

Private Sub cmdCancel_Click()
    If mRunning Then mCancelRequested = True: Exit Sub
    Unload Me

End Sub

Private Function HAMU_FormRangeText(ByVal rg As range) As String

    If rg Is Nothing Then
        HAMU_FormRangeText = ""
        Exit Function
    End If

    HAMU_FormRangeText = _
        "'" & rg.Worksheet.name & "'!" & _
        rg.address(False, False)

End Function

Private Sub UserForm_Initialize()
    HAMU_ThemeForm Me
End Sub


























































Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub













































Public Property Get IsRunning() As Boolean
 IsRunning = mRunning
End Property
Public Property Get CancelRequested() As Boolean
 CancelRequested = mCancelRequested
End Property
Private Sub FinishRun()
 If mRunning Then HAMU_EndToolCommand
 mRunning = False
 cmdRun.Enabled = True: cmdSourceSelect.Enabled = True: cmdTargetSelect.Enabled = True
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 If mRunning Then
  Cancel = True: mCancelRequested = True
 Else
  Set mSource = Nothing: Set mTarget = Nothing: mCatalog = Empty
 End If
End Sub


