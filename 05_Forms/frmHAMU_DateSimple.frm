VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_DateSimple 
   Caption         =   "HAMU | Tarih Seç"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4110
   OleObjectBlob   =   "frmHAMU_DateSimple.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_DateSimple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public result As Date
Private mMonth As Date
Private mHandlers As Collection
Private mUpdating As Boolean
Public Sub Setup(ByVal initial As Date)
    If initial = 0 Then initial = Date
    result = initial
    mMonth = DateSerial(year(initial), month(initial), 1)
    RenderDays
End Sub
Private Sub RenderDays()
    Dim i As Long, day As Date, Button As MSForms.label, handler As clsHAMU_DayButton
    mUpdating = True
    cboMonth.ListIndex = month(mMonth) - 1
    cboYear.value = CStr(year(mMonth))
    mUpdating = False
    lblDate.Caption = Format$(result, "dd.mm.yyyy")
    If mHandlers Is Nothing Then Set mHandlers = New Collection
    lblMonth.Caption = cboMonth.List(cboMonth.ListIndex) & " " & CStr(year(mMonth))
    lblSelected.Caption = Format$(result, "dd.mm.yyyy") & vbCrLf & HAMU_DateDistance(result)
    day = mMonth - Weekday(mMonth, vbMonday) + 1
    For i = 1 To 42
        Set Button = Me.controls("day" & i)
        Button.Caption = CStr(VBA.day(day))
        Button.Enabled = True
        Button.ControlTipText = Format$(day, "dd.mm.yyyy")
        Button.ForeColor = RGB(48, 57, 74)
        If month(day) <> month(mMonth) Then Button.ForeColor = RGB(145, 152, 160)
        Button.BackColor = RGB(255, 255, 255)
        If day = result Then
            Button.BackColor = RGB(30, 135, 170): Button.ForeColor = vbWhite
        End If
        If mHandlers.count < i Then
            Set handler = New clsHAMU_DayButton
            Set handler.Button = Button
            Set handler.owner = Me
            mHandlers.Add handler
        Else
            Set handler = mHandlers(i)
        End If
        handler.DateValue = day
        day = day + 1
    Next
End Sub
Public Sub SelectDate(ByVal value As Date)
    result = value
    RenderDays
End Sub
Public Function GetDate(Optional ByVal initial As Date = 0) As Date
    Setup initial
    Me.show vbModal
    GetDate = result
    ReleaseHandlers
    Unload Me
End Function
Private Sub cmdPrev_Click()
    mMonth = DateAdd("m", -1, mMonth): RenderDays
End Sub
Private Sub cmdNext_Click()
    mMonth = DateAdd("m", 1, mMonth): RenderDays
End Sub
Private Sub cmdToday_Click()
    result = Date: mMonth = DateSerial(year(Date), month(Date), 1): RenderDays
End Sub
Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> 0 Then ReleaseHandlers
    If CloseMode = 0 Then
        Cancel = True: result = 0: ReleaseHandlers: Me.Hide
    End If
End Sub
Private Sub UserForm_Initialize()
    Me.Caption = HAMU_Text("HAMU | Tarih Seç")
    mUpdating = True
    Dim months As Variant, monthName As Variant, year As Long
    months = Array(HAMU_Text("Ocak"), HAMU_Text("Şubat"), HAMU_Text("Mart"), HAMU_Text("Nisan"), HAMU_Text("Mayıs"), HAMU_Text("Haziran"), HAMU_Text("Temmuz"), HAMU_Text("Ağustos"), HAMU_Text("Eylül"), HAMU_Text("Ekim"), HAMU_Text("Kasım"), HAMU_L("Aralık", "December"))
    For Each monthName In months: cboMonth.AddItem monthName: Next
    For year = 1900 To 2100: cboYear.AddItem CStr(year): Next
    HAMU_ThemeForm Me
    lblDate.Font.Size = 14: lblDate.ForeColor = RGB(22, 109, 166)
    mUpdating = False
End Sub

Private Sub cboMonth_Change()
    If mUpdating Or cboMonth.ListIndex < 0 Then Exit Sub
    ChangeMonthYear
End Sub
Private Sub cboYear_Change()
    If mUpdating Then Exit Sub
    ' Wait until a typed year is complete; dropdown selections update immediately.
    If Len(Trim$(CStr(cboYear.value))) <> 4 Then Exit Sub
    If Not IsNumeric(cboYear.value) Then Exit Sub
    ChangeMonthYear
End Sub
Private Sub cboYear_AfterUpdate()
    If mUpdating Then Exit Sub
    ChangeMonthYear
End Sub
Private Sub ChangeMonthYear()
    Dim year As Long, month As Long, day As Long, maxDay As Long
    If Not IsNumeric(cboYear.value) Then
        lblHint.Caption = HAMU_Text("Yılı sayı olarak girin."): Exit Sub
    End If
    year = CLng(cboYear.value)
    If year < 100 Or year > 9999 Then
        lblHint.Caption = HAMU_Text("100–9999 arasında bir yıl girin."): Exit Sub
    End If
    month = cboMonth.ListIndex + 1
    If month < 1 Then Exit Sub
    day = VBA.day(result)
    If year = 9999 And month = 12 Then maxDay = 31 Else maxDay = VBA.day(DateSerial(year, month + 1, 0))
    If day > maxDay Then day = maxDay
    result = DateSerial(year, month, day)
    mMonth = DateSerial(year, month, 1)
    lblHint.Caption = "": RenderDays
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub




Private Sub ReleaseHandlers()
 Dim item As clsHAMU_DayButton
 If mHandlers Is Nothing Then Exit Sub
 For Each item In mHandlers
  Set item.Button = Nothing: Set item.owner = Nothing
 Next
 Set mHandlers = Nothing
End Sub

