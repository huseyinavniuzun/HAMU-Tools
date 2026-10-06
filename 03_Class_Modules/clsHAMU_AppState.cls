VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHAMU_AppState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mCaptured As Boolean
Private mSecurity As Long
Private mScreen As Boolean, mEvents As Boolean, mAlerts As Boolean
Private mCalc As Variant, mStatus As Variant
Private mHasCalculation As Boolean
Public Sub Capture()
    mSecurity = Application.AutomationSecurity
    mScreen = Application.ScreenUpdating
    mEvents = Application.EnableEvents
    mAlerts = Application.DisplayAlerts
    mHasCalculation = False
    If Not Application.ActiveWorkbook Is Nothing Then
        mCalc = Application.Calculation
        If Not IsError(mCalc) Then
            If IsNumeric(mCalc) Then
                Select Case CLng(mCalc)
                    Case xlCalculationAutomatic, xlCalculationManual, xlCalculationSemiautomatic
                        mHasCalculation = True
                End Select
            End If
        End If
    End If
    Dim statusOwner As Object
    Set statusOwner = Application
    mStatus = statusOwner.StatusBar
    mCaptured = True
End Sub
Public Sub Restore()
    Dim statusOwner As Object
    Set statusOwner = Application
    If Not mCaptured Then Exit Sub
    On Error Resume Next
    If Application.AutomationSecurity <> mSecurity Then Application.AutomationSecurity = mSecurity
    If Application.ScreenUpdating <> mScreen Then Application.ScreenUpdating = mScreen
    If Application.EnableEvents <> mEvents Then Application.EnableEvents = mEvents
    If Application.DisplayAlerts <> mAlerts Then Application.DisplayAlerts = mAlerts
    If mHasCalculation And Not Application.ActiveWorkbook Is Nothing Then
        If Application.Calculation <> CLng(mCalc) Then Application.Calculation = CLng(mCalc)
    End If
    Dim currentStatus As Variant
    currentStatus = statusOwner.StatusBar
    If VarType(mStatus) = vbBoolean Or IsEmpty(mStatus) Or IsNull(mStatus) Then
        If VarType(currentStatus) <> vbBoolean Then ResetStatusBar statusOwner
    ElseIf LCase$(Trim$(CStr(mStatus))) = "false" Then
        If VarType(currentStatus) <> vbBoolean Then ResetStatusBar statusOwner
    ElseIf VarType(currentStatus) <> VarType(mStatus) Then
        Application.StatusBar = mStatus
    ElseIf CStr(currentStatus) <> CStr(mStatus) Then
        Application.StatusBar = mStatus
    End If
    mCaptured = False
End Sub
Private Sub Class_Terminate()
    On Error Resume Next
    Restore
End Sub


Public Property Get EventsBefore() As Boolean
 EventsBefore = mEvents
End Property
Public Property Get ScreenBefore() As Boolean
 ScreenBefore = mScreen
End Property

Private Sub ResetStatusBar(ByVal statusOwner As Object)
 Dim reset As Variant, actual As Variant
 reset = False
 CallByName statusOwner, "StatusBar", VbLet, reset
 actual = CallByName(statusOwner, "StatusBar", VbGet)
 If VarType(actual) = vbString Then
  If LCase$(CStr(actual)) = "false" Then statusOwner.StatusBar = Empty
 End If
End Sub
