Attribute VB_Name = "modHAMU_Forms"
Option Explicit
Option Private Module
Public gHAMU_InputCancelled As Boolean

Public Sub HAMU_ThemeForm(ByVal f As Object)
    On Error Resume Next
    f.BackColor = RGB(255, 255, 255)
    f.Font.name = "Tahoma"
    f.Font.Size = 8.5
    HAMU_ThemeControls f.controls
    f.controls("lblTitle").Font.Size = 12
    f.controls("lblTitle").ForeColor = RGB(22, 109, 166)
    f.controls("lblMonth").Font.Size = 12
End Sub
Private Sub HAMU_ThemeControls(ByVal controls As Object)
    Dim c As Object
    For Each c In controls
        On Error Resume Next
        If c.ControlTipText = c.name Then c.ControlTipText = ""
        c.Font.name = "Tahoma"
        c.Font.Size = 8.5
        c.ForeColor = RGB(48, 57, 74)
        Select Case TypeName(c)
            Case "CommandButton"
                c.BackColor = RGB(229, 237, 245)
                If Len(c.ControlTipText) = 0 Then c.ControlTipText = c.Caption
            Case "TextBox", "ComboBox", "ListBox"
                c.BackColor = vbWhite
                If Len(c.ControlTipText) = 0 Then c.ControlTipText = HAMU_Text("Bu alanın yanındaki açıklamaya göre seçim veya giriş yapın.")
            Case "Label", "CheckBox", "OptionButton", "Frame"
                c.BackColor = RGB(255, 255, 255)
                If Len(c.ControlTipText) = 0 Then c.ControlTipText = c.Caption
        End Select
        If TypeName(c) = "Frame" Then HAMU_ThemeControls c.controls
        On Error GoTo 0
    Next
End Sub

Public Function HAMU_InputText(ByVal Prompt As String, Optional ByVal title As String = HAMU_APP_NAME, Optional ByVal Default As String = "") As String
    Dim f As frmHAMU_Parameter
    Set f = New frmHAMU_Parameter
    f.Setup Prompt, title, Default, False
    f.show
    gHAMU_InputCancelled = Not f.accepted
    If f.accepted Then HAMU_InputText = f.ValueText
    Unload f
End Function

Public Function HAMU_InputValue(ByVal Prompt As String, Optional ByVal title As String = HAMU_APP_NAME, Optional ByVal Default As Variant = "", Optional ByVal Left As Variant, Optional ByVal Top As Variant, Optional ByVal HelpFile As Variant, Optional ByVal HelpContextID As Variant, Optional ByVal ValueType As Long = 2) As Variant
    Dim f As frmHAMU_Parameter
    If ValueType = 8 Then Err.Raise vbObjectError + 320, , HAMU_Text("Aralık seçiminde HAMU_PromptRange kullanın.")
    Set f = New frmHAMU_Parameter
    f.Setup Prompt, title, CStr(Default), (ValueType = 1)
    f.show
    If f.accepted Then
        If ValueType = 1 Then
            HAMU_InputValue = CDbl(f.ValueText)
        Else
            HAMU_InputValue = f.ValueText
        End If
    Else
        HAMU_InputValue = False
    End If
    Unload f
End Function




Public Function HAMU_ChooseIndex(ByVal title As String, ByVal description As String, ByVal choices As Variant) As Long
    Dim f As frmHAMU_Choice
    Set f = New frmHAMU_Choice
    f.Setup title, description, choices
    f.show vbModal
    If f.accepted Then HAMU_ChooseIndex = f.cboChoice.ListIndex + 1
    Unload f
End Function

Public Function HAMU_Text(ByVal text As String) As String
    HAMU_Text = HAMU_Translate(text)
    HAMU_Text = Replace$(HAMU_Text, "F9 kullanın", "Ctrl+Alt+F9")
    HAMU_Text = Replace$(HAMU_Text, "with F9", "with Ctrl+Alt+F9")
End Function

