VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHAMU_DayButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents Button As MSForms.label
Attribute Button.VB_VarHelpID = -1
Public owner As Object
Public DateValue As Date
Private Sub Button_Click()
    owner.SelectDate DateValue
End Sub

