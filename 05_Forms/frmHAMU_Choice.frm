VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_Choice 
   Caption         =   "HAMU Tools"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3825
   OleObjectBlob   =   "frmHAMU_Choice.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_Choice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public accepted As Boolean
Public Sub Setup(ByVal title As String, ByVal description As String, ByVal choices As Variant)
    Dim item As Variant
    Me.Caption = title: lblDescription.Caption = description
    cboChoice.Clear
    For Each item In choices: cboChoice.AddItem CStr(item): Next
    cboChoice.ListIndex = 0: accepted = False
End Sub

Private Sub cmdOK_Click()
    accepted = True: Me.Hide
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: accepted = False: Me.Hide
End Sub
Private Sub UserForm_Initialize()
    HAMU_ThemeForm Me
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub


