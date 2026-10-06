VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_BackupResult 
   Caption         =   "HAMU | Yedek Hazır"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7290
   OleObjectBlob   =   "frmHAMU_BackupResult.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_BackupResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mPath As String
Public Sub Setup(ByVal path As String)
    mPath = path: txtPath.text = path
End Sub
Private Sub cmdOpen_Click()
    On Error GoTo Failed
    Dim folder As String
    folder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(mPath)
    Me.Hide
    HAMU_OpenFolder folder
    Exit Sub
Failed:
    HAMU_ShowError "HAMU_Backup", Err.number, Err.description
    Me.show vbModal
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    HAMU_ThemeForm Me
End Sub

Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub


