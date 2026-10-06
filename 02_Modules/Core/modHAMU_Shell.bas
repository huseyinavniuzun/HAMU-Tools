Attribute VB_Name = "modHAMU_Shell"
Option Explicit
Option Private Module
#If VBA7 Then
Private Declare PtrSafe Function ShellExecuteW Lib "shell32.dll" (ByVal hwnd As LongPtr, ByVal operation As LongPtr, ByVal file As LongPtr, ByVal parameters As LongPtr, ByVal directory As LongPtr, ByVal show As Long) As LongPtr
#Else
Private Declare Function ShellExecuteW Lib "shell32.dll" (ByVal hwnd As Long, ByVal operation As Long, ByVal file As Long, ByVal parameters As Long, ByVal directory As Long, ByVal show As Long) As Long
#End If
Public Function HAMU_OpenFolder(ByVal folder As String) As Boolean
    Dim fso As Object, operation As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Err.Raise 76, , HAMU_Text("Yedek klasörü bulunamadı.")
    folder = fso.GetAbsolutePathName(folder): operation = "open"
#If VBA7 Then
    Dim result As LongPtr
#Else
    Dim result As Long
#End If
    result = ShellExecuteW(Application.hwnd, StrPtr(operation), StrPtr(folder), 0, 0, 1)
    If result <= 32 Then Err.Raise 5, , HAMU_Text("Klasör açılamadı. Sonuç ekranındaki yolu kopyalayıp Dosya Gezgini'ne yapıştırabilirsiniz.")
    HAMU_OpenFolder = True
End Function

Public Sub HAMU_OpenProfile(ByVal address As String)
 On Error GoTo Failed
 Dim operation As String
 If address <> "https://huseyinavniuzun.com" And address <> "https://github.com/huseyinavniuzun" And address <> "mailto:hamu@huseyinavniuzun.com" Then Exit Sub
 operation = "open"
#If VBA7 Then
 Dim result As LongPtr
#Else
 Dim result As Long
#End If
 result = ShellExecuteW(Application.hwnd, StrPtr(operation), StrPtr(address), 0, 0, 1)
 If result <= 32 Then Err.Raise 5
 Exit Sub
Failed:
 HAMU_ShowInfo address
End Sub

