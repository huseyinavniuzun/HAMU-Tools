VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmHAMU_ItemPicker 
   Caption         =   "HAMU | Öğe Seçici"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10305
   OleObjectBlob   =   "frmHAMU_ItemPicker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmHAMU_ItemPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mItems() As Variant        ' [1..n, 1..2] -> (T?r, ?sim)
Private mExists() As Boolean       ' [1..n]
Private mN As Long
Private mTemplPath As String

Public Sub LoadData(ByVal templPath As String, ByRef items As Variant, ByRef existsArr() As Boolean)
    mTemplPath = templPath
    mItems = items
    mExists = existsArr
    mN = UBound(items, 1)
    RenderList ""
    Me.optRename.value = True
    Me.txtSuf.text = "_2"
    Me.chHepsiniSec.value = False
End Sub

Private Sub txtAra_Change()
    RenderList Me.txtAra.text
End Sub
Private Sub RenderList(ByVal q As String)
    Dim arr() As Variant, r As Long, k As Long
    Dim typ As String, FullName As String, disp As String
    Dim showRow As Boolean, ql As String

    ql = LCase$(Trim$(q))

    With Me.lst
        .Clear
        .columnCount = 4
        .ColumnWidths = "25 pt;85 pt;280 pt;55 pt"
        .MultiSelect = 1 ' fmMultiSelectMulti
    End With

    If mN <= 0 Then
        Me.lblInfo.Caption = Dir$(mTemplPath) & HAMU_Text(" | 0 öğe")
        Exit Sub
    End If

    ReDim arr(1 To mN, 1 To 4)
    k = 0

    For r = 1 To mN
        typ = CStr(mItems(r, 1))          ' "LAMBDA" veya "STYLE"
        FullName = CStr(mItems(r, 2))

        If ql = "" Then
            showRow = True
        Else
            showRow = (InStr(1, LCase$(typ), ql, vbTextCompare) > 0) _
                   Or (InStr(1, LCase$(FullName), ql, vbTextCompare) > 0)
        End If

        If showRow Then
            k = k + 1
            arr(k, 1) = r                  ' 0. sütuna koyaca??m?z orijinal index
            arr(k, 2) = typ
            disp = FullName: If Len(disp) > 46 Then disp = Left$(disp, 45) & HAMU_Text("…")
            arr(k, 3) = disp
            arr(k, 4) = IIf(mExists(r), "Evet", HAMU_Text("Hayır"))
        End If
    Next r

 If k > 0 Then

    Dim outArr() As Variant, r2 As Long
    ReDim outArr(1 To k, 1 To 4)
    For r2 = 1 To k
        outArr(r2, 1) = arr(r2, 1)
        outArr(r2, 2) = arr(r2, 2)
        outArr(r2, 3) = arr(r2, 3)
        outArr(r2, 4) = arr(r2, 4)
    Next r2

    Me.lst.List = outArr
End If
    Me.lblInfo.Caption = Dir$(mTemplPath) & " | " & k & " / " & mN & HAMU_Text(" öğe")
End Sub
Private Sub chHepsiniSec_Click()
    Dim i As Long, sel As Boolean
    sel = (Me.chHepsiniSec.value = True)
    For i = 0 To Me.lst.ListCount - 1
        Me.lst.selected(i) = sel
    Next i
End Sub

Private Sub btnTransfer_Click()
    Dim sel() As Long, i As Long, c As Long
    If Me.lst.ListCount = 0 Then Exit Sub

    For i = 0 To Me.lst.ListCount - 1
        If Me.lst.selected(i) Then c = c + 1
    Next i
    If c = 0 Then
        MsgBox HAMU_Text("Aktarılacak öğe seçmediniz."), vbExclamation
        Exit Sub
    End If

    ReDim sel(0 To c - 1): c = 0
    For i = 0 To Me.lst.ListCount - 1
        If Me.lst.selected(i) Then
            sel(c) = CLng(Me.lst.List(i, 0))
            c = c + 1
        End If
    Next i

    Dim pol As String
    If Me.optOverwrite.value Then
        pol = HAMU_Text("Üzerine")
    ElseIf Me.optSkip.value Then
        pol = HAMU_Text("Atla")
    Else
        pol = "Yeniden"
    End If

    IP_OnTransfer pol, Me.txtSuf.text, sel
    Unload Me
End Sub

Private Sub btnClose_Click()
    IP_OnClose
    Unload Me
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    IP_OnClose
    On Error GoTo 0
End Sub
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Const K_ESC As Long = vbKeyEscape
    If KeyCode = K_ESC Then
        On Error Resume Next
        IP_OnClose      ' template'i kapat
        On Error GoTo 0
        Unload Me       ' formu kapat
    End If
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = HAMU_Text("HAMU | Öğe Seçici")
    lblTop.Caption = HAMU_Text("Aktarılacak öğeleri seçin")
    optOverwrite.Caption = HAMU_Text("Üzerine yaz")
    optRename.Caption = HAMU_Text("Yeni ad ver")
    optSkip.Caption = HAMU_Text("Atla")
    chHepsiniSec.Caption = HAMU_Text("Tümünü seç")
    fraPol.Caption = HAMU_Text("Çakışma durumunda")
    Frame1.Caption = HAMU_Text("Ara")
    btnTransfer.Caption = HAMU_Text("Aktar")
    btnClose.Caption = HAMU_Text("Kapat")

    HAMU_ThemeForm Me
End Sub



























































Private Sub UserForm_Activate()
 HAMU_LocalizeForm Me
End Sub












































