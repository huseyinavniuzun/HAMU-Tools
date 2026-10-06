Attribute VB_Name = "modHAMU_Color"
Option Explicit
Option Private Module
Public Sub HAMU_CountByColor()
    On Error GoTo HAMU_CommandFailed
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_CountByColor") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    HAMU_ColorReport False
    Exit Sub
HAMU_CommandFailed:
    HAMU_ShowError "HAMU_CountByColor", Err.number, Err.description
End Sub
Public Sub HAMU_SumByColor()
    On Error GoTo HAMU_CommandFailed
    ' HAMU entry guard: also covers direct macro calls.
    If Not HAMU_HasWorkbook() Then
        HAMU_ShowInfo HAMU_Text("Önce bir çalışma kitabı açın veya Ctrl+N ile oluşturun."): Exit Sub
    End If
    If HAMU_NeedsWorksheet("HAMU_SumByColor") And TypeName(ActiveSheet) <> "Worksheet" Then
        HAMU_ShowInfo HAMU_Text("Bu komut için bir çalışma sayfasına geçin."): Exit Sub
    End If
    HAMU_ColorReport True
    Exit Sub
HAMU_CommandFailed:
    HAMU_ShowError "HAMU_SumByColor", Err.number, Err.description
End Sub
Private Sub HAMU_ColorReport(ByVal sumValues As Boolean)
    Dim Data As range, sample As range, result As Double, count As Long
    Dim title As String
    On Error GoTo Failed
    If sumValues Then title = HAMU_Text("Renge Göre Topla") Else title = HAMU_Text("Renge Göre Say")
    Set Data = HAMU_PromptRange("" & title & HAMU_Text(": hesaplanacak hücreleri seçin."))
    If Data Is Nothing Then Exit Sub
    Set sample = HAMU_PromptRange(HAMU_Text("Eşleştirilecek dolgu rengini taşıyan örnek hücreyi seçin."))
    If sample Is Nothing Then Exit Sub
    HAMU_ColorAggregate Data, sample.Cells(1, 1), sumValues, result, count
    HAMU_ShowInfo title & vbCrLf & HAMU_Text("Eşleşen hücre: ") & count & IIf(sumValues, vbCrLf & "Toplam: " & CStr(result), "") & vbCrLf & HAMU_Text("Görüntülenen dolgu rengi kullanılır; koşullu biçimlendirme dahildir.")
    Exit Sub
Failed:
    HAMU_ShowError title, Err.number, Err.description
End Sub
Public Sub HAMU_ColorAggregate(ByVal Data As range, ByVal sample As range, ByVal sumValues As Boolean, ByRef total As Double, ByRef count As Long)
    Dim c As range, color As Long
    Set Data = HAMU_BoundedRange(Data)
    If Data.CountLarge > 250000 Then Err.Raise 5, , HAMU_L("Renk hesabı için en fazla 250.000 hücre seçin.", "Select at most 250,000 cells for colour aggregation.")
    total = 0: count = 0
    color = sample.DisplayFormat.Interior.color
    For Each c In Data.Cells
        If c.DisplayFormat.Interior.color = color Then
            count = count + 1
            If sumValues Then
                If Not IsError(c.Value2) And Not IsEmpty(c.Value2) Then
                    If VarType(c.Value2) <> vbString And VarType(c.Value2) <> vbBoolean Then
                        If IsNumeric(c.Value2) Then total = total + CDbl(c.Value2)
                    End If
                End If
            End If
        End If
    Next
End Sub

