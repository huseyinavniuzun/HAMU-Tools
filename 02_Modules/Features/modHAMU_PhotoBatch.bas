Attribute VB_Name = "modHAMU_PhotoBatch"
Option Explicit
Option Private Module

Public Sub HAMU_ImportCellPhotoBatch(ByVal target As Range, ByVal files As Collection)
    Dim state As New clsHAMU_AppState, block As Collection, first As Long, last As Long, i As Long
    Dim number As Long, description As String, cancelKey As Long, cancelCaptured As Boolean
    On Error GoTo Failed
    state.Capture
    cancelKey = Application.EnableCancelKey: cancelCaptured = True
    If target Is Nothing Then Err.Raise 5, , HAMU_L("Başlangıç hücresini seçin.", "Select the starting cell.")
    If files Is Nothing Then Err.Raise 5
    If files.Count = 0 Then Exit Sub
    If files.Count > target.Worksheet.Rows.Count - target.Row + 1 Then Err.Raise 5, , HAMU_L("Fotoğraflar bu alana sığmıyor.", "The photos do not fit the destination.")
    If target.Worksheet.ProtectContents Then Err.Raise 1004, , HAMU_L("Sayfa korumalı. Önce korumayı kaldırın.", "Unprotect the destination sheet first.")
    Dim merges As Variant
    merges = target.Cells(1, 1).Resize(files.Count, 1).MergeCells
    If IsNull(merges) Then Err.Raise 5, , HAMU_L("Birleştirilmemiş hücreler seçin.", "Select unmerged cells.")
    If merges Then Err.Raise 5, , HAMU_L("Birleştirilmemiş hücreler seçin.", "Select unmerged cells.")
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.EnableCancelKey = xlErrorHandler
    ' Bound temporary memory without keeping a worker or image cache alive.
    For first = 1 To files.Count Step 50
        last = first + 49: If last > files.Count Then last = files.Count
        Set block = New Collection
        For i = first To last: block.Add CStr(files(i)): Next
        HAMU_CopyCellPhotoBlock target.Offset(first - 1, 0), block
        Set block = Nothing
    Next
    Application.EnableCancelKey = cancelKey
    state.Restore
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    If cancelCaptured Then Application.EnableCancelKey = cancelKey
    state.Restore
    Err.Raise number, "HAMU_ImportCellPhotoBatch", description
End Sub

Private Sub HAMU_CopyCellPhotoBlock(ByVal target As Range, ByVal files As Collection)
    Dim fso As Object, manifest As Object, shell As Object, source As Workbook, destination As Range
    Dim folder As String, manifestPath As String, bookPath As String, seedPath As String, script As String, exitCode As Long
    Dim number As Long, description As String, i As Long, height As Variant, row As Range
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.BuildPath(fso.GetSpecialFolder(2).Path, "HAMU_Photos_" & fso.GetTempName)
    fso.CreateFolder folder
    manifestPath = fso.BuildPath(folder, "files.txt"): bookPath = fso.BuildPath(folder, "HAMU_" & fso.GetTempName & ".xlsx")
    seedPath = fso.BuildPath(folder, "seed.xlsx")
    HAMU_PhotoSaveSeed seedPath
    Set manifest = fso.CreateTextFile(manifestPath, True, True)
    For i = 1 To files.Count: manifest.WriteLine CStr(files(i)): Next
    manifest.Close: Set manifest = Nothing
    script = "$manifest='" & Replace$(manifestPath, "'", "''") & "';$output='" & Replace$(bookPath, "'", "''") & "';$seed64=[Convert]::ToBase64String([IO.File]::ReadAllBytes('" & Replace$(seedPath, "'", "''") & "'));" & vbLf & HAMU_PhotoPackageScript
    Set shell = CreateObject("WScript.Shell")
    exitCode = shell.Run("powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand " & HAMU_PhotoEncode(script), 0, True)
    If exitCode <> 0 Or Not fso.FileExists(bookPath) Then Err.Raise 5, , HAMU_L("Fotoğraflar hazırlanamadı. Dosyaların okunabildiğini kontrol edin.", "The photos could not be prepared. Check that the files are readable.")
    Set source = Workbooks.Open(FileName:=bookPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    source.Windows(1).Visible = False
    Set destination = target.Cells(1, 1).Resize(files.Count, 1)
    source.Worksheets(1).Range("A1").Resize(files.Count, 1).Copy
    destination.PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    source.Close False: Set source = Nothing
    height = destination.EntireRow.RowHeight
    If IsNull(height) Then
        For Each row In destination.Rows
            If row.RowHeight < 100 Then row.RowHeight = 100
        Next
    ElseIf height < 100 Then
        destination.EntireRow.RowHeight = 100
    End If
Clean:
    On Error Resume Next
    Application.CutCopyMode = False
    If Not source Is Nothing Then source.Close False
    If Not manifest Is Nothing Then manifest.Close
    ' Delete only files created by this block; never recurse through a user folder.
    If Len(bookPath) > 0 Then Kill bookPath
    If Len(manifestPath) > 0 Then Kill manifestPath
    If Len(seedPath) > 0 Then Kill seedPath
    If Len(folder) > 0 Then RmDir folder
    On Error GoTo 0
    If number <> 0 Then Err.Raise number, "HAMU_CopyCellPhotoBlock", description
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    Resume Clean
End Sub

Private Function HAMU_PhotoEncode(ByVal text As String) As String
    Dim stream As Object, node As Object, doc As Object, bytes As Variant
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "unicode": stream.Open: stream.WriteText text
    stream.Position = 0: stream.Type = 1: stream.Position = 2: bytes = stream.Read: stream.Close
    Set doc = CreateObject("MSXML2.DOMDocument.6.0"): Set node = doc.createElement("base64")
    node.DataType = "bin.base64": node.nodeTypedValue = bytes
    HAMU_PhotoEncode = Replace$(Replace$(node.Text, vbCr, ""), vbLf, "")
End Function

Private Sub HAMU_PhotoSaveSeed(ByVal path As String)
    Dim doc As Object, node As Object, stream As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0"): Set node = doc.createElement("base64")
    node.DataType = "bin.base64": node.Text = HAMU_PhotoSeed
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1: stream.Open: stream.Write node.nodeTypedValue: stream.SaveToFile path, 2: stream.Close
End Sub

Private Function HAMU_PhotoSeed() As String
 Dim value As String
 value = value & "UEsDBBQAAAAIACtmW10S3NTdmwEAAAsHAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbMVUO2/bMBDeC/Q/CFwDiU6GoigsZ2iSrW0QpN1Z8mwR5gu8k2v/+55o1wUCO1acFF1EifxeOp1uer32rlpBRhtDKy6biagg6GhsWLTi++Nd/VFUSCoY"
 value = value & "5WKAVmwAxfXs/bvp4yYBVswO2IqOKH2SEnUHXmETEwQ+mcfsFfFjXsik9FItQF5NJh+kjoEgUE2DhphNb2CuekfV7Zq3t0lSWIjq8xY3WLXC+oE/7MuDjAwOn1BUSs5qRXwuV8E8yVXvMjXMLBjsbMILBhxxGE6OG+x437iY2Rqo7lWmr8oz"
 value = value & "Sq6d/BXz8meMy+Z5kQMp43xuNZioe8+UBlMGZbADIO+asjZe2XBx2r+AUZbl8o2D7PVfmOPqP+Ug7lTYXl9fiiJzwhBp4wDf+vMX0RPOHkgZRepftN6XnfaJCNnq7oZh5eaHcj08gBuRx2MNa83QgbcaePyvjvYye9aLrP7SznBCyr2mPp9r"
 value = value & "ueeP9374U9Myks/0LaP4OU/ugvscE/LwHvVyR2btwK4TC0EmO9KRpV/duzCMcQPmgLcsdZv9BlBLAwQUAAAACAArZltdtVUwI+sAAABMAgAACwAAAF9yZWxzLy5yZWxzrZLNasMwDIDvg72D0b1R2sEYo04vY9DbGNkDaLbyQxLL2G6Xvv28"
 value = value & "w9gCXelhR8vSp09C2908jerIIfbiNKyLEhQ7I7Z3rYa3+nn1AComcpZGcazhxBF21e3N9pVHSrkodr2PKlNc1NCl5B8Ro+l4oliIZ5d/GgkTpfwMLXoyA7WMm7K8x/CbAdWCqfZWQ9jbO1D1yfM1bGma3vCTmMPELp1pgTwndpbtyodcH1Kf"
 value = value & "p1E1hZaTBivmJYcjkvdFRgOeN9pcb/T3tDhxIkuJ0Ejgyz5fGZeE1v+5omXGj8084oeE4V1k+HbBxQ1Un1BLAwQUAAAACAArZltdQquXg8kDAADkCQAADwAAAHhsL3dvcmtib29rLnhtbK1WXW+rOBB9X2n/A+LdxeYb1PSKANGt1F5Vbdru"
 value = value & "SpEqB5yCCpg1pkl11f++YxqS9LYP2e5FicH25MyZmTMmp982daU9M9GVvJno5ATrGmsynpfN40S/nc+Qr2udpE1OK96wif7COv3b2Z9/nK65eFpy/qQBQNNN9ELKNjSMLitYTbsT3rIGdlZc1FTCVDwaXSsYzbuCMVlXhomxa9S0bPQ3hFAc"
 value = value & "g8FXqzJjCc/6mjXyDUSwikqg3xVl241odXYMXE3FU9+ijNctQCzLqpQvA6iu1Vl4/thwQZcVhL0hjrYR8HHhSzAM5ugJtj64qstM8I6v5AlAb0l/iJ9gg5B3Kdh8zMFxSDYk4blUNdxDuV/EcndY7h6M4P+NRvAezvwimrNDM/Wz01VZsbs3"
 value = value & "6Wq0bX/QWlWq0rWKdjLNS8nyie7BlK/ZuwXRt9O+rGBiYcf0dONsJ+croeVsRftKzoHYCD/RTWxaGCtLEEZUSSYaKlnMGwk6/E2aG7DjgkPo2jX7py8F6wbpwQ6MNAvpsruistB6UU30OFzcdsBvUfQde1lEbZtQSRcXPKPVYs7qdvE9urx9"
 value = value & "iFlVXRVc8gebsoAsfd+2HM+meUbpijiZ63rEypaW7y8OJE0/xvIfRE0zlSljF87b869pg6hEOBb0SgoNns+TCyjeDX2GUoJg8m2nn0OtiPXQZCIkDz9j13X9yE1R6rsRsmfER0HsJshxp9HU932SEutVV90aZpz2stiWUUFPdNv7ZOuSbsYd"
 value = value & "gsO+zPc0fuLthT4ZxutVBawEdFeydbfXk5pqm/uyyfl6oiNiQlAv76frYXZf5rIAnQXY3q19Z+VjAYyJ5yhD6BvFDBglgevEnuUjE0cBsl3soMBxLeT6Dsg0SGySRgMj44DSUKTxrjVDt9zQlxUlcMSrtSHJ0B2h8iHOc6IQPlibB9bmgbU5"
 value = value & "+BudgAYz6CV1G2ADgs1AWbCNvOjkcAcZlxAMsXHk4cBGOLUcZPuBiUCiJortxEwdL03SqfP6O0/boZvCg47PCirkXNDsCV5712w1pR1TqVABAc9DslPHn2ILKILqZsgmAUbTqWsjJ5lBV5EkTp3ZnqwKf/VFvr4x/JpR2cM5oEgP81CNs+3q"
 value = value & "bnFrtq3TOwfhdaICOcLwBqKv2JHGs7sjDeMfl/PLI20v0vnD/exY4+hymkTH20fX19Hf8/Sv0YXxaUJ/LXhC7ABbaYQsK7aR7c085M+g3yzbs2PHnqYEe/uCV+vs+Wv1Nm1jVGR8+GrYnk+q/go83P5X0zom9y+nvVKNsb92aGf/AlBLAwQU"
 value = value & "AAAACAArZltd+MHrFVUBAACABQAAGgAAAHhsL19yZWxzL3dvcmtib29rLnhtbC5yZWxzvZRLT8MwDMfvSHyHKneatsAGaN0uE9IOXMbgHqXuQ2uTKnEH+/a4G3t0lEhIFafWj/z9s61kMvusSm8DxhZaxSz0A+aBkjopVBazt9XzzQPzLAqV"
 value = value & "iFIriNkWLJtNr68mSygF0iGbF7X1SEXZmOWI9RPnVuZQCevrGhRFUm0qgWSajNdCrkUGPAqCETfnGmza0fQWSczMIqH6q20NP7SrQhptdYq+1BXXaVrIVjUc80thbpJlIfN3UTbwiqaR2BggVWEyQKpBsblAQWnt76ZNs4c0n/AZ7ye7/YWs"
 value = value & "p+s93lzLpgKFPc1zpLNnUDtz7wxdDONBp+MeiosjGnIWH9qsbQ6AJ56jy/LdJ3LBhP8M49zQ6G8biiIeBpcbOuyH9PtWdBZ2kdwPOZZd4y+AIiGEE1T17XFx3A3KgduSnqQjwN52lX8c9Mq0Urb33nRTDkS8825OvwBQSwMEFAAAAAgAK2Zb"
 value = value & "XRru6NQqAgAAngQAABgAAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWydlGtr2zAUhr8P9h807XPiW9JcsF2yXFhhg7Gt3WdFPo5FLMuTFCel7L/vSJnTQhmEgmXdrOd9dXTk9PYka9KBNkI1GY2GISXQcFWIZpfR+5+bwZQSY1lTsFo1kNFH"
 value = value & "MPQ2f/8uPSq9NxWAJUhoTEYra9t5EBhegWRmqFpocKZUWjKLXb0LTKuBFX6RrIM4DG8CyURDz4S5voahylJwWCl+kNDYM0RDzSz6N5VoTU+T/BqcZHp/aAdcyRYRW1EL++ihlEg+v9s1SrNtjfs+RSPGyUnjE2NJehk//kpJCq6VUaUdIvmf"
 value = value & "59fbnwWzgPEL6fX+r8JEIwxAJ9wBPqPiN7LGF1b8DEveCLu5wFy49Pwgiow+RbNlOFkvksEkno4Ho0USDT6t17PBeJNMp5Mk2SxGN39onhYCT9gtJhrKjC7HNMhTT38QcDQv2sSy7Q+ogVtAgYgSl5tbpfZu8g6HQsQZ/4HDMaw6WEJdeyox"
 value = value & "v18IBBfqy3avtvGp802TAkp2qO13dfwMYldZlEWST4Z58bgCwzE7UXgYeypXtfFvIgVeMwyHZCdfH0Vhq4zG0+EkCmfJBCn8YKySv84TkTd1Xu8trJhlearVkWC+OPctc7cvmSPM+5iGPeJi7X/O0JKjuCDgEFDSSfd1nnb5x4fFl/v1hzTo"
 value = value & "nDwWVOwjcrbQsh18ZXonGkNqKD1zQok+a/q2Va1vIX+rLDrqexUmCmjXQ9elUrbvuM1e/iz5X1BLAwQUAAAACAArZltdNOqFBKkBAACUAwAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQyLnhtbJ2TS4vbMBDH74V+B6G740fsvIizbDYN3UNh"
 value = value & "Kd29K/LYFrE0RlI2CaXfvWOn8RZyCQs2zN8e/eap5cNJN+wdrFNoch6PIs7ASCyUqXL++msbzDhzXphCNGgg52dw/GH19cvyiHbvagDPiGBczmvv20UYOlmDFm6ELRj6U6LVwpO0VehaC6LoD+kmTKJoEmqhDL8QFvYeBpalkrBBedBg/AVi"
 value = value & "oRGe8ne1at2VpuU9OC3s/tAGEnVLiJ1qlD/3UM60XDxXBq3YNVT3KU6FZCdLT0Lv+Bqm/34TSStp0WHpR0T+l/Nt+fNwHgo5kG7rvwsTp9SAd9UN8AOVfJKVDazkAzb+JGwywLp22cVBFTn/PYmzeD1fp8E6GsdBup4+BbM0yoKnSRTN4sm3"
 value = value & "x/F0+4evloWiCXeHmYUy548xD1fLnv6m4Oj+s1m3izvEfSeeKUbUuYY3vtt+8C+WFVCKQ+N/4vE7qKr2tPgZZdiNclGcN+Ak7RZhRkk2BN0IL8huRQU/hK2UcayBsveacmYvmN722PYWIXfoPeqrqqk9YDtF/SgR/VV02Q73afUXUEsDBBQA"
 value = value & "AAAIACtmW11JKSLk7gYAABMiAAATAAAAeGwvdGhlbWUvdGhlbWUxLnhtbO1aW48bNRR+R+I/WPOeziWZXKqmKFdKu9uudrdFPDoTJ+PGMx7Zzu5GqBIqT7wgIQHiBYk3HhACBBIVL7zwTyq1gvIj8Hgmk3Hi6YVuoYjdlXZjz/cdf3PO8fHJ"
 value = value & "JFfeOosIOEGMYxp3LfeSYwEUB3SK43nXun08rrUtwAWMp5DQGHWtFeLWW1fffOMKvCxCFCEg+TG/DLtWKERy2bZ5IKchv0QTFMtrM8oiKOSQze0pg6fSbkRsz3GadgRxbIEYRtLsrdkMBwgcp9TffrCurs2PiPwTC55OBIQdBWrNMkdhpws3"
 value = value & "/cdXfEAYOIGka8mVpvT0GJ0JCxDIhbzQtRz1Y9lXr9gFiYgKbok3Vj85LydMF57isfmkIDojr91wC/teZn8XN2qnv4U9BYBBIO/U3cG6ftNpezm2BMpeGmx3Wm5dx5fs13ftd5p9r6Hh6xt8Y/cex53R0NfwjQ3e38H3HK/fqWt4f4Nv7uAb"
 value = value & "o17LG2l4BQoJjhe76Gar3W7m6AIyo+SaEd5pNp3WMIdvUHYpuzJ+LKpyLYJ3KRtLgAouFDgGYpWgGQwkrpcIysEQ84TAlQUSGFMupx3PdWXiNRyv+FUeh5cRLLGzqYDvTKV6AA8YTkTXui6tWiXIowcPHt7/6eH9nx9++OHD+9+BPTwPhYF3"
 value = value & "DcbzMu/J15/8+eUH4I8fv3ry6WdmPC/jH3/70eNffn2aeaHJ+vz7xz99/+iLj3//5lMDvMfgpAw/xhHi4CY6BYc0kjdoWABN2IsxjkOINQYMJdIAHIlQA95cQWLC9ZHuwjtMVgoT8O3lXU3rUciWAhuAN8JIA+5TSvqUGW/nRrpW+XaW8dy8"
 value = value & "OFuWcYcQnpjWHmwFeLRMZMpjk8lBiDSZB0RGG85RjARIr9EFQgbaexhrft3HAaOczgR4D4M+xEaXHOOJMJOu4UjGZWUSKEOt+Wb/DuhTYjI/RCc6Um4LSEwmEdHc+DZcChgZFcOIlJF7UIQmkUcrFmgO50JGeo4IBaMp4tzEucVWmtwbUJYs"
 value = value & "Y9j3ySrSkUzghQm5ByktI4d0MQhhlBg14zgsY9/hC5miEBxQYRRB9R2SjmUcYFwZ7jsYiRfb1rdlBTInSHplyUxbAlF9P67IDCKT8R6LtOraY9iYHf3lXEvtPYQIPIVThMDtd0x4mlCz6OuhrCrXkMk316Geq+k4RjxtlM5MNXgPcy1lj9Cc"
 value = value & "VujZX20VnhWMI8iqLN9c6CkzmjBsLKW3SLDQSilm6aY1i7jFI/hcVg9CqKVVOubmfF2x+EX3mOTc/Rsc9MIcWdif2zfHkCBzwhxDDPZM5VZSlmZKup0UbWnkzfRNuwmDvdXvRDh+VvNzEzKWNs//Ru/zyrqe8+93qurKdpdThfsP9jZDuIwP"
 value = value & "kDxOLlqbi9bm/9jaVO3li4bmoqG5aGj+sYZm08PY5Uc9ykpU+dxnhgk5EiuC9rjqfrjc+9OxnFQDRSoeMyWhfJkvp+HmDKrXgFHxLhbhUQgTuYyrVpjz3PScg4Ry2TpZlbZV/7WM9uk0f4rnrp9sSgIUm3nHL+Zltyay2WZr8xi0MK9Gc14W"
 value = value & "4Cujzy+itJguom4Q0ao/nwjXOS8VHYOKtvs0FXYpKvJwAjB9LO43MkUy3WRKT9M4Zfx1dM890lXO1G/bM9xep3FukdZElNJNF1FKw1AeHtvT5xzrTsccas8oo9V+FbG2d2sDifUROE01tVI7AUy61ky+dZIvo0Qa5GmpgmQed61A5J7+O6Ul"
 value = value & "YVwMIQ8zmLqUOSDCAjFAcCSTvRwHEpfEdeSmeV3FeWkQXjdx9naU0WyGAlExsxnKa5kR49WXBKcDupSij8LpKZiQJTuE0lF+y02jO8VcFKGeYlbK7o0Xt+pVvhe1D4A2exSSJIT5kVKu5hlcvS7klO5DKd2+K9vkwsl8fB7H7rNJW1Wz4gRp"
 value = value & "VZaxV3fKl1TVzap8Y7HrtJ2nHxMvfyKUpLXN0upmaVWHxzl2BKXlmhV+8yqj+ZLHwXbW2qXGUo12Ptumk7sy84eyXV2SbIbEcqQkJwdMaZ/Q6Sp/SXi2S7J7WpcBEh+iGcDTM1kyTc7JPzwuithhtkB6eBVEo1d1Yo7fFJ6C7D6bXDDWPXtB"
 value = value & "Vm25yYA4K1bO8FnAiqqRe8o2eVG+92NwsP5oNyunanZdos8EWDLctd53/F5j4PmDmtP2R7VGveHU2n6vXuv5ft0d+a4z7Hv3pDwRRq6fBXAMI0xW+Tcg1PzOtyCi9RuWSwGNbKreTdiKrL4F4XrGb0Gk1y3pFSnLG7kNr+cNaoOh26w1vGGz"
 value = value & "1m7Ve7WB1xx6PVnJm+PePQucKLDbHw7HY9+rNQcS13B6fq3Xrw9qzfao743dUWPoSHAeiDOx/r/OUaXr6l9QSwMEFAAAAAgAK2ZbXWgGWC+aAgAAawYAAA0AAAB4bC9zdHlsZXMueG1spZXfa9swEMffB/sfhN5d2W6cJcF2aZoaClsZNIO9"
 value = value & "KraciOqHkZTU2dj/vpOdxAkd22ifdLo7fe4rnySnN60UaMeM5VplOLoKMWKq1BVX6wx/WxbBBCPrqKqo0IpleM8svsk/fkit2wv2tGHMIUAom+GNc82MEFtumKT2SjdMQaTWRlIHU7MmtjGMVtYvkoLEYTgmknKFe8JMlv8DkdQ8b5ug1LKh"
 value = value & "jq+44G7fsTCS5exhrbShKwFS22hES9RGYxOj1hyLdN5XdSQvjba6dlfAJbquecley52SKaHlQALy20hRQsL4Yu+teSNpRAzbcd8+nKe1Vs6iUm+Vg2YC22929qz0iyp8yDv7rDy1P9COCvBEmORpqYU2yEFl5pPAo6hkfcZt47RFj9QY/eIj"
 value = value & "NZVc7PtY3C3eUGPhIPS8cefrdnEASA5N8U7Sl+4GCyAuxElujHtHnkJfHTOqgAk62Mt9A7oUHMEe0+X9I3tt6D6Kk7MF3QB1V9pUcOSHD3V05algtYMFhq83fnS6IT7oHLQgTytO11pR4ZHHFQcDsCUT4slfi+/1BbutkdrKQrqHKsNwwfzu"
 value = value & "jyYIOpg9pp94/jmtZ78bi9r6kn9Cd4Uu6Ccv8gchw4/+CooBgVZbLhxXfxAMzKodtHZR5+/kZRVgVKymW+GWp2CGB/sLq/hWxqesr3yn3SFrsD/7TkVjX4O17rN13Yi2hmf45/3803RxX8TBJJxPgtE1S4JpMl8EyehuvlgU0zAO736dvQzv"
 value = value & "eBe6ywxNiUYzKyDLHDZ7EP80+DJ8Nunld98PZJ9rn8bj8DaJwqC4DqNgNKaTYDK+ToIiieLFeDS/T4rkTHvyxvcjJFE0iE9mjksmuGKX8pfnXmgSTP+yCXLsBBn+EvlvUEsDBBQAAAAIACtmW10jGDVdggEAALECAAAPAAAAeGwvbWV0YWRh"
 value = value & "dGEueG1sbZFPb+IwEMXvlfY7WL4Hh4Aoi5JUFVrUSvTSUrRX40yIVf+Tx2FBq/3u6yQtLW0vtsej+fm95/zmqBU5gEdpTUHHo5QSMMJW0uwL+rxZJXNKMHBTcWUNFPQESG/KH1e5hsArHjiJAIMFbUJwC8ZQNKA5jqwDEzu19ZqHWPo9Q+eB"
 value = value & "V9gABK1YlqYzprk0dCAsjspXXzBaCm/R1mEkrGa2rqWAL6DxNfNSNJ0aWp6FbU4OkAjbmhCNfWoQw3V083v9eL+8296un39RoqV5ap2zPkC1PScSdaYxFGHdqcMQxzHArVLvxZarFrCvNfg99Cd0SoaV9Ng/Trz9s7TqqZH1UAsF3K+GcN4v"
 value = value & "llZrMK83HFHuzdC04EXPZWXOLgyWed2G1sPD23d8Y+xDBruXModjWGPod9J6WdC/E8jmaSamCZ9WWTKt5rvk53g6T2Cyg5mYpDDLJv/idPdFC3/YkTiU9loiY1h7Iuvw7FJQmR+6fM76PonxggyJHN6QPeNi5oPn8j9QSwMEFAAAAAgAK2Zb"
 value = value & "Xc/+3jy9AAAABAEAABwAAAB4bC9yaWNoRGF0YS9yaWNoVmFsdWVSZWwueG1sZc8xSwQxEAXgXvA/hOnd2dtCZNnsNYdgK2ofktlLIMksM7lD/705tBHLx4OP95bjZ8nmSqKJq4XDMIKh6jmkerbw/vb88ARGm6vBZa5k4YsUjuv93SLJxw+X"
 value = value & "L/RKWU1XqlqIre0zovpIxelQkhdW3trguSBvW/KEugu5oJGolYzTOE14s643SyjDjzXLP413qr3bWIprPcr5lzyxvxSqrWPjI3bDtf5GY9oV1qVnI3MKFuQlHADXBf9sX78BUEsDBBQAAAAIACtmW10ybSmGmgAAALsAAAAbAAAAeGwvcmlj"
 value = value & "aERhdGEvcmRyaWNodmFsdWUueG1sJY5BDsIgFET3Jt6B/L2FmhiNKXRjPIEegMCvJSn8ho9Eby/qZmY282aG8RUXUTFzoKSh7xQITI58SA8N99t1dwLBxSZvF0qo4Y0Mo9luhlwvtljR6ok1zKWsZynZzRgtdzG4TExT6RxFSdMUHEpeM1rP"
 value = value & "M2KJi9yr/ihzcLNvGBCOnqm0A2AaWTSiaqkaNcj69cPPZf7Ld9l8AFBLAwQUAAAACAArZltdsgMWTMQAAAACAQAAJAAAAHhsL3JpY2hEYXRhL3JkcmljaHZhbHVlc3RydWN0dXJlLnhtbE2OS2vDQAyE74X+h0X3epNcWoLtHAqFQKDQxzks"
 value = value & "smyL7CNIsmn/fbdQ2h4GhpmR+NrDR4puJVEuuYNtswFHGcvAeerg/e3p7gGcWshDiCVTB5+kcOhvb1pZX00WtEVIXX2StYPZ7Lr3XnGmFLRJjFK0jNZgSb6MIyN5vQqFQWciS9HvNtt7L4zzECyAw7JkqxjQt+qqOceCIR5TmKhGF1cRz7K+"
 value = value & "UNyffovjQNl4ZBL4PmHwP8vHEPFZeOL8V3it+o/efwFQSwMEFAAAAAgAK2ZbXUP6hnVlAQAAowQAACAAAAB4bC9yaWNoRGF0YS9yZFJpY2hWYWx1ZVR5cGVzLnhtbK2U32vCMBDH3wf7H0reNeqDG2Iro1uhMMZA9zyuaVqDl6QkqbT//U6n"
 value = value & "7BfCxvqWXC6f7+V+ZLnqNEZ76byyJmbT8YRF0ghbKlPH7GWTjW5Z5AOYEtAaGbNeerZKrq+Wbr/pG+lzU9mIGMbHbBtCs+Dci63U4MdaCWe9rcJYWM1tVSkhuW+chNJvpQwa+WwyveFOiW0JAWbsHbTQ4gfLNtLQYWWdhkBbV3MNbtc2I2I3"
 value = value & "EFShUIWegJM5i7RY5LWxDgqkkLszt/sN9nuAkzkpKcOSZY22AEyWO9lnCLU/riIDmjRe1xIr8qno4GR66AS2pcyc1ZlCyaI9YEv2KeOX/FJAkR7e45S35usNTmJfFO+VbxD6dXBUq4vSf0IenzUQ6pjSYVjrtniEQuIwtLtAGSvaQA0/DDAX"
 value = value & "Q5FOJR0GloKxRgnAZ0dt7kL/RPaBqpsieK9ooOGQxrz8B5V/jBM/Txj/9Lkkb1BLAwQUAAAACAArZltduiplBTwBAAB1AgAAEQAAAGRvY1Byb3BzL2NvcmUueG1snZLLasMwEEX3hf6D0d6WH5C0wnbog9BFUwp1aMlOSJNE1JKFpMbxv3XX"
 value = value & "H6vtJE5Cs+rS3DPHdwalk60svQ0YKyqVoSgIkQeKVVyoVYbmxdS/QZ51VHFaVgoy1IBFk/z6KmWasMrAq6k0GCfAeq1JWcJ0htbOaYKxZWuQ1AYtodpwWRlJXftpVlhT9klXgOMwHGEJjnLqKO6Evh6MaK/kbFDqL1P2As4wlCBBOYujIMJH"
 value = value & "1oGR9uJAn5yQUrhGw0X0EA701ooBrOs6qJMebftH+GP2/Nav6gvV3YoBylPOCDNAXWXyp59vC41Q3t1GCW++mL+k+CTuTllS62bt1ZcC+H1zceIvle6X3ZmAe21JslvpkLwnD4/FFOVxGI/8KPTjcRHeknBMkmTRlTibPwrl/if/Nh4Eed/7"
 value = value & "/KHkv1BLAwQUAAAACAArZltd33oyOJEBAAA3AwAAEAAAAGRvY1Byb3BzL2FwcC54bWydk8FO4zAQhu9IvEPkO3VaVghVjhEqIA6LtlIL91ln0lo4tmUPUcsL7GPwDDzDsu+Fk6ohXQ5I3Gbmn/z+PJ6Ii01tsgZD1M4WbDzKWYZWuVLbVcHu"
 value = value & "lzcn5yyLBLYE4ywWbIuRXcjjIzEPzmMgjTFLFjYWbE3kp5xHtcYa4ijJNimVCzVQSsOKu6rSCq+ceqrREp/k+RnHDaEtsTzxvSHbOU4b+q5p6VTLFx+WW5/8pLj03mgFlG4p77QKLrqKsuuNQiP4UBTJaIHqKWjaylzwYSoWCgzOkrGswEQU"
 value = value & "/KMgbhHaoc1BhyhFQ9MGFbmQRf2cxjZh2W+I2OIUrIGgwRLbte2SLjY+UpB//4B5e/33UkO2gG06CcLbq+C93oXDz4ax/iEnXUMKDht5z5TiQ9qlJoPxVzWHQF/BdwxsgNsxjod8h9LkE/oe4r9jZ672YNOYeR/91PYx3vuluwLC/dAPi2Kx"
 value = value & "hoBleqf+UfqCuE3IwbT9szXYFZb7ns9CuyIPu/9Ajs9G+Wmed5uxrwn+sfHyHVBLAwQUAAAACAArZltdLyzzyLcAAAAkAQAAJwAAAHhsL3JpY2hEYXRhL19yZWxzL3JpY2hWYWx1ZVJlbC54bWwucmVsc43PsWoDMQwG4L3QdzDaa911CKWc"
 value = value & "L0sJZC3JAwhb5zM9y8Z2QvL2NXRpoENHSfzfj6b9LW7qyqWGJAZGPYBisckF8QbOp8PLG6jaSBxtSdjAnSvs5+en6ZM3aj1U15Cr6opUA2tr+R2x2pUjVZ0yS78sqURqfSweM9kv8oyvw7DD8tuA+cFUR2egHN0I6nTP/B87LUuw/JHsJbK0"
 value = value & "PyowxN7dQSqemwGtMbIL9LMfdRYPOE/48Nv8DVBLAQIUABQAAAAIACtmW10S3NTdmwEAAAsHAAATAAAAAAAAAAAAAACAAQAAAABbQ29udGVudF9UeXBlc10ueG1sUEsBAhQAFAAAAAgAK2ZbXbVVMCPrAAAATAIAAAsAAAAAAAAAAAAAAIAB"
 value = value & "zAEAAF9yZWxzLy5yZWxzUEsBAhQAFAAAAAgAK2ZbXUKrl4PJAwAA5AkAAA8AAAAAAAAAAAAAAIAB4AIAAHhsL3dvcmtib29rLnhtbFBLAQIUABQAAAAIACtmW134wesVVQEAAIAFAAAaAAAAAAAAAAAAAACAAdYGAAB4bC9fcmVscy93b3Jr"
 value = value & "Ym9vay54bWwucmVsc1BLAQIUABQAAAAIACtmW10a7ujUKgIAAJ4EAAAYAAAAAAAAAAAAAACAAWMIAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54bWxQSwECFAAUAAAACAArZltdNOqFBKkBAACUAwAAGAAAAAAAAAAAAAAAgAHDCgAAeGwvd29y"
 value = value & "a3NoZWV0cy9zaGVldDIueG1sUEsBAhQAFAAAAAgAK2ZbXUkpIuTuBgAAEyIAABMAAAAAAAAAAAAAAIABogwAAHhsL3RoZW1lL3RoZW1lMS54bWxQSwECFAAUAAAACAArZltdaAZYL5oCAABrBgAADQAAAAAAAAAAAAAAgAHBEwAAeGwvc3R5"
 value = value & "bGVzLnhtbFBLAQIUABQAAAAIACtmW10jGDVdggEAALECAAAPAAAAAAAAAAAAAACAAYYWAAB4bC9tZXRhZGF0YS54bWxQSwECFAAUAAAACAArZltdz/7ePL0AAAAEAQAAHAAAAAAAAAAAAAAAgAE1GAAAeGwvcmljaERhdGEvcmljaFZhbHVl"
 value = value & "UmVsLnhtbFBLAQIUABQAAAAIACtmW10ybSmGmgAAALsAAAAbAAAAAAAAAAAAAACAASwZAAB4bC9yaWNoRGF0YS9yZHJpY2h2YWx1ZS54bWxQSwECFAAUAAAACAArZltdsgMWTMQAAAACAQAAJAAAAAAAAAAAAAAAgAH/GQAAeGwvcmljaERh"
 value = value & "dGEvcmRyaWNodmFsdWVzdHJ1Y3R1cmUueG1sUEsBAhQAFAAAAAgAK2ZbXUP6hnVlAQAAowQAACAAAAAAAAAAAAAAAIABBRsAAHhsL3JpY2hEYXRhL3JkUmljaFZhbHVlVHlwZXMueG1sUEsBAhQAFAAAAAgAK2ZbXboqZQU8AQAAdQIAABEA"
 value = value & "AAAAAAAAAAAAAIABqBwAAGRvY1Byb3BzL2NvcmUueG1sUEsBAhQAFAAAAAgAK2ZbXd96MjiRAQAANwMAABAAAAAAAAAAAAAAAIABEx4AAGRvY1Byb3BzL2FwcC54bWxQSwECFAAUAAAACAArZltdLyzzyLcAAAAkAQAAJwAAAAAAAAAAAAAA"
 value = value & "gAHSHwAAeGwvcmljaERhdGEvX3JlbHMvcmljaFZhbHVlUmVsLnhtbC5yZWxzUEsFBgAAAAAQABAASQQAAM4gAAAAAA=="
 HAMU_PhotoSeed = value
End Function

Private Function HAMU_PhotoPackageScript() As String
 Dim value As String
 value = value & "$ErrorActionPreference='Stop'" & vbLf
 value = value & "Add-Type -AssemblyName System.IO.Compression" & vbLf
 value = value & "$paths=@([IO.File]::ReadAllLines($manifest,[Text.Encoding]::Unicode))" & vbLf
 value = value & "if($paths.Count -eq 0){throw 'No images'}" & vbLf
 value = value & "$seedStream=New-Object IO.MemoryStream(,[Convert]::FromBase64String($seed64))" & vbLf
 value = value & "$seedZip=New-Object IO.Compression.ZipArchive($seedStream,[IO.Compression.ZipArchiveMode]::Read,$false)" & vbLf
 value = value & "$fileStream=[IO.File]::Open($output,[IO.FileMode]::CreateNew,[IO.FileAccess]::ReadWrite)" & vbLf
 value = value & "$zip=New-Object IO.Compression.ZipArchive($fileStream,[IO.Compression.ZipArchiveMode]::Create,$false)" & vbLf
 value = value & "function Add-TextEntry([string]$name,[string]$text){" & vbLf
 value = value & " $entry=$zip.CreateEntry($name,[IO.Compression.CompressionLevel]::NoCompression)" & vbLf
 value = value & " $writer=New-Object IO.StreamWriter($entry.Open(),[Text.UTF8Encoding]::new($false))" & vbLf
 value = value & " try{$writer.Write($text)}finally{$writer.Dispose()}" & vbLf
 value = value & "}" & vbLf
 value = value & "try{" & vbLf
 value = value & " $dynamic=@('[Content_Types].xml','xl/worksheets/sheet1.xml','xl/metadata.xml','xl/richData/rdrichvalue.xml','xl/richData/richValueRel.xml','xl/richData/_rels/richValueRel.xml.rels')" & vbLf
 value = value & " foreach($entry in $seedZip.Entries){" & vbLf
 value = value & "  if($dynamic -contains $entry.FullName -or $entry.FullName.StartsWith('xl/media/')){continue}" & vbLf
 value = value & "  $dest=$zip.CreateEntry($entry.FullName,[IO.Compression.CompressionLevel]::NoCompression)" & vbLf
 value = value & "  $srcStream=$entry.Open();$destStream=$dest.Open()" & vbLf
 value = value & "  try{$srcStream.CopyTo($destStream)}finally{$srcStream.Dispose();$destStream.Dispose()}" & vbLf
 value = value & " }" & vbLf
 value = value & " $reader=New-Object IO.StreamReader($seedZip.GetEntry('[Content_Types].xml').Open())" & vbLf
 value = value & " try{[xml]$types=$reader.ReadToEnd()}finally{$reader.Dispose()}" & vbLf
 value = value & " $extensions=@{}" & vbLf
 value = value & " foreach($item in $types.DocumentElement.ChildNodes){if($item.LocalName -eq 'Default'){$extensions[$item.Extension]=$true}}" & vbLf
 value = value & " $mime=@{png='image/png';jpg='image/jpeg';jpeg='image/jpeg';bmp='image/bmp';tif='image/tiff';tiff='image/tiff';gif='image/gif';webp='image/webp'}" & vbLf
 value = value & " $rows=New-Object Text.StringBuilder" & vbLf
 value = value & " $future=New-Object Text.StringBuilder" & vbLf
 value = value & " $metadata=New-Object Text.StringBuilder" & vbLf
 value = value & " $values=New-Object Text.StringBuilder" & vbLf
 value = value & " $rels=New-Object Text.StringBuilder" & vbLf
 value = value & " $images=New-Object Text.StringBuilder" & vbLf
 value = value & " for($i=0;$i -lt $paths.Count;$i++){" & vbLf
 value = value & "  $number=$i+1;$extension=[IO.Path]::GetExtension($paths[$i]).TrimStart('.').ToLowerInvariant()" & vbLf
 value = value & "  if(-not $mime.ContainsKey($extension)){throw 'Unsupported image format'}" & vbLf
 value = value & "  if(-not $extensions.ContainsKey($extension)){" & vbLf
 value = value & "   $node=$types.CreateElement('Default',$types.DocumentElement.NamespaceURI)" & vbLf
 value = value & "   $node.SetAttribute('Extension',$extension);$node.SetAttribute('ContentType',$mime[$extension])" & vbLf
 value = value & "   [void]$types.DocumentElement.AppendChild($node);$extensions[$extension]=$true" & vbLf
 value = value & "  }" & vbLf
 value = value & "  $entry=$zip.CreateEntry(""xl/media/image$number.$extension"",[IO.Compression.CompressionLevel]::NoCompression)" & vbLf
 value = value & "  $sourceStream=[IO.File]::OpenRead($paths[$i]);$destStream=$entry.Open()" & vbLf
 value = value & "  try{$sourceStream.CopyTo($destStream)}finally{$sourceStream.Dispose();$destStream.Dispose()}" & vbLf
 value = value & "  [void]$rows.Append(('<row r=""{0}""><c r=""A{0}"" t=""e"" vm=""{0}""><v>#VALUE!</v></c></row>' -f $number))" & vbLf
 value = value & "  [void]$future.Append(('<bk><extLst><ext uri=""{{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}}""><xlrd:rvb i=""{0}""/></ext></extLst></bk>' -f $i))" & vbLf
 value = value & "  [void]$metadata.Append(('<bk><rc t=""1"" v=""{0}""/></bk>' -f $i))" & vbLf
 value = value & "  [void]$values.Append(('<rv s=""0""><v>{0}</v><v>5</v></rv>' -f $i))" & vbLf
 value = value & "  [void]$rels.Append(('<Relationship Id=""rId{0}"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"" Target=""../media/image{0}.{1}""/>' -f $number,$extension))" & vbLf
 value = value & "  [void]$images.Append(('<rel r:id=""rId{0}""/>' -f $number))" & vbLf
 value = value & " }" & vbLf
 value = value & " $count=$paths.Count" & vbLf
 value = value & " Add-TextEntry '[Content_Types].xml' $types.OuterXml" & vbLf
 value = value & " Add-TextEntry 'xl/worksheets/sheet1.xml' ('<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><dimension ref=""A1:A'+$count+'""/><sheetData>'+$rows.ToString()+'</sheetData></worksheet>')" & vbLf
 value = value & " Add-TextEntry 'xl/metadata.xml' ('<metadata xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:xlrd=""http://schemas.microsoft.com/office/spreadsheetml/2017/richdata""><metadataTypes count=""1""><metadataType name=""XLRICHVALUE"" minSupportedVersion=""120000"" copy=""1"" pasteAll=""1"" pasteValues=""1"" merge=""1"" splitFirst=""1"" rowColShift=""1"" clearFormats=""1"" clearComments=""1"" assign=""1"" coerce=""1""/></metadataTypes><futureMetadata name=""XLRICHVALUE"" count=""'+$count+'"">'+$future.ToString()+'</futureMetadata><valueMetadata count=""'+$count+'"">'+$metadata.ToString()+'</valueMetadata></metadata>')" & vbLf
 value = value & " Add-TextEntry 'xl/richData/rdrichvalue.xml' ('<rvData xmlns=""http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"" count=""'+$count+'"">'+$values.ToString()+'</rvData>')" & vbLf
 value = value & " Add-TextEntry 'xl/richData/richValueRel.xml' ('<richValueRels xmlns=""http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">'+$images.ToString()+'</richValueRels>')" & vbLf
 value = value & " Add-TextEntry 'xl/richData/_rels/richValueRel.xml.rels' ('<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">'+$rels.ToString()+'</Relationships>')" & vbLf
 value = value & "}finally{$zip.Dispose();$fileStream.Dispose();$seedZip.Dispose();$seedStream.Dispose()}" & vbLf
 HAMU_PhotoPackageScript = value
End Function
