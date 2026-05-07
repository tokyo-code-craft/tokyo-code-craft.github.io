---
---

# sample

```vba
Sub ImportTSV_Fast()
    Dim filePath As String
    Dim fileNum As Integer
    Dim allText As String
    Dim lines() As String
    Dim fields() As String
    Dim result() As String
    Dim i As Long, j As Long
    Dim rowCount As Long, colCount As Long
    Dim ws As Worksheet

    filePath = Application.GetOpenFilename("TSVファイル,*.tsv,テキスト,*.txt")
    If filePath = "False" Then Exit Sub

    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' ① ファイルを一括読み込み
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
        allText = Space(LOF(fileNum))
        Get #fileNum, , allText
    Close #fileNum

    ' ② 行・列に分割
    lines = Split(allText, vbCrLf)
    ' 末尾の空行対策
    If lines(UBound(lines)) = "" Then ReDim Preserve lines(UBound(lines) - 1)
    
    rowCount = UBound(lines) + 1
    fields = Split(lines(0), vbTab)
    colCount = UBound(fields) + 1

    ' ③ 2次元配列に格納
    ReDim result(1 To rowCount, 1 To colCount)
    For i = 0 To rowCount - 1
        fields = Split(lines(i), vbTab)
        For j = 0 To UBound(fields)
            result(i + 1, j + 1) = fields(j)
        Next j
    Next i

    ' ④ 配列をシートに一括貼り付け（最速）
    ws.Cells.Clear
    ws.Range("A1").Resize(rowCount, colCount).Value = result

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "完了：" & rowCount & "行 × " & colCount & "列"
End Sub

```