---
---

# filter

```vba
Sub DeleteRows_ArrayMethod()

    Dim ws          As Worksheet
    Dim srcData     As Variant
    Dim result()    As Variant
    Dim keepCount   As Long
    Dim totalRows   As Long
    Dim totalCols   As Long
    Dim i           As Long
    Dim j           As Long
    Dim destRow     As Long

    Set ws = ThisWorkbook.Sheets("Sheet1")  ' ←シート名を変更

    ' --- 設定 ---
    Const TARGET_COL As Long = 1             ' 削除判定する列番号（A列=1）
    Const DELETE_VAL As String = "削除対象"  ' 削除する値
    Const HEADER_ROW As Long = 1             ' ヘッダー行数

    ' --- 高速化設定 ---
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' --- 全データを配列に一括読み込み（ヘッダー含む） ---
    totalRows = ws.Cells(ws.Rows.Count, TARGET_COL).End(xlUp).Row
    totalCols = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    srcData = ws.Range(ws.Cells(1, 1), ws.Cells(totalRows, totalCols)).Value

    ' --- 残すデータ行数をカウント（ヘッダー除く） ---
    keepCount = 0
    For i = HEADER_ROW + 1 To totalRows
        If srcData(i, TARGET_COL) <> DELETE_VAL Then
            keepCount = keepCount + 1
        End If
    Next i

    ' --- データ行だけ result に詰める ---
    If keepCount > 0 Then
        ReDim result(1 To keepCount, 1 To totalCols)
        destRow = 1
        For i = HEADER_ROW + 1 To totalRows
            If srcData(i, TARGET_COL) <> DELETE_VAL Then
                For j = 1 To totalCols
                    result(destRow, j) = srcData(i, j)
                Next j
                destRow = destRow + 1
            End If
        Next i
    End If

    ' --- ヘッダー行より下だけクリア ---
    ws.Range(ws.Cells(HEADER_ROW + 1, 1), ws.Cells(totalRows, totalCols)).ClearContents

    ' --- データを書き戻し（ヘッダー行の次から） ---
    If keepCount > 0 Then
        ws.Cells(HEADER_ROW + 1, 1).Resize(keepCount, totalCols).Value = result
    End If

    ' --- 高速化設定を戻す ---
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "完了：" & (totalRows - HEADER_ROW - keepCount) & " 行削除 / " & keepCount & " 行残存"

End Sub
```
