---
---

# sample

```vba
Sub FilterAndCopyAToB()
    Dim ws As Worksheet
    Dim tbl As Range
    Dim cell As Range
    
    Set ws = ActiveSheet
    Set tbl = ws.Range("A1:E100") ' 表の範囲を指定
    
    ' フィルタ適用（例：C列が"東京"の行を抽出）
    tbl.AutoFilter Field:=3, Criteria1:="東京"
    
    ' フィルタ後の表示行のA列をB列にコピー
    For Each cell In tbl.Columns(1).SpecialCells(xlCellTypeVisible)
        If cell.Row > 1 Then
            ws.Cells(cell.Row, 2).Value = cell.Value
        End If
    Next cell
    
    ' フィルタ解除（必要に応じて）
    ' ws.AutoFilterMode = False
    
    MsgBox "コピー完了"
End Sub
```
