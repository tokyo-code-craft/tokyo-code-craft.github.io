---
---

# sample

```vba
' 指定セルを起点に1グループ（ticket行＋子行）を返す
' 戻り値：Rangeオブジェクト（グループ全体）
Function GetGroup(startCell As Range) As Range

    Dim ws       As Worksheet
    Dim typeCol  As Long
    Dim lastRow  As Long
    Dim i        As Long

    Set ws      = startCell.Worksheet
    typeCol     = startCell.Column
    lastRow     = ws.Cells(ws.Rows.Count, typeCol).End(xlUp).Row

    ' startCellがticketでなければ何も返さない
    If startCell.Value <> "ticket" Then
        Set GetGroup = Nothing
        Exit Function
    End If

    ' ticket行の次の行から、次のticket or 末尾まで走査
    Dim groupEnd As Long
    groupEnd = startCell.Row  ' 最低でもticket行自身

    For i = startCell.Row + 1 To lastRow
        If ws.Cells(i, typeCol).Value = "ticket" Then Exit For
        groupEnd = i
    Next i

    Set GetGroup = ws.Range(ws.Cells(startCell.Row, typeCol), _
                            ws.Cells(groupEnd, typeCol))

End Function
```
