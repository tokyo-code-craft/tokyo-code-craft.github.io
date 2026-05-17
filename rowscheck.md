---
---

# sample

```vba
Sub DeleteBatsuWithErrorCheck()

    Dim ws       As Worksheet
    Dim wsLog    As Worksheet
    Dim typeCol  As Long
    Dim lastRow  As Long
    Dim i        As Long
    Dim logRow   As Long

    Set ws     = ActiveSheet
    typeCol    = 1   ' ← タイプ列（A列=1）、必要に応じて変更
    lastRow    = ws.Cells(ws.Rows.Count, typeCol).End(xlUp).Row

    ' ログシートの準備
    If Not SheetExists("エラーログ") Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "エラーログ"
    End If
    Set wsLog = Worksheets("エラーログ")
    wsLog.Cells.Clear
    wsLog.Cells(1, 1).Value = "行番号"
    wsLog.Cells(1, 2).Value = "内容"
    wsLog.Cells(1, 3).Value = "理由"
    logRow = 2

    ' グループ収集と判定（下から削除するため削除対象をリストアップ）
    Dim deleteRows() As Long
    Dim deleteCount  As Long
    deleteCount = 0
    ReDim deleteRows(1 To lastRow)

    i = 2
    Do While i <= lastRow

        If ws.Cells(i, typeCol).Value = "ticket" Then

            ' グループの子行を収集
            Dim groupStart As Long
            Dim groupEnd   As Long
            Dim j          As Long
            Dim firstChild As String

            groupStart = i + 1
            groupEnd   = groupStart - 1
            firstChild = ""

            j = groupStart
            Do While j <= lastRow
                If ws.Cells(j, typeCol).Value = "ticket" Then Exit Do
                If firstChild = "" Then firstChild = ws.Cells(j, typeCol).Value
                groupEnd = j
                j = j + 1
            Loop

            ' 子行がない場合はスキップ
            If groupEnd >= groupStart Then

                If firstChild = "×" Then
                    ' エラーグループ：ログに記録
                    Dim k As Long
                    For k = groupStart To groupEnd
                        wsLog.Cells(logRow, 1).Value = k
                        wsLog.Cells(logRow, 2).Value = ws.Cells(k, typeCol).Value
                        wsLog.Cells(logRow, 3).Value = "×→○の順（エラー）"
                        logRow = logRow + 1
                    Next k

                ElseIf firstChild = "○" Then
                    ' 正常グループ：×行を削除リストに追加
                    For k = groupStart To groupEnd
                        If ws.Cells(k, typeCol).Value = "×" Then
                            deleteCount = deleteCount + 1
                            deleteRows(deleteCount) = k
                        End If
                    Next k
                End If

            End If

            i = groupEnd + 1

        Else
            i = i + 1
        End If

    Loop

    ' 削除は下から（行ズレ防止）
    Dim d As Long
    For d = deleteCount To 1 Step -1
        ws.Rows(deleteRows(d)).Delete
    Next d

    ' 結果報告
    Dim msg As String
    msg = deleteCount & " 行の×を削除しました。"
    If logRow > 2 Then
        msg = msg & vbCrLf & (logRow - 2) & " 行のエラーを「エラーログ」シートに記録しました。"
    End If
    MsgBox msg, vbInformation

End Sub

' --------------------------------
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
```
