---
---

# 多重vlookup

```vba
Sub UpdateAmounts()
    Dim wsA As Worksheet, wsB As Worksheet, wsTbl As Worksheet
    Set wsA  = Sheets("A")
    Set wsB  = Sheets("B")
    Set wsTbl = Sheets("変換テーブル") ' 変換テーブルのシート

    Application.ScreenUpdating = False ' 画面更新を止める（体感速度UP）

    Dim lastRow As Long
    lastRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row

    Dim rngConv  As Range : Set rngConv  = wsTbl.Range("A:B") ' 変換テーブル範囲
    Dim rngPrice As Range : Set rngPrice = wsB.Range("B:C")   ' 金額テーブル範囲

    Dim i As Long
    For i = 2 To lastRow
        Dim keyA As String
        keyA = CStr(wsA.Cells(i, 1).Value)
        If keyA = "" Then GoTo Continue

        ' ① 変換テーブルに照合
        Dim converted As Variant
        converted = Application.VLookup(keyA, rngConv, 2, False)

        ' ② ヒットすれば変換後のキーを、なければ元のキーを使う
        Dim keyLookup As String
        If IsError(converted) Then
            keyLookup = keyA
        Else
            keyLookup = CStr(converted)
        End If

        ' ③ Bシートから金額取得
        Dim result As Variant
        result = Application.VLookup(keyLookup, rngPrice, 2, False)
        If Not IsError(result) Then
            wsA.Cells(i, 2).Value = result
        End If
Continue:
    Next i

    Application.ScreenUpdating = True
    MsgBox "更新完了"
End Sub
```
