---
---

# sample

```vba
Sub DeleteNextSubGroup()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' --- 列インデックス ---
    Const COL_GROUP_ID     As Integer = 1  ' groupId列
    Const COL_SUB_GROUP_ID As Integer = 2  ' 小groupId列
    Const COL_PRODUCT_NAME As Integer = 3  ' 商品名列
    Const COL_GROUP_COUNT  As Integer = 4  ' groupIdの数列
    Const COL_TYPE         As Integer = 5  ' 種類列（ticket / item）
    Const COL_DATE         As Integer = 6  ' 日付列
    ' ----------------------

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' =====================================================
    ' Step1: 全体を走査してgroupId・小groupId単位の辞書を構築
    '        groupDict(gId)(sgId) = ticket行のRange
    '        groupDict("1")("1") = ws.Rows(2)  → groupId=1, 小groupId=1のticket行
    '        groupDict("1")("2") = ws.Rows(4)  → groupId=1, 小groupId=2のticket行
    ' =====================================================
    Dim groupDict As Object
    Set groupDict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To lastRow
        If ws.Cells(i, COL_GROUP_COUNT).Value = 4 Then
            If ws.Cells(i, COL_TYPE).Value = "ticket" Then
                Dim gId As String
                Dim sgId As String
                gId  = CStr(ws.Cells(i, COL_GROUP_ID).Value)
                sgId = CStr(ws.Cells(i, COL_SUB_GROUP_ID).Value)

                If Not groupDict.Exists(gId) Then
                    Set groupDict(gId) = CreateObject("Scripting.Dictionary")
                End If

                ' ticket行のRangeをそのまま格納
                Set groupDict(gId)(sgId) = ws.Rows(i)
            End If
        End If
    Next i

    ' =====================================================
    ' Step2: groupDictをループして削除対象の小groupIdを特定
    '        同一groupId内の1枚目・2枚目のticketをそれぞれticketRange1・ticketRange2に格納して比較
    '        ※判定ロジックは要件に応じて変更すること
    '        　ticketRange1.Cells(1, COL_DATE)等で各列にアクセス可能
    '
    '        deleteTargets(gId) = 削除対象の小groupId
    '        deleteTargets("1") = "2"  → groupId=1の中で小groupId=2を削除
    '        deleteTargets("3") = "4"  → groupId=3の中で小groupId=4を削除
    ' =====================================================
    Dim deleteTargets As Object
    Set deleteTargets = CreateObject("Scripting.Dictionary")

    Dim gKey As Variant
    For Each gKey In groupDict.Keys
        Dim subDict As Object
        Set subDict = groupDict(gKey)

        ' 1枚目・2枚目のticket行をそれぞれ取得
        Dim sgKeys As Variant
        sgKeys = subDict.Keys

        Dim ticketRange1 As Range  ' 1枚目のticket行
        Dim ticketRange2 As Range  ' 2枚目のticket行
        Set ticketRange1 = subDict(sgKeys(0))
        Set ticketRange2 = subDict(sgKeys(1))

        ' ※ここの判定ロジックは要件に応じて変更すること
        ' 　例）日付が新しい方を削除
        If CDate(ticketRange1.Cells(1, COL_DATE).Value) > CDate(ticketRange2.Cells(1, COL_DATE).Value) Then
            deleteTargets(CStr(gKey)) = CStr(sgKeys(0))  ' 1枚目が新しい → 1枚目を削除
        Else
            deleteTargets(CStr(gKey)) = CStr(sgKeys(1))  ' 2枚目が新しい → 2枚目を削除
        End If
    Next gKey

    ' =====================================================
    ' Step3: AutoFilterで削除対象の小groupIdを一括抽出して削除
    '        deleteTargetsの値を配列にまとめてフィルター条件に渡す
    '        → 1回のフィルターで全削除対象をまとめて削除
    ' =====================================================
    ' deleteTargetsの値（削除対象の小groupId）を配列に変換
    Dim delSgIds() As String
    ReDim delSgIds(0 To deleteTargets.Count - 1)
    Dim idx As Integer
    idx = 0
    Dim delGKey As Variant
    For Each delGKey In deleteTargets.Keys
        delSgIds(idx) = deleteTargets(delGKey)
        idx = idx + 1
    Next delGKey

    ' AutoFilterで削除対象の小groupIdを抽出
    ws.AutoFilterMode = False
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, COL_DATE)).AutoFilter _
        Field:=COL_SUB_GROUP_ID, _
        Criteria1:=delSgIds, _
        Operator:=xlFilterValues

    ' 抽出された行をまとめて削除（ヘッダー行を除く）
    ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, COL_DATE)) _
        .SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ' フィルター解除
    ws.AutoFilterMode = False

    MsgBox "削除完了しました。", vbInformation
End Sub
```
