---
---

# sample

```vba
' ============================================================
' 列定数定義
' ============================================================
Private Const COL_UNIQUE_ID     As Integer = 1  ' A列: UniqueId
Private Const COL_GROUP_ID      As Integer = 2  ' B列: groupId
Private Const COL_TYPE          As Integer = 3  ' C列: タイプ
Private Const COL_QTY           As Integer = 4  ' D列: 数
Private Const COL_NAME          As Integer = 5  ' E列: ticket名 または 商品名
Private Const COL_PRICE         As Integer = 6  ' F列: 割引対象商品（ticket行） または 金額（商品行）


' ============================================================
' 呼び出し元のイメージ
' groupRangeを走査してticketRowを特定し、SplitGroupに渡す
' ============================================================
Sub CallerExample(groupRange As Range)

    Dim ticketRow As Range
    Dim r As Range
    For Each r In groupRange.Rows
        If r.Cells(1, COL_TYPE).Value = "ticket" Then
            Set ticketRow = r
        End If
    Next r

    ' ticket行が見つからない場合はスキップ
    If ticketRow Is Nothing Then Exit Sub

    SplitGroup groupRange, ticketRow

End Sub


' ============================================================
' グループ分割処理
' 呼び出し元から groupRange（1グループ分のRange）と
' ticketRow（ticket行のRange）を受け取る
' ============================================================
Sub SplitGroup(groupRange As Range, ticketRow As Range)

    Dim ws As Worksheet
    Set ws = groupRange.Worksheet

    ' --- 1. ticket行に一致する商品行をプールに収集 ---
    ' ticket行のF列（割引対象商品）と商品行のF列（金額）が一致する行のみ収集する
    '
    ' 【注意】未初期化の動的配列をByRefで渡してReDimする挙動は環境によって
    '         不安定になる場合がある。CollectProductRows呼び出し後に
    '         productCountが正しく取得できているか動作確認が必要。
    Dim ticketPrice As String
    ticketPrice = ticketRow.Cells(1, COL_PRICE).Value

    Dim productRows() As Range
    Dim productQtys() As Integer
    Dim productCount As Integer
    CollectProductRows groupRange, ticketPrice, productRows, productQtys, productCount

    ' マッチする商品行が見つからない場合はスキップ
    If productCount = 0 Then Exit Sub

    ' --- 2. ticket数を取得 ---
    Dim ticketNum As Integer
    ticketNum = ticketRow.Cells(1, COL_QTY).Value

    ' ticket数が1以下なら分割不要
    If ticketNum <= 1 Then Exit Sub

    ' --- 3. 新規グループIDの開始値を決定 ---
    ' シート全体の現在の最大groupIdの次の値から採番する
    Dim nextGroupId As Integer
    nextGroupId = GetMaxGroupId(ws) + 1

    ' --- 4. 元グループのticket数を1に書き換え ---
    ticketRow.Cells(1, COL_QTY).Value = 1

    ' --- 5. 元グループへの商品割り当て ---
    ' 商品プールの先頭から1つ消費して元グループの商品数を1にする
    ' 残数が0になった行は後で削除するためproductQtysで管理する
    Dim poolIndex As Integer
    poolIndex = 0  ' 商品プールの現在位置（上から順にマッチング）

    ' 元グループの商品は先頭商品行から1消費
    productQtys(poolIndex) = productQtys(poolIndex) - 1
    productRows(poolIndex).Cells(1, COL_QTY).Value = 1  ' 元グループに残す数=1

    ' 消費後に残数0なら次の商品へ進める
    If productQtys(poolIndex) = 0 Then poolIndex = poolIndex + 1

    ' --- 6. 新グループを追加 ---
    ' ticket数-1 個の新グループを追加（元グループが1つ目になる）
    ' 新グループは常にgroupRangeの真下のinsertRowに挿入する
    ' 挿入のたびに直前の行が下にずれるため、常に同じinsertRowに挿入し続ければよい
    ' → 関数終了後にgroupId昇順＋ticketが先頭になるキーで並べ替えることで整列する
    Dim insertRow As Long
    insertRow = groupRange.Row + groupRange.Rows.Count

    Dim i As Integer
    For i = 1 To ticketNum - 1

        ' 商品プールが枯渇していたらスキップ（商品不足）
        If poolIndex >= productCount Then Exit For

        ' ticket行・商品行をinsertRowに挿入（順不同で問題なし、並べ替えで整列する）
        InsertGroupRow ws, ticketRow, insertRow, nextGroupId
        InsertGroupRow ws, productRows(poolIndex), insertRow, nextGroupId

        ' 消費を記録し残数0なら次の商品へ進める
        productQtys(poolIndex) = productQtys(poolIndex) - 1
        If productQtys(poolIndex) = 0 Then poolIndex = poolIndex + 1

        nextGroupId = nextGroupId + 1
    Next i

    ' --- 7. 残数0になった商品行をシートから削除 ---
    ' productQtys(j) = 0 の行は全消費済みなので削除する
    ' 下から削除することで行番号のずれを防ぐ
    Dim j As Integer
    For j = productCount - 1 To 0 Step -1
        If productQtys(j) = 0 Then
            productRows(j).Delete Shift:=xlUp
        End If
    Next j

End Sub


' ============================================================
' ユーティリティ：groupRangeからticketに一致する商品行をプールに収集する
' ticketPrice  : ticket行のF列（割引対象商品）の値
' productRows  : 一致した商品行のRangeを格納する配列（ByRefで呼び出し元に反映）
' productQtys  : 一致した商品行の初期残数を格納する配列（ByRefで呼び出し元に反映）
' productCount : 収集した商品行の件数（ByRefで呼び出し元に反映）
'
' 【注意】未初期化の動的配列をByRefで渡してReDimする挙動は環境によって
'         不安定になる場合がある。呼び出し後にproductCountが
'         正しく取得できているか動作確認が必要。
' ============================================================
Private Sub CollectProductRows(groupRange As Range, ticketPrice As String, _
                                productRows() As Range, productQtys() As Integer, _
                                productCount As Integer)
    productCount = 0
    Dim r As Range
    For Each r In groupRange.Rows
        If r.Cells(1, COL_TYPE).Value = "商品" Then
            If r.Cells(1, COL_PRICE).Value = ticketPrice Then
                ReDim Preserve productRows(productCount)
                ReDim Preserve productQtys(productCount)
                Set productRows(productCount) = r
                productQtys(productCount) = r.Cells(1, COL_QTY).Value
                productCount = productCount + 1
            End If
        End If
    Next r
End Sub


' ============================================================
' ユーティリティ：指定行にsourceRowをコピーして挿入し、groupIdと数を書き換える
' ============================================================
Private Sub InsertGroupRow(ws As Worksheet, sourceRow As Range, insertRow As Long, groupId As Integer)
    ws.Rows(insertRow).Insert Shift:=xlDown
    sourceRow.Copy ws.Cells(insertRow, 1)
    ws.Cells(insertRow, COL_GROUP_ID).Value = groupId
    ws.Cells(insertRow, COL_QTY).Value = 1
End Sub


' ============================================================
' ユーティリティ：シート全体から現在の最大groupIdを取得
' ============================================================
Private Function GetMaxGroupId(ws As Worksheet) As Integer
    Dim maxId As Integer
    maxId = 0
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_GROUP_ID).End(xlUp).Row
    Dim i As Long
    For i = 1 To lastRow
        Dim gid As Integer
        gid = ws.Cells(i, COL_GROUP_ID).Value
        If gid > maxId Then maxId = gid
    Next i
    GetMaxGroupId = maxId
End Function
```
