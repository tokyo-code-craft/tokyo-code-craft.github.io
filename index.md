---
---

# Excel ①_ThisWorkbook

```vba
'==========================================================================
' 【貼り付け先】VBAエディタ左ツリーの「ThisWorkbook」をダブルクリック
'==========================================================================

' -----------------------------------------------------------------------
' Workbook_Open：ファイルを開いたときに自動実行されるイベントプロシージャ
' -----------------------------------------------------------------------
Private Sub Workbook_Open()

    ' 当月のシート名を組み立てる（例：「2026年4月」）
    Dim wsName As String
    wsName = Year(Now()) & "年" & Month(Now()) & "月"

    ' 当月シートがまだ存在しない場合のみ新規作成する
    ' ※毎月初めてファイルを開いたときだけ生成される
    If Not SheetExists(wsName) Then
        CreateMonthSheet Year(Now()), Month(Now())
    End If

    ' サンプルシートと祝日シートを書き込み禁止にする
    ' UserInterfaceOnly:=True にすることでVBAからの操作は引き続き可能
    Worksheets("サンプル").Protect Password:="", UserInterfaceOnly:=True
    Worksheets("祝日").Protect Password:="", UserInterfaceOnly:=True

    ' 当月シートをアクティブ（最前面）にして表示する
    Worksheets(wsName).Activate

End Sub

' -----------------------------------------------------------------------
' SheetExists：指定した名前のシートがブック内に存在するか確認する関数
' 引数  sheetName : 確認したいシート名
' 戻り値           : 存在すれば True、存在しなければ False
' -----------------------------------------------------------------------
Function SheetExists(sheetName As String) As Boolean

    Dim ws As Worksheet

    ' 存在しないシートを参照するとエラーになるため、エラーを無視して取得を試みる
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0

    ' ws に何かセットされていれば存在する
    SheetExists = Not ws Is Nothing

End Function
```

# ②_Module1_標準モジュール
```vba
'==========================================================================
' 【貼り付け先】VBAエディタで「挿入 → 標準モジュール」を追加し、そこに貼り付け
'==========================================================================

Option Explicit  ' 変数の宣言を強制する（未宣言変数によるバグを防ぐ）

' -----------------------------------------------------------------------
' CreateCurrentMonthSheet：当月シートを作成するショートカット
' 使い方： Alt+F8 → このプロシージャを選択して実行
' -----------------------------------------------------------------------
Sub CreateCurrentMonthSheet()
    CreateMonthSheet Year(Now()), Month(Now())
End Sub

' -----------------------------------------------------------------------
' CreateNextMonthSheet：翌月シートを作成するショートカット
' 使い方： Alt+F8 → このプロシージャを選択して実行
' -----------------------------------------------------------------------
Sub CreateNextMonthSheet()
    ' DateSerial で翌月1日を算出（12月→翌年1月も正しく処理される）
    Dim nextMonth As Date
    nextMonth = DateSerial(Year(Now()), Month(Now()) + 1, 1)
    CreateMonthSheet Year(nextMonth), Month(nextMonth)
End Sub

' -----------------------------------------------------------------------
' SetupSampleCheckboxes：サンプルシートにチェックボックスを配置する
' 使い方： Alt+F8 → このプロシージャを選択して実行（初回1回だけでOK）
' ※ サンプルシートはコピー元なので、実際の入力には使わないこと
' -----------------------------------------------------------------------
Sub SetupSampleCheckboxes()

    ' ── 定数（CreateMonthSheet と同じレイアウト） ────────────────
    Const ROW_FIRST  As Long = 4   ' チェック項目の開始行
    Const ROW_LAST   As Long = 11  ' チェック項目の最終行（項目数8：4〜11行目）
    Const COL_START  As Long = 2   ' B列（1日目）

    If Not SheetExists("サンプル") Then
        MsgBox "「サンプル」シートが見つかりません。", vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Worksheets("サンプル")
    ws.Unprotect Password:=""

    ' 既存のチェックボックスをすべて削除してから再配置する
    Dim cb As CheckBox
    For Each cb In ws.CheckBoxes
        cb.Delete
    Next cb

    ' サンプルシートは2026年1月（31日）で固定
    ' 土日・祝日以外の平日セルにチェックボックスを配置する
    Dim holidays As Collection
    Set holidays = LoadHolidays(2026, 1)  ' 祝日シートから2026年1月の祝日を取得

    Dim d As Integer
    Dim row As Long
    Dim col As Long

    For d = 1 To 31  ' 1月は31日

        Dim curDate As Date
        curDate = DateSerial(2026, 1, d)
        Dim wd As Integer
        wd = Weekday(curDate, vbSunday)

        Dim isSat  As Boolean: isSat  = (wd = 7)
        Dim isSun  As Boolean: isSun  = (wd = 1)
        Dim isHol  As Boolean: isHol  = IsHoliday(curDate, holidays)
        Dim isRest As Boolean: isRest = isSat Or isSun Or isHol

        col = COL_START + d - 1

        For row = ROW_FIRST To ROW_LAST

            Dim targetCell As Range
            Set targetCell = ws.Cells(row, col)

            If Not isRest Then
                ' 平日セルのみチェックボックスをセル中央に配置する
                Const CB_SIZE2 As Double = 12
                Dim cbLeft2 As Double
                Dim cbTop2  As Double
                cbLeft2 = targetCell.Left + (targetCell.Width  - CB_SIZE2) / 2
                cbTop2  = targetCell.Top  + (targetCell.Height - CB_SIZE2) / 2

                Dim newCb As CheckBox
                Set newCb = ws.CheckBoxes.Add( _
                    cbLeft2, cbTop2, CB_SIZE2, CB_SIZE2)

                newCb.Caption     = ""                 ' ラベル文字なし
                newCb.Value       = xlOff              ' 初期状態：未チェック
                newCb.LinkedCell  = targetCell.Address ' セルに連動
                newCb.PrintObject = True               ' 印刷時にも表示

                ' LinkedCell のTRUE/FALSEテキストを非表示にする
                targetCell.Font.Color = targetCell.Interior.Color
            End If

        Next row
    Next d

    ' シート保護を再設定
    ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True

    MsgBox "サンプルシートにチェックボックスを配置しました！", vbInformation, "セットアップ完了"

End Sub

' -----------------------------------------------------------------------
' LoadHolidays：「祝日」シートから対象月の祝日日付をCollectionに読み込む
' 引数  yr : 対象年
'       mo : 対象月
' 戻り値   : 祝日の日付文字列を格納した Collection
'            （IsHoliday関数でキーとして使うため "yyyy/mm/dd" 形式で格納）
' -----------------------------------------------------------------------
Function LoadHolidays(yr As Integer, mo As Integer) As Collection

    Dim col As New Collection

    ' 「祝日」シートが存在しない場合は空のCollectionを返す
    If Not SheetExists("祝日") Then
        Set LoadHolidays = col
        Exit Function
    End If

    Dim wsH As Worksheet
    Set wsH = Worksheets("祝日")

    Dim lastRow As Long
    ' A列の最終行を取得（ヘッダー行を除く）
    lastRow = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow  ' 2行目からデータ行

        Dim cellVal As Variant
        cellVal = wsH.Cells(i, 1).Value

        ' 日付型セルのみ処理（空白・文字列は無視）
        If IsDate(cellVal) Then
            Dim hDate As Date
            hDate = CDate(cellVal)

            ' 対象年月の祝日のみCollectionに追加する
            If Year(hDate) = yr And Month(hDate) = mo Then
                ' キーとして "yyyy/mm/dd" 文字列を使い、重複エラーを回避する
                Dim key As String
                key = Format(hDate, "yyyy/mm/dd")
                On Error Resume Next
                col.Add hDate, key  ' 同じ日付が重複登録されてもエラーにしない
                On Error GoTo 0
            End If
        End If

    Next i

    Set LoadHolidays = col

End Function

' -----------------------------------------------------------------------
' IsHoliday：指定した日付が祝日Collectionに含まれているか判定する関数
' 引数  checkDate : 確認したい日付
'       holidays  : LoadHolidays で取得した Collection
' 戻り値           : 祝日なら True、違えば False
' -----------------------------------------------------------------------
Function IsHoliday(checkDate As Date, holidays As Collection) As Boolean

    Dim key As String
    key = Format(checkDate, "yyyy/mm/dd")

    ' CollectionはキーでのExists確認ができないため、On Errorで代用する
    Dim dummy As Date
    On Error Resume Next
    dummy = holidays(key)
    IsHoliday = (Err.Number = 0)  ' エラーなし＝キーが存在する＝祝日
    On Error GoTo 0

End Function

' -----------------------------------------------------------------------
' CreateMonthSheet：指定した年・月のチェックシートを生成するメインプロシージャ
' 引数  yr : 生成したいシートの年（例：2026）
'       mo : 生成したいシートの月（例：5）
' -----------------------------------------------------------------------
Sub CreateMonthSheet(yr As Integer, mo As Integer)

    ' ── レイアウト定数 ──────────────────────────────────────────────
    Const ROW_TITLE  As Long = 1   ' タイトル行
    Const ROW_DATE   As Long = 2   ' 日付ヘッダー行
    Const ROW_WEEK   As Long = 3   ' 曜日ヘッダー行
    Const ROW_FIRST  As Long = 4   ' チェック項目の開始行
    Const COL_LABEL  As Long = 1   ' A列（チェック項目名）
    Const COL_START  As Long = 2   ' B列（1日目）

    ' ── チェック項目（追加・変更はここだけ修正すればOK） ────────────
    Dim checkItems(0 To 7) As String
    checkItems(0) = "朝礼・申し送り確認"
    checkItems(1) = "メール・チャット確認"
    checkItems(2) = "スケジュール確認"
    checkItems(3) = "タスク進捗更新"
    checkItems(4) = "書類・資料提出"
    checkItems(5) = "顧客対応ログ記録"
    checkItems(6) = "経費・申請処理"
    checkItems(7) = "終業報告・引継ぎ"

    Dim itemCount As Integer
    itemCount = 8  ' checkItems の要素数と合わせる

    ' ── 変数宣言 ────────────────────────────────────────────────────
    Dim wsNew       As Worksheet
    Dim wsName      As String
    Dim daysInMonth As Integer
    Dim col         As Long
    Dim row         As Long
    Dim d           As Integer
    Dim wd          As Integer   ' 曜日番号（vbSunday基準：1=日〜7=土）
    Dim lastRow     As Long

    wsName  = yr & "年" & mo & "月"
    lastRow = ROW_FIRST + itemCount - 1

    ' ── 既存シートの確認 ─────────────────────────────────────────
    If SheetExists(wsName) Then
        MsgBox wsName & " のシートはすでに存在します。", vbInformation
        Worksheets(wsName).Activate
        Exit Sub
    End If

    ' ── サンプルシートの存在確認 ──────────────────────────────────
    If Not SheetExists("サンプル") Then
        MsgBox "「サンプル」シートが見つかりません。削除しないでください。", vbCritical
        Exit Sub
    End If

    ' ── 祝日リストを「祝日」シートから読み込む ────────────────────
    ' 祝日シートがない場合は空のCollectionが返るだけで処理は続行する
    Dim holidays As Collection
    Set holidays = LoadHolidays(yr, mo)

    ' ── サンプルシートをコピーしてリネーム ───────────────────────
    Worksheets("サンプル").Copy After:=Worksheets(Worksheets.Count)
    Set wsNew = ActiveSheet
    wsNew.Name = wsName
    wsNew.Unprotect Password:=""

    ' ── 対象月の日数を取得 ────────────────────────────────────────
    daysInMonth = Day(DateSerial(yr, mo + 1, 0))

    ' ── タイトル更新 ──────────────────────────────────────────────
    wsNew.Cells(ROW_TITLE, COL_LABEL).Value = wsName & "　業務チェックシート"

    ' ── 曜日の日本語表記（vbSunday基準：1=日〜7=土） ────────────
    Dim weekdayJP(1 To 7) As String
    weekdayJP(1) = "日": weekdayJP(2) = "月": weekdayJP(3) = "火"
    weekdayJP(4) = "水": weekdayJP(5) = "木": weekdayJP(6) = "金"
    weekdayJP(7) = "土"

    ' ── 既存の日付列をリセット ────────────────────────────────────
    wsNew.Range(wsNew.Cells(ROW_DATE, COL_START), _
                wsNew.Cells(lastRow, COL_START + 30)).ClearContents
    wsNew.Range(wsNew.Cells(ROW_DATE, COL_START), _
                wsNew.Cells(lastRow, COL_START + 30)).ClearFormats
    wsNew.Range(wsNew.Cells(ROW_DATE, COL_START), _
                wsNew.Cells(lastRow, COL_START + 30)).Interior.ColorIndex = 0

    ' ── 既存チェックボックスを削除（サンプルの残骸） ─────────────
    Dim cb As CheckBox
    For Each cb In wsNew.CheckBoxes
        cb.Delete
    Next cb

    ' ── 日ごとの処理（ヘッダー設定 + チェックボックス配置） ────────
    For d = 1 To daysInMonth

        Dim curDate As Date
        curDate = DateSerial(yr, mo, d)
        wd  = Weekday(curDate, vbSunday)
        col = COL_START + d - 1

        Dim isSat  As Boolean: isSat  = (wd = 7)
        Dim isSun  As Boolean: isSun  = (wd = 1)
        Dim isHol  As Boolean: isHol  = IsHoliday(curDate, holidays)  ' 祝日判定
        Dim isRest As Boolean: isRest = isSat Or isSun Or isHol       ' 休日判定（土日＋祝日）

        ' -------- 日付ヘッダー（ROW_DATE 行）の書式設定 --------
        With wsNew.Cells(ROW_DATE, col)
            .Value        = DateSerial(yr, mo, d)  ' 日付型で格納
            .NumberFormat = "d"                     ' 表示は日のみ
            .Font.Name = "メイリオ": .Font.Size = 9: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            If isSat Then
                ' 土曜：青系
                .Interior.Color = RGB(169, 196, 232)
                .Font.Color     = RGB(31, 78, 121)
            ElseIf isSun Then
                ' 日曜：赤系
                .Interior.Color = RGB(244, 204, 204)
                .Font.Color     = RGB(139, 0, 0)
            ElseIf isHol Then
                ' 祝日：赤系（日曜と同じ色でわかりやすく）
                .Interior.Color = RGB(244, 204, 204)
                .Font.Color     = RGB(139, 0, 0)
            Else
                ' 平日：濃紺
                .Interior.Color = RGB(47, 84, 150)
                .Font.Color     = RGB(255, 255, 255)
            End If
        End With

        ' -------- 曜日ヘッダー（ROW_WEEK 行）の書式設定 --------
        With wsNew.Cells(ROW_WEEK, col)
            .Value = weekdayJP(wd)
            .Font.Name = "メイリオ": .Font.Size = 9: .Font.Bold = True
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            If isSat Then
                .Interior.Color = RGB(169, 196, 232)
                .Font.Color     = RGB(31, 78, 121)
            ElseIf isSun Or isHol Then
                ' 日曜・祝日：赤系
                .Interior.Color = RGB(244, 204, 204)
                .Font.Color     = RGB(139, 0, 0)
            Else
                .Interior.Color = RGB(47, 84, 150)
                .Font.Color     = RGB(255, 255, 255)
            End If
        End With

        ' -------- チェックセル（ROW_FIRST〜lastRow）の処理 --------
        For row = ROW_FIRST To lastRow

            Dim isEven As Boolean: isEven = ((row Mod 2) = 0)
            Dim targetCell As Range
            Set targetCell = wsNew.Cells(row, col)

            If isRest Then
                ' 休日（土日・祝日）：グレー塗りつぶし・入力ロック・チェックボックスなし
                targetCell.Interior.Color = RGB(191, 191, 191)
                targetCell.Value          = ""
                targetCell.Locked         = True
            Else
                ' 平日：ロック解除 & チェックボックスを配置する
                targetCell.Interior.Color = IIf(isEven, RGB(237, 243, 251), RGB(255, 255, 255))
                targetCell.Locked         = False

                ' チェックボックスをセル中央に配置する
                ' フォームコントロールはチェック図形部分が常に左端に描画されるため、
                ' コントロール自体を図形サイズ（約12pt角）に絞り、セル中央に座標指定する
                Const CB_SIZE As Double = 12   ' チェック図形の実寸（ポイント）
                Dim cbLeft As Double
                Dim cbTop  As Double
                cbLeft = targetCell.Left + (targetCell.Width  - CB_SIZE) / 2  ' 水平中央
                cbTop  = targetCell.Top  + (targetCell.Height - CB_SIZE) / 2  ' 垂直中央

                Dim newCb As CheckBox
                Set newCb = wsNew.CheckBoxes.Add( _
                    cbLeft, cbTop, CB_SIZE, CB_SIZE)

                newCb.Caption     = ""                 ' ラベル文字なし
                newCb.Value       = xlOff              ' 初期状態：未チェック
                newCb.LinkedCell  = targetCell.Address ' TRUE/FALSE をセルに連動
                newCb.PrintObject = True               ' 印刷時にも表示

                ' LinkedCell のTRUE/FALSEテキストを背景色と同色にして非表示にする
                targetCell.Font.Color = targetCell.Interior.Color
            End If

        Next row

        wsNew.Columns(col).ColumnWidth = 6

    Next d

    ' ── 余剰列を非表示 ────────────────────────────────────────────
    For col = COL_START + daysInMonth To COL_START + 30
        wsNew.Columns(col).Hidden = True
    Next col

    ' ── シート保護を再設定 ────────────────────────────────────────
    wsNew.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True

    ' ── ウィンドウ枠固定・ズーム設定 ─────────────────────────────
    wsNew.Activate
    ActiveWindow.FreezePanes = False
    wsNew.Cells(ROW_FIRST, COL_START).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.Zoom = 90

    MsgBox wsName & " のシートを作成しました！", vbInformation, "チェックシート生成"

End Sub

' -----------------------------------------------------------------------
' SheetExists：指定した名前のシートが存在するか確認する関数
' -----------------------------------------------------------------------
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function
```
