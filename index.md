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

Option Explicit

' -----------------------------------------------------------------------
' CreateCurrentMonthSheet：当月シートを作成するショートカット
' 使い方： Alt+F8 → このプロシージャを選択して実行
' -----------------------------------------------------------------------
Sub CreateCurrentMonthSheet()
    CreateMonthSheet Year(Now()), Month(Now())
End Sub

' -----------------------------------------------------------------------
' SetupSampleCheckboxes：サンプルシートにチェックボックスを配置する
' 使い方： Alt+F8 → このプロシージャを選択して実行（初回1回だけでOK）
' ※ サンプルシートはコピー元なので、実際の入力には使わないこと
' -----------------------------------------------------------------------
Sub SetupSampleCheckboxes()

    Const ROW_FIRST As Long = 4  ' チェック項目の開始行
    Const COL_START As Long = 2  ' B列（1日目）
    Const COL_LABEL As Long = 1  ' A列（チェック項目名）

    If Not SheetExists("サンプル") Then
        MsgBox "「サンプル」シートが見つかりません。", vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Worksheets("サンプル")
    ws.Unprotect Password:=""

    ' サンプルシートの項目数をA列から取得する
    Dim itemCount As Integer
    itemCount = 0
    Do While ws.Cells(ROW_FIRST + itemCount, COL_LABEL).Value <> ""
        itemCount = itemCount + 1
    Loop

    Dim lastRow As Long
    lastRow = ROW_FIRST + itemCount - 1

    ' 既存のチェックボックスをすべて削除してから再配置する
    Dim cb As CheckBox
    For Each cb In ws.CheckBoxes
        cb.Delete
    Next cb

    ' サンプルシートは2026年1月（31日）で固定
    Dim holidays As Collection
    Set holidays = LoadHolidays(2026, 1)

    Dim d As Integer, row As Long, col As Long

    For d = 1 To 31

        Dim curDate As Date
        curDate = DateSerial(2026, 1, d)
        Dim wd As Integer
        wd = Weekday(curDate, vbSunday)

        Dim isRest As Boolean
        isRest = (wd = 7) Or (wd = 1) Or IsHoliday(curDate, holidays)

        col = COL_START + d - 1

        For row = ROW_FIRST To lastRow

            Dim targetCell As Range
            Set targetCell = ws.Cells(row, col)

            If Not isRest Then
                ' 平日セルにチェックボックスをセル中央に配置する
                Const CB_SIZE2 As Double = 12
                Dim cbLeft2 As Double, cbTop2 As Double
                cbLeft2 = targetCell.Left + (targetCell.Width  - CB_SIZE2) / 2
                cbTop2  = targetCell.Top  + (targetCell.Height - CB_SIZE2) / 2

                Dim newCb As CheckBox
                Set newCb = ws.CheckBoxes.Add(cbLeft2, cbTop2, CB_SIZE2, CB_SIZE2)
                newCb.Caption     = ""
                newCb.Value       = xlOff
                newCb.LinkedCell  = targetCell.Address
                newCb.PrintObject = True
                targetCell.Font.Color = targetCell.Interior.Color
            End If

        Next row
    Next d

    ws.Protect Password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
    MsgBox "サンプルシートにチェックボックスを配置しました！", vbInformation, "セットアップ完了"

End Sub

' -----------------------------------------------------------------------
' LoadHolidays：「祝日」シートから対象月の祝日をCollectionに読み込む
' -----------------------------------------------------------------------
Function LoadHolidays(yr As Integer, mo As Integer) As Collection

    Dim col As New Collection

    If Not SheetExists("祝日") Then
        Set LoadHolidays = col
        Exit Function
    End If

    Dim wsH As Worksheet
    Set wsH = Worksheets("祝日")

    Dim lastRow As Long
    lastRow = wsH.Cells(wsH.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        Dim cellVal As Variant
        cellVal = wsH.Cells(i, 1).Value
        If IsDate(cellVal) Then
            Dim hDate As Date
            hDate = CDate(cellVal)
            If Year(hDate) = yr And Month(hDate) = mo Then
                Dim key As String
                key = Format(hDate, "yyyy/mm/dd")
                On Error Resume Next
                col.Add hDate, key
                On Error GoTo 0
            End If
        End If
    Next i

    Set LoadHolidays = col

End Function

' -----------------------------------------------------------------------
' IsHoliday：指定した日付が祝日Collectionに含まれているか判定する
' -----------------------------------------------------------------------
Function IsHoliday(checkDate As Date, holidays As Collection) As Boolean
    Dim dummy As Date
    On Error Resume Next
    dummy = holidays(Format(checkDate, "yyyy/mm/dd"))
    IsHoliday = (Err.Number = 0)
    On Error GoTo 0
End Function

' -----------------------------------------------------------------------
' CreateMonthSheet：指定した年・月のチェックシートを生成するメインプロシージャ
' 引数  yr : 生成したいシートの年（例：2026）
'       mo : 生成したいシートの月（例：5）
' -----------------------------------------------------------------------
Sub CreateMonthSheet(yr As Integer, mo As Integer)

    Const ROW_TITLE As Long = 1  ' タイトル行
    Const ROW_DATE  As Long = 2  ' 日付ヘッダー行
    Const ROW_WEEK  As Long = 3  ' 曜日ヘッダー行
    Const ROW_FIRST As Long = 4  ' チェック項目の開始行
    Const COL_LABEL As Long = 1  ' A列（チェック項目名）
    Const COL_START As Long = 2  ' B列（1日目）

    ' ── サンプルシートの確認 ────────────────────────────────────
    If Not SheetExists("サンプル") Then
        MsgBox "「サンプル」シートが見つかりません。削除しないでください。", vbCritical
        Exit Sub
    End If

    ' ── チェック項目数をサンプルシートのA列から取得 ──────────────
    ' 項目の追加・変更はサンプルシートのA列を編集するだけでOK
    Dim wsSample As Worksheet
    Set wsSample = Worksheets("サンプル")
    wsSample.Unprotect Password:=""

    Dim itemCount As Integer
    itemCount = 0
    Do While wsSample.Cells(ROW_FIRST + itemCount, COL_LABEL).Value <> ""
        itemCount = itemCount + 1
    Loop

    If itemCount = 0 Then
        MsgBox "サンプルシートにチェック項目が見つかりません。" & vbCrLf & _
               "A列（" & ROW_FIRST & "行目以降）に項目を入力してください。", vbCritical
        wsSample.Protect Password:="", UserInterfaceOnly:=True
        Exit Sub
    End If

    wsSample.Protect Password:="", UserInterfaceOnly:=True

    ' ── 変数宣言 ─────────────────────────────────────────────────
    Dim wsNew       As Worksheet
    Dim wsName      As String
    Dim daysInMonth As Integer
    Dim col         As Long
    Dim row         As Long
    Dim d           As Integer
    Dim wd          As Integer
    Dim lastRow     As Long

    wsName  = yr & "年" & mo & "月"
    lastRow = ROW_FIRST + itemCount - 1

    ' ── 既存シートの確認 ─────────────────────────────────────────
    If SheetExists(wsName) Then
        MsgBox wsName & " のシートはすでに存在します。", vbInformation
        Worksheets(wsName).Activate
        Exit Sub
    End If

    ' ── 祝日リストを読み込む ──────────────────────────────────────
    Dim holidays As Collection
    Set holidays = LoadHolidays(yr, mo)

    ' ── サンプルをコピーしてリネーム ─────────────────────────────
    ' 色・書式・列幅・行高さはすべてサンプルから引き継ぐ
    Worksheets("サンプル").Copy After:=Worksheets(Worksheets.Count)
    Set wsNew = ActiveSheet
    wsNew.Name = wsName
    wsNew.Unprotect Password:=""

    daysInMonth = Day(DateSerial(yr, mo + 1, 0))

    ' ── タイトルのテキストだけ書き換える ────────────────────────
    ' 色・書式はサンプルのものをそのまま使う
    wsNew.Cells(ROW_TITLE, COL_LABEL).Value = wsName & "　業務チェックシート"

    ' ── 曜日の日本語表記 ─────────────────────────────────────────
    Dim weekdayJP(1 To 7) As String
    weekdayJP(1) = "日": weekdayJP(2) = "月": weekdayJP(3) = "火"
    weekdayJP(4) = "水": weekdayJP(5) = "木": weekdayJP(6) = "金"
    weekdayJP(7) = "土"

    ' ── サンプルから平日・土曜・日曜の色をあらかじめ取得しておく ────
    ' 2026年1月のサンプルシートを基準に各曜日の列を特定する
    ' Jan 1=木(COL+0)、Jan 3=土(COL+2)、Jan 4=日(COL+3)
    Dim clrWdayBg  As Long, clrWdayFg  As Long   ' 平日の背景色・文字色
    Dim clrSatBg   As Long, clrSatFg   As Long   ' 土曜の背景色・文字色
    Dim clrSunBg   As Long, clrSunFg   As Long   ' 日曜・祝日の背景色・文字色
    Dim clrRestBg  As Long                        ' 休日チェックセルの背景色

    clrWdayBg = wsNew.Cells(ROW_DATE, COL_START).Interior.Color      ' 1月1日（木）
    clrWdayFg = wsNew.Cells(ROW_DATE, COL_START).Font.Color
    clrSatBg  = wsNew.Cells(ROW_DATE, COL_START + 2).Interior.Color  ' 1月3日（土）
    clrSatFg  = wsNew.Cells(ROW_DATE, COL_START + 2).Font.Color
    clrSunBg  = wsNew.Cells(ROW_DATE, COL_START + 3).Interior.Color  ' 1月4日（日）
    clrSunFg  = wsNew.Cells(ROW_DATE, COL_START + 3).Font.Color
    clrRestBg = wsNew.Cells(ROW_FIRST, COL_START + 3).Interior.Color ' 1月4日（日）のチェックセル

    ' ── 日付列の値をリセット（書式・色は後で曜日に応じて上書きする） ──
    wsNew.Range(wsNew.Cells(ROW_DATE, COL_START), _
                wsNew.Cells(lastRow, COL_START + 30)).ClearContents

    ' ── チェックボックスを削除（サンプルの残骸） ─────────────────
    Dim cb As CheckBox
    For Each cb In wsNew.CheckBoxes
        cb.Delete
    Next cb

    ' ── 日ごとの処理 ─────────────────────────────────────────────
    For d = 1 To daysInMonth

        Dim curDate As Date
        curDate = DateSerial(yr, mo, d)
        wd  = Weekday(curDate, vbSunday)
        col = COL_START + d - 1

        Dim isSat  As Boolean: isSat  = (wd = 7)
        Dim isSun  As Boolean: isSun  = (wd = 1)
        Dim isHol  As Boolean: isHol  = IsHoliday(curDate, holidays)
        Dim isRest As Boolean: isRest = isSat Or isSun Or isHol

        ' 曜日に応じたヘッダー色を決定する
        Dim hBg As Long, hFg As Long
        If isSat Then
            hBg = clrSatBg:  hFg = clrSatFg
        ElseIf isSun Or isHol Then
            hBg = clrSunBg:  hFg = clrSunFg
        Else
            hBg = clrWdayBg: hFg = clrWdayFg
        End If

        ' -------- 日付ヘッダー：値・書式・色を設定 --------
        With wsNew.Cells(ROW_DATE, col)
            .Value          = DateSerial(yr, mo, d)
            .NumberFormat   = "d"
            .Interior.Color = hBg
            .Font.Color     = hFg
        End With

        ' -------- 曜日ヘッダー：値・色を設定 --------
        With wsNew.Cells(ROW_WEEK, col)
            .Value          = weekdayJP(wd)
            .Interior.Color = hBg
            .Font.Color     = hFg
        End With

        ' -------- チェックセル --------
        For row = ROW_FIRST To lastRow

            Dim targetCell As Range
            Set targetCell = wsNew.Cells(row, col)

            If isRest Then
                ' 休日：サンプルの日曜チェックセルの色を使う
                targetCell.Interior.Color = clrRestBg
                targetCell.Value          = ""
                targetCell.Locked         = True
            Else
                ' 平日：サンプルの同じ行の平日列（1月1日=木曜）から背景色を引き継ぐ
                ' ※ これをしないとサンプルで土日だった列の灰色が残ってしまう
                targetCell.Interior.Color = wsNew.Cells(row, COL_START).Interior.Color
                targetCell.Locked = False

                Const CB_SIZE As Double = 12
                Dim cbLeft As Double, cbTop As Double
                cbLeft = targetCell.Left + (targetCell.Width  - CB_SIZE) / 2
                cbTop  = targetCell.Top  + (targetCell.Height - CB_SIZE) / 2

                Dim newCb As CheckBox
                Set newCb = wsNew.CheckBoxes.Add(cbLeft, cbTop, CB_SIZE, CB_SIZE)
                newCb.Caption     = ""
                newCb.Value       = xlOff
                newCb.LinkedCell  = targetCell.Address
                newCb.PrintObject = True
                targetCell.Font.Color = targetCell.Interior.Color
            End If

        Next row

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

# 祝日

| 日付 | 祝日名 |
|---|---|
| 2026/1/1 | 元日 |
| 2026/1/12 | 成人の日 |
| 2026/2/11 | 建国記念の日 |
| 2026/2/23 | 天皇誕生日 |
| 2026/3/20 | 春分の日 |
| 2026/4/29 | 昭和の日 |
| 2026/5/3 | 憲法記念日 |
| 2026/5/4 | みどりの日 |
| 2026/5/5 | こどもの日 |
| 2026/5/6 | 振替休日 |
| 2026/7/20 | 海の日 |
| 2026/8/11 | 山の日 |
| 2026/9/21 | 敬老の日 |
| 2026/9/22 | 国民の休日 |
| 2026/9/23 | 秋分の日 |
| 2026/10/12 | スポーツの日 |
| 2026/11/3 | 文化の日 |
| 2026/11/23 | 勤労感謝の日 |
| 2027/1/1 | 元日 |
| 2027/1/11 | 成人の日 |
| 2027/2/11 | 建国記念の日 |
| 2027/2/23 | 天皇誕生日 |
| 2027/3/21 | 春分の日 |
| 2027/4/29 | 昭和の日 |
| 2027/5/3 | 憲法記念日 |
| 2027/5/4 | みどりの日 |
| 2027/5/5 | こどもの日 |
| 2027/7/19 | 海の日 |
| 2027/8/11 | 山の日 |
| 2027/9/20 | 敬老の日 |
| 2027/9/23 | 秋分の日 |
| 2027/10/11 | スポーツの日 |
| 2027/11/3 | 文化の日 |
| 2027/11/23 | 勤労感謝の日 |
