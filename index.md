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
