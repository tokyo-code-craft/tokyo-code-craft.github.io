---
---

# sample

```vba
Sub ImportMultipleFiles()
    Dim vFiles As Variant
    Dim i As Long
    Dim wbSrc As Workbook

    ' 複数選択を許可してダイアログを開く
    vFiles = Application.GetOpenFilename( _
        FileFilter:="対象ファイル (*.xlsx;*.csv;*.tsv;*.txt),*.xlsx;*.csv;*.tsv;*.txt", _
        Title:="読み込むファイルを選択（複数可）", _
        MultiSelect:=True)

    ' キャンセル時（Falseが返る）は終了
    If Not IsArray(vFiles) Then Exit Sub

    Application.ScreenUpdating = False

    ' 選択した各ファイルをループ処理
    For i = LBound(vFiles) To UBound(vFiles)
        Set wbSrc = Workbooks.Open(vFiles(i))

        ' ↓ここに1ファイルごとの処理を書く（例：先頭シートを転記元にする）
        ' MsgBox wbSrc.Name & " を読み込みました"

        wbSrc.Close SaveChanges:=False
    Next i

    Application.ScreenUpdating = True
    MsgBox UBound(vFiles) - LBound(vFiles) + 1 & " 件のファイルを処理しました"
End Sub
```
