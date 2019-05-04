Const MaxCell As Long = 1048576
Const StartRange As String = "C2"

' ----------------------------------------------------------
' 日付コピー
'
' ----------------------------------------------------------
Sub Macro1()

    ' 終点セル取得
    With ActiveSheet.UsedRange
        MaxRow = .Rows.Count
        MaxCol = .Columns.Count
    End With

    ' 開始位置に移動
    Range(StartRange).Select

    Do
        Dim SourceRow As Long

        ' 未入力セルの１つ上のセルに移動
        Selection.End(xlDown).Select

        ' コピー元セル位置取得
        SourceRow = Selection.Row
        SourceCol = Selection.Column

        ' コピー元セルが最終セルだったら、ループを抜ける
        If SourceRow > MaxRow Or SourceRow >= MaxCell Then Exit Do

        ' コピー元セルと隣のDセルを選択
        Range(Cells(SourceRow, SourceCol), Cells(SourceRow, SourceCol + 1)).Select

        ' コピー元データを保持
        'Selection.Copy

        ' TODO: コピー元セルの１つ下のセル絶対位置取得
        ' TODO: コピー元セルの１つ下のセルに移動
        'Range("C90").Select

        ' 次の入力済セルまでカーソル移動
        'Range(Selection, Selection.End(xlDown)).Select

        ' TODO: コピー先セルの１つ上のセルに移動
        'Range("C90:C100").Select
        'ActiveSheet.Paste

    Loop
End Sub


Sub Test1()
    With ActiveSheet.UsedRange
        MaxRow = .Rows.Count
        MaxCol = .Columns.Count
    End With
End Sub

Sub Test2()
    ' C1を選択
    ' Cells(1, 3).Select

    ' C2:D4を選択
    Range(Cells(2, 3), Cells(4, 4)).Select
End Sub
