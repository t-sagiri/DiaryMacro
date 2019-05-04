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
    End With

    ' 開始位置に移動
    Range(StartRange).Select

    Do
        ' 未入力セルの１つ上のセルに移動
        Selection.End(xlDown).Select

        ' コピー元セル位置取得
        SourceRow = Selection.Row
        SourceCol = Selection.Column

        ' 未入力セルの１つ上のセルに移動
        Selection.End(xlDown).Select

        ' 次の入力済みセルLine取得
        NextRow = Selection.Row

        ' コピー先セル位置取得
        DestRow = NextRow - 1
        DestEndCol = Selection.Column + 1

        ' コピー元セルと隣のDセルを選択
        Range(Cells(SourceRow, SourceCol), Cells(SourceRow, SourceCol + 1)).Select

        ' コピー元データを保持
        Selection.Copy

        ' 次の入力済セルまでカーソル移動
        Range(Cells(SourceRow + 1, SourceCol), Cells(DestRow, DestEndCol)).Select

        ' コピー先セルの１つ上のセルに移動
        ActiveSheet.Paste

    Loop While NextRow < MaxRow
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
