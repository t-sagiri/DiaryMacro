Attribute VB_Name = "Module1"

' ----------------------------------------------------------
' 日付コピー
'
' ----------------------------------------------------------
Sub Macro1()
Attribute Macro1.VB_Description = "日付コピー"
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
Const MaxCell As Long = 1048576
Const StartRange As String = "C2"

    ' 開始位置に移動
    Range("C2").Select

    ' 未入力セルの１つ上のセルに移動
    Selection.End(xlDown).Select

    ' TODO: カレントセルが最終セルだったら、ループを抜ける

    ' TODO: カレントセルの絶対位置取得
    ' TODO: カレントセルと隣のDセルを選択
    Range("C89:D89").Select

    ' コピー元データを保持
    Selection.Copy

    ' TODO: コピー元セルの１つ下のセル絶対位置取得
    ' TODO: コピー元セルの１つ下のセルに移動
    Range("C90").Select

    ' 次の入力済セルまでカーソル移動
    Range(Selection, Selection.End(xlDown)).Select

    ' TODO: コピー先セルの１つ上のセルに移動
    Range("C90:C100").Select
    ActiveSheet.Paste
End Sub
