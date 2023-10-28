Attribute VB_Name = "Module1"
Sub MasterProcedure()
    ' マスタープロシージャで他のプロシージャを呼び出す
    CopySheet ' 作業用シートを作成
    QuoteValuesInColumnsAtoV 'ダブルクォーテーション設定
'    QuoteValuesInColumnT　'T行のみにダブルクォーテーション設定
    Make_Csv_Quote  'ダブルクォーテーション付きでファイル出力
End Sub
'#############################################################################
' 作業用シートを作成
'　copy_sheet
'#############################################################################
Sub CopySheet()
    Dim ws As Worksheet
    ' シート「固定資産台帳」が存在するか確認
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("固定資産台帳")
    On Error GoTo 0
    ' シート「固定資産台帳」が存在する場合、コピーして「固定資産台帳bak」という名前で保存
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' 警告を非表示にする
        ws.Copy Before:=ThisWorkbook.Sheets(1)
        Application.DisplayAlerts = True ' 警告を非表示にする
        ' 新しく作成されたシートに名前を設定
        ActiveSheet.Name = "固定資産台帳_作業用"
        MsgBox "シート「固定資産台帳」が「固定資産台帳_作業用」としてコピーされました。"
    Else
        MsgBox "シート「固定資産台帳」が存在しません。"
    End If
End Sub
'#############################################################################
' ダブルクォーテーション設定
'　quote_values_A2V
'#############################################################################
Sub QuoteValuesInColumnsAtoV()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Cell As Range
    Dim Column As Range
   ' ワークシートを指定
    Set ws = ThisWorkbook.Worksheets("固定資産台帳_作業用") ' シート名を適切なものに変更
    ' データの最終行を取得
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ' A列からAR列のセル内の値をダブルクォーテーションで囲む
     For Each Column In ws.Range("A4:V" & LastRow).Columns
        For Each Cell In Column.Cells
            If Not IsEmpty(Cell.Value) Then
                Cell.Value = """" & Cell.Value & """"
            End If
        Next Cell
    Next Column
    MsgBox "A列からV列のセル内の値がダブルクォーテーションで囲まれました。"
End Sub
'#############################################################################
' ダブルクォーテーション設定
'  quote_values_T
'#############################################################################
Sub QuoteValuesInColumnT()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Cell As Range
    ' ワークシートを指定
    Set ws = ThisWorkbook.Worksheets("固定資産台帳_作業用") ' シート名を適切なものに変更
    ' データの最終行を取得
    LastRow = ws.Cells(ws.Rows.Count, "T").End(xlUp).Row
    ' Q列のセル内の値をダブルクォーテーションで囲む
    For Each Cell In ws.Range("T4:T" & LastRow)
        If Not IsEmpty(Cell.Value) Then
            Cell.Value = """" & Cell.Value & """"
        End If
    Next Cell
    MsgBox "T列のセル内の値がダブルクォーテーションで囲まれました。"
End Sub
'#############################################################################
' ダブルクォーテーション付きでファイル出力
'　make_csv_quote
'#############################################################################
Sub Make_Csv_Quote()
    Dim i As Integer, j As Integer, FileNumber As Integer, LR As Integer, LC As Integer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("固定資産台帳_作業用")  ' ワークシートを設定
    Dim OutputCsv As String, OutputTxt As String
    OutputCsv = Environ("USERPROFILE") & "\Desktop\sample.csv"
    OutputTxt = Environ("USERPROFILE") & "\Desktop\sample.txt"
    ' 最初の1行を削除
    On Error Resume Next ' エラーを無視して次の行に進む
    ws.Rows("1:1").Delete Shift:=xlUp ' 1から3行を削除して上にシフト
    On Error GoTo 0 ' エラーハンドリングを元に戻す
    ' W列以降の列削除
    Dim DeleteRange As Range
    Set DeleteRange = ws.Range("W:ZZ") ' 削除したい列の範囲を指定
    DeleteRange.Delete Shift:=xlToLeft
    ' T列が空の行削除
    Dim lRow As Integer
    lRow = Cells(Rows.Count, "T").End(xlUp).Row  'T列の最終行を取得
    For i = lRow To 1 Step -1 '最終行から1行目まで繰り返す。
        If IsEmpty(Cells(i, "T").Value) Then  'T列の空白行を判定
                Rows(i).Delete '該当（空白）する行削除
        End If
    Next i
    Dim NewWorkbook As Workbook
    Dim CSVFilePath As String
    ' 新しいワークブックを作成
    Set NewWorkbook = Workbooks.Add
    ' 新しいワークブックにデータをコピー
    ws.Copy Before:=NewWorkbook.Sheets(1)
    ' 新しいワークブックをCSVファイルに保存
    NewWorkbook.SaveAs OutputCsv, xlCSV
    ' 新しいワークブックを閉じる（変更あり）
    NewWorkbook.Close SaveChanges:=False
    ' メッセージを表示
    MsgBox "「固定資産台帳_作業用」シートの内容がCSVファイルにエクスポートされました。", vbInformation
'最初の項目行２行を削除
    Workbooks.OpenText Filename:=OutputCsv
    Rows("1:2").Delete
    ActiveWorkbook.Save
    ActiveWindow.Close True
' 作業用シートを削除
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' 警告を非表示にする
        ws.Delete
        Application.DisplayAlerts = True
        MsgBox sheetName & " シートが削除されました。", vbInformation
    Else
        MsgBox sheetName & " シートは存在しません。", vbExclamation
    End If
End Sub
