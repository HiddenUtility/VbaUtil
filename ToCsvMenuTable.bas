Attribute VB_Name = "ToCsvMenuTable"
Option Explicit
Sub Main()

    Dim ws As Worksheet
    Set ws = Sheets("MENU")
    Dim end_row As Long
    end_row = GetEndRowNumber(ws)
    
    Dim out_ws As Worksheet
    Dim sheet_name As String
    Dim schema As String
    Dim table_name As String
    Dim i As Long
    For i = 2 To end_row
        sheet_name = Cells(i, 1)
        schema = Cells(i, 2)
        table_name = Cells(i, 3)
        Set out_ws = Sheets(sheet_name)
        Call ExportSheetToCSV(out_ws, schema, table_name)
    Next i
    

End Sub


Private Function GetAllSheetNames(ByVal wb As Workbook) As Variant
    ' Workbookの全てのシート名を取得する
    Dim i As Long, numSheets As Long
    numSheets = wb.Sheets.Count
    ReDim sheetnames(1 To numSheets) As String
    For i = 1 To numSheets
        sheetnames(i) = wb.Sheets(i).Name
    Next i
    GetAllSheetNames = sheetnames
End Function

Private Sub ExportSheetToCSV(sh As Worksheet, schema As String, table_name As String)
    ' シートのデータをCSV形式でエクスポートする
    ' CSVファイルのパスを作成する
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\" & schema & "." & table_name & ".csv"

    ' CSVファイルを保存する
    sh.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs csvPath, xlCSV
    ActiveWorkbook.Close 'ブックを閉じる
    Application.DisplayAlerts = True
End Sub


Private Function GetEndRowNumber(ws As Worksheet, Optional RowNumber As Long = 1, Optional ColNumber As Long = 1, Optional Limit As Long = 1048576) As Long
    If RowNumber < 1 Or RowNumber > 1048576 Then Err.Raise 1, , "RowNumberは1-1048576の値しかとれません。"
    If ColNumber < 1 Or ColNumber > 16384 Then Err.Raise 1, , "ColNumberは1-16384の値しかとれません。"
    Do
        If Cells(RowNumber, ColNumber) = 0 Then Exit Do
        RowNumber = RowNumber + 1
        If RowNumber > Limit Then Err.Raise 2, , "限界に到達しました。"
        
    Loop

    GetEndRowNumber = RowNumber - 1
    
End Function

