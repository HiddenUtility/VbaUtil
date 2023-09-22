Attribute VB_Name = "toCsvAllSheets"
Option Explicit
Sub Main()

    Dim sheetnames As Variant
    sheetnames = GetAllSheetNames(ThisWorkbook)
    Dim sheetname As Variant
    Dim sh As Worksheet
    For Each sheetname In sheetnames
        If sheetname <> "MENU" Then
            Set sh = Sheets(sheetname)
            Call ExportSheetToCSV(sh)
        End If
    Next sheetname

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

Private Sub ExportSheetToCSV(sh As Worksheet)
    ' シートのデータをCSV形式でエクスポートする
    ' CSVファイルのパスを作成する
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\" & sh.Name & ".csv"

    ' CSVファイルを保存する
    sh.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs csvPath, xlCSV
    ActiveWorkbook.Close 'ブックを閉じる
    Application.DisplayAlerts = True
End Sub
