Attribute VB_Name = "toCsv"
Option Explicit

Private Sub ExportSheetToCSV(sh As Worksheet)
    ' シートのデータをCSV形式でエクスポートする
    ' CSVファイルのパスを作成する
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\" & sh.Name & ".csv"
    ' CSVファイルを保存する
    sh.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & csvPath, xlCSV
    ActiveWorkbook.Close 'ブックを閉じる
    Application.DisplayAlerts = True
End Sub

Private Sub ExportActiveSheetToCSV()
    ' シートのデータをCSV形式でエクスポートする
    Dim sh As Worksheet
    Set sh = ActiveSheet
    ' CSVファイルのパスを作成する
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\" & sh.Name & ".csv"
    ' CSVファイルを保存する
    sh.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & csvPath, xlCSV
    ActiveWorkbook.Close 'ブックを閉じる
    Application.DisplayAlerts = True
End Sub
