Attribute VB_Name = "wsutil"
Option Explicit

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

