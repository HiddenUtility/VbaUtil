Attribute VB_Name = "cellutil"
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

Private Function GetEndColNumber(ws As Worksheet, Optional RowNumber As Long = 1, Optional ColNumber As Long = 1, Optional Limit As Long = 16384) As Long
    If RowNumber < 1 Or RowNumber > 1048576 Then Err.Raise 1, , "RowNumberは1-1048576の値しかとれません。"
    If ColNumber < 1 Or ColNumber > 16384 Then Err.Raise 1, , "ColNumberは1-16384の値しかとれません。"
    Do
        If Cells(RowNumber, ColNumber) = 0 Then Exit Do
        ColNumber = ColNumber + 1
        If ColNumber > Limit Then Err.Raise 2, , "限界に到達しました。"
    Loop

    GetEndColNumber = ColNumber - 1
    
End Function

Function GetLastRow(ws As Worksheet, Optional Column As String = "A") As Long
    GetLastRow = ws.Cells(ws.Rows.Count, Column).End(xlUp).Row
End Function

