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
    ' Workbook�̑S�ẴV�[�g�����擾����
    Dim i As Long, numSheets As Long
    numSheets = wb.Sheets.Count
    ReDim sheetnames(1 To numSheets) As String
    For i = 1 To numSheets
        sheetnames(i) = wb.Sheets(i).Name
    Next i
    GetAllSheetNames = sheetnames
End Function

Private Sub ExportSheetToCSV(sh As Worksheet)
    ' �V�[�g�̃f�[�^��CSV�`���ŃG�N�X�|�[�g����
    ' CSV�t�@�C���̃p�X���쐬����
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\" & sh.Name & ".csv"

    ' CSV�t�@�C����ۑ�����
    sh.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs csvPath, xlCSV
    ActiveWorkbook.Close '�u�b�N�����
    Application.DisplayAlerts = True
End Sub
