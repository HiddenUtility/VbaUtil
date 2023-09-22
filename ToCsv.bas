Attribute VB_Name = "toCsv"
Option Explicit

Private Sub ExportSheetToCSV(sh As Worksheet)
    ' �V�[�g�̃f�[�^��CSV�`���ŃG�N�X�|�[�g����
    ' CSV�t�@�C���̃p�X���쐬����
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\" & sh.Name & ".csv"
    ' CSV�t�@�C����ۑ�����
    sh.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & csvPath, xlCSV
    ActiveWorkbook.Close '�u�b�N�����
    Application.DisplayAlerts = True
End Sub

Private Sub ExportActiveSheetToCSV()
    ' �V�[�g�̃f�[�^��CSV�`���ŃG�N�X�|�[�g����
    Dim sh As Worksheet
    Set sh = ActiveSheet
    ' CSV�t�@�C���̃p�X���쐬����
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\" & sh.Name & ".csv"
    ' CSV�t�@�C����ۑ�����
    sh.Copy
    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & csvPath, xlCSV
    ActiveWorkbook.Close '�u�b�N�����
    Application.DisplayAlerts = True
End Sub
