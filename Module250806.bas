Attribute VB_Name = "Module250806"
Option Explicit

Public fileNameSuffix As String ' �ǉ� 20250806 by maruyama

Function MakeBaseNamePL5(ws As Worksheet, ext As String) As String
    Dim sdCell As String: sdCell = mt.Config("Config", "sdCell")        ' �o�ד��̃Z���A�h���X
    Dim ccCell As String: ccCell = mt.Config("Config", "ccCell")        ' �o�א於�i�����j�̃Z���A�h���X

    If fileNameSuffix <> "" Then
        fileNameSuffix = " " & fileNameSuffix
    End If

    MakeBaseNamePL5 = Format(ws.Range(sdCell), "yyyymmdd") & " " & Left(StrConv(Replace(ws.Range(ccCell), " ", ""), vbProperCase), 3) & " Packing list " & ws.Name & fileNameSuffix & "." & ext
End Function

' ws: �R�s�[�����[�N�V�[�g
' savePath: �ۑ���t���p�X�i��: "C:\\temp\\test.xlsx"�j
' fileFormat: �ۑ��`���i��: xlOpenXMLWorkbook�j
Sub CopySheetToNewWorkbook(ws As Worksheet, savePath As String, Optional FileFormat As XlFileFormat = xlOpenXMLWorkbook)
    Dim newWb As Workbook
    ws.Copy
    Set newWb = ActiveWorkbook
    Application.DisplayAlerts = False
    newWb.SaveAs Filename:=savePath, FileFormat:=FileFormat
    Application.DisplayAlerts = True
    newWb.Close SaveChanges:=False
End Sub

