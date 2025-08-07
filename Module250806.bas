Attribute VB_Name = "Module250806"
Option Explicit

Public fileNameSuffix As String ' 追加 20250806 by maruyama

Function MakeBaseNamePL5(ws As Worksheet, ext As String) As String
    Dim sdCell As String: sdCell = mt.Config("Config", "sdCell")        ' 出荷日のセルアドレス
    Dim ccCell As String: ccCell = mt.Config("Config", "ccCell")        ' 出荷先名（国名）のセルアドレス

    If fileNameSuffix <> "" Then
        fileNameSuffix = " " & fileNameSuffix
    End If

    MakeBaseNamePL5 = Format(ws.Range(sdCell), "yyyymmdd") & " " & Left(StrConv(Replace(ws.Range(ccCell), " ", ""), vbProperCase), 3) & " Packing list " & ws.Name & fileNameSuffix & "." & ext
End Function

' ws: コピー元ワークシート
' savePath: 保存先フルパス（例: "C:\\temp\\test.xlsx"）
' fileFormat: 保存形式（例: xlOpenXMLWorkbook）
Sub CopySheetToNewWorkbook(ws As Worksheet, savePath As String, Optional FileFormat As XlFileFormat = xlOpenXMLWorkbook)
    Dim newWb As Workbook
    ws.Copy
    Set newWb = ActiveWorkbook
    Application.DisplayAlerts = False
    newWb.SaveAs Filename:=savePath, FileFormat:=FileFormat
    Application.DisplayAlerts = True
    newWb.Close SaveChanges:=False
End Sub

' ファイル名に使えない文字が含まれていないかチェックする関数
Function ContainsInvalidFileNameChars(s As String) As String
    Dim invalidChars As String
    invalidChars = "\/:*?""<>|"
    Dim i As Integer, found As String
    found = ""
    For i = 1 To Len(invalidChars)
        If InStr(s, Mid(invalidChars, i, 1)) > 0 Then
            found = found & Mid(invalidChars, i, 1)
        End If
    Next i
    ContainsInvalidFileNameChars = found
End Function
