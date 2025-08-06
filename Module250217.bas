Attribute VB_Name = "Module250217"
Option Explicit

Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As RECT) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public wsOCR As Worksheet

Function GetWidth(rc As RECT) As Integer
    GetWidth = rc.Right - rc.Left
End Function

Function GetHeight(rc As RECT) As Integer
    GetHeight = rc.Bottom - rc.Top
End Function

' Excelアプリケーションウィンドウの位置とサイズを取得する
Function GetWindowPositionAndSize(winTitle As String) As RECT
    Dim hWnd As LongPtr
    Dim appRect As RECT

    ' Excelのウィンドウハンドルを取得
    hWnd = FindWindow("XLMAIN", winTitle)
    
    ' ウィンドウの位置とサイズを取得
    GetWindowRect hWnd, appRect
    
    GetWindowPositionAndSize = appRect
End Function

Sub SaveWindowPositionAndSize()
    Dim appRect As RECT
    RECT = GetWindowPositionAndSize()
    
    ' セルに保存
    With ThisWorkbook.Sheets("Sheet1")
        .Range("A1").Value = appRect.Left
        .Range("A2").Value = appRect.Top
        .Range("A3").Value = appRect.Right - appRect.Left ' Width
        .Range("A4").Value = appRect.Bottom - appRect.Top ' Height
    End With
End Sub

' Excelアプリケーションウィンドウの位置とサイズを設定する
Sub RestoreWindowPositionAndSize(rc As RECT)
    Dim hWnd As LongPtr
    Dim X As Long, Y As Long, Width As Long, Height As Long
    
    X = rc.Left
    Y = rc.Top
    Width = GetWidth(rc)
    Height = GetHeight(rc)
    
    ' Excelのウィンドウハンドルを取得
    hWnd = FindWindow("XLMAIN", Application.Caption)
    
    ' ウィンドウの位置とサイズを設定
    MoveWindow hWnd, X, Y, Width, Height, True
End Sub

' 日付を表す文字列をDate型オブジェクトに変換する
Function CovertToDate(str As String) As Date
On Error GoTo ErrorHandler
    
    If InStr(str, ".") = 0 Then
        CovertToDate = DateSerial(CInt(Left(str, 4)), CInt(Mid(str, 5, 2)), CInt(Right(str, 2)))
    Else
        Dim part As Variant: part = Split(str, ".")
        CovertToDate = DateSerial(CInt(part(2)), CInt(part(1)), CInt(part(0)))
    End If
    
    Exit Function
    
ErrorHandler:
    CovertToDate = Date
End Function

Sub JumpToErrorOfOCR()
    ' エラー出力先テーブルの基準セル
    Dim tl2 As Range    ' TopLeft
    Dim inRange As Range, foundCell As Range, ErrorCol As Long
    Dim targetData As Variant
    With New MyTool
        With .FindTable("ErrorOfOCR")
            Set tl2 = .Range.Cells(1, 1) ' エラーテーブルのタイトル行左端列
            If Not .DataBodyRange Is Nothing Then
                Set inRange = Intersect(ActiveCell, .DataBodyRange)
                If Not inRange Is Nothing Then
                    ErrorCol = inRange.Column - .DataBodyRange.Columns(1).Column
                
                    If ErrorCol > 0 And ActiveCell <> "" Then
                        targetData = .DataBodyRange.Cells(inRange.Row - .DataBodyRange.Rows(1).Row + 1, ErrorCol + 1)
                        
                        Set foundCell = wsOCR.Cells.Find(What:=targetData, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
                        If Not foundCell Is Nothing Then
                            wsOCR.Activate
                            foundCell.Select
                        End If
                    Else
                        targetData = .DataBodyRange.Cells(inRange.Row - .DataBodyRange.Rows(1).Row + 1, 1)
                        
                        If IsNumeric(targetData) And targetData > 0 Then
                            Call wsOCR.Parent.Activate
                            Call Application.Goto(Cells(targetData, 1), True)
                            ActiveCell.EntireRow.Select
                        End If
                    End If
                End If
            End If
        End With
    End With
End Sub

Sub TestMatch()
    Debug.Print ExtractMaterialName("20250110_8 LN 126.7 RY L YA_OCR_SG.csv")
End Sub

Function ExtractMaterialName(text As String) As String
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim pattern As String

    ' 正規表現パターン
    pattern = "(.*)(?=_OCR)"

    ' 正規表現オブジェクトの作成
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.pattern = pattern

    ' 一致する部分文字列を取得
    If regex.Test(Mid(text, 10)) Then
        Set matches = regex.Execute(Mid(text, 10))
        If matches.Count > 0 Then
            Set match = matches(0)
            ExtractMaterialName = match.Value
        End If
    End If
End Function
