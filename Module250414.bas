Attribute VB_Name = "Module250414"
Option Explicit

Public HasBlanks As Boolean

Sub ModifyWindowSize()
Attribute ModifyWindowSize.VB_Description = "ウィンドウサイズを保存します"
Attribute ModifyWindowSize.VB_ProcData.VB_Invoke_Func = "w\n14"
    With New MyTool
        Dim menuB As New Collection
        menuB.Add "Me"
        menuB.Add "CoC"
        menuB.Add "PL"
        menuB.Add "OCR"
                
        Dim TargetWindow As String: TargetWindow = .MenuBox(menuB, "どのウィンドウが対象ですか？" & vbLf & "選択肢に含まれていなければキャンセルを！" & vbCrLf & vbCrLf, "ウィンドウサイズ調整")
        If TargetWindow <> "" Then
            Call .PushAppSize(ThisWorkbook, TargetWindow)
        End If
    End With
End Sub

Function InsertBeforeExtension(filePath As String, insertText As String) As String
    Dim baseName As String
    Dim extension As String
    Dim dotPos As Long

    dotPos = InStrRev(filePath, ".")
    If dotPos > 0 Then
        InsertBeforeExtension = Left(filePath, dotPos - 1) & insertText & Mid(filePath, dotPos)
    Else
        InsertBeforeExtension = baseName & insertText
    End If
End Function

' パッキングリストの特定セルが空白かどうかを調査する
Sub HasBlankCells(ws As Worksheet)
    Call HasBlankCells_B(ws)
End Sub

Sub HasBlankCells_B(ws As Worksheet)

    HasBlanks = False
    If ws Is Nothing Then Exit Sub
    
    ' "MADE IN JAPAN"と入力されたセル一覧を取得する
    Dim mijs As New Collection
    Set mijs = mt.SearchStringFromS(ws, "MADE IN JAPAN", xlPart, "No.1")
    
    ' "T O T A L"と入力されたセル一覧を取得する
    Dim totals As New Collection
    Set totals = mt.SearchStringFromS(ws, "T O T A L", xlWhole, "Lot No.")
    
    If mijs.Count = totals.Count Then
        
        ' 空白チェック
        Dim isBlankCell As Range, rTotal As Range
        Dim AlertString As String: AlertString = ""
        Dim i As Integer
        For i = 1 To mijs.Count
            AlertString = MakeAlertString(mijs(i).Offset(-1, 0), AlertString)
            AlertString = MakeAlertString(totals(i).Offset(0, 1), AlertString)
            AlertString = MakeAlertString(totals(i).Offset(0, 2), AlertString)
            AlertString = MakeAlertString(totals(i).Offset(1, 3), AlertString)
            AlertString = MakeAlertString(totals(i).Offset(1, 4), AlertString)
            
            If AlertString <> "" Then AlertString = AlertString & vbCrLf
        Next i
    
        If AlertString <> "" Then
            Call MsgBox("空白確認対象セルに空白があります。元のファイルを確認して下さい。" & vbCrLf & AppendAfterCrLf(AlertString), vbOKOnly, "アップロード自動化")
            HasBlanks = True
        End If
    Else
        Call MsgBox("""MADE IN JAPAN"" と ""T O T A L"" の数が合いませんでしたので空白チェックは実施しません。", vbOKOnly, "空白チェック")
    End If

End Sub

Function MakeAlertString(r As Range, s As String) As String
    If r = "" Then
        If Not (s = "" Or Right(s, Len(vbCrLf)) = vbCrLf) Then s = s & ", "
        MakeAlertString = s & r.Address(False, False)
    Else
        MakeAlertString = s
    End If
End Function

Sub HasBlankCells_A(ws As Worksheet)

    HasBlanks = False
    If ws Is Nothing Then Exit Sub
    
    ' "T O T A L"と入力されたセル（ページの基準）一覧を取得する
    Dim totals As New Collection
    Set totals = mt.SearchStringFromS(ws, "T O T A L", xlWhole, "Lot No.")
    
    ' 設定文字列から空白チェック対象のセルアドレスを取得し、基準セルからのオフセットを配列にセット
    Dim IsBlanks(1 To 5, 1 To 2) As Long
    Dim cellList As String: cellList = mt.Config("Config", "IsBlankCells")
    Dim CellAddress As Variant: CellAddress = Split(cellList, ",")
    Dim cellBase As Range: Set cellBase = totals.Item(1)
    Dim cellTemp As Range
    Dim i As Integer
    For i = LBound(CellAddress) To UBound(CellAddress)
        Set cellTemp = ws.Range(CellAddress(i))
         IsBlanks(i + 1, 1) = cellTemp.Row - cellBase.Row
         IsBlanks(i + 1, 2) = cellTemp.Column - cellBase.Column
    Next i

    ' 空白チェック
    Dim isBlankCell As Range, rTotal As Range
    Dim CellAddress As String: CellAddress = ""
    Dim MsgString As String
    For Each rTotal In totals
        Msg
        For i = 1 To 5
            Set isBlankCell = rTotal.Offset(IsBlanks(i, 1), IsBlanks(i, 2))
            If isBlankCell.Value = "" Then
                If Not (CellAddress = "" Or Right(CellAddress, Len(vbCrLf)) = vbCrLf) Then CellAddress = CellAddress & ", "
                CellAddress = CellAddress & isBlankCell.Address(False, False)
            End If
        Next i
        
        If CellAddress <> "" Then CellAddress = CellAddress & vbCrLf
    Next

    If CellAddress <> "" Then
        If vbYes <> MsgBox("空白確認対象セルに空白があります。このまま続けますか？" & vbCrLf & AppendAfterCrLf(CellAddress), vbYesNo, "アップロード自動化") Then HasBlanks = True
    End If

End Sub

Sub CloseAndDelete(wb As Workbook)
    Dim killBookName As String: killBookName = mt.GetLocalFullPath(wb)
    
    wb.Close SaveChanges:=False
    Do While Dir(killBookName) <> ""
        Kill killBookName
        Call mt.mtSleep(1000)
    Loop
End Sub

Sub DeleteCellsPL_B(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    
    Init
    
    ' ワークブックを別名保存する（デスクトップへ！）
    With ws.Parent
        Dim sheetName As String: sheetName = ws.Name
        Dim srcPath As String: srcPath = mt.GetLocalFullPath(ws.Parent)
        Dim dstPath As String: dstPath = ChangeFileExtension(mt.GetDesktopPath() & "\tmp" & .Name, "xlsx")
        
        ' マクロなしブックとして保存する
        Application.DisplayAlerts = False: .SaveAs Filename:=dstPath, FileFormat:=xlOpenXMLWorkbook: Application.DisplayAlerts = False
        
        .Close SaveChanges:=False
        'FileCopy srcPath, dstPath
        
        Dim wbNew As Workbook: Set wbNew = Workbooks.Open(dstPath)
        Set ws = wbNew.Worksheets(sheetName)
        
        ' auld-CheckData-NEW
        ws.UsedRange.Value = ws.UsedRange.Value ' 関数を削除する
    End With
    
    Dim cellList As New Collection
    Set cellList = mt.SearchStringFromS(ws, "T O T A L", xlWhole, "Lot No.")
    
    ' 計算結果（右隣セル）が空白となっている"T O T A L"を検索
    Dim rTotal As Range
    For Each rTotal In cellList
        If rTotal.Offset(0, 1) = 0 Then Exit For
    Next
    
    ' 該当する書式のページ番号（ページの左上に当たる）を検索
    If Not rTotal Is Nothing Then
        Dim rNo As Range: Set rNo = rTotal.Offset(0, -2)
        Do While rNo.Row <> 1
            If rNo Like "No.*" Then
                Exit Do
            End If
            Set rNo = rNo.Offset(-1, 0)
        Loop
        
        ' 書式の最下段検索
        Dim rEnd As Range: Set rEnd = ws.Cells(Rows.Count, rNo.Column).End(xlUp)
        
        ' 行を削除する
        ' ws.Range(rNo.Offset(-1, 0).Row & ":" & rEnd.Row).Delete
        If rNo.Row <> 1 Then ws.Range(rNo.Offset(-1, 0).Row & ":" & rEnd.Row).Delete
    End If
    
    ws.Range(mt.Config("Config", "DelColPL")).Delete    ' 列を削除する
    
    ' ↑ここまでで不要なセルを削除している！
    
    ' auld-CheckData-NEW
    Call HasBlankCells(ws)
End Sub

Function MakeBaseNamePL4(ws As Worksheet, ext As String) As String
    Dim sdCell As String: sdCell = mt.Config("Config", "sdCell")        ' 出荷日のセルアドレス
    Dim ccCell As String: ccCell = mt.Config("Config", "ccCell")        ' 出荷先名（国名）のセルアドレス
    MakeBaseNamePL4 = Format(ws.Range(sdCell), "yyyymmdd") & " " & Left(StrConv(Replace(ws.Range(ccCell), " ", ""), vbProperCase), 3) & " Packing list " & ws.Name & "." & ext
End Function

Function ChangeFileExtension(filePath As String, newExtension As String) As String
    Dim dotPos As Long
    Dim basePath As String

    ' 拡張子の前のドットの位置を取得
    dotPos = InStrRev(filePath, ".")

    If dotPos > 0 Then
        ' 拡張子を除いた部分を取得
        basePath = Left(filePath, dotPos - 1)
    Else
        ' 拡張子がない場合はそのまま
        basePath = filePath
    End If

    ' 新しい拡張子を追加（ピリオドを自動で付ける）
    If Left(newExtension, 1) <> "." Then
        newExtension = "." & newExtension
    End If

    ChangeFileExtension = basePath & newExtension
End Function

Function AppendAfterCrLf(srcString As String) As String
    Dim i As Long
    Dim l As Long: l = 1
    Dim tempChar As String
    Dim lineBuffer As String

    AppendAfterCrLf = "No." & l & "："

    ' 1文字ずつ処理
    For i = 1 To Len(srcString)
        tempChar = Mid(srcString, i, 1)
        lineBuffer = lineBuffer & tempChar

        ' 改行コードの検出（vbCrLfは2文字なので、直前の2文字で判定）
        If Right(lineBuffer, 2) = vbCrLf Then
            l = l + 1
            AppendAfterCrLf = AppendAfterCrLf & lineBuffer & "No." & l & "："
            lineBuffer = ""
        End If
    Next i
End Function

