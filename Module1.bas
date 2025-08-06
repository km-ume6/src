Attribute VB_Name = "Module1"
Option Explicit

' 20250521 機能追加 auld-CheckData-NEW
' 特定のセルが空白の場合はアラートとする
' 20250519
' パッキングリスト特別処理追加
' 処理対象シートからマテリアルコードを取得する
Public MatCode As String
Public MatAddress As String
' 20250515
' 違和感行へのジャンプが機能していなかったのを修正


Private wbCoC As Workbook
Private wbPL As Workbook
'Private wbOCR As Workbook
Public mt As MyTool
Private NewNamePL As String ' 追加 20250515 by maruyama

Sub Init()
    If mt Is Nothing Then
        Set mt = New MyTool
    End If
    
    MatAddress = "A17"  ' 20250520 追加
End Sub

' アップロードファイルテーブルにファイルパスを書込む
'
' 引数
' sType ：ファイルの種類（CoC/PL/OCR）
' sPath ：ファイル名（フルパス）
Sub AddUploadList(sType As String, sPath As String)
    Init
    Dim dstTable As ListObject: Set dstTable = mt.FindTable("UploadList")
    If Not dstTable Is Nothing Then
        
        'パスが存在していれば何もしないように・・・
        Dim ret As Range: Set ret = Nothing
        If Not dstTable.DataBodyRange Is Nothing Then
            Set ret = dstTable.DataBodyRange.Find(What:=sPath, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False)
        End If
        
        If ret Is Nothing Then
            With dstTable.ListRows
                .Add
                .Item(.Count).Range(1).Value = sType
                .Item(.Count).Range(2).Value = sPath
            End With
        End If
    End If
End Sub

' ドロップダウンリストを表示する
Public Sub SendKeys2(Keys As String, Optional Wait As Boolean = False)
    Init
    
    Call mt.KeyDown(vbKeyMenu)
    Call mt.KeyDown(vbKeyDown)
    Call mt.KeyUp(vbKeyDown)
    Call mt.KeyUp(vbKeyMenu)
    
    DoEvents
End Sub

Sub JumpToRow()
    Call Init
    
    Dim wb As Workbook
    Select Case ActiveSheet.Name
        Case "CoC"
            Set wb = mt.IsOpened(mt.GetFileName(mt.Config("Config", ActiveSheet.Name & "_Path")))
        Case "PL"
            ' 変更 20250515 by maruyama
            'Set wb = mt.IsOpened(mt.ChangeExt(mt.GetFileName(mt.Config("Config", ActiveSheet.Name & "_Path")), "pck"))
            Set wb = mt.IsOpened(NewNamePL & ".pck", True)
        Case Else
            Exit Sub
    End Select
    
    If Not wb Is Nothing Then
        Dim targetRow As Variant: targetRow = ActiveSheet.Cells(ActiveCell.Row, 1)
        If IsNumeric(targetRow) = True And targetRow > 0 Then
            Call wb.Activate
            ActiveSheet.Cells(targetRow, Columns.Count).End(xlToLeft).Select
        End If
    End If
End Sub

' チェック対象ワークブックを開く
'
' 引数
' configKey     ：フォルダ保存用キー
' fdTitle       ：Application.FileDialogに渡す
' fdFilters     ：Application.FileDialogに渡す
' fdFilterExt   ：Application.FileDialogに渡す
'
' 返値
' vbOK          ：ブックを開いた
' vbCansel      ：ファイルダイアログでキャンセル
Function OpenBook(configKey As String, fdTitle As String, fdFilterDesc As String, fdFilterExt As String) As VbMsgBoxResult
    
    OpenBook = vbCancel
    
    ' ファイルダイアログを作成
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' ダイアログのタイトルとフィルターを設定
    Dim selectedFile As String
    With fd
        
        ' 初期フォルダを設定する
        Dim InitPath As String: InitPath = mt.Config("Config", configKey)
        If InitPath <> "" Then .InitialFileName = mt.GetParentFolderName(InitPath) & "\"
        
        .title = fdTitle
        .Filters.Clear
        .Filters.Add fdFilterDesc, fdFilterExt
        If .Show = True Then
        ' ユーザーがファイルを選択した場合
            Dim wb As Workbook
            selectedFile = .SelectedItems(1) ' 選択されたファイルのパスを取得
            
            ' 同名ブックが開かれていたら閉じる
            Set wb = mt.IsOpened(mt.GetFileName(selectedFile))
            If Not wb Is Nothing Then wb.Close
            
            Set wb = Workbooks.Open(selectedFile, ReadOnly:=True)  ' 選択されたファイルを開く
            OpenBook = vbOK
        End If
    End With

End Function

' チェック対象シートを選択する
'
' 引数
' tt    :target type <"CoC", "PL", "OCR">
'
' 返値は対象のワークシートオブジェクト
Function SelectTarget(tt As String) As Worksheet
    Set SelectTarget = Nothing
    
    Dim fdT As String, fdFD As String, fdFE As String
    Select Case tt
        Case "CoC"
            fdT = "CoCファイルを選択してください"
            fdFD = "CoCファイル"
            fdFE = "*.csv"
        Case "PL"
            fdT = "Packing Listファイルを選択してください"
            fdFD = "Packing Listファイル"
            'fdFE = "*.xlsx;*.xlsm"
            fdFE = "*.xls*"
        Case "OCR"
            fdT = "OCRリストファイルを選択してください"
            fdFD = "OCRリストファイル"
            fdFE = "*.csv;*.xlsx"
    End Select
    
    Dim wbTarget As Workbook: Set wbTarget = Nothing
    Dim wbTemp As Workbook
    Dim flagOpen As Boolean: flagOpen = False
    
    If Workbooks.Count = 1 Then
        flagOpen = True
    ElseIf Workbooks.Count = 2 Then
        For Each wbTemp In Workbooks
            If wbTemp.Name <> ThisWorkbook.Name Then
                Exit For
            End If
        Next
        
        If vbYes = MsgBox("このブック [" & wbTemp.Name & "] が対象ですか？", vbYesNo Or vbQuestion, "アップロード自動化") Then
            Set wbTarget = wbTemp
        Else
            flagOpen = True
        End If
    Else
        Dim menuB As New Collection
        For Each wbTemp In Workbooks
            If wbTemp.Name <> ThisWorkbook.Name Then
                menuB.Add wbTemp.Name
            End If
        Next
                
        Dim wbn As String: wbn = mt.MenuBox(menuB, "どのブックが対象ですか？" & vbLf & "選択肢に含まれていなければキャンセルを！" & vbCrLf & vbCrLf, "アップロード自動化")
        If wbn <> "" Then
            Set wbTarget = Workbooks(wbn)
        Else
            flagOpen = True
        End If
    End If
    
    If flagOpen Then
        If vbOK = OpenBook(tt & "_Path", fdT, fdFD, fdFE) Then
            Set wbTarget = ActiveWorkbook
        End If
    End If
    
    ' ここまででワークブックが特定されれば、ワークシートの特定に進む
    If Not wbTarget Is Nothing Then
        
        Call mt.Config("Config", tt & "_Path", wbTarget.Path & "\" & wbTarget.Name)  ' 対象ワークブック名（フルパス）を保存する
        Call mt.PopAppSize(wbTarget, tt)
        
        If wbTarget.Worksheets.Count = 1 Then
        ' シートがひとつしかないとき
            If vbYes = MsgBox("このシート [" & wbTarget.Worksheets(1).Name & "] が対象ですか？", vbYesNo Or vbQuestion, "アップロード自動化") Then
                Set SelectTarget = wbTarget.Worksheets(1)
            End If
        Else
        ' シートが複数あるときは選択肢から選ばせる
            Dim menuS As New Collection
            Dim wsTemp As Worksheet
            For Each wsTemp In wbTarget.Worksheets
                menuS.Add wsTemp.Name
            Next
                    
            Dim wsn As String: wsn = mt.MenuBox(menuS, "どのシートが対象ですか？", "アップロード自動化")
            If wsn <> "" Then
                Set SelectTarget = wbTarget.Worksheets(wsn)
            End If
        End If
        
        ' パッキングリストの特別処理
        If tt = "PL" Then
            ' auld-CheckData-NEW
            Call DeleteCellsPL(SelectTarget)
        End If
    End If
End Function

' CoC違和感チェック
Sub CheckCoC()
    
    Call Init
    
    Dim wsTarget As Worksheet: Set wsTarget = SelectTarget("CoC")
    If Not wsTarget Is Nothing Then
        
        ' 処理結果出力先テーブルの基準セル
        Dim tl As Range    ' TopLeft
        With mt.FindTable("CheckCoC")
            Set tl = .Range.Cells(1, 1) ' タイトル行
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete ' 処理結果出力テーブルをクリア
            End If
        End With
        
        ' 画面描画用のフラグをセット
        mt.DrawScroll = mt.Config("Config", "DrawScroll")
        Dim ds As Boolean: ds = mt.Config("Config", "DrawSelect")
    
        Dim rCell As Range, lCell As Range, l As Long, s As String, d As Integer
        Dim sc As Range: Set sc = wsTarget.Range("A1").CurrentRegion.Columns(1)
        Dim flagNotice As Boolean, productName As String
        Dim cntNotice As Integer: cntNotice = 0
        For l = sc.Row To sc.Rows.Count    ' 処理対象の行数でループする
            flagNotice = False
            
            ' 左端セルを選択（CoC）
            Set lCell = sc.Cells(l, 1)
            If ds Then
                wsTarget.Parent.Activate
                lCell.Select
            End If
            
            ' 行の右端までアクティブセルを動かす
            Set rCell = sc.Cells(l, Columns.Count).End(xlToLeft)
            Call mt.ScrollCell(lCell, rCell)
            
            If l > 3 Then
                ' その行が表す製品が６インチなのか８インチなのか？
                ' １列目の先頭文字で判断する。
                productName = sc.Cells(l, 1)
                s = Left(productName, 1)
                If IsNumeric(s) = True Then
                     d = Val(s)
                                    
                    ' 違和感チェック
                    If mt.EndWithString(rCell, ";;;;;") = True Then
                    ' 右端セルの内容がセミコロン5連続、かつ８インチ行のとき
                        If d = 8 Then
                            flagNotice = True
                        End If
                    Else
                    ' 右端セルの内容がセミコロン5連続ではなく、かつ６インチ行のとき
                        If d = 6 Then
                            flagNotice = True
                        End If
                    End If
                End If
            Else
                If l = 1 Then
                    flagNotice = Not mt.EndWithString(rCell, ";;;;;")
                Else
                    If Right(rCell, 1) = ";" Then flagNotice = True
                End If
            End If
            
            ' 違和感を表明する
            If flagNotice = True Then
                cntNotice = cntNotice + 1
                tl.Offset(cntNotice, 0) = l
                tl.Offset(cntNotice, 1) = s & " インチ"
                tl.Offset(cntNotice, 2) = "違和感あり"
            End If
                    
            ' 行の左端までアクティブセルを動かす（CoC）
            Call mt.ScrollCell(rCell, lCell)
            
            If ds Then
                ThisWorkbook.Activate
                tl.Offset(l, 0).Select
                mt.ReDraw 1
            End If
        Next
        
        If cntNotice = 0 Then
            MsgBox prompt:="問題ありませんでした", title:="CoCチェック結果"
            
            ' アップロードエントリー
            Call AddUploadList("CoC", mt.GetLocalFullPath(wsTarget.Parent))
        Else
            MsgBox prompt:="違和感ありです。元のファイルを修正してもらってください。", title:="CoCチェック結果"
            If vbYes = MsgBox(prompt:="ブックを閉じても良いですか？", Buttons:=vbYesNo, title:="CoCチェック結果") Then wsTarget.Parent.Close
        End If
    End If
End Sub

' ワークシートをテキストに変換して開き直す
' 引数
' ws    ：ワークシートオブジェクト
' ext   ：拡張子
Function SelectWorksheet(ws As Worksheet, ext As String, Optional spFlag As Boolean = False) As Worksheet
    
    If Not ws Is Nothing Then
        Dim wb As Workbook: Set wb = ws.Parent
        
        ' auld-CheckData-NEW
        If HasBlanks = True Then
            HasBlanks = False
            Call CloseAndDelete(wb)
            Set wb = Nothing
            Set SelectWorksheet = Nothing
            Exit Function
        End If
        
        If Not wb Is Nothing Then
            wb.Activate
            Set SelectWorksheet = ActiveSheet
            'Call mt.PushAppSize
            
            If Not mt.GetExtensionName(wb.Name) = ext Then
                Dim newName As String
                If spFlag <> True Then
                    newName = mt.ChangeExt(wb.Name, ext)
                Else
                    newName = mt.MakeBaseNamePL(ws, ext)    ' パッキングリストファル名変換
                End If
                Dim newFile As String: newFile = mt.GetDesktopPath() & "\" & newName
                
                Dim wbTemp As Workbook: Set wbTemp = mt.IsOpened(newName)
                If Not wbTemp Is Nothing Then wbTemp.Close SaveChanges:=False
                
                ws.Copy
                
                ' CSV作成
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs Filename:=newFile, FileFormat:=xlCSV, CreateBackup:=False
                Application.DisplayAlerts = True
                ActiveWorkbook.Close
                
                If Dir(newFile) <> "" Then
                    Call Workbooks.OpenText(newFile, 932, 1)
                    'Call mt.PopAppSize
                
                    Set SelectWorksheet = ActiveSheet
                Else
                    Call MsgBox(prompt:="ファイルが保存されていません。やり直してください！", Buttons:=vbOK Or vbCritical, title:="自動アップロード")
                End If
                
                Dim killBookName As String: If spFlag = True Then killBookName = mt.GetLocalFullPath(wb)
                
                If spFlag <> True Then
                    wb.Close
                Else
                ' パッキングリスト処理のときはワークブックを削除する
                    wb.Close SaveChanges:=False
                    Do While Dir(killBookName) <> ""
                        Kill killBookName
                        Call mt.mtSleep(1000)
                    Loop
                    
                End If
            End If
        Else
            wb.Close
        End If
    End If
    
End Function

' PackingList違和感チェック
Sub CheckPL()
    
    Call Init
    
    NewNamePL = ""  ' 追加 20250515 by maruyama, auld-CheckData-NEW

    ' ファイル名変換の有効・無効フラグ取得
    Dim f As Boolean
    If mt.Config("Config", "ConvertFileNamePL") = "0" Then f = False Else f = True
    
    Dim wsTarget As Worksheet: Set wsTarget = SelectWorksheet(SelectTarget("PL"), "pck", f)
    
    If Not wsTarget Is Nothing Then
        
        NewNamePL = wsTarget.Name     ' 追加 20250515 by maruyama, auld-CheckData-NEW
        
        Call mt.PopAppSize(wsTarget.Parent, "PL")
        
        ' 処理結果出力先テーブルの基準セル
        Dim tl As Range    ' TopLeft
        With mt.FindTable("CheckPL")
            Set tl = .Range.Cells(1, 1) ' タイトル行
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete ' 処理結果出力テーブルをクリア
            End If
        End With
        
        Dim ds As Boolean: ds = mt.Config("Config", "DrawSelect")
        Dim r As Range, l As Long, cntSafeWord As Integer
        Dim sc As Range: Set sc = wsTarget.Range("A1").CurrentRegion.Columns(1)
        Dim cntNotice As Integer: cntNotice = 0
        For l = sc.Row To sc.Rows.Count    ' 処理対象の行数でループする
            
            ' 右端セルを選択
            Set r = sc.Cells(l, Columns.Count).End(xlToLeft)
            If ds = True Then
                wsTarget.Parent.Activate
                r.Select
            End If
            
            ' 違和感表明
            If Right(r, 1) <> "," Then
            ' 行末がカンマでなければ違和感あり！
            
                ' 例外処理（対象行にその単語が含まれていれば不問とする）
                cntSafeWord = 0
                cntSafeWord = cntSafeWord + InStr(r, "Measurement")
                cntSafeWord = cntSafeWord + InStr(r, "T O T A L")
                
                If cntSafeWord = 0 Then
                    cntNotice = cntNotice + 1
                    tl.Offset(cntNotice, 0) = l
                    tl.Offset(cntNotice, 1) = "違和感あり"
                    r.Select
                End If
            End If
            
            If ds = True Then
                ThisWorkbook.Activate
                tl.Offset(l, 0).Select
                mt.ReDraw 1
            End If
        Next
        
        ' 最下行から遡って連続カンマを削除する
        For l = sc.Rows.Count To sc.Row Step -1
            If mt.IsAllSameCharacters(sc.Cells(l)) = 7 Then
                sc.Cells(l) = ""
            Else
                Exit For
            End If
        Next l
        
        If cntNotice = 0 Then
            MsgBox prompt:="問題ありませんでした", title:="PackingListチェック結果"
            
            ' アップロードエントリー
            Call AddUploadList("PL", mt.GetLocalFullPath(wsTarget.Parent))
        Else
            MsgBox prompt:="違和感ありです。元のファイルを修正してもらってください。", title:="PackingListチェック結果"
            If vbYes = MsgBox(prompt:="ブックを閉じても良いですか？", Buttons:=vbYesNo, title:="PackingListチェック結果") Then wsTarget.Parent.Close
        End If
        
    End If
End Sub


' OCRリスト違和感チェック
Sub CheckOCR()
    
    Call Init
    
    Dim wsTarget As Worksheet: Set wsTarget = SelectWorksheet(SelectTarget("OCR"), "csv")
    
    If Not wsTarget Is Nothing Then
        
        ' 処理結果出力先テーブルの基準セル
        Dim tl As Range    ' TopLeft
        With mt.FindTable("CheckOCR")
            Set tl = .Range.Cells(1, 1) ' タイトル行
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete ' 処理結果出力テーブルをクリア
            End If
        End With
        
        ' 2025/02/17 追加
        ' ↓↓↓↓↓↓↓↓
        Set wsOCR = wsTarget
        
        Dim ParentCaption As String: ParentCaption = wsOCR.Parent.Parent.Caption
        Dim DeliveryDate As Date: DeliveryDate = CovertToDate(Left(ParentCaption, 8))
        Dim RowOfError As Integer: RowOfError = 1
        Dim IsError As Boolean
        
        ' エラー出力先テーブルの基準セル
        Dim tl2 As Range    ' TopLeft
        With mt.FindTable("ErrorOfOCR")
            Set tl2 = .Range.Cells(1, 1) ' タイトル行
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete ' 処理結果出力テーブルをクリア
            End If
        End With
        
        ' Materialチェック用文字列生成
        Dim materialName As String: materialName = ExtractMaterialName(ParentCaption)
        If materialName = "" Then MsgBox "ファイル名がおかしいかも！"
        ' ↑↑↑↑↑↑↑↑
        
        Dim ds As Boolean: ds = mt.Config("Config", "DrawSelect")
        Dim r As Range, l As Long
        Dim sc As Range: Set sc = wsTarget.Range("A1").CurrentRegion.Columns(1)
        Dim idDic As New Dictionary
        Dim idDayDic As New Dictionary
        Dim lotDic As New Dictionary
        Dim lotDayDic As New Dictionary
        Dim s As String, cnt As Integer
        For l = sc.Row To sc.Rows.Count    ' 処理対象の行数でループする
            Set r = sc.Cells(l, 1)
            If Not r.Offset(0, 1) = "WaferID" Then
                
                ' WaferID で処理
                s = Left(r.Offset(0, 1), 6)
                If idDic.Exists(s) = True Then
                    idDic.Item(s) = idDic.Item(s) + 1
                Else
                    Call idDic.Add(s, 1)
                    Call idDayDic.Add(s, r.Offset(0, 4))
                End If
                
                ' 2025/02/17 処理追加
                ' ↓↓↓↓↓↓↓↓↓↓
                IsError = False
                                
                ' WaferIDチェック
                If Len(sc.Cells(l, 2)) <> 12 Then
                    IsError = True
                    tl2.Offset(RowOfError, 1) = sc.Cells(l, 2)
                End If
                
                ' Materialチェック
                'If InStr(ParentCaption, sc.Cells(l, 3)) = 0 Then
                If materialName <> sc.Cells(l, 3) Then
                    IsError = True
                    tl2.Offset(RowOfError, 2) = sc.Cells(l, 3)
                End If
                
                ' LotNoチェック
                If Len(sc.Cells(l, 4)) <> 16 Then
                    IsError = True
                    tl2.Offset(RowOfError, 3) = sc.Cells(l, 4)
                End If
                
                ' 日付チェック
                If DeliveryDate <> CovertToDate(sc.Cells(l, 5)) Then
                    IsError = True
                    tl2.Offset(RowOfError, 4) = sc.Cells(l, 5)
                End If
                
                If IsError = True Then
                    tl2.Offset(RowOfError, 0) = l
                    RowOfError = RowOfError + 1
                End If
                ' ↑↑↑↑↑↑↑↑↑↑
            End If
        Next
        
        ' けっかはっぴょーーーーう①
        Dim vKey As Variant, sumWafer As Integer
        l = 1
        sumWafer = 0
        For Each vKey In idDic.Keys
            tl.Offset(l, 0) = Split(vKey, "")
            tl.Offset(l, 1) = idDic.Item(vKey)
            tl.Offset(l, 2) = idDayDic.Item(vKey)
            sumWafer = sumWafer + idDic.Item(vKey)
            
            l = l + 1
        Next
        tl.Offset(l, 0) = "合計枚数"
        tl.Offset(l, 1) = sumWafer
        l = l + 1
        
        ' アップロードエントリー
        'Call AddUploadList("OCR", mt.GetLocalFullPath(wsTarget.Parent))
        ' 2025/02/17 変更（エラーがなかったときだけエントリーする）
        With mt.FindTable("ErrorOfOCR")
            If .DataBodyRange Is Nothing Then
                Call AddUploadList("OCR", mt.GetLocalFullPath(wsTarget.Parent))
            End If
        End With
    End If
    
End Sub

' アップロード実行
Sub UploadFiles()
    Call Init
    Call AutoUpload
End Sub

Sub DeleteUploadFileList()
    Init
    
    Dim flTable As ListObject: Set flTable = mt.FindTable("UploadList")
    If Not flTable Is Nothing Then
        If Not flTable.DataBodyRange Is Nothing Then
            flTable.DataBodyRange.Delete
        End If
    End If
End Sub

Sub AutoUpload()
    
    Init
    
    Dim patternStr As String: patternStr = mt.Config("ConfigFTP", "パターン")
    Dim flTable As ListObject: Set flTable = mt.FindTable("UploadList")
    If patternStr = "" Or flTable Is Nothing Then Exit Sub
    
    Dim v As Variant: v = Split(patternStr, ",")
    Dim i As Integer, s1 As String, s2 As String, wb As Workbook
    Dim fileList As New Collection
    Dim lr As ListRow
    For i = LBound(v) To UBound(v)
        
        s1 = v(i)
        s2 = ""
        
        For Each lr In flTable.ListRows
            If lr.Range(1) = s1 Then
                s2 = lr.Range(2)
                If Len(Dir(s2)) > 0 Then
                    Call fileList.Add(s2)
                    Set wb = mt.IsOpened(mt.GetFileName(s2))
                    If Not wb Is Nothing Then wb.Close
                End If
            End If
        Next
        
        If s2 = "" Then
            MsgBox "ファイルが揃っていません！"
            Exit Sub
        End If
    
    Next i
    
    ' コマンドのオプション
    Dim optionCmd As String: optionCmd = "cd " & mt.Config("ConfigFTP", "アップロード先")
    
    'コマンドファイル作成
    Open "auld.txt" For Output As #1
    If optionCmd <> "" Then Print #1, optionCmd
    For i = 1 To fileList.Count
        Print #1, "put " & mt.NormalizePath(fileList(i))
    Next i
    Print #1, "quit"
    Close #1
    
    ' 鍵ファイル名を取得
    Dim keyFileName As String: keyFileName = mt.Config("ConfigFTP", "PPK")
    If Len(Dir(keyFileName)) = 0 Then
        keyFileName = mt.GetParentFolderName(mt.GetLocalFullPath(ThisWorkbook)) & "\" & keyFileName
        If Len(Dir(keyFileName)) <> 0 Then
            keyFileName = mt.NormalizePath(keyFileName)
        Else
            MsgBox "鍵ファイルが見つかりません"
            Exit Sub
        End If
    End If
    
    ' アップロード
    Dim objWSH As Object: Set objWSH = CreateObject("WScript.Shell")
    Dim sCmd As String: sCmd = """C:\Program Files\PuTTY\psftp.exe"" -i " & keyFileName & " " & mt.Config("ConfigFTP", "User") & "@" & mt.Config("ConfigFTP", "Server") & " -bc -be -b auld.txt"
    Dim objExec As Object: Set objExec = objWSH.Exec(sCmd)  ' WshScriptExecオブジェクト取得
    objExec.StdIn.Close
    
    ' ログ表示
    Dim textStream As String: textStream = objExec.StdOut.ReadAll
    If objExec.ExitCode <> 0 Then
        textStream = textStream & vbCrLf & "エラー発生！" & vbCrLf
        textStream = textStream & objExec.StdErr.ReadAll
    End If
    MsgBox prompt:=textStream, title:="アップロードの結果"
    
    'オブジェクトを破棄
    Set objWSH = Nothing
    Set objExec = Nothing

    Call DeleteUploadFileList
    Kill "auld.txt"
    
End Sub

' パッキングリストの不要なセルを削除する
' ①書式内の"T O T A L"と入力された右隣のセルの計算結果が0（空白）の場合、該当ページ以降の行を削除する
' ②列H～Kを削除する
' 引数
' ws    ：対象となるワークシート
Sub DeleteCellsPL(ws As Worksheet)
    'Call DeleteCellsPL_B(ws)
    Call DeleteCellsPL_B(ws)
End Sub

Sub DeleteCellsPL_A(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    
    Init
    
    ' ワークブックを別名保存する（デスクトップへ！）
    With ws.Parent
        Dim sheetName As String: sheetName = ws.Name
        Dim srcPath As String: srcPath = mt.GetLocalFullPath(ws.Parent)
        Dim dstPath As String: dstPath = mt.GetDesktopPath() & "\tmp" & .Name
        .Close SaveChanges:=False
        FileCopy srcPath, dstPath
        
        Dim wbNew As Workbook: Set wbNew = Workbooks.Open(dstPath)
        Set ws = wbNew.Worksheets(sheetName)
    End With
    
    Dim cellList As New Collection
    Set cellList = mt.SearchStringFromS(ws, "T O T A L", xlWhole, "Lot No.")
    
    ' 計算結果（右隣セル）が0（空白）となっている"T O T A L"を検索
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
End Sub
