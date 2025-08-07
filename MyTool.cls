VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 以下を参照設定すること！
' Microsoft Scripting Runtime
' Microsoft VBScript Regular Expressions 5.5

Option Explicit

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private Declare PtrSafe Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_EXTENDEDKEY = &H1   'キー押下
Const KEYEVENTF_KEYUP = &H2         'キー解放

Private ds As Boolean
Private appLeft As Integer
Private appTop As Integer
Private appWidth As Integer
Private appHeight As Integer

Private wso As Object
Private pathDictionary As Dictionary

' プロパティ
Public Property Let DrawScroll(f As Boolean)
    ds = f
End Property

Public Property Get DrawScroll() As Boolean
    DrawScroll = ds
End Property

' APIラッパー
Function mtGetKeyState(kc As Long) As Integer
    mtGetKeyState = GetKeyState(kc)
End Function

Public Sub KeyDown(keyCode As Byte)
    keybd_event keyCode, 0, KEYEVENTF_EXTENDEDKEY, 0
End Sub

Public Sub KeyUp(keyCode As Byte)
    keybd_event keyCode, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End Sub

Public Sub SendKeys2(Keys As String, Optional Wait As Boolean = False)
    If wso Is Nothing Then
        Set wso = CreateObject("WScript.Shell")   ' WshShellオブジェクトを生成
    End If
    
    Call wso.SendKeys(Keys, Wait)
End Sub

Public Sub mtSleep(st As Long)
    Call Sleep(st)
End Sub

' 文字列（S1）が文字列（S2）で終わっているのか？
Function EndWithString(s1, s2) As Boolean
    EndWithString = False
    If Len(s1) >= Len(s2) Then
        If Right(s1, Len(s2)) = s2 Then
            EndWithString = True
        End If
    End If
End Function

' パッキンリストのファイル名を作る特別処理
'
' 引数
' ws    ：ファイル名を作るときの部品を含むワークシート
' ext   ：ファイルの拡張子（これまでの経緯により）
Function MakeBaseNamePL(ws As Worksheet, ext As String) As String
    ' MakeBaseNamePL = MakeBaseNamePL1(ws, ext)
    ' MakeBaseNamePL = MakeBaseNamePL2(ws, ext)
    ' MakeBaseNamePL = MakeBaseNamePL3(ws, ext)
    ' MakeBaseNamePL = MakeBaseNamePL4(ws, ext)   ' 新関数！20250514 by maruyama
    MakeBaseNamePL = MakeBaseNamePL5(ws, ext)   ' 新関数！20250806 by maruyama
End Function

Function MakeBaseNamePL1(ws As Worksheet, ext As String) As String
    Dim productClass As String: productClass = ""
    With New RegExp
        .pattern = "\([A-Z]-\d..\)"
        .Global = True
        Dim mc As MatchCollection: Set mc = .Execute(ws.Name)
        If mc.Count > 0 Then
            ' 丸カッコ内から品種情報文字列を抽出してハイフンで分割
            Dim v As Variant: v = Split(ExtractContents(ws.Name, "()"), "-")
            
            productClass = Config("RuleTable", v(0) & "PC" & v(1))      ' 品種名作成
        Else
            productClass = Config("RuleTable", InputBox(prompt:="シート名から品種名を作ることができませんでしたので、ここで入力してください。", title:="チェックPL"))
        End If
    End With
    
    Dim sdCell As String: sdCell = mt.Config("Config", "sdCell")        ' 出荷日のセルアドレス
    Dim ccCell As String: ccCell = mt.Config("Config", "ccCell")        ' 出荷先名（国名）のセルアドレス
    MakeBaseNamePL1 = Format(ws.Range(sdCell), "yyyymmdd") & " " & Left(StrConv(Replace(ws.Range(ccCell), " ", ""), vbProperCase), 3) & " Packing list " & productClass & "." & ext
End Function

Function MakeBaseNamePL2(ws As Worksheet, ext As String) As String
    Dim productClass As String: productClass = ""
    Dim sepPos As Integer: sepPos = InStr(1, ws.Name, "(", vbTextCompare)
    If sepPos > 1 Then
        productClass = Config("RuleTable", Left(ws.Name, sepPos - 1))   ' 品種名作成
    End If
        
    If productClass = "" Then
        productClass = Config("RuleTable", InputBox(prompt:="シート名から品種名を作ることができませんでしたので、ここで入力してください。", title:="チェックPL"))
    End If
    
    Dim sdCell As String: sdCell = mt.Config("Config", "sdCell")        ' 出荷日のセルアドレス
    Dim ccCell As String: ccCell = mt.Config("Config", "ccCell")        ' 出荷先名（国名）のセルアドレス
    MakeBaseNamePL2 = Format(ws.Range(sdCell), "yyyymmdd") & " " & Left(StrConv(Replace(ws.Range(ccCell), " ", ""), vbProperCase), 3) & " Packing list " & productClass & "." & ext
End Function

Function MakeBaseNamePL3(ws As Worksheet, ext As String) As String
    Dim productClass As String: productClass = ""
    Dim sepPos As Integer: sepPos = InStr(1, ws.Name, "(", vbTextCompare)
    
    ' (の位置が７文字目より後ろの場合は品種連結（）と判定し、-を探す
    If sepPos > 7 Then sepPos = InStr(1, ws.Name, "-", vbTextCompare)
    
    If sepPos > 1 Then
        productClass = Config("RuleTable", Left(ws.Name, sepPos - 1))   ' 品種名作成
    End If
        
    If productClass = "" Then
        productClass = Config("RuleTable", InputBox(prompt:="シート名から品種名を作ることができませんでしたので、ここで入力してください。", title:="チェックPL"))
    End If
    
    Dim sdCell As String: sdCell = mt.Config("Config", "sdCell")        ' 出荷日のセルアドレス
    Dim ccCell As String: ccCell = mt.Config("Config", "ccCell")        ' 出荷先名（国名）のセルアドレス
    MakeBaseNamePL3 = Format(ws.Range(sdCell), "yyyymmdd") & " " & Left(StrConv(Replace(ws.Range(ccCell), " ", ""), vbProperCase), 3) & " Packing list " & productClass & "." & ext
End Function

' 文字列中の括弧記号で囲まれた部分を取り出す
'
' 引数
' s1    ：抽出元文字列
' s2    ：()、[]、<>など括弧の始と閉じるをセットにした文字列
Function ExtractContents(s1 As String, s2 As String) As String
    If s1 <> "" And Len(s2) = 2 Then
        ExtractContents = Mid(s1, InStr(s1, Left(s2, 1)) + 1, (InStr(s1, Right(s2, 1)) - InStr(s1, Left(s2, 1)) - 1))
    End If
End Function

' パスの正規化
' フルパスを表す文字列に半角スペースが含まれているときはダブルクォーテーションで囲んだ文字列を返す
'
' 引数
' s ：フルパスを表す文字列
Function NormalizePath(s As String) As String
    NormalizePath = s
    If NormalizePath <> "" Then
        If InStr(NormalizePath, " ") <> 0 Then
        ' 半角スペースが含まれていた
            If Not (Left(NormalizePath, 1) = """" And Right(NormalizePath, 1) = """") Then
            ' ダブルクォーテーションで囲まれていない
                NormalizePath = """" & NormalizePath & """"
            End If
        End If
    End If
End Function

' テーブルのデータ部をクリアする
Function ClearTable(tn As String) As Range
    With FindTable(tn)
        Set ClearTable = .Range.Cells(1, 1) ' テーブルの左上（見出しの左端）
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
End Function

' 拡張子を変えてファイルをコピーする
'
' 引数
' fullPath  ：対象ファイル（フルパス）  例）C:\Test\Sample.csv
' newExt    ：新しい拡張子              例）txt
'                                           C:\Test\Sample.csv → C:\Test\Sample.txt
Function ChangeFileExt(fullPath As String, newExt As String)
    Dim newFullPath As String: newFullPath = Left(fullPath, Len(fullPath) - Len(mt.GetExtensionName(fullPath))) & newExt
    FileCopy fullPath, newFullPath
    ChangeFileExt = newFullPath
End Function

' Sharepointにあるワークブックのフルパスをローカルパスに変換する
' デフォルト設定で同期していることを前提にしている
'
' 引数
' wb    ：ワークブックオブジェクト
Function GetLocalFullPath(wb As Workbook) As String
    GetLocalFullPath = GetLocalFullPath_V3(wb)
End Function

Function GetLocalFullPath_V3(wb As Workbook) As String
    
    GetLocalFullPath_V3 = ""
    
    If Not wb.Path Like "http*" Then
        GetLocalFullPath_V3 = wb.Path & "\" & wb.Name
        Exit Function
    Else
        If pathDictionary.Exists(wb.FullName) = True Then
            GetLocalFullPath_V3 = pathDictionary.Item(wb.FullName)
        Else
            ' ユーザーフォルダ以下からDirコマンドを使いダミーファイルを検索する
            Dim otf As String: otf = "" & Environ("Temp") & "\" & "Dir.result" & ""
            Dim cmd As String: cmd = Environ("ComSpec") & " /c dir /B /S """ & Environ("USERPROFILE") & "\" & wb.Name & """" & " > " & otf
            With CreateObject("WScript.Shell")
                Call .Run(cmd, 1, True)
            End With
            
            ' Dirの結果分析
            Dim s As String
            Open otf For Input As #1
            Do Until EOF(1)
                Line Input #1, s
                
                If IsBookOpened(s) = True Then
                    GetLocalFullPath_V3 = s
                    Call pathDictionary.Add(wb.FullName, GetLocalFullPath_V3)
                    Exit Do
                End If
            Loop
            
            Close #1
            Kill otf
        End If
    End If

End Function

Function IsBookOpened(fn As Variant) As Boolean
    On Error Resume Next
    
    Open fn For Append As #1
    Close #1
    
    If Err.Number > 0 Then
        IsBookOpened = True
    Else
        IsBookOpened = False
    End If
End Function

Function GetLocalFullPath_V2(wb As Workbook) As String
    
    GetLocalFullPath_V2 = ""
    
    If Not wb.Path Like "http*" Then
        GetLocalFullPath_V2 = wb.Path & "\" & wb.Name
        Exit Function
    Else
        If pathDictionary.Exists(wb.FullName) = True Then
            GetLocalFullPath_V2 = pathDictionary.Item(wb.FullName)
        Else
            ' ダミーファイルを作る（ユニークな名前に変更して保存）
            Dim prevName As String: prevName = wb.Name
            Dim uniqueName As String: uniqueName = Format(Now, "yyyymmddhhmmss") & "." & mt.GetExtensionName(wb.Name)
            wb.SaveAs wb.Path & "\" & uniqueName
            wb.SaveAs wb.Path & "\" & prevName
            
            ' ユーザーフォルダ以下からDirコマンドを使いダミーファイルを検索する
            Dim otf As String: otf = "" & Environ("Temp") & "\" & "Dir.result" & ""
            Dim cmd As String: cmd = Environ("ComSpec") & " /c dir /B /S """ & Environ("USERPROFILE") & "\" & uniqueName & """" & " > " & otf
            With CreateObject("WScript.Shell")
                Call .Run(cmd, 1, True)
            End With
            
            ' 検索結果を読込む
            Dim result As String: result = ReadAllText(otf)
            
            ' Dirの結果分析
            With New RegExp
                .pattern = ".*" & uniqueName
                .Global = True
                Dim mc As MatchCollection: Set mc = .Execute(result)
                If mc.Count > 1 Then
                    MsgBox "パス変換失敗？"
                Else
                    GetLocalFullPath_V2 = GetParentFolderName(mc.Item(0)) & "\" & wb.Name
                    Call pathDictionary.Add(wb.FullName, GetLocalFullPath_V2)
                    Call DeleteFile(mc.Item(0))
                End If
            End With
        End If
    End If
    
End Function

Function ReadAllText(fn As String) As String
   With CreateObject("Scripting.FileSystemObject")
        With .GetFile(fn).OpenAsTextStream
            ReadAllText = .ReadAll
            .Close
        End With
    End With
    Kill fn
End Function

Function GetLocalFullPath_V1(wb As Workbook) As String
    
    GetLocalFullPath_V1 = ""
    
    If Not wb.Path Like "http*" Then
        GetLocalFullPath_V1 = wb.Path & "\" & wb.Name
        Exit Function
    Else
        
        ' ユーザーフォルダ以下からDirコマンドを使いThisWorkbook.Nameを検索する
        Dim cmd As String: cmd = "%ComSpec% /c dir /B /S """ & Environ("USERPROFILE") & "\" & wb.Name & """"
        Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
        Dim eo As Object: Set eo = wsh.Exec(cmd)
        Dim result As String: result = eo.StdOut.ReadAll
        
        ' Dirの結果分析
        With New RegExp
        
            .pattern = ".*" & wb.Name
            .Global = True
            Dim mc As MatchCollection: Set mc = .Execute(result)
            If mc.Count > 1 Then
            ' 同名のブックが複数ある場合
                
                ' 一番短いパスを探す
                Dim m As match
                Dim minPath As String: minPath = wb.FullName
                For Each m In mc
                    If minPath > m.Value Then
                        minPath = m.Value
                    End If
                Next
                
                ' 共通部分（フォルダ）を特定
                minPath = Left(minPath, InStrRev(minPath, "\"))
                
                ' 共通部分とファイル名を省くとサブフォルダ文字列となる
                Dim dic As New Dictionary, s As String, pos As Integer
                For Each m In mc
                    s = Mid(m.Value, Len(minPath) + 1)
                    pos = InStrRev(s, "\")
                    If pos > 0 Then
                        dic.Add Left(s, pos - 1), m.Value
                    Else
                    ' サブフォルダ無し
                        dic.Add wb.Name, m.Value
                    End If
                Next
                
                ' 元のフルパス（URL）にサブフォルダ文字列が含まれているかどうかで判定する
                Dim k As Variant
                For Each k In dic.Keys
                    If InStr(wb.Path, k) <> 0 Then
                        GetLocalFullPath_V1 = dic.Item(k)
                        Exit For
                    End If
                Next
                
                ' 上記一致がなければ同期先直下を採択
                If GetLocalFullPath_V1 = "" Then GetLocalFullPath_V1 = dic.Item(wb.Name)
            Else
                GetLocalFullPath_V1 = mc.Item(0)
            End If
        End With
        
        Set eo = Nothing
        Set wsh = Nothing
        
    End If
    
End Function

Function DeleteFile(s As String)
    With New FileSystemObject
        .DeleteFile s, True
    End With
End Function

Function GetAbsolutePathName(s As String)
    With New FileSystemObject
        GetAbsolutePathName = .GetAbsolutePathName(s)
    End With
End Function

Function GetParentFolderName(s As String)
    With New FileSystemObject
        GetParentFolderName = .GetParentFolderName(s)
    End With
End Function

Function GetFileName(s As String)
    With New FileSystemObject
        GetFileName = .GetFileName(s)
    End With
End Function

Function GetExtensionName(s As String) As String
    With New FileSystemObject
        GetExtensionName = .GetExtensionName(s)
    End With
End Function

Function GetBaseName(s As String) As String
    With New FileSystemObject
        GetBaseName = .GetBaseName(s)
    End With
End Function

' ブックがオープンされているか調べる
'
' 引数
' wbn   :検索するブック名
'
' 返値
' Workbookオブジェクト
'
' 20250521 追加
' 前方からの一部一致に対応
Function IsOpened(wbn As String, Optional sp As Boolean = False) As Workbook
    Set IsOpened = Nothing
    
    If sp = True Then wbn = InsertBeforeExtension(wbn, "*")
    
    Dim wb As Workbook
    For Each wb In Workbooks
        If wb.Name Like wbn Then
            Set IsOpened = wb
            Exit For
        End If
    Next
End Function

' Excelのウインドウサイズを復元
Sub PopAppSize(Optional wb As Workbook = Nothing, Optional ks As String = "")
    Dim t, l, w, h
    If Not wb Is Nothing Then
        Dim ss As String: ss = Config("Config", "WindowSize_" & ks)
        If ss <> "" Then
            Dim v As Variant: v = Split(ss, ";")
            t = v(0)
            l = v(1)
            w = v(2)
            h = v(3)
            wb.Activate
        Else
            Exit Sub
        End If
    Else
        t = appTop
        l = appLeft
        w = appWidth
        h = appHeight
    End If
    
    ActiveWindow.WindowState = xlNormal
    'Application.Top = t
    'Application.Left = l
    'Application.Width = w
    'Application.Height = h
    
    Dim appRect As RECT
    appRect.Top = Val(t)
    appRect.Left = Val(l)
    appRect.Right = Val(l) + Val(w)
    appRect.Bottom = Val(t) + Val(h)
    
    Call RestoreWindowPositionAndSize(appRect)
End Sub

' Excelのウインドウサイズを保存
Sub PushAppSize(Optional wb As Workbook = Nothing, Optional ks As String = "")
    If Not wb Is Nothing Then wb.Activate
    
    'appTop = Application.Top
    'appLeft = Application.Left
    'appWidth = Application.Width
    'appHeight = Application.Height
    
    Dim appRect As RECT: appRect = GetWindowPositionAndSize(wb.Parent.Caption)
    appTop = appRect.Top
    appLeft = appRect.Left
    appWidth = GetWidth(appRect)
    appHeight = GetHeight(appRect)
    
    If Not wb Is Nothing Then Call Config("Config", "WindowSize_" & ks, appTop & ";" & appLeft & ";" & appWidth & ";" & appHeight)
End Sub

' 画面再描画を待ってみる
'
' 引数
' st    : APIのSleepに渡す
Sub ReDraw(st As Integer)
    DoEvents
    Sleep st
End Sub

' 拡張子を変更する
' 文字列終端から見てひとつめのドットより右側を拡張子と判定している。
Function ChangeExt(fn As String, ext As String) As String
    ChangeExt = Left(fn, InStrRev(fn, ".")) & ext
End Function

' デスクトップのパスを取得する
Function GetDesktopPath() As String
    ' 20250528 変更 auld-CheckData-NEW
    ' 開発環境のデスクトップはOndriveで共有されていてややこしい！
    ' 設定値があればそこをワークフォルダとする
    GetDesktopPath = Config("Config", "WorkFolder")
    
    If GetDesktopPath = "" Then
        Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
        GetDesktopPath = wsh.SpecialFolders("Desktop")
    End If
    
    Set wsh = Nothing
End Function

' アクティブセルを動かす！
' セルをチェックしている様を視覚的に表すために作った関数
'
' 引数
' startCell :開始セル
' endCell   :終了セル
Sub ScrollCell(startCell As Range, endCell As Range)
    Call ScrollCell_v2(startCell, endCell)
End Sub
' v2
' 行・列対応
Sub ScrollCell_v2(startCell As Range, endCell As Range)
    If ds = True Then
        startCell.Parent.Select
        startCell.Select
        
        Dim dx As Integer, dy As Integer
        Dim ws As Worksheet: Set ws = startCell.Parent
        
        Do While ActiveCell.Row <> endCell.Row Or ActiveCell.Column <> endCell.Column
            If ActiveCell.Row > endCell.Row Then
                dx = ActiveCell.Row - 1
            ElseIf ActiveCell.Row < endCell.Row Then
                dx = ActiveCell.Row + 1
            Else
                dx = ActiveCell.Row
            End If
            
            If ActiveCell.Column > endCell.Column Then
                dy = ActiveCell.Column - 1
            ElseIf ActiveCell.Column < endCell.Column Then
                dy = ActiveCell.Column + 1
            Else
                dy = ActiveCell.Column
            End If
            
            ws.Cells(dx, dy).Select
            DoEvents
            Call Sleep(1)
        Loop
    End If
End Sub
' v1
' 列のスクロールのみ
Sub ScrollCell_v1(startCell As Range, endCell As Range)
    If ds = True Then
        Dim s As Integer, r As Range
        If startCell.Column <= endCell.Column Then
            s = 1
            Set r = startCell
        Else
            s = -1
            Set r = endCell
        End If
        
        Dim i As Long
        For i = startCell.Column To endCell.Column Step s
            r.Offset(0, i - 1).Select
            DoEvents
            Call Sleep(1)
        Next i
    End If
End Sub

' コレクションの中身をメニューとして数値入力で選ぶ
' 実体はInputBoxを使った選択処理
'
' 引数
' menu:     選択肢
' prompt:   InputBoxへの引数
' title:    InputBoxへの引数
'
' 返値は選択されたCollectionの要素そのもの
Function MenuBox(menu As Collection, prompt As String, title As String) As Variant
    
    ' 選択肢文字列を作る
    Dim m As Variant
    Dim n As Integer: n = 1
    Dim s As String
    prompt = prompt & "下記から選んで数字を入力してください"
    For Each m In menu
        s = Format(n, "(0) ")
        prompt = prompt & Chr(13) & s & m
        n = n + 1
    Next
    
    ' 選択処理
    Dim loopFlag As Boolean: loopFlag = True
    Do While loopFlag
        If title <> "" Then
            s = InputBox(prompt, title, 1)
        Else
            s = InputBox(prompt, "", 1)
        End If
        
        If s <> "" Then
            If IsNumeric(s) Then
                n = Val(s)
                If n >= 1 And n <= menu.Count Then
                    loopFlag = False
                    MenuBox = menu(n)
                Else
                    MsgBox "無効な数字！"
                End If
            Else
                MsgBox "数字を入力して！"
            End If
        Else
        ' キャンセルされた！
            loopFlag = False
            MenuBox = ""
        End If
    Loop

End Function

' 設定値テーブルからパスを取得する。もしスペースを含むパスだった場合はダブルクォートで囲んで返す。
'
' 引数
' tn        ：テーブル名
' keyStr    ：設定値名
Function GetPathFromConfig(tn As String, keyStr As String)
    GetPathFromConfig = NormalizePath(Config(tn, keyStr))
End Function

' 指定したテーブルの指定した文字列の右隣のセルの内容を設定値として読んだり書いたりする。
'
' 引数
' tn        ：指定するテーブル名
' keyStr    ：指定する文字列
' ofst      ：通常は右隣だが更に右へオフセット可能
' valStr    ：書込む設定値
Function Config(tn As String, keyStr As String, Optional valStr As String = "", Optional ofst As Integer = 0) As String
    Dim tnTable As ListObject: Set tnTable = FindTable(tn)
    If tnTable Is Nothing Then
        MsgBox "テーブル[" & tn & "]が存在しません！"
        Exit Function
    End If
    
    Dim ret As Range
    Set ret = tnTable.DataBodyRange.Find(What:=keyStr, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False, MatchByte:=False)
    If Not (ret Is Nothing) Then
        If valStr <> "" Then
            ret.Offset(0, 1 + ofst) = valStr
        Else
            Config = ret.Offset(0, 1 + ofst)
        End If
    Else
    ' keyStrが存在しないときはテーブルに追加する！
        With tnTable.ListRows
            .Add
            .Item(.Count).Range(1).Value = keyStr
            .Item(.Count).Range(2).Value = valStr
        End With
        Config = valStr
    End If
End Function

' 開かれている全てのブックの全シートからテーブルを探す
'
' 引数
' tn    :テーブル名
Function FindTable(tn As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If tn = lo.Name Then
                Set FindTable = lo
                Exit Function
            End If
        Next
    Next
End Function

' 指定のシートから対象文字列を検索し、見つかったセルのコレクションを返す
'
' 引数
' ws    : ワークシートオブジェクト
' s1    : 検索する文字列
' s2    : 検索範囲を列指定にするためのオプション
'       : この文字列をシート内で探し、見つかったセルを含む列を検索範囲とする
Function SearchStringFromS(ws As Worksheet, s1 As String, la As XlLookAt, Optional s2 As String = "") As Collection
    Dim cell As Range, firstAddress As String, nextLoop As Boolean
    Dim myCollection As New Collection
    
    ' 列内検索を準備
    Dim lookinRange As Range
    Dim skipFind As Boolean: skipFind = False
    If s2 <> "" Then
        Set cell = ws.Cells.Find(What:=s2, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByColumns, MatchCase:=False)
        If Not cell Is Nothing Then
            Set lookinRange = ws.Columns(cell.Column)
            skipFind = True
        End If
    End If
    
    ' シート内検索を準備
    If skipFind <> True Then
        Set lookinRange = ws.Cells
    End If
    
    Set cell = lookinRange.Find(What:=s1, LookIn:=xlValues, LookAt:=la, MatchCase:=False)
    
    If Not cell Is Nothing Then
        firstAddress = cell.Address
        myCollection.Add cell
        Do
            Set cell = lookinRange.Cells.FindNext(cell)
            
            ' ループ続行判定
            nextLoop = False
            If Not cell Is Nothing Then
                If cell.Address <> firstAddress Then
                    myCollection.Add cell
                    nextLoop = True
                End If
            End If
        Loop While nextLoop = True
    End If
    
    Set SearchStringFromS = myCollection
End Function

Private Sub Class_Initialize()
    Set pathDictionary = New Dictionary
End Sub

Function IsAllSameCharacters(targetString As String) As Integer
    
    If Len(targetString) = 0 Then
        IsAllSameCharacters = 0
        Exit Function
    End If

    Dim firstChar As String
    firstChar = Mid(targetString, 1, 1)
    
    Dim i As Integer
    For i = 1 To Len(targetString)
        If Mid(targetString, i, 1) <> firstChar Then
            IsAllSameCharacters = 0
            Exit Function
        End If
    Next i

    IsAllSameCharacters = i - 1
End Function

