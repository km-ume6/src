Attribute VB_Name = "Module250414"
Option Explicit

Public HasBlanks As Boolean

Sub ModifyWindowSize()
Attribute ModifyWindowSize.VB_Description = "�E�B���h�E�T�C�Y��ۑ����܂�"
Attribute ModifyWindowSize.VB_ProcData.VB_Invoke_Func = "w\n14"
    With New MyTool
        Dim menuB As New Collection
        menuB.Add "Me"
        menuB.Add "CoC"
        menuB.Add "PL"
        menuB.Add "OCR"
                
        Dim TargetWindow As String: TargetWindow = .MenuBox(menuB, "�ǂ̃E�B���h�E���Ώۂł����H" & vbLf & "�I�����Ɋ܂܂�Ă��Ȃ���΃L�����Z�����I" & vbCrLf & vbCrLf, "�E�B���h�E�T�C�Y����")
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

' �p�b�L���O���X�g�̓���Z�����󔒂��ǂ����𒲍�����
Sub HasBlankCells(ws As Worksheet)
    Call HasBlankCells_B(ws)
End Sub

Sub HasBlankCells_B(ws As Worksheet)

    HasBlanks = False
    If ws Is Nothing Then Exit Sub
    
    ' "MADE IN JAPAN"�Ɠ��͂��ꂽ�Z���ꗗ���擾����
    Dim mijs As New Collection
    Set mijs = mt.SearchStringFromS(ws, "MADE IN JAPAN", xlPart, "No.1")
    
    ' "T O T A L"�Ɠ��͂��ꂽ�Z���ꗗ���擾����
    Dim totals As New Collection
    Set totals = mt.SearchStringFromS(ws, "T O T A L", xlWhole, "Lot No.")
    
    If mijs.Count = totals.Count Then
        
        ' �󔒃`�F�b�N
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
            Call MsgBox("�󔒊m�F�ΏۃZ���ɋ󔒂�����܂��B���̃t�@�C�����m�F���ĉ������B" & vbCrLf & AppendAfterCrLf(AlertString), vbOKOnly, "�A�b�v���[�h������")
            HasBlanks = True
        End If
    Else
        Call MsgBox("""MADE IN JAPAN"" �� ""T O T A L"" �̐��������܂���ł����̂ŋ󔒃`�F�b�N�͎��{���܂���B", vbOKOnly, "�󔒃`�F�b�N")
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
    
    ' "T O T A L"�Ɠ��͂��ꂽ�Z���i�y�[�W�̊�j�ꗗ���擾����
    Dim totals As New Collection
    Set totals = mt.SearchStringFromS(ws, "T O T A L", xlWhole, "Lot No.")
    
    ' �ݒ蕶���񂩂�󔒃`�F�b�N�Ώۂ̃Z���A�h���X���擾���A��Z������̃I�t�Z�b�g��z��ɃZ�b�g
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

    ' �󔒃`�F�b�N
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
        If vbYes <> MsgBox("�󔒊m�F�ΏۃZ���ɋ󔒂�����܂��B���̂܂ܑ����܂����H" & vbCrLf & AppendAfterCrLf(CellAddress), vbYesNo, "�A�b�v���[�h������") Then HasBlanks = True
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
    
    ' ���[�N�u�b�N��ʖ��ۑ�����i�f�X�N�g�b�v�ցI�j
    With ws.Parent
        Dim sheetName As String: sheetName = ws.Name
        Dim srcPath As String: srcPath = mt.GetLocalFullPath(ws.Parent)
        Dim dstPath As String: dstPath = ChangeFileExtension(mt.GetDesktopPath() & "\tmp" & .Name, "xlsx")
        
        ' �}�N���Ȃ��u�b�N�Ƃ��ĕۑ�����
        Application.DisplayAlerts = False: .SaveAs Filename:=dstPath, FileFormat:=xlOpenXMLWorkbook: Application.DisplayAlerts = False
        
        .Close SaveChanges:=False
        'FileCopy srcPath, dstPath
        
        Dim wbNew As Workbook: Set wbNew = Workbooks.Open(dstPath)
        Set ws = wbNew.Worksheets(sheetName)
        
        ' auld-CheckData-NEW
        ws.UsedRange.Value = ws.UsedRange.Value ' �֐����폜����
    End With
    
    Dim cellList As New Collection
    Set cellList = mt.SearchStringFromS(ws, "T O T A L", xlWhole, "Lot No.")
    
    ' �v�Z���ʁi�E�׃Z���j���󔒂ƂȂ��Ă���"T O T A L"������
    Dim rTotal As Range
    For Each rTotal In cellList
        If rTotal.Offset(0, 1) = 0 Then Exit For
    Next
    
    ' �Y�����鏑���̃y�[�W�ԍ��i�y�[�W�̍���ɓ�����j������
    If Not rTotal Is Nothing Then
        Dim rNo As Range: Set rNo = rTotal.Offset(0, -2)
        Do While rNo.Row <> 1
            If rNo Like "No.*" Then
                Exit Do
            End If
            Set rNo = rNo.Offset(-1, 0)
        Loop
        
        ' �����̍ŉ��i����
        Dim rEnd As Range: Set rEnd = ws.Cells(Rows.Count, rNo.Column).End(xlUp)
        
        ' �s���폜����
        ' ws.Range(rNo.Offset(-1, 0).Row & ":" & rEnd.Row).Delete
        If rNo.Row <> 1 Then ws.Range(rNo.Offset(-1, 0).Row & ":" & rEnd.Row).Delete
    End If
    
    ws.Range(mt.Config("Config", "DelColPL")).Delete    ' ����폜����
    
    ' �������܂łŕs�v�ȃZ�����폜���Ă���I
    
    ' auld-CheckData-NEW
    Call HasBlankCells(ws)
End Sub

Function MakeBaseNamePL4(ws As Worksheet, ext As String) As String
    Dim sdCell As String: sdCell = mt.Config("Config", "sdCell")        ' �o�ד��̃Z���A�h���X
    Dim ccCell As String: ccCell = mt.Config("Config", "ccCell")        ' �o�א於�i�����j�̃Z���A�h���X
    MakeBaseNamePL4 = Format(ws.Range(sdCell), "yyyymmdd") & " " & Left(StrConv(Replace(ws.Range(ccCell), " ", ""), vbProperCase), 3) & " Packing list " & ws.Name & "." & ext
End Function

Function ChangeFileExtension(filePath As String, newExtension As String) As String
    Dim dotPos As Long
    Dim basePath As String

    ' �g���q�̑O�̃h�b�g�̈ʒu���擾
    dotPos = InStrRev(filePath, ".")

    If dotPos > 0 Then
        ' �g���q���������������擾
        basePath = Left(filePath, dotPos - 1)
    Else
        ' �g���q���Ȃ��ꍇ�͂��̂܂�
        basePath = filePath
    End If

    ' �V�����g���q��ǉ��i�s���I�h�������ŕt����j
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

    AppendAfterCrLf = "No." & l & "�F"

    ' 1����������
    For i = 1 To Len(srcString)
        tempChar = Mid(srcString, i, 1)
        lineBuffer = lineBuffer & tempChar

        ' ���s�R�[�h�̌��o�ivbCrLf��2�����Ȃ̂ŁA���O��2�����Ŕ���j
        If Right(lineBuffer, 2) = vbCrLf Then
            l = l + 1
            AppendAfterCrLf = AppendAfterCrLf & lineBuffer & "No." & l & "�F"
            lineBuffer = ""
        End If
    Next i
End Function

