Attribute VB_Name = "Module1"
Option Explicit

' 20250521 �@�\�ǉ� auld-CheckData-NEW
' ����̃Z�����󔒂̏ꍇ�̓A���[�g�Ƃ���
' 20250519
' �p�b�L���O���X�g���ʏ����ǉ�
' �����ΏۃV�[�g����}�e���A���R�[�h���擾����
Public MatCode As String
Public MatAddress As String
' 20250515
' ��a���s�ւ̃W�����v���@�\���Ă��Ȃ������̂��C��


Private wbCoC As Workbook
Private wbPL As Workbook
'Private wbOCR As Workbook
Public mt As MyTool
Private NewNamePL As String ' �ǉ� 20250515 by maruyama

Sub Init()
    If mt Is Nothing Then
        Set mt = New MyTool
    End If
    
    MatAddress = "A17"  ' 20250520 �ǉ�
End Sub

' �A�b�v���[�h�t�@�C���e�[�u���Ƀt�@�C���p�X��������
'
' ����
' sType �F�t�@�C���̎�ށiCoC/PL/OCR�j
' sPath �F�t�@�C�����i�t���p�X�j
Sub AddUploadList(sType As String, sPath As String)
    Init
    Dim dstTable As ListObject: Set dstTable = mt.FindTable("UploadList")
    If Not dstTable Is Nothing Then
        
        '�p�X�����݂��Ă���Ή������Ȃ��悤�ɁE�E�E
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

' �h���b�v�_�E�����X�g��\������
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
            ' �ύX 20250515 by maruyama
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

' �`�F�b�N�Ώۃ��[�N�u�b�N���J��
'
' ����
' configKey     �F�t�H���_�ۑ��p�L�[
' fdTitle       �FApplication.FileDialog�ɓn��
' fdFilters     �FApplication.FileDialog�ɓn��
' fdFilterExt   �FApplication.FileDialog�ɓn��
'
' �Ԓl
' vbOK          �F�u�b�N���J����
' vbCansel      �F�t�@�C���_�C�A���O�ŃL�����Z��
Function OpenBook(configKey As String, fdTitle As String, fdFilterDesc As String, fdFilterExt As String) As VbMsgBoxResult
    
    OpenBook = vbCancel
    
    ' �t�@�C���_�C�A���O���쐬
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' �_�C�A���O�̃^�C�g���ƃt�B���^�[��ݒ�
    Dim selectedFile As String
    With fd
        
        ' �����t�H���_��ݒ肷��
        Dim InitPath As String: InitPath = mt.Config("Config", configKey)
        If InitPath <> "" Then .InitialFileName = mt.GetParentFolderName(InitPath) & "\"
        
        .title = fdTitle
        .Filters.Clear
        .Filters.Add fdFilterDesc, fdFilterExt
        If .Show = True Then
        ' ���[�U�[���t�@�C����I�������ꍇ
            Dim wb As Workbook
            selectedFile = .SelectedItems(1) ' �I�����ꂽ�t�@�C���̃p�X���擾
            
            ' �����u�b�N���J����Ă��������
            Set wb = mt.IsOpened(mt.GetFileName(selectedFile))
            If Not wb Is Nothing Then wb.Close
            
            Set wb = Workbooks.Open(selectedFile, ReadOnly:=True)  ' �I�����ꂽ�t�@�C�����J��
            OpenBook = vbOK
        End If
    End With

End Function

' �`�F�b�N�ΏۃV�[�g��I������
'
' ����
' tt    :target type <"CoC", "PL", "OCR">
'
' �Ԓl�͑Ώۂ̃��[�N�V�[�g�I�u�W�F�N�g
Function SelectTarget(tt As String) As Worksheet
    Set SelectTarget = Nothing
    
    Dim fdT As String, fdFD As String, fdFE As String
    Select Case tt
        Case "CoC"
            fdT = "CoC�t�@�C����I�����Ă�������"
            fdFD = "CoC�t�@�C��"
            fdFE = "*.csv"
        Case "PL"
            fdT = "Packing List�t�@�C����I�����Ă�������"
            fdFD = "Packing List�t�@�C��"
            'fdFE = "*.xlsx;*.xlsm"
            fdFE = "*.xls*"
        Case "OCR"
            fdT = "OCR���X�g�t�@�C����I�����Ă�������"
            fdFD = "OCR���X�g�t�@�C��"
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
        
        If vbYes = MsgBox("���̃u�b�N [" & wbTemp.Name & "] ���Ώۂł����H", vbYesNo Or vbQuestion, "�A�b�v���[�h������") Then
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
                
        Dim wbn As String: wbn = mt.MenuBox(menuB, "�ǂ̃u�b�N���Ώۂł����H" & vbLf & "�I�����Ɋ܂܂�Ă��Ȃ���΃L�����Z�����I" & vbCrLf & vbCrLf, "�A�b�v���[�h������")
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
    
    ' �����܂łŃ��[�N�u�b�N�����肳���΁A���[�N�V�[�g�̓���ɐi��
    If Not wbTarget Is Nothing Then
        
        Call mt.Config("Config", tt & "_Path", wbTarget.Path & "\" & wbTarget.Name)  ' �Ώۃ��[�N�u�b�N���i�t���p�X�j��ۑ�����
        Call mt.PopAppSize(wbTarget, tt)
        
        If wbTarget.Worksheets.Count = 1 Then
        ' �V�[�g���ЂƂ����Ȃ��Ƃ�
            If vbYes = MsgBox("���̃V�[�g [" & wbTarget.Worksheets(1).Name & "] ���Ώۂł����H", vbYesNo Or vbQuestion, "�A�b�v���[�h������") Then
                Set SelectTarget = wbTarget.Worksheets(1)
            End If
        Else
        ' �V�[�g����������Ƃ��͑I��������I�΂���
            Dim menuS As New Collection
            Dim wsTemp As Worksheet
            For Each wsTemp In wbTarget.Worksheets
                menuS.Add wsTemp.Name
            Next
                    
            Dim wsn As String: wsn = mt.MenuBox(menuS, "�ǂ̃V�[�g���Ώۂł����H", "�A�b�v���[�h������")
            If wsn <> "" Then
                Set SelectTarget = wbTarget.Worksheets(wsn)
            End If
        End If
        
        ' �p�b�L���O���X�g�̓��ʏ���
        If tt = "PL" Then
            ' auld-CheckData-NEW
            Call DeleteCellsPL(SelectTarget)
        End If
    End If
End Function

' CoC��a���`�F�b�N
Sub CheckCoC()
    
    Call Init
    
    Dim wsTarget As Worksheet: Set wsTarget = SelectTarget("CoC")
    If Not wsTarget Is Nothing Then
        
        ' �������ʏo�͐�e�[�u���̊�Z��
        Dim tl As Range    ' TopLeft
        With mt.FindTable("CheckCoC")
            Set tl = .Range.Cells(1, 1) ' �^�C�g���s
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete ' �������ʏo�̓e�[�u�����N���A
            End If
        End With
        
        ' ��ʕ`��p�̃t���O���Z�b�g
        mt.DrawScroll = mt.Config("Config", "DrawScroll")
        Dim ds As Boolean: ds = mt.Config("Config", "DrawSelect")
    
        Dim rCell As Range, lCell As Range, l As Long, s As String, d As Integer
        Dim sc As Range: Set sc = wsTarget.Range("A1").CurrentRegion.Columns(1)
        Dim flagNotice As Boolean, productName As String
        Dim cntNotice As Integer: cntNotice = 0
        For l = sc.Row To sc.Rows.Count    ' �����Ώۂ̍s���Ń��[�v����
            flagNotice = False
            
            ' ���[�Z����I���iCoC�j
            Set lCell = sc.Cells(l, 1)
            If ds Then
                wsTarget.Parent.Activate
                lCell.Select
            End If
            
            ' �s�̉E�[�܂ŃA�N�e�B�u�Z���𓮂���
            Set rCell = sc.Cells(l, Columns.Count).End(xlToLeft)
            Call mt.ScrollCell(lCell, rCell)
            
            If l > 3 Then
                ' ���̍s���\�����i���U�C���`�Ȃ̂��W�C���`�Ȃ̂��H
                ' �P��ڂ̐擪�����Ŕ��f����B
                productName = sc.Cells(l, 1)
                s = Left(productName, 1)
                If IsNumeric(s) = True Then
                     d = Val(s)
                                    
                    ' ��a���`�F�b�N
                    If mt.EndWithString(rCell, ";;;;;") = True Then
                    ' �E�[�Z���̓��e���Z�~�R����5�A���A���W�C���`�s�̂Ƃ�
                        If d = 8 Then
                            flagNotice = True
                        End If
                    Else
                    ' �E�[�Z���̓��e���Z�~�R����5�A���ł͂Ȃ��A���U�C���`�s�̂Ƃ�
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
            
            ' ��a����\������
            If flagNotice = True Then
                cntNotice = cntNotice + 1
                tl.Offset(cntNotice, 0) = l
                tl.Offset(cntNotice, 1) = s & " �C���`"
                tl.Offset(cntNotice, 2) = "��a������"
            End If
                    
            ' �s�̍��[�܂ŃA�N�e�B�u�Z���𓮂����iCoC�j
            Call mt.ScrollCell(rCell, lCell)
            
            If ds Then
                ThisWorkbook.Activate
                tl.Offset(l, 0).Select
                mt.ReDraw 1
            End If
        Next
        
        If cntNotice = 0 Then
            MsgBox prompt:="��肠��܂���ł���", title:="CoC�`�F�b�N����"
            
            ' �A�b�v���[�h�G���g���[
            Call AddUploadList("CoC", mt.GetLocalFullPath(wsTarget.Parent))
        Else
            MsgBox prompt:="��a������ł��B���̃t�@�C�����C�����Ă�����Ă��������B", title:="CoC�`�F�b�N����"
            If vbYes = MsgBox(prompt:="�u�b�N����Ă��ǂ��ł����H", Buttons:=vbYesNo, title:="CoC�`�F�b�N����") Then wsTarget.Parent.Close
        End If
    End If
End Sub

' ���[�N�V�[�g���e�L�X�g�ɕϊ����ĊJ������
' ����
' ws    �F���[�N�V�[�g�I�u�W�F�N�g
' ext   �F�g���q
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
                    newName = mt.MakeBaseNamePL(ws, ext)    ' �p�b�L���O���X�g�t�@�����ϊ�
                End If
                Dim newFile As String: newFile = mt.GetDesktopPath() & "\" & newName
                
                Dim wbTemp As Workbook: Set wbTemp = mt.IsOpened(newName)
                If Not wbTemp Is Nothing Then wbTemp.Close SaveChanges:=False
                
                ws.Copy
                
                ' CSV�쐬
                Application.DisplayAlerts = False
                ActiveWorkbook.SaveAs Filename:=newFile, FileFormat:=xlCSV, CreateBackup:=False
                Application.DisplayAlerts = True
                ActiveWorkbook.Close
                
                If Dir(newFile) <> "" Then
                    Call Workbooks.OpenText(newFile, 932, 1)
                    'Call mt.PopAppSize
                
                    Set SelectWorksheet = ActiveSheet
                Else
                    Call MsgBox(prompt:="�t�@�C�����ۑ�����Ă��܂���B��蒼���Ă��������I", Buttons:=vbOK Or vbCritical, title:="�����A�b�v���[�h")
                End If
                
                Dim killBookName As String: If spFlag = True Then killBookName = mt.GetLocalFullPath(wb)
                
                If spFlag <> True Then
                    wb.Close
                Else
                ' �p�b�L���O���X�g�����̂Ƃ��̓��[�N�u�b�N���폜����
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

' PackingList��a���`�F�b�N
Sub CheckPL()
    
    Call Init
    
    NewNamePL = ""  ' �ǉ� 20250515 by maruyama, auld-CheckData-NEW

    ' �t�@�C�����ϊ��̗L���E�����t���O�擾
    Dim f As Boolean
    If mt.Config("Config", "ConvertFileNamePL") = "0" Then f = False Else f = True
    
    Dim wsTarget As Worksheet: Set wsTarget = SelectWorksheet(SelectTarget("PL"), "pck", f)
    
    If Not wsTarget Is Nothing Then
        
        NewNamePL = wsTarget.Name     ' �ǉ� 20250515 by maruyama, auld-CheckData-NEW
        
        Call mt.PopAppSize(wsTarget.Parent, "PL")
        
        ' �������ʏo�͐�e�[�u���̊�Z��
        Dim tl As Range    ' TopLeft
        With mt.FindTable("CheckPL")
            Set tl = .Range.Cells(1, 1) ' �^�C�g���s
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete ' �������ʏo�̓e�[�u�����N���A
            End If
        End With
        
        Dim ds As Boolean: ds = mt.Config("Config", "DrawSelect")
        Dim r As Range, l As Long, cntSafeWord As Integer
        Dim sc As Range: Set sc = wsTarget.Range("A1").CurrentRegion.Columns(1)
        Dim cntNotice As Integer: cntNotice = 0
        For l = sc.Row To sc.Rows.Count    ' �����Ώۂ̍s���Ń��[�v����
            
            ' �E�[�Z����I��
            Set r = sc.Cells(l, Columns.Count).End(xlToLeft)
            If ds = True Then
                wsTarget.Parent.Activate
                r.Select
            End If
            
            ' ��a���\��
            If Right(r, 1) <> "," Then
            ' �s�����J���}�łȂ���Έ�a������I
            
                ' ��O�����i�Ώۍs�ɂ��̒P�ꂪ�܂܂�Ă���Εs��Ƃ���j
                cntSafeWord = 0
                cntSafeWord = cntSafeWord + InStr(r, "Measurement")
                cntSafeWord = cntSafeWord + InStr(r, "T O T A L")
                
                If cntSafeWord = 0 Then
                    cntNotice = cntNotice + 1
                    tl.Offset(cntNotice, 0) = l
                    tl.Offset(cntNotice, 1) = "��a������"
                    r.Select
                End If
            End If
            
            If ds = True Then
                ThisWorkbook.Activate
                tl.Offset(l, 0).Select
                mt.ReDraw 1
            End If
        Next
        
        ' �ŉ��s����k���ĘA���J���}���폜����
        For l = sc.Rows.Count To sc.Row Step -1
            If mt.IsAllSameCharacters(sc.Cells(l)) = 7 Then
                sc.Cells(l) = ""
            Else
                Exit For
            End If
        Next l
        
        If cntNotice = 0 Then
            MsgBox prompt:="��肠��܂���ł���", title:="PackingList�`�F�b�N����"
            
            ' �A�b�v���[�h�G���g���[
            Call AddUploadList("PL", mt.GetLocalFullPath(wsTarget.Parent))
        Else
            MsgBox prompt:="��a������ł��B���̃t�@�C�����C�����Ă�����Ă��������B", title:="PackingList�`�F�b�N����"
            If vbYes = MsgBox(prompt:="�u�b�N����Ă��ǂ��ł����H", Buttons:=vbYesNo, title:="PackingList�`�F�b�N����") Then wsTarget.Parent.Close
        End If
        
    End If
End Sub


' OCR���X�g��a���`�F�b�N
Sub CheckOCR()
    
    Call Init
    
    Dim wsTarget As Worksheet: Set wsTarget = SelectWorksheet(SelectTarget("OCR"), "csv")
    
    If Not wsTarget Is Nothing Then
        
        ' �������ʏo�͐�e�[�u���̊�Z��
        Dim tl As Range    ' TopLeft
        With mt.FindTable("CheckOCR")
            Set tl = .Range.Cells(1, 1) ' �^�C�g���s
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete ' �������ʏo�̓e�[�u�����N���A
            End If
        End With
        
        ' 2025/02/17 �ǉ�
        ' ����������������
        Set wsOCR = wsTarget
        
        Dim ParentCaption As String: ParentCaption = wsOCR.Parent.Parent.Caption
        Dim DeliveryDate As Date: DeliveryDate = CovertToDate(Left(ParentCaption, 8))
        Dim RowOfError As Integer: RowOfError = 1
        Dim IsError As Boolean
        
        ' �G���[�o�͐�e�[�u���̊�Z��
        Dim tl2 As Range    ' TopLeft
        With mt.FindTable("ErrorOfOCR")
            Set tl2 = .Range.Cells(1, 1) ' �^�C�g���s
            If Not .DataBodyRange Is Nothing Then
                .DataBodyRange.Delete ' �������ʏo�̓e�[�u�����N���A
            End If
        End With
        
        ' Material�`�F�b�N�p�����񐶐�
        Dim materialName As String: materialName = ExtractMaterialName(ParentCaption)
        If materialName = "" Then MsgBox "�t�@�C�������������������I"
        ' ����������������
        
        Dim ds As Boolean: ds = mt.Config("Config", "DrawSelect")
        Dim r As Range, l As Long
        Dim sc As Range: Set sc = wsTarget.Range("A1").CurrentRegion.Columns(1)
        Dim idDic As New Dictionary
        Dim idDayDic As New Dictionary
        Dim lotDic As New Dictionary
        Dim lotDayDic As New Dictionary
        Dim s As String, cnt As Integer
        For l = sc.Row To sc.Rows.Count    ' �����Ώۂ̍s���Ń��[�v����
            Set r = sc.Cells(l, 1)
            If Not r.Offset(0, 1) = "WaferID" Then
                
                ' WaferID �ŏ���
                s = Left(r.Offset(0, 1), 6)
                If idDic.Exists(s) = True Then
                    idDic.Item(s) = idDic.Item(s) + 1
                Else
                    Call idDic.Add(s, 1)
                    Call idDayDic.Add(s, r.Offset(0, 4))
                End If
                
                ' 2025/02/17 �����ǉ�
                ' ��������������������
                IsError = False
                                
                ' WaferID�`�F�b�N
                If Len(sc.Cells(l, 2)) <> 12 Then
                    IsError = True
                    tl2.Offset(RowOfError, 1) = sc.Cells(l, 2)
                End If
                
                ' Material�`�F�b�N
                'If InStr(ParentCaption, sc.Cells(l, 3)) = 0 Then
                If materialName <> sc.Cells(l, 3) Then
                    IsError = True
                    tl2.Offset(RowOfError, 2) = sc.Cells(l, 3)
                End If
                
                ' LotNo�`�F�b�N
                If Len(sc.Cells(l, 4)) <> 16 Then
                    IsError = True
                    tl2.Offset(RowOfError, 3) = sc.Cells(l, 4)
                End If
                
                ' ���t�`�F�b�N
                If DeliveryDate <> CovertToDate(sc.Cells(l, 5)) Then
                    IsError = True
                    tl2.Offset(RowOfError, 4) = sc.Cells(l, 5)
                End If
                
                If IsError = True Then
                    tl2.Offset(RowOfError, 0) = l
                    RowOfError = RowOfError + 1
                End If
                ' ��������������������
            End If
        Next
        
        ' �������͂��҂�[�[�[�[���@
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
        tl.Offset(l, 0) = "���v����"
        tl.Offset(l, 1) = sumWafer
        l = l + 1
        
        ' �A�b�v���[�h�G���g���[
        'Call AddUploadList("OCR", mt.GetLocalFullPath(wsTarget.Parent))
        ' 2025/02/17 �ύX�i�G���[���Ȃ������Ƃ������G���g���[����j
        With mt.FindTable("ErrorOfOCR")
            If .DataBodyRange Is Nothing Then
                Call AddUploadList("OCR", mt.GetLocalFullPath(wsTarget.Parent))
            End If
        End With
    End If
    
End Sub

' �A�b�v���[�h���s
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
    
    Dim patternStr As String: patternStr = mt.Config("ConfigFTP", "�p�^�[��")
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
            MsgBox "�t�@�C���������Ă��܂���I"
            Exit Sub
        End If
    
    Next i
    
    ' �R�}���h�̃I�v�V����
    Dim optionCmd As String: optionCmd = "cd " & mt.Config("ConfigFTP", "�A�b�v���[�h��")
    
    '�R�}���h�t�@�C���쐬
    Open "auld.txt" For Output As #1
    If optionCmd <> "" Then Print #1, optionCmd
    For i = 1 To fileList.Count
        Print #1, "put " & mt.NormalizePath(fileList(i))
    Next i
    Print #1, "quit"
    Close #1
    
    ' ���t�@�C�������擾
    Dim keyFileName As String: keyFileName = mt.Config("ConfigFTP", "PPK")
    If Len(Dir(keyFileName)) = 0 Then
        keyFileName = mt.GetParentFolderName(mt.GetLocalFullPath(ThisWorkbook)) & "\" & keyFileName
        If Len(Dir(keyFileName)) <> 0 Then
            keyFileName = mt.NormalizePath(keyFileName)
        Else
            MsgBox "���t�@�C����������܂���"
            Exit Sub
        End If
    End If
    
    ' �A�b�v���[�h
    Dim objWSH As Object: Set objWSH = CreateObject("WScript.Shell")
    Dim sCmd As String: sCmd = """C:\Program Files\PuTTY\psftp.exe"" -i " & keyFileName & " " & mt.Config("ConfigFTP", "User") & "@" & mt.Config("ConfigFTP", "Server") & " -bc -be -b auld.txt"
    Dim objExec As Object: Set objExec = objWSH.Exec(sCmd)  ' WshScriptExec�I�u�W�F�N�g�擾
    objExec.StdIn.Close
    
    ' ���O�\��
    Dim textStream As String: textStream = objExec.StdOut.ReadAll
    If objExec.ExitCode <> 0 Then
        textStream = textStream & vbCrLf & "�G���[�����I" & vbCrLf
        textStream = textStream & objExec.StdErr.ReadAll
    End If
    MsgBox prompt:=textStream, title:="�A�b�v���[�h�̌���"
    
    '�I�u�W�F�N�g��j��
    Set objWSH = Nothing
    Set objExec = Nothing

    Call DeleteUploadFileList
    Kill "auld.txt"
    
End Sub

' �p�b�L���O���X�g�̕s�v�ȃZ�����폜����
' �@��������"T O T A L"�Ɠ��͂��ꂽ�E�ׂ̃Z���̌v�Z���ʂ�0�i�󔒁j�̏ꍇ�A�Y���y�[�W�ȍ~�̍s���폜����
' �A��H�`K���폜����
' ����
' ws    �F�ΏۂƂȂ郏�[�N�V�[�g
Sub DeleteCellsPL(ws As Worksheet)
    'Call DeleteCellsPL_B(ws)
    Call DeleteCellsPL_B(ws)
End Sub

Sub DeleteCellsPL_A(ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    
    Init
    
    ' ���[�N�u�b�N��ʖ��ۑ�����i�f�X�N�g�b�v�ցI�j
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
    
    ' �v�Z���ʁi�E�׃Z���j��0�i�󔒁j�ƂȂ��Ă���"T O T A L"������
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
End Sub
