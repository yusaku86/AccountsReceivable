Attribute VB_Name = "severalTimesGroup"
'// ����������O���[�v�}�X�^�o�^
Option Explicit

'// ����������O���[�v�}�X�^�ꗗ�\��
Public Sub severalTimesGroupIndex()
    
    '// �ύX������ꍇ�͕ۑ����Ȃ����m�F
    If WorksheetFunction.CountIf(Columns(7), True) + WorksheetFunction.CountIf(Columns(7), "NEW") > 0 Then
        If MsgBox("�ύX���j������܂����A��낵���ł���?", vbQuestion + vbYesNo, "���Z�O���[�v�}�X�^�o�^") = vbNo Then
            Exit Sub
        End If
    End If
        
    Application.ScreenUpdating = False
    
    With Sheets("����������O���[�v�}�X�^")
        .Unprotect
        .Range(.Cells(10, 3), .Cells(Rows.Count, 5)).Clear
        .Columns(7).ClearContents
        
        '// �w�b�_�[�̐ݒ�
        .Cells(10, 3).value = "id"
        .Cells(10, 4).value = "����������O���[�v��"
        .Cells(10, 5).value = "�������`"
        .Rows(10).RowHeight = 50
        
        With .Range(.Cells(10, 3), .Cells(10, 5))
            .Interior.ColorIndex = 14
            .Font.Color = vbWhite
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
    End With
    
    '// �폜�`�F�b�N�{�b�N�X���폜
    Dim chkController As New checkBoxController
    chkController.deleteChk Sheets("����������O���[�v�}�X�^")
    
    Dim con As ADODB.Connection: Set con = connectDb(ThisWorkbook.Path & "\database\customers.xlsx")

    '// ���Z�O���[�v�}�X�^
    Dim severalTimesRs As New ADODB.Recordset
    severalTimesRs.Open "SELECT * FROM [several_times_payment_groups$] ORDER BY id", con, adOpenStatic, adLockOptimistic
    
    '// �ꗗ�\��
    Dim i As Long: i = 11
    
    Do Until severalTimesRs.EOF
        Cells(i, 3).value = severalTimesRs!ID
        Cells(i, 4).value = severalTimesRs!Name
        Cells(i, 5).value = severalTimesRs!account
        
        '// �폜�`�F�b�N�{�b�N�X�ǉ�
        chkController.add Cells(i, 4), "chk" & severalTimesRs!ID
        
        severalTimesRs.MoveNext
        i = i + 1
    Loop
    
    severalTimesRs.Close
    Set severalTimesRs = Nothing

    
    With Sheets("����������O���[�v�}�X�^")
    
        '// �\���e�[�u���ɐݒ�
        .ListObjects.add(xlSrcRange, .Range(.Cells(10, 3), .Cells(Rows.Count, 3).End(xlUp).Offset(, 2)), , xlYes, , "TableStyleLight2").Name = "several_times_groups"
    
        '// �t�H���g�ݒ�E�Z���̃��b�N�Ȃ�
        .Cells.Font.Name = "Meiryo UI"
        .Range(.Cells(11, 4), .Cells(Rows.Count, 5).End(xlUp)).Font.Color = vbBlue
        .Range(.Columns(4), .Columns(5)).HorizontalAlignment = xlCenter
        .Range(.Columns(4), .Columns(5)).EntireColumn.AutoFit
        
        .Columns(3).Hidden = True
        .Columns(7).Hidden = True
        
        .Shapes("btnEdit").Visible = True
        .Shapes("imgEdit").Visible = True
        .Shapes("btnAdd").Visible = True
        .Shapes("imgAdd").Visible = True
        
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        
        .Cells.Locked = True
        .Protect
    End With
    
    Set chkController = Nothing
    
End Sub

'// �ҏW�̂��߂ɃV�[�g�̕ی����
Public Sub unprotectToEditSeveralTimes()

    With Sheets("����������O���[�v�}�X�^")
        .Unprotect
        .Shapes("btnRegister").Visible = True
        .Shapes("imgRegister").Visible = True
        .Shapes("btnDelete").Visible = True
        .Shapes("imgDelete").Visible = True
        .Range(.Cells(11, 4), .Cells(Rows.Count, 5).End(xlUp)).Font.Color = vbBlack
    End With
    
End Sub

'/**
 '* �f�[�^���ύX���ꂽ�O���[�v���X�V�E�V�K�ǉ�
'**/
Public Sub registerSeveralTimesGroups()

    Dim severalTimesRs As New Recordset
    
    severalTimesRs.CursorLocation = adUseClient
    severalTimesRs.Open "SELECT * FROM [several_times_payment_groups$] ORDER BY id", connectDb(ThisWorkbook.Path & "\database\customers.xlsx"), adOpenStatic, adLockOptimistic
    Dim i As Long
    
    '// �l���ύX���ꂽ���R�[�h�̂ݍX�V
    For i = 11 To Sheets("����������O���[�v�}�X�^").Cells(Rows.Count, 3).End(xlUp).Row
        If Sheets("����������O���[�v�}�X�^").Cells(i, 7).value <> True And Sheets("����������O���[�v�}�X�^").Cells(i, 7).value <> "NEW" Then
            GoTo Continue
        End If
        
        '// �o���f�[�V����
        If validate(i) = False Then
            Exit Sub
        End If
        
        '// �V�K�ǉ�
        If Sheets("����������O���[�v�}�X�^").Cells(i, 7).value = "NEW" Then
            Call addGroup(i, severalTimesRs)
        '// �X�V
        Else
            Call updateGroup(i, severalTimesRs)
        End If
        
Continue:
    Next
    
    With Sheets("����������O���[�v�}�X�^")
        .Unprotect
        .Columns(7).ClearContents
    End With

End Sub

'/**
 '* �V�K�O���[�v�ǉ��̂��߂̍s�}��
'**/
Public Sub insertRowForNewSeveralTimesGroup()

    Sheets("����������O���[�v�}�X�^").Unprotect
    
    If Sheets("����������O���[�v�}�X�^").Cells(11, 4).value = "" Then: Exit Sub

    '// ���̍s�̏����������p��
    Sheets("����������O���[�v�}�X�^").Rows(11).Insert copyorigin:=xlFormatFromRightOrBelow
    
    With Sheets("����������O���[�v�}�X�^")
        .Range("several_times_groups").Font.Color = vbBlack
        .Cells(11, 4).Select
        .Cells(11, 7).value = "NEW"
        
        '// �u�o�^�v�{�^���\��
        .Shapes("btnRegister").Visible = True
        .Shapes("imgRegister").Visible = True
    End With
    
End Sub

'/**
 '* ���͂����l�̃o���f�[�V����
'**/
Private Function validate(ByVal targetRow As Long) As Boolean

    validate = False
    
    '// ����������O���[�v�������͂���Ă��邩
    If Cells(targetRow, 4).value = "" Then
        MsgBox "�O���[�v���͕K�{���ڂł��B", vbExclamation, "����������O���[�v�}�X�^�o�^"
        Cells(targetRow, 4).Select
        Exit Function
    End If

    '// �������`�����͂���Ă��邩
    If Cells(targetRow, 5).value = "" Then
        MsgBox "�������`�͕K�{���ڂł��B", vbExclamation, "����������O���[�v�}�X�^�o�^"
        Cells(targetRow, 5).Select
        Exit Function
    End If
    
    '// �������`�����p�J�i��
    Dim reg As New RegController
    If reg.pregMatch(Cells(targetRow, 5).value, "^[�-���\-\(\)\�i\�j\.a-zA-Z]+$") = False Then
        MsgBox "�������`�͔��p�J�i�A�܂��͔��p�A���t�@�x�b�g�œ��͂��Ă��������B", vbExclamation, "������O���[�v�}�X�^�o�^"
        Cells(targetRow, 5).Select
        Set reg = Nothing
        Exit Function
    End If

    validate = True

    Set reg = Nothing

End Function

'/**
 '* �V�K����������O���[�v��ǉ�����
'**/
Private Sub addGroup(ByVal targetRow As Long, ByVal severalTimesRs As Recordset)

    severalTimesRs.Sort = "id DESC"

    '// �V�K���Z�O���[�v��id
    Dim nextId As Long: nextId = severalTimesRs!ID + 1
    Sheets("����������O���[�v�}�X�^").Cells(targetRow, 3).value = nextId
    
    severalTimesRs.AddNew
    
    '// �e���ڂ̒l��ǉ�
    severalTimesRs!ID = nextId
    severalTimesRs!Name = Sheets("����������O���[�v�}�X�^").Cells(targetRow, 4).value
    severalTimesRs!account = Sheets("����������O���[�v�}�X�^").Cells(targetRow, 5).value
        
    '// �폜�`�F�b�N�{�b�N�X�̒ǉ�
    Dim chkController As New checkBoxController
    chkController.add Cells(targetRow, 4), "chk" & Cells(targetRow, 3).value
    
    severalTimesRs.Update
    
End Sub

'/**
 '* �f�[�^�x�[�X�̒l���X�V����
'**/
Private Sub updateGroup(ByVal targetRow As Long, ByVal severalTimesRs As Recordset)
    
    severalTimesRs.filter = "id = " & Sheets("����������O���[�v�}�X�^").Cells(targetRow, 3).value
    
    severalTimesRs!Name = Sheets("����������O���[�v�}�X�^").Cells(targetRow, 4).value
    severalTimesRs!account = Sheets("����������O���[�v�}�X�^").Cells(targetRow, 5).value
    
    severalTimesRs.Update
    
End Sub

'/**
 '* �`�F�b�N�{�b�N�X�Ƀ`�F�b�N�������Ă���O���[�v���폜
'**/
Public Sub deleteSeveralTimesGroups()

    If MsgBox("�`�F�b�N�����O���[�v���폜���܂�����낵���ł���?", vbQuestion + vbYesNo, "����������O���[�v�}�X�^�o�^") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '// DB�Ƃ��Ďg�p���Ă���G�N�Z���u�b�N(�G�N�Z����DB�Ƃ��Ďg�p����ƃ��R�[�h�Z�b�g��Delete���\�b�h�����s�ł��Ȃ����߁A�u�b�N���J���s���폜����)
    Dim dbBook As Workbook: Set dbBook = Workbooks.Open(ThisWorkbook.Path & "\database\customers.xlsx")
    
    ThisWorkbook.Sheets("����������O���[�v�}�X�^").Activate
    Dim i As Long
    Dim deleteRow As Long
    
    '// �������J�n����O�̍ŏI�s
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 3).End(xlUp).Row
    
    '// �`�F�b�N�{�b�N�X�Ƀ`�F�b�N�������Ă�����f�[�^�폜 & �s�폜
    For i = 11 To Cells(Rows.Count, 3).End(xlUp).Row
        '// �s�폜����ƍŏI�s�̒l���ύX����A�`�F�b�N�{�b�N�X�̒l���擾�ł��Ȃ��Ȃ邽�߁A�������J�n����O�̍ŏI�s - �폜�����s����i���������烋�[�v�𔲂���
        If i > lastRow Then
            Exit For
        End If
    
        '// �V�K�O���[�v�͓o�^����܂Ń`�F�b�N�{�b�N�X���Ȃ��̂łƂ΂�
        If Sheets("����������O���[�v�}�X�^").Cells(i, 3).value = "" Then
            GoTo Continue
        End If
        
        If Sheets("����������O���[�v�}�X�^").CheckBoxes("chk" & Cells(i, 3).value) = 1 Then
        
            '// DB�̃f�[�^�폜
            deleteRow = WorksheetFunction.Match(Sheets("����������O���[�v�}�X�^").Cells(i, 3).value, dbBook.Sheets("several_times_payment_groups").Columns(1), 0)
            dbBook.Sheets("several_times_payment_groups").Rows(deleteRow).Delete
            
            '// �`�F�b�N�{�b�N�X�폜 & �V�[�g�u�����}�X�^�v�̍s�폜
            Sheets("����������O���[�v�}�X�^").CheckBoxes("chk" & Cells(i, 3).value).Delete
            Sheets("����������O���[�v�}�X�^").Rows(i).Delete
            
            i = i - 1
            lastRow = lastRow - 1
        End If
        
Continue:
    Next
        
    dbBook.Close True
    
    Set dbBook = Nothing

End Sub
