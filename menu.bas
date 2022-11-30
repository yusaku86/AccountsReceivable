Attribute VB_Name = "menu"
'// �����Ƀ��j���[�Ŏg�p���郂�W���[��
Option Explicit

'// ��s���ו\��
Public Sub showAccountStatement()

    Sheets("��s����").Activate

End Sub

'// �����}�X�^�o�^��ʂ��z�[������\��
Public Sub showCustomersFromHome()

    Sheets("�����}�X�^").Activate
    
    With Sheets("�����}�X�^")
        .Unprotect
        .Shapes("btnEdit").Visible = False
        .Shapes("imgEdit").Visible = False
        .Shapes("btnAdd").Visible = False
        .Shapes("imgAdd").Visible = False
        .Shapes("btnReset").Visible = False
        .Shapes("imgReset").Visible = False
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        .Protect
    End With

End Sub

'// �����}�X�^�o�^��ʂ�\��
Public Sub showCustomers()

    Sheets("�����}�X�^").Activate

End Sub

'// ���Z�O���[�v�}�X�^�o�^��ʕ\��
Public Sub showConbinedGroups()

    Sheets("���Z�O���[�v�}�X�^").Activate
    
    With Sheets("���Z�O���[�v�}�X�^")
        .Unprotect
        .Shapes("btnEdit").Visible = False
        .Shapes("imgEdit").Visible = False
        .Shapes("btnAdd").Visible = False
        .Shapes("imgAdd").Visible = False
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        .Protect
    End With

End Sub

'// ����������O���[�v�}�X�^�o�^��ʕ\��
Public Sub showSeveralTimesGroups()

    Sheets("����������O���[�v�}�X�^").Activate
    
    With Sheets("����������O���[�v�}�X�^")
        .Unprotect
        .Shapes("btnEdit").Visible = False
        .Shapes("imgEdit").Visible = False
        .Shapes("btnAdd").Visible = False
        .Shapes("imgAdd").Visible = False
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        .Protect
    End With

End Sub

'// �z�[����ʂ�\��
Public Sub showHome()
    
    '// ��s���ׂ܂��͐ݒ肩��z�[����ʂɈړ�����ꍇ
    If ActiveSheet.Name = "��s����" Or ActiveSheet.Name = "�ݒ�" Then
        GoTo Show
    End If
    
    ActiveSheet.Unprotect
        
    '// �����}�X�^����z�[����ʂɈړ�����ꍇ
    If ActiveSheet.Name = "�����}�X�^" Then
        If WorksheetFunction.CountIf(Columns(10), True) + WorksheetFunction.CountIf(Columns(10), "NEW") > 0 Then
            If MsgBox("�ύX���j������܂�����낵���ł���?", vbQuestion + vbYesNo, "�����}�X�^�o�^") = vbNo Then
                Exit Sub
            End If
        End If
        
        '// ���������N���A
        Range(Cells(6, 3), Cells(8, 3)).ClearContents
        
        '// �e�[�u������
        On Error Resume Next
        ActiveSheet.ListObjects(1).Unlist
        On Error GoTo 0
        
        '// �\�����Ă������������N���A
        Range(Cells(11, 1), Cells(Rows.Count, 10)).Clear
                
    '// ���Z�O���[�v�}�X�^�E������O���[�v�}�X�^����z�[����ʂɈړ�����ꍇ
    Else
        If WorksheetFunction.CountIf(Columns(7), True) + WorksheetFunction.CountIf(Columns(7), "NEW") > 0 Then
            If MsgBox("�ύX���j������܂�����낵���ł���?", vbQuestion + vbYesNo, ActiveSheet.Name & "�o�^") = vbNo Then
                Exit Sub
            End If
        End If
        
        '// �e�[�u������
        On Error Resume Next
        ActiveSheet.ListObjects(1).Unlist
        On Error GoTo 0
    
        '// �\�����Ă��������N���A
        Range(Cells(11, 3), Cells(Rows.Count, 5)).Clear
        Columns(7).ClearContents
    End If
     
    Dim chkController As New checkBoxController
    
    '// �`�F�b�N�{�b�N�X�폜
    chkController.deleteChk ActiveSheet
                    
    Set chkController = Nothing
    
    ActiveSheet.Protect
 
Show:
    Sheets("�z�[��").Activate

End Sub

'// �u�ݒ�v��\��
Public Sub showSetting()

    Sheets("�ݒ�").Activate
    

End Sub

'// �t�@�C�������
Public Sub closeFile()

    If MsgBox("�I�����Ă�낵���ł���?", vbQuestion + vbYesNo, "�R�݉^�����|������t�@�C��") = vbNo Then
        Exit Sub
    End If
    
    ThisWorkbook.Close True

End Sub
