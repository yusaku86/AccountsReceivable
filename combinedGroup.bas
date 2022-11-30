Attribute VB_Name = "combinedGroup"
'// ���Z�O���[�v�}�X�^�o�^
Option Explicit

'// ���Z�O���[�v�}�X�^�ꗗ�\��
Public Sub combinedGroupIndex()
    
    '// �ύX������ꍇ�͕ۑ����Ȃ����m�F
    If WorksheetFunction.CountIf(Columns(7), True) + WorksheetFunction.CountIf(Columns(7), "NEW") > 0 Then
        If MsgBox("�ύX���j������܂����A��낵���ł���?", vbQuestion + vbYesNo, "���Z�O���[�v�}�X�^�o�^") = vbNo Then
            Exit Sub
        End If
    End If
        
    Application.ScreenUpdating = False
    
    With Sheets("���Z�O���[�v�}�X�^")
        .Unprotect
        .Range(.Cells(10, 3), .Cells(Rows.Count, 5)).Clear
        .Columns(7).ClearContents
        
        '// �w�b�_�[�̐ݒ�
        .Cells(10, 3).value = "id"
        .Cells(10, 4).value = "���Z�O���[�v��"
        .Cells(10, 5).value = "�������`"
        
        With .Range(.Cells(10, 3), .Cells(10, 5))
            .Interior.ColorIndex = 47
            .Font.Color = vbWhite
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
    End With
    
    '// �폜�`�F�b�N�{�b�N�X���폜
    Dim chkController As New checkBoxController
    chkController.deleteChk Sheets("���Z�O���[�v�}�X�^")
    
    Dim con As ADODB.Connection: Set con = connectDb(ThisWorkbook.Path & "\database\customers.xlsx")

    '// ���Z�O���[�v�}�X�^
    Dim combinedRs As New ADODB.Recordset
    combinedRs.Open "SELECT * FROM [combined_groups$] ORDER BY id", con, adOpenStatic, adLockOptimistic
    
    '// �ꗗ�\��
    Dim i As Long: i = 11
    
    Do Until combinedRs.EOF
        Cells(i, 3).value = combinedRs!ID
        Cells(i, 4).value = combinedRs!Name
        Cells(i, 5).value = combinedRs!account
        
        '// �폜�`�F�b�N�{�b�N�X�ǉ�
        chkController.add Cells(i, 4), "chk" & combinedRs!ID
        
        combinedRs.MoveNext
        i = i + 1
    Loop
    
    combinedRs.Close
    Set combinedRs = Nothing

    
    With Sheets("���Z�O���[�v�}�X�^")
    
        '// �\���e�[�u���ɐݒ�
        .ListObjects.add(xlSrcRange, .Range(.Cells(10, 3), .Cells(Rows.Count, 3).End(xlUp).Offset(, 2)), , xlYes, , "TableStyleLight1").Name = "combined_groups"
    
        '// �t�H���g�ݒ�E�Z���̃��b�N�Ȃ�
        .Cells.Font.Name = "Meiryo UI"
        .Range(.Cells(11, 4), .Cells(Rows.Count, 5).End(xlUp)).Font.Color = vbBlue
        .Range(.Columns(4), .Columns(5)).HorizontalAlignment = xlCenter
        
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
Public Sub unprotectToEditCombined()

    With Sheets("���Z�O���[�v�}�X�^")
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
Public Sub registerCombinedGroups()

    Dim combinedRs As New Recordset
    
    combinedRs.CursorLocation = adUseClient
    combinedRs.Open "SELECT * FROM [combined_groups$] ORDER BY id", connectDb(ThisWorkbook.Path & "\database\customers.xlsx"), adOpenStatic, adLockOptimistic
    Dim i As Long
    
    '// �l���ύX���ꂽ���R�[�h�̂ݍX�V
    For i = 11 To Sheets("���Z�O���[�v�}�X�^").Cells(Rows.Count, 3).End(xlUp).Row
        If Sheets("���Z�O���[�v�}�X�^").Cells(i, 7).value <> True And Sheets("���Z�O���[�v�}�X�^").Cells(i, 7).value <> "NEW" Then
            GoTo Continue
        End If
        
        '// �o���f�[�V����
        If validate(i) = False Then
            Exit Sub
        End If
        
        '// �V�K�ǉ�
        If Sheets("���Z�O���[�v�}�X�^").Cells(i, 7).value = "NEW" Then
            Call addGroup(i, combinedRs)
        '// �X�V
        Else
            Call updateGroup(i, combinedRs)
        End If
        
Continue:
    Next
    
    With Sheets("���Z�O���[�v�}�X�^")
        .Unprotect
        .Columns(7).ClearContents
    End With

End Sub

'/**
 '* �V�K�O���[�v�ǉ��̂��߂̍s�}��
'**/
Public Sub insertRowForNewCombinedGroup()

    Sheets("���Z�O���[�v�}�X�^").Unprotect
    
    If Sheets("���Z�O���[�v�}�X�^").Cells(11, 4).value = "" Then: Exit Sub

    '// ���̍s�̏����������p��
    Sheets("���Z�O���[�v�}�X�^").Rows(11).Insert copyorigin:=xlFormatFromRightOrBelow
    
    With Sheets("���Z�O���[�v�}�X�^")
        .Range("combined_groups").Font.Color = vbBlack
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
    
    '// ���Z�O���[�v�������͂���Ă��邩
    If Cells(targetRow, 4).value = "" Then
        MsgBox "����於�͕K�{���ڂł��B", vbExclamation, "���Z�O���[�v�}�X�^�o�^"
        Cells(targetRow, 4).Select
        Exit Function
    End If

    '// �������`�����͂���Ă��邩
    If Cells(targetRow, 5).value = "" Then
        MsgBox "�������`�͕K�{���ڂł��B", vbExclamation, "���Z�O���[�v�}�X�^�o�^"
        Cells(targetRow, 5).Select
        Exit Function
    End If
    
    '// �������`�����p�J�i��
    Dim reg As New RegController
    If reg.pregMatch(Cells(targetRow, 5).value, "^[�-���\-\(\)\�i\�j\.a-zA-Z]+$") = False Then
        MsgBox "�������`�͔��p�J�i�A�܂��͔��p�A���t�@�x�b�g�œ��͂��Ă��������B", vbExclamation, "���Z�O���[�v�}�X�^�o�^"
        Cells(targetRow, 5).Select
        Set reg = Nothing
        Exit Function
    End If

    validate = True

    Set reg = Nothing

End Function

'/**
 '* �V�K���Z�O���[�v��ǉ�����
'**/
Private Sub addGroup(ByVal targetRow As Long, ByVal combinedRs As Recordset)

    combinedRs.Sort = "id DESC"

    '// �V�K���Z�O���[�v��id
    Dim nextId As Long: nextId = combinedRs!ID + 1
    Sheets("���Z�O���[�v�}�X�^").Cells(targetRow, 3).value = nextId
    
    combinedRs.AddNew
    
    '// �e���ڂ̒l��ǉ�
    combinedRs!ID = nextId
    combinedRs!Name = Sheets("���Z�O���[�v�}�X�^").Cells(targetRow, 4).value
    combinedRs!account = Sheets("���Z�O���[�v�}�X�^").Cells(targetRow, 5).value
        
    '// �폜�`�F�b�N�{�b�N�X�̒ǉ�
    Dim chkController As New checkBoxController
    chkController.add Cells(targetRow, 4), "chk" & Cells(targetRow, 3).value
    
    combinedRs.Update
    
End Sub

'/**
 '* �f�[�^�x�[�X�̒l���X�V����
'**/
Private Sub updateGroup(ByVal targetRow As Long, ByVal combinedRs As Recordset)
    
    combinedRs.filter = "id = " & Sheets("���Z�O���[�v�}�X�^").Cells(targetRow, 3).value
    
    combinedRs!Name = Sheets("���Z�O���[�v�}�X�^").Cells(targetRow, 4).value
    combinedRs!account = Sheets("���Z�O���[�v�}�X�^").Cells(targetRow, 5).value
    
    combinedRs.Update
    
End Sub

'/**
 '* �`�F�b�N�{�b�N�X�Ƀ`�F�b�N�������Ă���O���[�v���폜
'**/
Public Sub deleteCombinedGroups()

    If MsgBox("�`�F�b�N�����O���[�v���폜���܂�����낵���ł���?", vbQuestion + vbYesNo, "���Z�O���[�v�}�X�^�o�^") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '// DB�Ƃ��Ďg�p���Ă���G�N�Z���u�b�N(�G�N�Z����DB�Ƃ��Ďg�p����ƃ��R�[�h�Z�b�g��Delete���\�b�h�����s�ł��Ȃ����߁A�u�b�N���J���s���폜����)
    Dim dbBook As Workbook: Set dbBook = Workbooks.Open(ThisWorkbook.Path & "\database\customers.xlsx")
    
    ThisWorkbook.Sheets("���Z�O���[�v�}�X�^").Activate
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
        If Sheets("���Z�O���[�v�}�X�^").Cells(i, 3).value = "" Then
            GoTo Continue
        End If
           
        If Sheets("���Z�O���[�v�}�X�^").CheckBoxes("chk" & Cells(i, 3).value) = 1 Then
            
            '// DB�̃f�[�^�폜
            deleteRow = WorksheetFunction.Match(Sheets("���Z�O���[�v�}�X�^").Cells(i, 3).value, dbBook.Sheets("combined_groups").Columns(1), 0)
            dbBook.Sheets("combined_groups").Rows(deleteRow).Delete
            
            '// �`�F�b�N�{�b�N�X�폜 & �V�[�g�u�����}�X�^�v�̍s�폜
            Sheets("���Z�O���[�v�}�X�^").CheckBoxes("chk" & Cells(i, 3).value).Delete
            Sheets("���Z�O���[�v�}�X�^").Rows(i).Delete
            
            i = i - 1
            lastRow = lastRow - 1
        End If
        
Continue:
    Next
        
    dbBook.Close True
    
    Set dbBook = Nothing

End Sub
