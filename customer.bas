Attribute VB_Name = "customer"
'// �ڋq�Ǘ��ɏ�����郂�W���[��
Option Explicit

'/**
 '* ����挟��
'**/
Public Sub searchCustomers()
    
    If confirmSave = False Then: Exit Sub
    
    Dim where As String

    '// �����R�[�h�̌������ɒl������ꍇ
    If Cells(6, 3).value <> "" Then
        where = " id LIKE '%" & Cells(6, 3).value & "%'"
    End If
    
    '// ����於�̌������ɒl������ꍇ
    If Cells(7, 3).value <> "" And where <> "" Then
        where = where & " OR customer_name LIKE '%" & Cells(7, 3).value & "%'"
    ElseIf Cells(7, 3).value <> "" Then
        where = where & " customer_name LIKE '%" & Cells(7, 3).value & "%'"
    End If
    
    '// �������`�̌������ɒl������ꍇ �� ���͒l�̃t���K�i�𔼊p�ɂ������̂Ō���
    If Cells(8, 3).value <> "" And where <> "" Then
        where = where & " OR customer_account LIKE '%" & StrConv(Application.GetPhonetic(Cells(8, 3).value), vbNarrow) & "%'"
    ElseIf Cells(8, 3).value <> "" Then
        where = where & " customer_account LIKE '%" & StrConv(Application.GetPhonetic(Cells(8, 3).value), vbNarrow) & "%'"
    End If

    If where <> "" Then
        where = " WHERE" & where
    End If
    
    Call index(where)
End Sub

'/**
 '* �������������Z�b�g���đS�����\��
'**/
Public Sub resetSearchWord()
    
    If confirmSave = False Then: Exit Sub

    Range(Cells(6, 3), Cells(8, 3)).value = ""
    
    Call index

End Sub

'/**
 '* �ύX��o�^���Ă��Ȃ�����悪����ꍇ�A�ۑ����邩�m�F����
'**/
Private Function confirmSave() As Boolean
    
    If WorksheetFunction.CountIf(Columns(10), True) + WorksheetFunction.CountIf(Columns(10), "NEW") = 0 Then
        confirmSave = True
        Exit Function
    End If
    
    If MsgBox("�ύX���j������܂����A��낵���ł���?", vbQuestion + vbYesNo, "�����}�X�^�o�^") = vbYes Then
        confirmSave = True
        Exit Function
    End If
    
    confirmSave = False

End Function

'/**
 '* �����ꗗ�\��
 '* @params where �����̒��o����
'**/
 Private Sub index(Optional ByVal where As String = "")
 
    Application.ScreenUpdating = False
 
    Dim con As ADODB.Connection: Set con = connectDb(ThisWorkbook.Path & "\database\customers.xlsx")
    
    '// ����惌�R�[�h�Z�b�g
    Dim customerRs As New ADODB.Recordset
    customerRs.Open "SELECT * FROM [customers$]" & where & " ORDER BY id", con, adOpenStatic, adLockOptimistic
 
    '// ���Z�O���[�v���R�[�h�Z�b�g
    Dim combinedRs As New Recordset
    combinedRs.CursorLocation = adUseClient
    combinedRs.Open "SELECT * FROM [combined_groups$]", con, adOpenStatic, adLockOptimistic
    
    '// ����������O���[�v���R�[�h�Z�b�g
    Dim severalTimesRs As New Recordset
    severalTimesRs.CursorLocation = adUseClient
    severalTimesRs.Open "SELECT * FROM [several_times_payment_groups$]", con, adOpenStatic, adLockOptimistic
    
    '// �V�[�g�̕ی�ƃZ���̃��b�N������
    Sheets("�����}�X�^").Unprotect
    Sheets("�����}�X�^").Cells.Locked = False
 
    '// �e�[�u���̐ݒ������
    On Error Resume Next
    Sheets("�����}�X�^").ListObjects(1).Unlist
    On Error GoTo 0
 
    '// �O��\�����Ă��������N���A
    If Sheets("�����}�X�^").Cells(Rows.Count, 2).End(xlUp).Row > 10 Then
        With Sheets("�����}�X�^")
            .Range(.Cells(11, 1), .Cells(Rows.Count, 2).End(xlUp).Offset(, 8)).Clear
        End With
    '// 1�����\������Ă��Ȃ��ꍇ �� �w�b�_�[��1�s���̍s���N���A(�e�[�u�������̌r�����c��ꍇ�����邽��)
    Else
        With Sheets("�����}�X�^")
            .Range(.Cells(11, 1), .Cells(11, 10)).Clear
        End With
    End If
    
    '// �폜�`�F�b�N�{�b�N�X���폜
    Dim chkController As New checkBoxController
    chkController.deleteChk Sheets("�����}�X�^")
    
    '// �����Ƀq�b�g��������悪0���������甲����
    If customerRs.RecordCount = 0 Then
        GoTo Break
    End If
    
    Dim i  As Long: i = 11
    
    '// �����̏����Z���ɓ���
    Do Until customerRs.EOF
        With Sheets("�����}�X�^")
            '// �ύX�O�̎����}�X�^��A��̃Z���ɓ��͂���
            .Cells(i, 1).value = customerRs!ID
            
            .Cells(i, 2).value = customerRs!ID
            .Cells(i, 3).value = customerRs!customer_name
            .Cells(i, 4).value = customerRs!customer_account
            .Cells(i, 5).value = customerRs!customer_site
            
            '// ���E�̗L��
            If customerRs!is_offset = True Then
                .Cells(i, 6).value = "�L"
            Else
                .Cells(i, 6).value = "��"
            End If
            
            '// ���Z�O���[�v�̓���
            If customerRs!combined_group <> "" And customerRs!combined_group <> 0 Then
                combinedRs.filter = "id = " & customerRs!combined_group
                .Cells(i, 7).value = customerRs!combined_group & ":" & combinedRs!Name
            End If
            
            '// ����������O���[�v�̓���
            If customerRs!several_times_payment_group <> "" And customerRs!several_times_payment_group <> 0 Then
                severalTimesRs.filter = "id = " & customerRs!several_times_payment_group
                .Cells(i, 8).value = customerRs!several_times_payment_group & ":" & severalTimesRs!Name
            End If

            '// �폜�`�F�b�N�{�b�N�X�ǉ�
            chkController.add .Cells(i, 2), "chk" & customerRs!ID
        
        End With
        i = i + 1
        customerRs.MoveNext
    Loop
    
    Set chkController = Nothing
    
    With Sheets("�����}�X�^")
        '// �����T�C�g��Ƀh���b�v�_�E���ݒ�
        .Range(.Cells(11, 5), .Cells(Rows.Count, 2).End(xlUp).Offset(, 3)).Validation.add _
            Type:=xlValidateList, Formula1:="����,���X,�����X"
    
        '// ���E�̗L����Ƀh���b�v�_�E���ݒ�
        .Range(.Cells(11, 6), .Cells(Rows.Count, 2).End(xlUp).Offset(, 4)).Validation.add _
            Type:=xlValidateList, Formula1:="�L,��"
    
        '// ���Z�O���[�v��Ƀh���b�v�_�E���ݒ�
        .Range(.Cells(11, 7), .Cells(Rows.Count, 2).End(xlUp).Offset(, 5)).Validation.add _
            Type:=xlValidateList, Formula1:=createDropDownList(combinedRs)
        
        '// ����������O���[�v��Ƀh���b�v�_�E���ݒ�
        .Range(.Cells(11, 8), .Cells(Rows.Count, 2).End(xlUp).Offset(, 6)).Validation.add _
            Type:=xlValidateList, Formula1:=createDropDownList(severalTimesRs)
    End With
    
Break:
    With Sheets("�����}�X�^")
        ' //�u�����v�̃V�[�g���e�[�u���ɐݒ�
        .ListObjects.add(xlSrcRange, .Range(.Cells(10, 2), .Cells(Rows.Count, 2).End(xlUp).Offset(, 6)), , xlYes, , "TableStyleLight1").Name = "customers"
        
        .Range(.Cells(11, 5), .Cells(Rows.Count, 2).End(xlUp).Offset(, 4)).HorizontalAlignment = xlCenter
        
        '// �����R�[�h���IME���[�h�𔼊p�p�����ɕύX
        With .Range(.Cells(11, 2), .Cells(Rows.Count, 2).End(xlUp)).Validation
            .Delete
            .add Type:=xlValidateInputOnly
            .IMEMode = xlIMEModeAlpha
        End With
        
        '// ����於��E�������`��IME���[�h����{����͂ɕύX(�����R�[�h�̃Z������ړ������Ƃ��ɓ��{����͂ɂȂ�悤��)
        With .Range(.Cells(11, 3), .Cells(Rows.Count, 4).End(xlUp)).Validation
            .Delete
            .add Type:=xlValidateInputOnly
            .IMEMode = xlIMEModeOn
        End With
        
        .Cells.Locked = True
        '// �������̃��b�N����
        .Range(.Cells(6, 3), Cells(8, 3)).Locked = False
        
        .Range("customers").Font.Color = vbBlue
        .Cells.Font.Name = "Meiryo UI"
        
        '// �����R�[�g���ύX���ꂽ�����m�F����A��ƃf�[�^�̒l���ύX���ꂽ�����m�F����J����\��
        .Columns(1).Hidden = True
        .Columns(10).Hidden = True
        
        '// �u���Z�b�g�v�E�u�ҏW�v�E�u�V�K�ǉ��v�{�^�����g�p�\�ɂ���
        .Shapes("btnReset").Visible = True
        .Shapes("imgReset").Visible = True
        .Shapes("btnEdit").Visible = True
        .Shapes("imgEdit").Visible = True
        .Shapes("btnAdd").Visible = True
        .Shapes("imgAdd").Visible = True
        
        '// �u�o�^�v�E�u�폜�v�{�^����.�g�p�s�ɂ���
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        .Protect
        
    End With

    customerRs.Close
    combinedRs.Close
    severalTimesRs.Close
    
    Set customerRs = Nothing
    Set combinedRs = Nothing
    Set severalTimesRs = Nothing

    Set con = Nothing
    
 End Sub
 
'/**
 '* ���Z�O���[�v�E����������O���[�v�̃h���b�v�_�E�����X�g�p�̕�����쐬
 '* �uid:�O���[�v���v�̌`�Ńh���b�v�_�E���ɕ\������
'**/
Private Function createDropDownList(ByVal rs As Recordset) As String

    rs.filter = adFilterNone
    rs.Sort = "id ASC"

    Dim dropDownList As String

    Do Until rs.EOF = True
        If dropDownList = "" Then
            dropDownList = rs!ID & ":" & rs!Name
        Else
            dropDownList = dropDownList & "," & rs!ID & ":" & rs!Name
        End If
        
        rs.MoveNext
    Loop
    
    createDropDownList = dropDownList

End Function

'/**
 '* �V�K�����ǉ��̂��߂̍s�}��
'**/
Public Sub insertRowForNewCustomer()

    Sheets("�����}�X�^").Unprotect
    
    If Sheets("�����}�X�^").Cells(11, 2).value = "" Then: Exit Sub
    
    '// ���̍s�̏����������p��
    Sheets("�����}�X�^").Rows(11).Insert copyorigin:=xlFormatFromRightOrBelow
    
    With Sheets("�����}�X�^")
        .Range("customers").Font.Color = vbBlack
        .Cells(11, 2).Select
        .Cells(11, 10).value = "NEW"
        
        '// �u�o�^�v�{�^���\��
        .Shapes("btnRegister").Visible = True
        .Shapes("imgRegister").Visible = True
    End With
    
End Sub

'/**
 '* �ҏW�̂��߂ɃZ���̃��b�N����������
'**/
Public Sub unProtectToEditCustomer()

    With Sheets("�����}�X�^")
        .Unprotect
        
        On Error Resume Next
        .Range("customers").Font.Color = vbBlack
        On Error GoTo 0
        
        '// �u�o�^�v�E�u�폜�v�{�^���\��
        .Shapes("btnRegister").Visible = True
        .Shapes("imgRegister").Visible = True
        .Shapes("btnDelete").Visible = True
        .Shapes("imgDelete").Visible = True
    End With
    
End Sub

'/**
 '* �f�[�^���ύX���ꂽ�����̂��X�V�E�V�K�ǉ�
'**/
Public Sub registerCustomers()

    Application.ScreenUpdating = False

    Dim customerRs As New Recordset
    
    customerRs.Open "SELECT * FROM [customers$] ORDER BY id", connectDb(ThisWorkbook.Path & "\database\customers.xlsx"), adOpenStatic, adLockOptimistic
    Dim i As Long
    
    '// �l���ύX���ꂽ���R�[�h�̂ݍX�V
    For i = 11 To Sheets("�����}�X�^").Cells(Rows.Count, 1).End(xlUp).Row
        If Sheets("�����}�X�^").Cells(i, 10).value <> True And Sheets("�����}�X�^").Cells(i, 10).value <> "NEW" Then
            GoTo Continue
        End If
        
        '// �o���f�[�V����
        If validate(i, customerRs) = False Then
            Exit Sub
        End If
        
        '// �V�K�ǉ�
        If Sheets("�����}�X�^").Cells(i, 10).value = "NEW" Then
            Call addCustomer(i, customerRs)
        '// �X�V
        Else
            Call updateCustomer(i, customerRs)
        End If
        
Continue:
    Next
    
    customerRs.Close
    Set customerRs = Nothing
    
    With Sheets("�����}�X�^")
        .Unprotect
        .Columns(10).ClearContents
    End With
    
    '// ������Ƃ��ĕۑ������f�[�^�𐔒l��
    Dim dbBook As Workbook: Set dbBook = Workbooks.Open(ThisWorkbook.Path & "\database\customers.xlsx")
    
    With dbBook.Sheets("customers")
        convertStr2Number .Range(.Cells(2, 1), .Cells(Rows.Count, 1).End(xlUp))
        convertStr2Number .Range(.Cells(2, 6), .Cells(Rows.Count, 1).End(xlUp).Offset(, 6))
    End With
    
    dbBook.Close True
    
    Set dbBook = Nothing
    
End Sub

'/**
 '* ���͂����l�̃o���f�[�V����
'**/
Private Function validate(ByVal targetRow As Long, ByVal customerRs As Recordset) As Boolean

    validate = False

    '// �����R�[�h�����͂���Ă��邩
    If Cells(targetRow, 2).value = "" Then
        MsgBox "�����R�[�h����͂��Ă��������B", vbExclamation, "�����}�X�^�o�^"
        Cells(targetRow, 2).Select
        Exit Function
    '// �����R�[�h��������
    ElseIf IsNumeric(Cells(targetRow, 2).value) = False Then
        MsgBox "�����R�[�h�ɂ͐�������͂��Ă��������B", vbExclamation, "�����}�X�^�o�^"
        Cells(targetRow, 2).Select
        Exit Function
    End If
    
    '// ����於�����͂���Ă��邩
    If Cells(targetRow, 3).value = "" Then
        MsgBox "����於�͕K�{���ڂł��B", vbExclamation, "�����}�X�^�o�^"
        Cells(targetRow, 3).Select
        Exit Function
    End If
    
    '// �������`�����͂���Ă��邩
    If Cells(targetRow, 4).value = "" Then
        If MsgBox("�������`�����͂���Ă��܂��񂪁A��낵���ł���?", vbQuestion + vbYesNo, "�����}�X�^�o�^") = vbNo Then
            Cells(targetRow, 4).Select
            Exit Function
        End If
    End If
    
    '// �������`�����p�J�i��
    Dim reg As New RegController
    If Cells(targetRow, 4).value <> "" And reg.pregMatch(Cells(targetRow, 4).value, "^[�-���\-\(\)\�i\�j\.a-zA-Z]+$") = False Then
        MsgBox "�������`�͔��p�J�^�J�i�A�܂��͔��p�A���t�@�x�b�g�œ��͂��Ă��������B", vbExclamation + vbYesNo, "���Z�O���[�v�}�X�^"
        Cells(targetRow, 4).Select
        Set reg = Nothing
        Exit Function
    End If
    
    '// �����T�C�g�����͂���Ă��邩
    If Cells(targetRow, 5).value = "" Then
        MsgBox "�����T�C�g�͕K�{���ڂł��B", vbExclamation, "�����}�X�^�o�^"
        Cells(targetRow, 5).Select
        Exit Function
    End If
    
    '// �����R�[�h���ύX���ꂽ�ꍇ �� �ύX��̎����R�[�h�Ńt�B���^�[�������A���Ɏg�p����Ă���ꍇ�͏����𒆒f����
    If Cells(targetRow, 1).value <> Cells(targetRow, 2).value Then
        customerRs.filter = "id = " & Cells(targetRow, 2).value
    
        If customerRs.RecordCount > 0 Then
            MsgBox "�����R�[�h " & Cells(targetRow, 2).value & " �͊��Ɏg�p����Ă��܂��B", vbExclamation, "�����}�X�^�o�^"
            Cells(targetRow, 2).Select
            Exit Function
        End If
    End If

    validate = True
    
    Set reg = Nothing

End Function

'/**
 '* �V�K������ǉ�����
'**/
Private Sub addCustomer(ByVal rowNumber As Long, ByVal customerRs As Recordset)

    customerRs.AddNew
    
    '// ��r�p�����R�[�h��A��̃Z���ɓ���
    Sheets("�����}�X�^").Cells(rowNumber, 1).value = Sheets("�����}�X�^").Cells(rowNumber, 2).value
    
    '// �e���ڂ̒l��ǉ�
    customerRs!ID = Sheets("�����}�X�^").Cells(rowNumber, 2).value
    customerRs!customer_name = Sheets("�����}�X�^").Cells(rowNumber, 3).value
    customerRs!customer_account = Sheets("�����}�X�^").Cells(rowNumber, 4).value
    customerRs!customer_site = Sheets("�����}�X�^").Cells(rowNumber, 5).value
    
    '// ���Z�O���[�v�����͂���Ă���ꍇ
    If Sheets("�����}�X�^").Cells(rowNumber, 7).value <> "" Then
        customerRs!combined_group = Split(Sheets("�����}�X�^").Cells(rowNumber, 7).value, ":")(0)
    End If
    
    '// ����������O���[�v�����͂���Ă���ꍇ
    If Sheets("�����}�X�^").Cells(rowNumber, 8).value <> "" Then
        customerRs!several_times_payment_group = Split(Sheets("�����}�X�^").Cells(rowNumber, 2).value, ":")(0)
    End If
    
    '// �폜�`�F�b�N�{�b�N�X�̒ǉ�
    Dim chkController As New checkBoxController
    chkController.add Cells(rowNumber, 2), "chk" & Cells(rowNumber, 2).value
    
    customerRs.Update
    
End Sub

'/**
 '* �f�[�^�x�[�X�̒l���X�V����
'**/
Public Sub updateCustomer(ByVal rowNumber As Long, ByVal customerRs As Recordset)
    
    customerRs.filter = "id = " & Sheets("�����}�X�^").Cells(rowNumber, 1).value
    
    customerRs!ID = Sheets("�����}�X�^").Cells(rowNumber, 2).value
    customerRs!customer_name = Sheets("�����}�X�^").Cells(rowNumber, 3).value
    customerRs!customer_account = Sheets("�����}�X�^").Cells(rowNumber, 4).value
    customerRs!customer_site = Sheets("�����}�X�^").Cells(rowNumber, 5).value
    customerRs!is_offset = Sheets("�����}�X�^").Cells(rowNumber, 6).value = "�L"
    
    
    If Sheets("�����}�X�^").Cells(rowNumber, 7).value <> "" Then
        customerRs!combined_group = Split(Sheets("�����}�X�^").Cells(rowNumber, 7).value, ":")(0)
    Else
        customerRs!combined_group = 0
    End If
    
    If Sheets("�����}�X�^").Cells(rowNumber, 8).value <> "" Then
        customerRs!several_times_payment_group = Split(Sheets("�����}�X�^").Cells(rowNumber, 8).value, ":")(0)
    Else
        customerRs!several_times_payment_group = 0
    End If
    
    customerRs.Update
    
    '// �`�F�b�N�{�b�N�X�̖��O�X�V
    Sheets("�����}�X�^").CheckBoxes("chk" & Cells(rowNumber, 1).value).Name = "chk" & Cells(rowNumber, 2).value
    
    '// �����R�[�h���ύX���ꂽ�����m�F���邽�߂̃Z���̒l��ύX
    Cells(rowNumber, 1).value = Cells(rowNumber, 2).value

End Sub

'/**
 '* �`�F�b�N�{�b�N�X�Ƀ`�F�b�N�������Ă���������폜
'**/
Public Sub deleteCustomers()

    If MsgBox("�`�F�b�N�{�b�N�X�Ƀ`�F�b�N���������������폜���܂�����낵���ł���?", vbQuestion + vbYesNo, "�����}�X�^�o�^") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '// DB�Ƃ��Ďg�p���Ă���G�N�Z���u�b�N(�G�N�Z����DB�Ƃ��Ďg�p����ƃ��R�[�h�Z�b�g��Delete���\�b�h�����s�ł��Ȃ����߁A�u�b�N���J���s���폜����)
    Dim dbBook As Workbook: Set dbBook = Workbooks.Open(ThisWorkbook.Path & "\database\customers.xlsx")
    
    ThisWorkbook.Sheets("�����}�X�^").Activate
    Dim i As Long
    Dim deleteRow As Long
    
    '// �������J�n����O�̍ŏI�s
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    '// �`�F�b�N�{�b�N�X�Ƀ`�F�b�N�������Ă�����f�[�^�폜 & �s�폜
    For i = 11 To Cells(Rows.Count, 2).End(xlUp).Row
        '// �s�폜����ƍŏI�s�̒l���ύX����A�`�F�b�N�{�b�N�X�̒l���擾�ł��Ȃ��Ȃ邽�߁A�������J�n����O�̍ŏI�s - �폜�����s����i���������烋�[�v�𔲂���
        If i > lastRow Then
            Exit For
        End If
            
        '// �V�K�����͓o�^����܂Ń`�F�b�N�{�b�N�X���Ȃ��̂łƂ΂�
        If Sheets("�����}�X�^").Cells(i, 1).value = "" Then
            GoTo Continue
        End If
        
        If Sheets("�����}�X�^").CheckBoxes("chk" & Cells(i, 1).value) = 1 Then
            
            '// DB�̃f�[�^�폜
            deleteRow = WorksheetFunction.Match(Sheets("�����}�X�^").Cells(i, 1).value, dbBook.Sheets("customers").Columns(1), 0)
            dbBook.Sheets("customers").Rows(deleteRow).Delete
            
            '// �`�F�b�N�{�b�N�X�폜 & �V�[�g�u�����}�X�^�v�̍s�폜
            Sheets("�����}�X�^").CheckBoxes("chk" & Cells(i, 1).value).Delete
            Sheets("�����}�X�^").Rows(i).Delete
            
            i = i - 1
            lastRow = lastRow - 1
        End If

Continue:
    Next
        
    dbBook.Close True
    
    Set dbBook = Nothing

End Sub
