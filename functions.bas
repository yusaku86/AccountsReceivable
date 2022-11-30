Attribute VB_Name = "functions"
Option Explicit

'// db�ڑ�
Public Function connectDb(ByVal dbBook As String) As ADODB.Connection

    Dim returnCon As New ADODB.Connection
    
    '// db�Ƃ��Ďg�p����t�@�C���ɐڑ�
    With returnCon
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open dbBook
    End With

    Set connectDb = returnCon

End Function

'/**
 '* ������Ƃ��ĕۑ�����Ă���f�[�^�𐔒l��
 '* @params targetRange �f�[�^�𐔒l������͈�
'**/
Public Sub convertStr2Number(ByVal targetRange As Range)
    
    targetRange.value = Evaluate(targetRange.Address & "*1")
    
End Sub

'// �_�C�A���O��\�����A�t�@�C����I������
Public Function selectFile(ByVal dialogTitle As String, ByVal initialFile As String, ByVal targetDiscription As String, ByVal targetExtension As String) As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = initialFile
        .AllowMultiSelect = False
        .Title = dialogTitle
        
        '// �I������t�@�C���̊g���q�ݒ�
        .filters.Clear
        .filters.add targetDiscription, targetExtension
        
        If .Show Then
            selectFile = .SelectedItems(1)
        End If
    End With
    
End Function

'// �t�@�C���Ɏw��̃V�[�g�����݂��邩�m�F
Public Function sheetExist(ByVal activeFile As Workbook, ByVal sheetName As String) As Boolean

    Dim sheet As Worksheet
    
    For Each sheet In activeFile.Worksheets
        If sheet.Name = sheetName Then
            sheetExist = True
            Exit Function
        End If
    Next
    
    sheetExist = False

End Function

'/**
 '* �������w��̌������o��(�����̂ڂ���)���������������߂�
 '* �k����������N�ɂȂ�ꍇ�͕Ԃ�l�Ƀ}�C�i�X��t����
 '*
 '* ��) passMonth [1], [-3] �� -10
'**/

Public Function passMonth(ByVal standardMonth As Long, ByVal passingMonths As Long)

    Dim returnMonth As Long: returnMonth = standardMonth + passingMonths

    If returnMonth <= 0 Then
        returnMonth = -(returnMonth + 12)
    ElseIf returnMonth > 12 Then
        returnMonth = returnMonth - 12
    End If
    
    passMonth = returnMonth

End Function

'// 2�̐����̂����A�ǂ��炪�萔���Ƃ��đÓ����𔻒�
Public Function compareNumbersAsCommision(ByVal number1 As Long, ByVal number2 As Long) As Long

    '// 2�̐����������ꍇ �� �O�҂�Ԃ�
    If number1 = number2 Then
        compareNumbersAsCommision = number1
        Exit Function
    End If
    
    '// �ǂ��炩��0�̏ꍇ��0�̕����萔���Ƃ��đÓ�
    If number1 = 0 Then
        compareNumbersAsCommision = number1
        Exit Function
    ElseIf number2 = 0 Then
        compareNumbersAsCommision = number2
        Exit Function
    End If
    
    '// �ǂ��炩�P�݂̂����̐����̏ꍇ�����̐����ł͂Ȃ������Ó�
    If 0 < number1 And number2 < 0 Then
        compareNumbersAsCommision = number1
        Exit Function
    ElseIf 0 < number2 And number1 < 0 Then
        compareNumbersAsCommision = number2
        Exit Function
    End If
    
    '// ��Βl�������������Ó�
    If Asc(number1) < Asc(number2) Then
        compareNumbersAsCommision = number1
    Else
        compareNumbersAsCommision = number2
    End If
    
End Function

'// 2�̐����̂����A�����z�ɑ΂��Ăǂ��炪����Ƃ��đÓ����𔻒�
Public Function compareNumbersAsSales(ByVal payment As Long, ByVal sales1 As Long, ByVal sales2 As Long) As Long

    '// 2�̐����������ꍇ �� �O�҂�Ԃ�
    If sales1 = sales2 Then
        compareNumbersAsSales = sales1
        Exit Function
    End If
    
    '// �����z�Ƃǂ��炩1�̐����̍��z��0�̏ꍇ -> ���z��0�ɂȂ�����Ó�
    If sales1 - payment = 0 Then
        compareNumbersAsSales = sales1
        Exit Function
    ElseIf sales2 - payment = 0 Then
        compareNumbersAsSales = sales2
        Exit Function
    End If
    
    '// �ǂ��炩�̓����z�Ƃ̍��z�����̐����ɂȂ�ꍇ �� ���ɂȂ�Ȃ������Ó�
    If 0 < sales1 - payment And sales2 - payment < 0 Then
        compareNumbersAsSales = sales1
        Exit Function
    ElseIf 0 < sales2 - payment And sales1 - payment < 0 Then
        compareNumbersAsSales = sales2
        Exit Function
    End If
    
    '// ��L�ȊO �� �����z�Ƃ̍��z�̐�Βl�������������Ó�
    If Asc(sales1 - payment) < Asc(sales2 - payment) Then
        compareNumbersAsSales = sales1
    Else
        compareNumbersAsSales = sales2
    End If
    
End Function



