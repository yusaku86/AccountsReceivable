Attribute VB_Name = "accountStatement"
'// ��s���׊֘A(���|������Ȃ�)
Option Explicit

'// �������ׂ̔N��I������t�H�[���N��
Public Sub openFormYear()

    Sheets("mode").Cells(1, 1).value = "IMPORT_STATEMENT"

    formYear.Show
    
End Sub

'// �O��Z�F�̐U�����ׂ����H���ăV�[�g�u��s���ׁv�֓\��t��(���C���v���V�[�W��)
Public Sub putBankStatement(ByVal targetYear As Long)
    
    Dim wsh As New WshShell
    
    '// �_�C�A���O��\�����Ď�荞�ރt�@�C����I�� selectFile [�_�C�A���O�^�C�g��], [�����\���t�H���_], [�g���q�i�荞�݃��b�Z�[�W], [�i�荞�ފg���q]
    Dim fileName As String: fileName = selectFile("��s���׎捞", wsh.SpecialFolders(4), "CSV�t�@�C��", "*.csv")
    
    Set wsh = Nothing
    
    If fileName = "" Then: Exit Sub
    
    '// �w�b�_�[�̐ݒ�
    Call createTableHeader
    
    '// �w�肵���t�@�C�����������`�����m�F
    If checkFile(fileName) = False Then
        Exit Sub
    End If
    
    Dim fso As New FileSystemObject
    fileName = fso.GetFileName(fileName)
    
    Set fso = Nothing
    
    Dim bankCsv As Workbook: Set bankCsv = Workbooks(fileName)

    With bankCsv.Sheets(1)
        
        Dim lastRow As Long: lastRow = .Cells(Rows.Count, 3).End(xlUp).Row - 1
    
        '// ���t���u00��00���v�ɕϊ�
        .Cells(2, 3).Formula = "=MID(D2," & Len(.Cells(2, 4).value) - 3 & ",2) & ""��"" & TEXT(RIGHT(D2,2),0) & ""��"""
        .Cells(2, 3).AutoFill .Range(Cells(2, 3), Cells(lastRow, 3))
        
        '// ���t����R�s�[���ē\��t��
        .Range(Cells(2, 3), Cells(lastRow, 3)).Copy
        ThisWorkbook.Sheets("��s����").Cells(6, 2).PasteSpecial xlPasteValues
        
        '// �������`���R�s�[���ē\��t��
        .Range(Cells(2, 8), Cells(lastRow, 8)).Copy
        ThisWorkbook.Sheets("��s����").Cells(6, 5).PasteSpecial xlPasteValues
        
        '// ���z���R�s�[���ē\��t��
        .Range(Cells(2, 5), Cells(lastRow, 5)).Copy
        ThisWorkbook.Sheets("��s����").Cells(6, 6).PasteSpecial xlPasteValues
    End With
    
    bankCsv.Close False
    Set bankCsv = Nothing
        
    With ThisWorkbook.Sheets("��s����")
    
        ' //�u��s���ׁv�̃V�[�g���e�[�u���ɐݒ�
        With .ListObjects.add(xlSrcRange, .Range(.Cells(5, 1), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 9)), , xlYes, , "TableStyleLight14")
            .Name = "account_statements"
        End With
    End With
    
    '// �������`��������R�[�h�Ǝ���於���擾
    Call fetchCodeAndName
    
    '// �萔���̌v�Z
    '// calculateHandlingCharge [�����N]
    Call calculateHandlingCharge(targetYear)
    
    '// �v�m�F�t���O�E�Čv�Z�t���O�̐ݒ�
    Call setFlags
    
    With Sheets("��s����")
    
        '// ���t�E����於�E�����R�[�h�E���l��̕����̑傫���ݒ�
        .Range(.Cells(6, 2), .Cells(Rows.Count, 2).End(xlUp)).Font.Size = 10
        .Range(.Cells(6, 4), .Cells(Rows.Count, 5).End(xlUp)).Font.Size = 10
        .Range(.Cells(6, 8), .Cells(Rows.Count, 8).End(xlUp)).Font.Size = 9
        
        .Range(.Cells(6, 2), .Cells(Rows.Count, 2).End(xlUp)).HorizontalAlignment = xlCenter
        .Range(.Cells(6, 9), .Cells(Rows.Count, 1).End(xlUp).Offset(, 8)).HorizontalAlignment = xlCenter
        
        .Range(.Columns(2), .Columns(8)).EntireColumn.AutoFit
        
        .Columns(11).Hidden = True
        .Cells(6, 1).Select
    End With
    
    Application.CutCopyMode = False
    
End Sub

'/**
 '* �w�b�_�[�̐ݒ�
'**/
Private Sub createTableHeader()

    With Sheets("��s����")
        .Range(.Cells(5, 1), .Cells(Rows.Count, 11)).Clear
        .Cells.Font.Name = "Meiryo UI"
        .Columns(1).HorizontalAlignment = xlCenter
        
        '// �w�b�_�[�̓��e
        .Cells(5, 1).value = "�v�m�F"
        .Cells(5, 2).value = "���t"
        .Cells(5, 3).value = "�����R�[�h"
        .Cells(5, 4).value = "����於"
        .Cells(5, 5).value = "�������`"
        .Cells(5, 6).value = "�U�����z"
        .Cells(5, 7).value = "���|���Ƃ̍��z"
        .Cells(5, 8).value = "���l"
        .Cells(5, 9).value = "�Čv�Z"
        '// ���Z�������ǂ�����������
        .Cells(5, 11).value = "combined"
        .Range(.Columns(6), .Columns(7)).NumberFormatLocal = "#,##0;[��]-#,##0"
        
        '// �w�b�_�[�̐F�Ȃ�
        With .Range(.Cells(5, 1), .Cells(5, 9))
            .Interior.ColorIndex = 50
            .Font.Color = vbWhite
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        .Rows(5).RowHeight = 50
    
        '// �w�b�_�[�̐���
        With .Cells(5, 1).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "������͊m�F���K�v����\���܂��B" & vbLf & vbLf & "�l��1�̂Ƃ� �I ��\�����܂��B�_�u���N���b�N�Ő؂�ւ��\�ł��B"
        End With
        
        With .Cells(5, 2).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "������͓����̂��������t��\���܂��B"
        End With
        
        With .Cells(5, 3).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "������͓����̂����������R�[�h��\���܂��B" & vbLf & vbLf & "����悪�o�^����Ă��Ȃ��ꍇ�A�u����悪������܂���v�ƕ\�����܂��B"
        End With
        
        With .Cells(5, 4).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "������͓����̂���������於��\���܂��B" & vbLf & vbLf & "����悪�o�^����Ă��Ȃ��ꍇ�A�u����悪������܂���v�ƕ\�����܂��B"
        End With
        
        With .Cells(5, 5).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "������͓����̂������������`��\���܂��B"
        End With
        
        With .Cells(5, 6).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "������͐U�荞�܂ꂽ���z��\���܂��B"
        End With
        
        With .Cells(5, 7).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "������͔��|���Ɠ����z�̍��z��\���܂��B" & vbLf & vbLf & "���G�ȓ����ɂ͑Ή��ł��Ă��Ȃ����߁A�蓮�ŏC�������肢���܂�m(__)m"
        End With
        
        With .Cells(5, 8).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "������͔��l��\���܂��B" & vbLf & vbLf & "���Z�����╡��������̏ꍇ�ȂǂɎg�p����܂��B"
        End With
    
        With .Cells(5, 9).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "������͍Čv�Z���s������\���܂��B" & vbLf & vbLf & "�_�u���N���b�N�Ő؂�ւ��\�ł��B"
        End With
        
    End With

End Sub

'/**
 '* �捞�O�Ƀt�@�C���̌`�������������𔻒f
'**/
Private Function checkFile(ByVal fileName As String) As Boolean

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(fileName)
    
    With targetFile.Sheets(1)
        If .Cells(1, 8).value <> "�²����" Then
            targetFile.Close False
            MsgBox "�w�肵���t�@�C���̌`��������������܂���B" & vbLf & "�O��Z�F��s�̐U���������ׂ��w�肵�Ă��������B", vbExclamation, "�������׎捞"
            checkFile = False
            Exit Function
        End If
    End With
    
    checkFile = True

End Function

'// �������`�����ƂɃf�[�^�x�[�X��������R�[�h������於���擾
Public Sub fetchCodeAndName()

    Workbooks.Open ThisWorkbook.Path & "\database\customers.xlsx"
    
    With ThisWorkbook.Sheets("��s����")

        '// �����R�[�h��E����於��XLOOKUP�Ńf�[�^�x�[�X�t�@�C������擾
        .Cells(6, 3).Formula = "=XLOOKUP([@�������`],[customers.xlsx]customers!$C:$C,[customers.xlsx]customers!$A:$A,""����悪������܂���"",0)"
        .Cells(6, 4).Formula = "=XLOOKUP([@�������`],[customers.xlsx]customers!$C:$C,[customers.xlsx]customers!$B:$B,""����悪������܂���"",0)"
        
        .Cells(6, 3).AutoFill .Range(.Cells(6, 3), .Cells(Rows.Count, 2).End(xlUp).Offset(, 1)), xlFillValues
        .Cells(6, 4).AutoFill .Range(.Cells(6, 4), .Cells(Rows.Count, 2).End(xlUp).Offset(, 2)), xlFillValues
        
        '// ����l�ɕϊ�
        .Range(.Cells(6, 3), .Cells(Rows.Count, 3).End(xlUp)).Copy
        .Cells(6, 3).PasteSpecial xlPasteValues
        .Range(.Cells(6, 4), .Cells(Rows.Count, 4).End(xlUp)).Copy
        .Cells(6, 4).PasteSpecial xlPasteValues
    
    End With
    
    Application.CutCopyMode = False
    
    Workbooks("customers.xlsx").Close False

End Sub

'/**
 '* �v�m�F�t���O�E�Čv�Z�t���O�̐ݒ�
'**/
Private Sub setFlags()

    With Sheets("��s����")
    
        '// �v�m�F��̓��e�ݒ�(���z���}�C�i�X�̏ꍇ��1000�~�𒴂���ꍇ��1�ɂ���)
        .Cells(6, 1).Formula = "=IFERROR(IF(OR([@���|���Ƃ̍��z]<0,[@���|���Ƃ̍��z]>1000,[@���|���Ƃ̍��z] =""""),1,0),0)"
        .Range(.Cells(6, 1), .Cells(Rows.Count, 1).End(xlUp)).Copy
        .Cells(6, 1).PasteSpecial xlPasteValues
        
        '// �v�m�F�t���O��Ƀh���b�v�_�E���ݒ�(0��1)
        .Range(.Cells(6, 1), .Cells(Rows.Count, 1).End(xlUp)).Validation.add _
            Type:=xlValidateList, Formula1:="0,1"
        
        '// �v�m�F�t���O�̃A�C�R���ݒ�
        Dim confirmFlag As IconSetCondition
        Set confirmFlag = .Range(.Cells(6, 1), .Cells(Rows.Count, 1).End(xlUp)).FormatConditions.AddIconSetCondition
        
        With confirmFlag
            .IconSet = ThisWorkbook.IconSets(xl3Symbols2)
            .ShowIconOnly = True
        
            .IconCriteria(1).Icon = xlIconNoCellIcon
            
            '// �v�m�F�t���O�̒l��1�̏ꍇ�͉��F���т�����}�[�N��\������
            With .IconCriteria(2)
                .Type = xlConditionValueNumber
                .value = 1
                .Operator = xlGreaterEqual
            End With
            
            With .IconCriteria(3)
                .Icon = xlIconNoCellIcon
                .Type = xlConditionValueNumber
                .value = 1
                .Operator = xlGreater
            End With
            
        End With
        
        '// �v�m�F�t���O�������Ă���s�𔖂��ԂŐF�t������悤���[����ݒ�
        With .Range(.Cells(6, 1), .Cells(Rows.Count, 1).Offset(, 8)).FormatConditions.add( _
            Type:=xlExpression, _
            Formula1:="=$A6 = 1" _
        )
            .Interior.ColorIndex = 40
            .StopIfTrue = False
        End With
    
        '// �Čv�Z�t���O��Ƀh���b�v�_�E���ݒ�(0��1)
        .Range(.Cells(6, 9), .Cells(Rows.Count, 1).End(xlUp).Offset(, 8)).Validation.add _
            Type:=xlValidateList, Formula1:="0,1"
            
        '// �Čv�Z�t���O�̃A�C�R���ݒ�
        Dim recalcFlag As IconSetCondition
        Set recalcFlag = .Range(.Cells(6, 9), .Cells(Rows.Count, 1).End(xlUp).Offset(, 8)).FormatConditions.AddIconSetCondition
        
        With recalcFlag
            .IconSet = ThisWorkbook.IconSets(xl3Flags)
            .ShowIconOnly = True
            
            .IconCriteria(1).Icon = xlIconNoCellIcon
            
            '// �Čv�Z�t���O�̒l��1�̏ꍇ�͗΂̊���\��
            With .IconCriteria(2)
                .Icon = xlIconGreenFlag
                .Type = xlConditionValueNumber
                .value = 1
                .Operator = xlGreaterEqual
            End With
            
            With .IconCriteria(3)
                .Icon = xlIconNoCellIcon
                .Type = xlConditionValueNumber
                .value = 1
                .Operator = xlGreater
            End With
        End With
        
    End With
    
    Set confirmFlag = Nothing
    Set recalcFlag = Nothing
    
End Sub

'/**
 '* �萔���Čv�Z
'**/
Public Sub openFormToRecalculate()
    
    Sheets("mode").Cells(1, 1).value = "RECALCULATE"
    
    formYear.Show vbModeless

End Sub

'/**
 '* �萔���v�Z
 '*
'**/
Public Sub calculateHandlingCharge(ByVal paymentYear As Long, Optional ByVal isRecalc As Boolean = False)

    '// �����֘A�̃f�[�^�x�[�X�ڑ�
    Dim customerCon As ADODB.Connection: Set customerCon = connectDb(ThisWorkbook.Path & "\database\customers.xlsx")
    
    '// ���|���̃f�[�^�x�[�X�ڑ�(���N�ƍ�N)
    Dim salesConThisYear As ADODB.Connection: Set salesConThisYear = connectDb(ThisWorkbook.Path & "\database\sales\" & paymentYear & ".xlsx")
    Dim salesConLastYear As ADODB.Connection: Set salesConLastYear = connectDb(ThisWorkbook.Path & "\database\sales\" & paymentYear - 1 & ".xlsx")
    
    '// ����惌�R�[�h�Z�b�g
    Dim customerRs As New ADODB.Recordset
    customerRs.Open "SELECT * FROM [customers$]", customerCon, adOpenStatic, adLockOptimistic
    
    '// ���ヌ�R�[�h�Z�b�g(���N�ƍ�N)
    Dim salesRsThisyear As New ADODB.Recordset
    salesRsThisyear.Open "SELECT * FROM [sales$]", salesConThisYear, adOpenStatic, adLockOptimistic
    
    Dim salesRsLastYear As New ADODB.Recordset
    salesRsLastYear.Open "SELECT * FROM [sales$]", salesConLastYear, adOpenStatic, adLockOptimistic
    
    '// ���Z�O���[�v���R�[�h�Z�b�g
    Dim groupRs As New Recordset
    groupRs.Open "SELECT * FROM [combined_groups$]", customerCon, adOpenStatic, adLockOptimistic
    
    '// ��s���׃��R�[�h�Z�b�g
    Dim bankRs As New Recordset
    bankRs.CursorLocation = adUseClient
    bankRs.Open "SELECT * FROM [��s����$A5:H]", connectDb(ThisWorkbook.FullName), adOpenStatic, adLockOptimistic
  
    '/**
     '* ���|���Ɠ����z�̍��z�v�Z
    '**/
    Dim rowNumber As Long: rowNumber = 6
    Dim customerId As Long
    Dim paymentMonth As Long
      
    Do Until Sheets("��s����").Cells(rowNumber, 2).value = ""
        '// ����悪������Ȃ��ꍇ�A���l�ɕ���������Ƃ�����̂͂Ƃ΂�
        If Sheets("��s����").Cells(rowNumber, 3).value = "����悪������܂���" Or InStr(1, Sheets("��s����").Cells(rowNumber, 8).value, "���������") > 0 Then
            GoTo Continue
        '// �Čv�Z�̏ꍇ�́A�Čv�Z�t���O�������Ă��Ȃ��ꍇ�A���Z�̏ꍇ�͂Ƃ΂�
        ElseIf isRecalc = True And Sheets("��s����").Cells(rowNumber, 9).value <> 1 Then
            GoTo Continue
        End If
        
        customerId = Sheets("��s����").Cells(rowNumber, 3).value
        customerRs.filter = "id = " & customerId
        
        '// ���E����̎����̏ꍇ �� ���l�ɓ���
        If customerRs!is_offset Then
            bankRs.filter = "�����R�[�h = " & customerId
            bankRs!���l = "���u���E�L�v�̎����ł��B"
            bankRs.Update
            bankRs.filter = adFilterNone
        End If
        
        paymentMonth = Split(Sheets("��s����").Cells(rowNumber, 2).value, "��")(0)
    
        '/**
         '* ���Z�̏ꍇ�̏���
        '**/
        If customerRs!combined_group <> 0 Then
            '// ���Z�O���[�vID�Ŏ������i��
            Dim combinedGroupId As Long: combinedGroupId = customerRs!combined_group
            customerRs.filter = "combined_group = " & combinedGroupId
            
            '// ���|���Ɠ����z�̍��z
            Dim combinedDiff As Long
            combinedDiff = calculateTotalSales(salesRsThisyear, salesRsLastYear, customerRs, paymentMonth) - Sheets("��s����").Cells(rowNumber, 6).value
                        
            '// ���Z�O���[�v�ɓo�^����Ă���������`�������O���[�v������ꍇ �� �S�Ă̍��Z�O���[�v�Ŕ��|���Ɠ����z�̍��z�����߁A
            '// �ł����z���������O���[�v���̗p
            groupRs.filter = "account = " & "'" & customerRs!customer_account & "'"
            
            If groupRs.RecordCount >= 2 Then
                Dim tmpDiff As Long
                
                Do Until groupRs.EOF
                    customerRs.filter = "combined_group = " & groupRs!ID
                    tmpDiff = calculateTotalSales(salesRsThisyear, salesRsLastYear, customerRs, paymentMonth) - Sheets("��s����").Cells(rowNumber, 6).value
                    
                    '// ���|���Ɠ����z�̍��z���Ó��ȕ����̗p
                    combinedDiff = compareNumbersAsCommision(combinedDiff, tmpDiff)
                    
                    If combinedDiff = tmpDiff Then
                        combinedGroupId = groupRs!ID
                    End If
                    
                    groupRs.MoveNext
                Loop
            End If
        
            '// ���Z�����ő}�����ꂽ�s��
            Dim insertedCount As Long
            
            '// calculateCombindedPayment [���|���Ɠ����z�̍��z], [������], [���Z�O���[�vID],[���ヌ�R�[�h�Z�b�g(���N)],[���ヌ�R�[�h�Z�b�g(��N)],[����惌�R�[�h�Z�b�g],[�������̍s�ԍ�]
            With Sheets("��s����")
                insertedCount = calculateCombinedPayment( _
                    combinedDiff, _
                    paymentMonth, _
                    combinedGroupId, _
                    salesRsThisyear, _
                    salesRsLastYear, _
                    customerRs, _
                    rowNumber _
                )
                rowNumber = rowNumber + insertedCount
            End With
            
            GoTo Continue
        End If
        
        '// ����������̂�������̏ꍇ
        If WorksheetFunction.CountIf(Sheets("��s����").Columns(4), Sheets("��s����").Cells(rowNumber, 4).value) > 1 Then
            Dim totalSales As Long
            Dim targetMonth As Long: targetMonth = getTargetMonth(customerRs!customer_site, paymentMonth)
            
            '// �Ώی�����N�̏ꍇ �� ��N�̔��|�����R�[�h�Z�b�g���g�p
            Dim salesRs As Recordset
            If targetMonth < 0 Then
                Set salesRs = salesRsLastYear
            Else
                Set salesRs = salesRsThisyear
            End If
            
            '// �����������قȂ�����̔��|����������ɕ�����ē��������ꍇ
            If customerRs!several_times_payment_group <> 0 Then
                customerRs.filter = "several_times_payment_group = " & customerRs!several_times_payment_group
                totalSales = calculateTotalSales(salesRsThisyear, salesRsLastYear, customerRs, paymentMonth)
            
            '// �����������1�̎����̔��|����������ɕ�����ē��������ꍇ�ɑΏی�����N�̏ꍇ
            Else
                salesRs.filter = "customer_id = " & customerId & " AND sales_month = " & Abs(targetMonth)
                totalSales = salesRs!sales
            End If
            
            '// calculateSeveralTimePayment [�������`], [���|��], [���|�����R�[�h�Z�b�g], [����惌�R�[�h�Z�b�g], [��s���׃��R�[�h�Z�b�g], [������]
            Call calculateSeveralTimePayment( _
                customerRs!customer_account, totalSales, salesRs, customerRs, bankRs, paymentMonth _
            )
            GoTo Continue
        End If
        
        '// ��L�ȊO�̓���
        With Sheets("��s����")
            '// calculateDifference [�����R�[�h], [�����z], [������], [�����T�C�g], [���ヌ�R�[�h�Z�b�g]
            .Cells(rowNumber, 7).value = _
                calculateDifference(customerId, .Cells(rowNumber, 6).value, paymentMonth, customerRs!customer_site, salesRsThisyear, salesRsLastYear)
        End With
        
        customerRs.filter = adFilterNone
Continue:
        rowNumber = rowNumber + 1
    Loop
    
    salesRsThisyear.Close
    salesRsLastYear.Close
    customerRs.Close
    groupRs.Close
    
    Set salesRsThisyear = Nothing
    Set salesRsLastYear = Nothing
    Set customerRs = Nothing
    Set groupRs = Nothing

End Sub

'/**
 '* ���Z�����̏ꍇ�̎萔���Ȃǂ����߂�
 '* @return ���s�}��������
 '* @params amountDiff       ���|���Ɠ����z�̍��z
 '* @params paymentMonth     ������
 '* @params combinedGroupId  ���Z�O���[�vID
 '* @params salesRsThisYear  ���ヌ�R�[�h�Z�b�g(���N)
 '* @params salesRsLastYear  ���ヌ�R�[�h�Z�b�g(��N)
 '* @pamras customerRs       ����惌�R�[�h�Z�b�g
 '* @params sheetRow         �V�[�g�u��s���ׁv�̌��ݏ������̍s�ԍ�
'**/
Private Function calculateCombinedPayment(ByVal amountDiff As Long, ByVal paymentMonth As Long, ByVal combinedGroupId As Long, ByVal salesRsThisyear As Recordset, ByVal salesRsLastYear As Recordset, ByVal customerRs As Recordset, ByVal sheetRow As Long) As Long

    Dim totalPayment As Long: totalPayment = Sheets("��s����").Cells(sheetRow, 6).value

    customerRs.filter = "combined_group = " & combinedGroupId
        
    Sheets("��s����").Cells(sheetRow, 7).value = amountDiff
    customerRs.MoveFirst
    
    Dim insertedCount As Long
    Dim counter As Long: counter = -1
    Dim salesRs As ADODB.Recordset
    
    '// ���Z�O���[�v�ɓo�^����Ă��镪�̉�Ж��Ȃǂ̃f�[�^����s���ׂɓ���
    Do Until customerRs.EOF
                
        '// �����R�[�h�ƑΏی��Ŕ���𒊏o getTargetMonth [�T�C�g],[������]
        Dim targetMonth As Long: targetMonth = getTargetMonth(customerRs!customer_site, paymentMonth)
        
        '// �Ώی�����N�̏ꍇ(�l���}�C�i�X�̏ꍇ) �� ��N�̔��|�����R�[�h�Z�b�g���g�p
        If targetMonth < 0 Then
            Set salesRs = salesRsLastYear
        Else
            Set salesRs = salesRsThisyear
        End If
        
        salesRs.filter = "customer_id = " & customerRs!ID & " AND sales_month = " & Abs(targetMonth)
        
        If salesRs.RecordCount = 0 Then: GoTo Continue
    
        counter = counter + 1
        
        '// ���Z�O���[�v�Ŕ��オ0�~�ȏ�̍ŏ��̏����̏ꍇ �� ���ォ��萔�������������z������z�Ƃ���
        If counter = 0 Then
            Sheets("��s����").Cells(sheetRow + counter, 6).value = salesRs!sales - amountDiff
        '// 2���ڈȍ~ �� �s�}�����ăf�[�^���� ���㍂������z�Ƃ��A�萔����0�Ƃ��Acombined���True�ɂ���
        Else
            Sheets("��s����").Rows(sheetRow + counter).Insert xlDown
            insertedCount = insertedCount + 1
            Sheets("��s����").Cells(sheetRow + counter, 6).value = salesRs!sales
            Sheets("��s����").Cells(sheetRow + counter, 7).value = 0
            Sheets("��s����").Cells(sheetRow + counter, 11).value = True
        End If
        
        '// ���t�E�����R�[�h�E����於�E�������`����(���t�͒l�������͂���ƃ��[�U�[��`�ɂȂ邽�߁A�R�s�[���ē\��t����)
        With Sheets("��s����")
            .Cells(sheetRow, 2).Copy Destination:=.Cells(sheetRow + counter, 2)
            .Cells(sheetRow + counter, 3).value = customerRs!ID
            .Cells(sheetRow + counter, 4).value = customerRs!customer_name
            .Cells(sheetRow + counter, 5).value = customerRs!customer_account
        End With
        
Continue:
        salesRs.filter = adFilterNone
        customerRs.MoveNext
    Loop
    
    '// �����������̔��オ���Z����Ă���ꍇ(1�̔���ɑ΂�������ł͂Ȃ��ꍇ) �� ���l�ɍ��Z�����z��\���A���Z���ǂ���������combined��̒l��True�ɂ���
    If counter > 0 Then
        Sheets("��s����").Cells(sheetRow, 8).value = "���Z�����z:" & Format(totalPayment, "#,##0") & "�~ " & Sheets("��s����").Cells(sheetRow, 8).value
        Sheets("��s����").Cells(sheetRow, 11).value = True
        
        
        With Sheets("��s����").Range(Cells(sheetRow, 1), Cells(sheetRow + counter, 9))
            .Interior.Color = RGB(144, 238, 144)
            .Font.Color = vbBlack
        End With
    End If
            
    customerRs.filter = adFilterNone
    
    '// ���s�}�����������Ԃ�l�ƂȂ�
    calculateCombinedPayment = insertedCount
    
End Function

'/**
 '* ���ɕ������������������̏���
 '* @params customerAccount �������`
 '* @params sales           ���|��
 '* @params customerRs      ����惌�R�[�h�Z�b�g
 '* @params bankRs          ��s���׃��R�[�h�Z�b�g
 '* @params paymentMonth    ������
'**/
Private Sub calculateSeveralTimePayment(ByVal customerAccount As String, ByVal sales As Long, ByVal salesRs As Recordset, ByVal customerRs As Recordset, ByVal bankRs As Recordset, ByVal paymentMonth As Long)
    
    '// �����z�̍��v
    Dim totalPayment As Long
    totalPayment = WorksheetFunction.SumIf(Sheets("��s����").Columns(5), customerAccount, Sheets("��s����").Columns(6))
    
   '/**
    '* �����z�Ɣ��|���̍��z�����߁A1000�~�ȏ�ł���΍��z������񐔂Ŋ���A���ꂼ��̎萔���Ƃ���
    '* 1000�~��菬�����ꍇ�͓����z���ő�̂Ƃ���̎萔���Ƃ���
   '**/
    Dim amountDiff As Long: amountDiff = sales - totalPayment
    bankRs.filter = "�������` = " & "'" & customerAccount & "'"
    
    '// ������
    Dim paymentCount As Long: paymentCount = WorksheetFunction.CountIf(Sheets("��s����").Columns(5), customerAccount)
    
    '// ���|���Ɠ����z���v�̍��z��0�̏ꍇ �� �Y���̎����̎萔����S��0�~�ɂ���
    If amountDiff = 0 Then
        Do While bankRs.EOF = False
            bankRs!���|���Ƃ̍��z = 0
            bankRs!���l = "��������� " & bankRs!���l
            bankRs.MoveNext
        Loop
    '// ���z�������񐔁~880�~�ȓ��̏ꍇ �� �萔��������񐔂Ŋ��������̂����ꂼ��̎萔���Ƃ���
    ElseIf 0 < amountDiff And amountDiff <= 880 * paymentCount Then
        Do While bankRs.EOF = False
            bankRs!���|���Ƃ̍��z = amountDiff / paymentCount
            bankRs!���l = "���������" & bankRs!���l
            bankRs.MoveNext
        Loop
    '// ����ȊO �� ���z������z��1�ԑ����Ƃ���̎萔���Ƃ��A����ȊO�̎萔����0�~�ɂ���
    Else
        bankRs.Sort = "�U�����z DESC"
        bankRs!���|���Ƃ̍��z = amountDiff
        bankRs!���l = "��������� " & bankRs!���l
        bankRs.MoveNext
        
        Do While bankRs.EOF = False
            bankRs!���|���Ƃ̍��z = 0
            bankRs!���l = "���������" & bankRs!���l
            bankRs.MoveNext
        Loop
    End If
    
    '// �����������قȂ�����̔��|����������ɕ����ē��������ꍇ
    If customerRs.RecordCount > 1 Then
        bankRs.MoveFirst
            
        Dim tmpSales As Long
        Dim validCustomerId As Long
        
        '// ����惌�R�[�h�Z�b�g�̃t�B���^�[:�قȂ�����̔��|������������Ă���ꍇ��, "several_times_payment_group = 0" �̂悤�ȃt�B���^�[�ɂȂ��Ă���
        Dim previousFilter As String: previousFilter = customerRs.filter
        
        '// �����̎����̔��|���Ɠ����z�̍��z���r���āA�ł��Ó��Ȕ��|���̎����̃R�[�h�Ɩ��O����͂���
        Do While bankRs.EOF = False
            
            Do While customerRs.EOF = False
                salesRs.filter = "customer_id = " & customerRs!ID & " AND sales_month = " & getTargetMonth(customerRs!customer_site, paymentMonth)
                
                '// �����z��2�̔��|������A�ǂ��炪���|���Ƃ��đÓ����𔻒f
                '// compareNumbersAsSales [�����z]. [���|��1], [���|��2]
                tmpSales = compareNumbersAsSales(bankRs!�U�����z, tmpSales, salesRs!sales)
                        
                If tmpSales = salesRs!sales Then
                    validCustomerId = customerRs!ID
                End If
                
                customerRs.MoveNext
            Loop
            
            '// �Ó����Ɣ��f���ꂽ�����̃R�[�h�Ɩ��O�𖾍ׂɓ���
            customerRs.MoveFirst
            customerRs.filter = "id = " & validCustomerId
            bankRs!�����R�[�h = customerRs!ID
            bankRs!����於 = customerRs!customer_name
        
            customerRs.filter = previousFilter
            bankRs.MoveNext
        Loop
    End If
                
    bankRs.MoveFirst
    bankRs.Update
    
    customerRs.filter = adFilterNone
    bankRs.filter = adFilterNone
    bankRs.Sort = ""
    
End Sub

'/**
 '* ��������̍��v�z�����߂�
 '* @params salesRsThisYear ���|�����R�[�h�Z�b�g(���N)
 '* @params salesRsLastYear ���|�����R�[�h�Z�b�g(��N)
 '* @params customerRs      ����惌�R�[�h�Z�b�g
 '* @params paymentMonth    ������
 '* @return ����̍��v�z
'**/
Private Function calculateTotalSales(ByVal salesRsThisyear As Recordset, ByVal salesRsLastYear As Recordset, customerRs As Recordset, paymentMonth As Long) As Long

    Dim targetMonth As Long
    Dim salesRs As ADODB.Recordset
    
    Do While customerRs.EOF = False
        '// ����̑Ώی����T�C�g���狁�߂�(��N�̏ꍇ�͑Ώی��Ƀ}�C�i�X����) ��) getTargetMonth [����], [1] �� -12
        targetMonth = getTargetMonth(customerRs!customer_site, paymentMonth)
    
        '// �Ώی�����N�̏ꍇ(�l���}�C�i�X�̏ꍇ) �� ��N�̔��|�����R�[�h�Z�b�g���g�p
        If targetMonth < 0 Then
            Set salesRs = salesRsLastYear
        Else
            Set salesRs = salesRsThisyear
        End If
        
        salesRs.filter = "customer_id = " & customerRs!ID & " AND sales_month = " & Abs(targetMonth)
        
        If salesRs.RecordCount = 0 Then
            GoTo Continue
        End If
        
        calculateTotalSales = calculateTotalSales + salesRs!sales

Continue:
        salesRs.filter = adFilterNone
        customerRs.MoveNext
    Loop

    customerRs.MoveFirst

    Set salesRs = Nothing

End Function

'/**
 '* �������ƃT�C�g���牽�����̔��オ�Ώۂ������߂�
'**/
Private Function getTargetMonth(ByVal site As String, ByVal paymentMonth As Long)

    If site = "����" Then
        getTargetMonth = passMonth(paymentMonth, -1)
    ElseIf site = "���X" Then
        getTargetMonth = passMonth(paymentMonth, -2)
    ElseIf site = "�����X" Then
        getTargetMonth = passMonth(paymentMonth, -3)
    End If
    
End Function

'/**
 '* ���|���ƐU�����z�̍��z���v�Z
 '*
'**/
Private Function calculateDifference(ByVal customerId As Long, ByVal amount As Long, ByVal paymentMonth As Long, ByVal site As String, ByVal salesRsThisyear As ADODB.Recordset, ByVal salesRsLastYear As ADODB.Recordset)

    Dim targetMonth As Long
    
    If site = "����" Then
        targetMonth = passMonth(paymentMonth, -1)
    ElseIf site = "���X" Then
        targetMonth = passMonth(paymentMonth, -2)
    ElseIf site = "�����X" Then
        targetMonth = passMonth(paymentMonth, -3)
    End If
    
    Dim salesRs As ADODB.Recordset
    
    '// �������ƃT�C�g���甄�|���̑Ώی������߂����ʂ��}�C�i�X�̏ꍇ(�������̔N���1�N�O���Ώۂ̏ꍇ)��1�N�O�̃��R�[�h�Z�b�g���g�p����
    If targetMonth < 0 Then
        Set salesRs = salesRsLastYear
    Else
        Set salesRs = salesRsThisyear
    End If
    
    salesRs.filter = "customer_id = " & customerId & " AND sales_month = " & Abs(targetMonth)
    If salesRs.RecordCount = 0 Then
        salesRs.filter = adFilterNone
        Exit Function
    End If

    '/**
     '* ���z���v�Z���A���z��1000�~�𒴂���ꍇ�͑O���̔��ォ��3�J���������̂ڂ�A���ꂼ�ꍷ�z���v�Z����
     '* ���̒��ō��z��1000�~�ȉ��̂��̂�������΃T�C�g�ɓo�^����Ă��锄�ォ������z�����������z�����z�Ƃ���
    '**/
    
    '// DB�ɓo�^����Ă���T�C�g�̔��㍂�Ɠ����z�̍��z
    Dim amountDiffDefault As Long: amountDiffDefault = salesRs!sales - amount
    Dim amountDiff As Long
    
    '// �T�C�g�ɓo�^����Ă��錎�̔��㍂�Ɠ����z�̍��z��0�~�ȏ�1000�~�ȉ��̏ꍇ
    If 0 <= amountDiffDefault And amountDiffDefault <= 1000 Then
        amountDiff = amountDiffDefault
        GoTo Break
    End If
    
    Dim i As Long
    
    '// ���z��1000�~�ȏ�̏ꍇ �� �������̑O�����ォ��3�J�������̂ڂ�A���ꂼ��̍��z���v�Z���A�ł��Ó��Ȃ��̂��̗p����
    For i = 1 To 3
        '// ����������i���������l��0�ȉ��̏ꍇ �� 1�N�O�̔��|�����R�[�h�Z�b�g���g�p����
        If paymentMonth - i < 1 Then
            Set salesRs = salesRsLastYear
        Else
            Set salesRs = salesRsThisyear
        End If
        
        salesRs.filter = "customer_id = " & customerId & " AND sales_month = " & paymentMonth - i
        
        If salesRs.RecordCount = 0 Then
            salesRs.filter = adFilterNone
            GoTo Continue
        End If
        
        amountDiff = salesRs!sales - amount
        
        If 0 <= amountDiff And amountDiff <= 1000 Then
            GoTo Break
        End If
Continue:
    Next

Break:
    salesRs.filter = adFilterNone
    
    '// ����������3�J�������̂ڂ��Ă��Ó��ȍ��z��������Ȃ��ꍇ�́A�T�C�g���̔��㍂�Ɠ����z�̍��z���g�p����
    If amountDiff < 0 Or 1000 < amountDiff Then
        amountDiff = amountDiffDefault
    End If
    
    calculateDifference = amountDiff
    
    Set salesRs = Nothing
    
End Function

'// �Ǘ����֓\��t��
Public Sub pasteToLedger()

    Application.ScreenUpdating = False

    If MsgBox("��s���ׂ��Ǘ����֓\��t���܂����A��낵���ł���?", vbQuestion + vbYesNo, "�R�݉^�����|������p�t�@�C��") = vbNo Then
        Exit Sub
    End If
    
    Dim filePath As String: filePath = Sheets("�ݒ�").Cells(8, 3).value
    
    Dim fso As New FileSystemObject
    
    '// �\��t���悪���݂��Ȃ���Δ�����
    If fso.FileExists(filePath) = False Then
        MsgBox "�\��t����Ƃ��Đݒ肳��Ă���t�@�C�������݂��܂���B", vbExclamation, "�R�݉^�����|������p�t�@�C��"
        Set fso = Nothing
        Exit Sub
    End If
    
    '/**
     '* �K�v�ȉӏ���\��t��
    '**/
    Dim ledgerFile As Workbook: Set ledgerFile = Workbooks.Open(filePath)
    
    '// �Ǘ����Ɂu���[�N2�v�̃V�[�g��������Δ�����
    If sheetExist(ledgerFile, "���[�N2") = False Then
        ledgerFile.Close False
        MsgBox "�Ǘ����Ƃ��Ďw�肳��Ă���t�@�C�����K�؂ł͂���܂���B" & vbLf & "�Ǘ����Ɂu���[�N2�v�̃V�[�g�����邱�Ƃ��m�F���Ă��������B", vbExclamation, "�R�݉^�����|������p�t�@�C��"
        
        Set ledgerFile = Nothing
        Set fso = Nothing
        Exit Sub
    End If
    
    ledgerFile.Sheets("���[�N2").Cells.Clear
        
    With ThisWorkbook.Sheets("��s����")
        '// ���t�\��t��
        .Range(.Cells(5, 2), .Cells(Rows.Count, 2).End(xlUp)).Copy
        ledgerFile.Sheets("���[�N2").Cells(1, 1).PasteSpecial xlPasteValues
    
        '// ���t���u0��0���v������t�̐��������ɕύX ��)10��3�� �� 3
        With ledgerFile.Sheets("���[�N2")
            .Cells(2, 2).Formula = "=DAY(A2)"
            .Cells(2, 2).AutoFill .Range(.Cells(2, 2), .Cells(Rows.Count, 1).End(xlUp).Offset(, 1))
            .Columns(2).Copy
            .Columns(2).PasteSpecial xlPasteValues
            .Columns(1).Delete
            .Cells(1, 1).value = "���t"
        End With
        
        '// �����R�[�h�E����於�\��t��
        .Range(.Cells(5, 3), .Cells(Rows.Count, 4).End(xlUp)).Copy
        ledgerFile.Sheets("���[�N2").Cells(1, 2).PasteSpecial xlPasteValues
        
        '// �U�����z�E���z�\��t��
        .Range(.Cells(5, 6), .Cells(Rows.Count, 6).End(xlUp).Offset(, 1)).Copy
        ledgerFile.Sheets("���[�N2").Cells(1, 4).PasteSpecial xlPasteValues
        
    End With
    
    Application.CutCopyMode = False
    
    ThisWorkbook.Sheets("��s����").Activate
    
    MsgBox "�Ǘ����ւ̓\��t�����������܂����B", vbInformation, "�R�݉^�����|������p�t�@�C��"

End Sub
