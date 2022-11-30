Attribute VB_Name = "sales"
'// ����֘A
Option Explicit

'// ����f�[�^�X�V
Public Sub updateSales()
    
    Application.ScreenUpdating = False
    
    Dim wsh As New WshShell
    
    '// ��荞�ރf�[�^��I�����A�t�@�C�����I������Ă��Ȃ��ꍇ�ƃt�@�C�����K�؏o�Ȃ��ꍇ�͔�����
    Dim fileName As String: fileName = selectFile("����f�[�^�捞", wsh.SpecialFolders(4), "CSV�t�@�C��", "*.csv")
    
    Set wsh = Nothing
    
    If fileName = "" Then
        Exit Sub
    ElseIf checkFile(fileName) = False Then
        Exit Sub
    End If
    
    Dim fso As New FileSystemObject
    
    Dim salesCsv As Workbook: Set salesCsv = Workbooks(fso.GetFileName(fileName))
    
    Set fso = Nothing
    
    '// ����N
    Dim salesYear As Long: salesYear = Format(salesCsv.Sheets(1).Cells(2, 1).value, "yyyy")
    Dim salesMonth As Long: salesMonth = Format(salesCsv.Sheets(1).Cells(2, 1).value, "m")
    
    '// ����̔N��DB�Ƃ��Ďg�p����Excel�t�@�C����ύX����
    Dim dbBookName As String: dbBookName = ThisWorkbook.Path & "\database\sales\" & salesYear & ".xlsx"
    Dim fc As New FileController
    fc.createFileIfNotExist (dbBookName)
    Set fc = Nothing
    
    '// DB�ڑ�
    Dim con As ADODB.Connection: Set con = connectDb(dbBookName)

    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [sales$] ORDER BY sales_id", con, adOpenStatic, adLockOptimistic
    
    '// �����id(���オ�܂�DB��1���o�^����Ă��Ȃ����id��1�ɂȂ�)
    Dim nextId As Long

    If rs.RecordCount = 0 Then
        nextId = 1
    Else
        rs.MoveLast
        nextId = rs![sales_id] + 1
    End If

    Dim i As Long

    '// ����f�[�^�̍X�V
    For i = 2 To salesCsv.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row

        '// �����R�[�h�ƔN���Ō������Ċ��Ƀf�[�^������΋��z���X�V���A������ΐV�K�ǉ�����
        rs.filter = _
            "customer_id = " & salesCsv.Sheets(1).Cells(i, 9).value & _
            " AND sales_year = " & salesYear & _
            " AND sales_month = " & salesMonth

        '// �V�K�ǉ�
        If rs.RecordCount = 0 Then
            rs.AddNew
            rs!sales_id = nextId
            rs!customer_id = salesCsv.Sheets(1).Cells(i, 9).value
            rs!sales_year = salesYear
            rs!sales_month = salesMonth
            nextId = nextId + 1
        End If

        rs!sales = salesCsv.Sheets(1).Cells(i, 12).value

        rs.Update

        rs.filter = adFilterNone
    Next

    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    salesCsv.Close False
    Set salesCsv = Nothing
    
    '// ����𐔒l��
    Call convertStrSales2Number(dbBookName)
    
    MsgBox "�捞���������܂����B", Title:="����f�[�^�捞"
    
End Sub

'// �I�������t�@�C�����K�؂��m�F
Private Function checkFile(ByVal fileName As String) As Boolean

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(fileName)

    With targetFile.Sheets(1).Cells(1, 1)
        If .value = "�J�n���t" _
            And .Offset(, 1).value = "�I�����t" _
            And .Offset(, 2).value = "�R�[�h" _
            And .Offset(, 3).value = "����" _
            And .Offset(, 4).value = "�R�[�h" _
            And .Offset(, 5).value = "�Ȗ�" _
            And .Offset(, 6).value = "�R�[�h" _
            And .Offset(, 7).value = "�⏕�Ȗ�" _
            And .Offset(, 8).value = "�R�[�h" _
            And .Offset(, 9).value = "����於" _
            And .Offset(, 10).value = "�J�z�z" _
            And .Offset(, 11).value = "�ؕ����z" _
            And .Offset(, 12).value = "�ݕ����z" _
            And .Offset(, 13).value = "�c��" Then
        
            checkFile = True
        Else
            targetFile.Close False
            MsgBox "�I�������t�@�C�����K�؂ł͂���܂���B", vbExclamation, "����f�[�^�捞"
            checkFile = False
        End If
    End With

End Function


'// ������Ƃ��ĕۑ�����Ă��锄��f�[�^�𐔒l�ɕϊ�
Private Sub convertStrSales2Number(ByVal dbBookName As String)
    
    Dim dbBook As Workbook: Set dbBook = Workbooks.Open(dbBookName)
    
    With dbBook.Sheets("sales")
        .Activate
        
        '// converStr2Number [���l������͈�]
        Call convertStr2Number(.Range(.Cells(2, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 5)))
    End With
    
    dbBook.Close True
    Set dbBook = Nothing
    
End Sub

