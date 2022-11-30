VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// �t�@�C���𐧌䂷��N���X
Option Explicit

Dim fso As FileSystemObject

'// �R���X�g���N�^
Private Sub Class_Initialize()
    
    Set fso = New FileSystemObject

End Sub

'// �w��̃t�@�C�������݂��Ȃ���΍쐬����
Public Sub createFileIfNotExist(ByVal fileName As String)

    If fso.FileExists(fileName) = True Then
        Exit Sub
    End If
    
    Dim salesBook As Workbook: Set salesBook = Workbooks.add
    
    '// �V�[�g���E�J�������ݒ�
    With salesBook.Sheets(1)
        .Name = "sales"
        .Cells(1, 1).Value = "sales_id"
        .Cells(1, 2).Value = "customer_id"
        .Cells(1, 3).Value = "sales"
        .Cells(1, 4).Value = "sales_year"
        .Cells(1, 5).Value = "sales_month"
    End With
    
    salesBook.SaveAs fileName
    salesBook.Close
    
    Set salesBook = Nothing
           
End Sub

'// �f�X�g���N�^
Private Sub Class_Terminate()

    Set fso = Nothing

End Sub