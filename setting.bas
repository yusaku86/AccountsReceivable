Attribute VB_Name = "setting"
'// �Ǘ����̓\��t����ݒ�
Option Explicit

'// �\��t�����ύX
Public Sub changePath()

    Dim filePath As String: filePath = selectFile("�\��t����ύX", "G:", "Excel�t�@�C��", "*.xls;*.xlsx;*.xlsm")

    If filePath = "" Then
        Exit Sub
    End If
    
    Sheets("�ݒ�").Cells(8, 3).value = filePath
    
    MsgBox "�\��t�����ύX���܂����B", Title:="�R�݉^�����|������p�t�@�C��"
    
End Sub
