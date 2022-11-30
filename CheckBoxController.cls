VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "checkBoxController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'// �`�F�b�N�{�b�N�X�𐧌䂷��N���X
Option Explicit

'�w��͈̔͂Ƀ`�F�b�N�{�b�N�X��ǉ�
Public Sub add(targetRange As Range, ByVal chkName As String)
    
    'chkName�Ɠ������O�̃`�F�b�N�{�b�N�X�����ɂ��邩�m�F���A����΃`�F�b�N�{�b�N�X��ǉ����Ȃ�
    If isExistChk(ActiveSheet, chkName) = True Then: Exit Sub
    
    '�`�F�b�N�{�b�N�X�ǉ�
    With ActiveSheet.CheckBoxes.add(targetRange.Left, targetRange.Top, 25, targetRange.Height)
        .Caption = ""
        .Name = chkName
    End With
    
End Sub

'// �w��̖��O�̃`�F�b�N�{�b�N�X�����݂��邩
Public Function isExistChk(ByVal targetSheet As Worksheet, ByVal chkName As String) As Boolean

    Dim tmpShape As Shape
    
    For Each tmpShape In targetSheet.Shapes
        If tmpShape.Name = chkName Then
            isExistChk = True
            Exit Function
        End If
    Next
            
    isExistChk = False
    
End Function

'// �V�[�g�ɑ��݂���`�F�b�N�{�b�N�X���폜
Public Sub deleteChk(ByVal targetSheet As Worksheet)

    Dim tmpShape As Shape
    
    For Each tmpShape In targetSheet.Shapes
        If tmpShape.Name Like "chk*" Then
            tmpShape.Delete
        End If
    Next

End Sub