VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formYear 
   Caption         =   "��s���׎捞"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "formYear.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// �������ׂ̔N��I������t�H�[��
Option Explicit

'// �t�H�[���N�����̏���
Private Sub UserForm_Initialize()

    Dim i As Long
    
    For i = Year(Now) - 5 To Year(Now)
        cmbYear.AddItem i
    Next
    
    cmbYear.Value = Year(Now)
    
End Sub

'// �u���s�v���������Ƃ��̏���
Private Sub cmdEnter_Click()

    Me.Hide

    Call putBankStatement(cmbYear.Value)

    Unload Me
    
End Sub
