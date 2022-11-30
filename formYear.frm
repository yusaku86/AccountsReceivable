VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formYear 
   Caption         =   "銀行明細取込"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "formYear.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// 入金明細の年を選択するフォーム
Option Explicit

'// フォーム起動時の処理
Private Sub UserForm_Initialize()

    Dim i As Long
    
    For i = Year(Now) - 5 To Year(Now)
        cmbYear.AddItem i
    Next
    
    cmbYear.Value = Year(Now)
    
End Sub

'// 「実行」を押したときの処理
Private Sub cmdEnter_Click()

    Me.Hide

    Call putBankStatement(cmbYear.Value)

    Unload Me
    
End Sub
