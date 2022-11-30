Attribute VB_Name = "setting"
'// 管理帳の貼り付け先設定
Option Explicit

'// 貼り付け先を変更
Public Sub changePath()

    Dim filePath As String: filePath = selectFile("貼り付け先変更", "G:", "Excelファイル", "*.xls;*.xlsx;*.xlsm")

    If filePath = "" Then
        Exit Sub
    End If
    
    Sheets("設定").Cells(8, 3).value = filePath
    
    MsgBox "貼り付け先を変更しました。", Title:="山岸運送売掛金回収用ファイル"
    
End Sub
