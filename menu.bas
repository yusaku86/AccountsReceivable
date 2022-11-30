Attribute VB_Name = "menu"
'// おもにメニューで使用するモジュール
Option Explicit

'// 銀行明細表示
Public Sub showAccountStatement()

    Sheets("銀行明細").Activate

End Sub

'// 取引先マスタ登録画面をホームから表示
Public Sub showCustomersFromHome()

    Sheets("取引先マスタ").Activate
    
    With Sheets("取引先マスタ")
        .Unprotect
        .Shapes("btnEdit").Visible = False
        .Shapes("imgEdit").Visible = False
        .Shapes("btnAdd").Visible = False
        .Shapes("imgAdd").Visible = False
        .Shapes("btnReset").Visible = False
        .Shapes("imgReset").Visible = False
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        .Protect
    End With

End Sub

'// 取引先マスタ登録画面を表示
Public Sub showCustomers()

    Sheets("取引先マスタ").Activate

End Sub

'// 合算グループマスタ登録画面表示
Public Sub showConbinedGroups()

    Sheets("合算グループマスタ").Activate
    
    With Sheets("合算グループマスタ")
        .Unprotect
        .Shapes("btnEdit").Visible = False
        .Shapes("imgEdit").Visible = False
        .Shapes("btnAdd").Visible = False
        .Shapes("imgAdd").Visible = False
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        .Protect
    End With

End Sub

'// 複数回入金グループマスタ登録画面表示
Public Sub showSeveralTimesGroups()

    Sheets("複数回入金グループマスタ").Activate
    
    With Sheets("複数回入金グループマスタ")
        .Unprotect
        .Shapes("btnEdit").Visible = False
        .Shapes("imgEdit").Visible = False
        .Shapes("btnAdd").Visible = False
        .Shapes("imgAdd").Visible = False
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        .Protect
    End With

End Sub

'// ホーム画面を表示
Public Sub showHome()
    
    '// 銀行明細または設定からホーム画面に移動する場合
    If ActiveSheet.Name = "銀行明細" Or ActiveSheet.Name = "設定" Then
        GoTo Show
    End If
    
    ActiveSheet.Unprotect
        
    '// 取引先マスタからホーム画面に移動する場合
    If ActiveSheet.Name = "取引先マスタ" Then
        If WorksheetFunction.CountIf(Columns(10), True) + WorksheetFunction.CountIf(Columns(10), "NEW") > 0 Then
            If MsgBox("変更が破棄されますがよろしいですか?", vbQuestion + vbYesNo, "取引先マスタ登録") = vbNo Then
                Exit Sub
            End If
        End If
        
        '// 検索欄をクリア
        Range(Cells(6, 3), Cells(8, 3)).ClearContents
        
        '// テーブル解除
        On Error Resume Next
        ActiveSheet.ListObjects(1).Unlist
        On Error GoTo 0
        
        '// 表示していた取引先情報をクリア
        Range(Cells(11, 1), Cells(Rows.Count, 10)).Clear
                
    '// 合算グループマスタ・複数回グループマスタからホーム画面に移動する場合
    Else
        If WorksheetFunction.CountIf(Columns(7), True) + WorksheetFunction.CountIf(Columns(7), "NEW") > 0 Then
            If MsgBox("変更が破棄されますがよろしいですか?", vbQuestion + vbYesNo, ActiveSheet.Name & "登録") = vbNo Then
                Exit Sub
            End If
        End If
        
        '// テーブル解除
        On Error Resume Next
        ActiveSheet.ListObjects(1).Unlist
        On Error GoTo 0
    
        '// 表示していた分をクリア
        Range(Cells(11, 3), Cells(Rows.Count, 5)).Clear
        Columns(7).ClearContents
    End If
     
    Dim chkController As New checkBoxController
    
    '// チェックボックス削除
    chkController.deleteChk ActiveSheet
                    
    Set chkController = Nothing
    
    ActiveSheet.Protect
 
Show:
    Sheets("ホーム").Activate

End Sub

'// 「設定」を表示
Public Sub showSetting()

    Sheets("設定").Activate
    

End Sub

'// ファイルを閉じる
Public Sub closeFile()

    If MsgBox("終了してよろしいですか?", vbQuestion + vbYesNo, "山岸運送売掛金回収ファイル") = vbNo Then
        Exit Sub
    End If
    
    ThisWorkbook.Close True

End Sub
