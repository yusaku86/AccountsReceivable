Attribute VB_Name = "severalTimesGroup"
'// 複数回入金グループマスタ登録
Option Explicit

'// 複数回入金グループマスタ一覧表示
Public Sub severalTimesGroupIndex()
    
    '// 変更がある場合は保存しないか確認
    If WorksheetFunction.CountIf(Columns(7), True) + WorksheetFunction.CountIf(Columns(7), "NEW") > 0 Then
        If MsgBox("変更が破棄されますが、よろしいですか?", vbQuestion + vbYesNo, "合算グループマスタ登録") = vbNo Then
            Exit Sub
        End If
    End If
        
    Application.ScreenUpdating = False
    
    With Sheets("複数回入金グループマスタ")
        .Unprotect
        .Range(.Cells(10, 3), .Cells(Rows.Count, 5)).Clear
        .Columns(7).ClearContents
        
        '// ヘッダーの設定
        .Cells(10, 3).value = "id"
        .Cells(10, 4).value = "複数回入金グループ名"
        .Cells(10, 5).value = "口座名義"
        .Rows(10).RowHeight = 50
        
        With .Range(.Cells(10, 3), .Cells(10, 5))
            .Interior.ColorIndex = 14
            .Font.Color = vbWhite
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        
    End With
    
    '// 削除チェックボックスを削除
    Dim chkController As New checkBoxController
    chkController.deleteChk Sheets("複数回入金グループマスタ")
    
    Dim con As ADODB.Connection: Set con = connectDb(ThisWorkbook.Path & "\database\customers.xlsx")

    '// 合算グループマスタ
    Dim severalTimesRs As New ADODB.Recordset
    severalTimesRs.Open "SELECT * FROM [several_times_payment_groups$] ORDER BY id", con, adOpenStatic, adLockOptimistic
    
    '// 一覧表示
    Dim i As Long: i = 11
    
    Do Until severalTimesRs.EOF
        Cells(i, 3).value = severalTimesRs!ID
        Cells(i, 4).value = severalTimesRs!Name
        Cells(i, 5).value = severalTimesRs!account
        
        '// 削除チェックボックス追加
        chkController.add Cells(i, 4), "chk" & severalTimesRs!ID
        
        severalTimesRs.MoveNext
        i = i + 1
    Loop
    
    severalTimesRs.Close
    Set severalTimesRs = Nothing

    
    With Sheets("複数回入金グループマスタ")
    
        '// 表をテーブルに設定
        .ListObjects.add(xlSrcRange, .Range(.Cells(10, 3), .Cells(Rows.Count, 3).End(xlUp).Offset(, 2)), , xlYes, , "TableStyleLight2").Name = "several_times_groups"
    
        '// フォント設定・セルのロックなど
        .Cells.Font.Name = "Meiryo UI"
        .Range(.Cells(11, 4), .Cells(Rows.Count, 5).End(xlUp)).Font.Color = vbBlue
        .Range(.Columns(4), .Columns(5)).HorizontalAlignment = xlCenter
        .Range(.Columns(4), .Columns(5)).EntireColumn.AutoFit
        
        .Columns(3).Hidden = True
        .Columns(7).Hidden = True
        
        .Shapes("btnEdit").Visible = True
        .Shapes("imgEdit").Visible = True
        .Shapes("btnAdd").Visible = True
        .Shapes("imgAdd").Visible = True
        
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        
        .Cells.Locked = True
        .Protect
    End With
    
    Set chkController = Nothing
    
End Sub

'// 編集のためにシートの保護解除
Public Sub unprotectToEditSeveralTimes()

    With Sheets("複数回入金グループマスタ")
        .Unprotect
        .Shapes("btnRegister").Visible = True
        .Shapes("imgRegister").Visible = True
        .Shapes("btnDelete").Visible = True
        .Shapes("imgDelete").Visible = True
        .Range(.Cells(11, 4), .Cells(Rows.Count, 5).End(xlUp)).Font.Color = vbBlack
    End With
    
End Sub

'/**
 '* データが変更されたグループを更新・新規追加
'**/
Public Sub registerSeveralTimesGroups()

    Dim severalTimesRs As New Recordset
    
    severalTimesRs.CursorLocation = adUseClient
    severalTimesRs.Open "SELECT * FROM [several_times_payment_groups$] ORDER BY id", connectDb(ThisWorkbook.Path & "\database\customers.xlsx"), adOpenStatic, adLockOptimistic
    Dim i As Long
    
    '// 値が変更されたレコードのみ更新
    For i = 11 To Sheets("複数回入金グループマスタ").Cells(Rows.Count, 3).End(xlUp).Row
        If Sheets("複数回入金グループマスタ").Cells(i, 7).value <> True And Sheets("複数回入金グループマスタ").Cells(i, 7).value <> "NEW" Then
            GoTo Continue
        End If
        
        '// バリデーション
        If validate(i) = False Then
            Exit Sub
        End If
        
        '// 新規追加
        If Sheets("複数回入金グループマスタ").Cells(i, 7).value = "NEW" Then
            Call addGroup(i, severalTimesRs)
        '// 更新
        Else
            Call updateGroup(i, severalTimesRs)
        End If
        
Continue:
    Next
    
    With Sheets("複数回入金グループマスタ")
        .Unprotect
        .Columns(7).ClearContents
    End With

End Sub

'/**
 '* 新規グループ追加のための行挿入
'**/
Public Sub insertRowForNewSeveralTimesGroup()

    Sheets("複数回入金グループマスタ").Unprotect
    
    If Sheets("複数回入金グループマスタ").Cells(11, 4).value = "" Then: Exit Sub

    '// 下の行の書式を引き継ぐ
    Sheets("複数回入金グループマスタ").Rows(11).Insert copyorigin:=xlFormatFromRightOrBelow
    
    With Sheets("複数回入金グループマスタ")
        .Range("several_times_groups").Font.Color = vbBlack
        .Cells(11, 4).Select
        .Cells(11, 7).value = "NEW"
        
        '// 「登録」ボタン表示
        .Shapes("btnRegister").Visible = True
        .Shapes("imgRegister").Visible = True
    End With
    
End Sub

'/**
 '* 入力した値のバリデーション
'**/
Private Function validate(ByVal targetRow As Long) As Boolean

    validate = False
    
    '// 複数回入金グループ名が入力されているか
    If Cells(targetRow, 4).value = "" Then
        MsgBox "グループ名は必須項目です。", vbExclamation, "複数回入金グループマスタ登録"
        Cells(targetRow, 4).Select
        Exit Function
    End If

    '// 口座名義が入力されているか
    If Cells(targetRow, 5).value = "" Then
        MsgBox "口座名義は必須項目です。", vbExclamation, "複数回入金グループマスタ登録"
        Cells(targetRow, 5).Select
        Exit Function
    End If
    
    '// 口座名義が半角カナか
    Dim reg As New RegController
    If reg.pregMatch(Cells(targetRow, 5).value, "^[ｧ-ﾝﾞﾟ\-\(\)\（\）\.a-zA-Z]+$") = False Then
        MsgBox "口座名義は半角カナ、または半角アルファベットで入力してください。", vbExclamation, "複数回グループマスタ登録"
        Cells(targetRow, 5).Select
        Set reg = Nothing
        Exit Function
    End If

    validate = True

    Set reg = Nothing

End Function

'/**
 '* 新規複数回入金グループを追加する
'**/
Private Sub addGroup(ByVal targetRow As Long, ByVal severalTimesRs As Recordset)

    severalTimesRs.Sort = "id DESC"

    '// 新規合算グループのid
    Dim nextId As Long: nextId = severalTimesRs!ID + 1
    Sheets("複数回入金グループマスタ").Cells(targetRow, 3).value = nextId
    
    severalTimesRs.AddNew
    
    '// 各項目の値を追加
    severalTimesRs!ID = nextId
    severalTimesRs!Name = Sheets("複数回入金グループマスタ").Cells(targetRow, 4).value
    severalTimesRs!account = Sheets("複数回入金グループマスタ").Cells(targetRow, 5).value
        
    '// 削除チェックボックスの追加
    Dim chkController As New checkBoxController
    chkController.add Cells(targetRow, 4), "chk" & Cells(targetRow, 3).value
    
    severalTimesRs.Update
    
End Sub

'/**
 '* データベースの値を更新する
'**/
Private Sub updateGroup(ByVal targetRow As Long, ByVal severalTimesRs As Recordset)
    
    severalTimesRs.filter = "id = " & Sheets("複数回入金グループマスタ").Cells(targetRow, 3).value
    
    severalTimesRs!Name = Sheets("複数回入金グループマスタ").Cells(targetRow, 4).value
    severalTimesRs!account = Sheets("複数回入金グループマスタ").Cells(targetRow, 5).value
    
    severalTimesRs.Update
    
End Sub

'/**
 '* チェックボックスにチェックが入っているグループを削除
'**/
Public Sub deleteSeveralTimesGroups()

    If MsgBox("チェックしたグループを削除しますがよろしいですか?", vbQuestion + vbYesNo, "複数回入金グループマスタ登録") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '// DBとして使用しているエクセルブック(エクセルをDBとして使用するとレコードセットのDeleteメソッドを実行できないため、ブックを開き行を削除する)
    Dim dbBook As Workbook: Set dbBook = Workbooks.Open(ThisWorkbook.Path & "\database\customers.xlsx")
    
    ThisWorkbook.Sheets("複数回入金グループマスタ").Activate
    Dim i As Long
    Dim deleteRow As Long
    
    '// 処理を開始する前の最終行
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 3).End(xlUp).Row
    
    '// チェックボックスにチェックが入っていたらデータ削除 & 行削除
    For i = 11 To Cells(Rows.Count, 3).End(xlUp).Row
        '// 行削除すると最終行の値が変更され、チェックボックスの値が取得できなくなるため、処理を開始する前の最終行 - 削除した行数をiが超えたらループを抜ける
        If i > lastRow Then
            Exit For
        End If
    
        '// 新規グループは登録するまでチェックボックスがないのでとばす
        If Sheets("複数回入金グループマスタ").Cells(i, 3).value = "" Then
            GoTo Continue
        End If
        
        If Sheets("複数回入金グループマスタ").CheckBoxes("chk" & Cells(i, 3).value) = 1 Then
        
            '// DBのデータ削除
            deleteRow = WorksheetFunction.Match(Sheets("複数回入金グループマスタ").Cells(i, 3).value, dbBook.Sheets("several_times_payment_groups").Columns(1), 0)
            dbBook.Sheets("several_times_payment_groups").Rows(deleteRow).Delete
            
            '// チェックボックス削除 & シート「取引先マスタ」の行削除
            Sheets("複数回入金グループマスタ").CheckBoxes("chk" & Cells(i, 3).value).Delete
            Sheets("複数回入金グループマスタ").Rows(i).Delete
            
            i = i - 1
            lastRow = lastRow - 1
        End If
        
Continue:
    Next
        
    dbBook.Close True
    
    Set dbBook = Nothing

End Sub
