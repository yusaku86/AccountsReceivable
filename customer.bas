Attribute VB_Name = "customer"
'// 顧客管理に書かわるモジュール
Option Explicit

'/**
 '* 取引先検索
'**/
Public Sub searchCustomers()
    
    If confirmSave = False Then: Exit Sub
    
    Dim where As String

    '// 取引先コードの検索欄に値がある場合
    If Cells(6, 3).value <> "" Then
        where = " id LIKE '%" & Cells(6, 3).value & "%'"
    End If
    
    '// 取引先名の検索欄に値がある場合
    If Cells(7, 3).value <> "" And where <> "" Then
        where = where & " OR customer_name LIKE '%" & Cells(7, 3).value & "%'"
    ElseIf Cells(7, 3).value <> "" Then
        where = where & " customer_name LIKE '%" & Cells(7, 3).value & "%'"
    End If
    
    '// 口座名義の検索欄に値がある場合 → 入力値のフリガナを半角にしたもので検索
    If Cells(8, 3).value <> "" And where <> "" Then
        where = where & " OR customer_account LIKE '%" & StrConv(Application.GetPhonetic(Cells(8, 3).value), vbNarrow) & "%'"
    ElseIf Cells(8, 3).value <> "" Then
        where = where & " customer_account LIKE '%" & StrConv(Application.GetPhonetic(Cells(8, 3).value), vbNarrow) & "%'"
    End If

    If where <> "" Then
        where = " WHERE" & where
    End If
    
    Call index(where)
End Sub

'/**
 '* 検索条件をリセットして全取引先表示
'**/
Public Sub resetSearchWord()
    
    If confirmSave = False Then: Exit Sub

    Range(Cells(6, 3), Cells(8, 3)).value = ""
    
    Call index

End Sub

'/**
 '* 変更を登録していない取引先がある場合、保存するか確認する
'**/
Private Function confirmSave() As Boolean
    
    If WorksheetFunction.CountIf(Columns(10), True) + WorksheetFunction.CountIf(Columns(10), "NEW") = 0 Then
        confirmSave = True
        Exit Function
    End If
    
    If MsgBox("変更が破棄されますが、よろしいですか?", vbQuestion + vbYesNo, "取引先マスタ登録") = vbYes Then
        confirmSave = True
        Exit Function
    End If
    
    confirmSave = False

End Function

'/**
 '* 取引先一覧表示
 '* @params where 取引先の抽出条件
'**/
 Private Sub index(Optional ByVal where As String = "")
 
    Application.ScreenUpdating = False
 
    Dim con As ADODB.Connection: Set con = connectDb(ThisWorkbook.Path & "\database\customers.xlsx")
    
    '// 取引先レコードセット
    Dim customerRs As New ADODB.Recordset
    customerRs.Open "SELECT * FROM [customers$]" & where & " ORDER BY id", con, adOpenStatic, adLockOptimistic
 
    '// 合算グループレコードセット
    Dim combinedRs As New Recordset
    combinedRs.CursorLocation = adUseClient
    combinedRs.Open "SELECT * FROM [combined_groups$]", con, adOpenStatic, adLockOptimistic
    
    '// 複数回入金グループレコードセット
    Dim severalTimesRs As New Recordset
    severalTimesRs.CursorLocation = adUseClient
    severalTimesRs.Open "SELECT * FROM [several_times_payment_groups$]", con, adOpenStatic, adLockOptimistic
    
    '// シートの保護とセルのロックを解除
    Sheets("取引先マスタ").Unprotect
    Sheets("取引先マスタ").Cells.Locked = False
 
    '// テーブルの設定を解除
    On Error Resume Next
    Sheets("取引先マスタ").ListObjects(1).Unlist
    On Error GoTo 0
 
    '// 前回表示していた分をクリア
    If Sheets("取引先マスタ").Cells(Rows.Count, 2).End(xlUp).Row > 10 Then
        With Sheets("取引先マスタ")
            .Range(.Cells(11, 1), .Cells(Rows.Count, 2).End(xlUp).Offset(, 8)).Clear
        End With
    '// 1件も表示されていない場合 → ヘッダーの1行下の行をクリア(テーブル下部の罫線が残る場合があるため)
    Else
        With Sheets("取引先マスタ")
            .Range(.Cells(11, 1), .Cells(11, 10)).Clear
        End With
    End If
    
    '// 削除チェックボックスを削除
    Dim chkController As New checkBoxController
    chkController.deleteChk Sheets("取引先マスタ")
    
    '// 検索にヒットした取引先が0件だったら抜ける
    If customerRs.RecordCount = 0 Then
        GoTo Break
    End If
    
    Dim i  As Long: i = 11
    
    '// 取引先の情報をセルに入力
    Do Until customerRs.EOF
        With Sheets("取引先マスタ")
            '// 変更前の取引先マスタをA列のセルに入力する
            .Cells(i, 1).value = customerRs!ID
            
            .Cells(i, 2).value = customerRs!ID
            .Cells(i, 3).value = customerRs!customer_name
            .Cells(i, 4).value = customerRs!customer_account
            .Cells(i, 5).value = customerRs!customer_site
            
            '// 相殺の有無
            If customerRs!is_offset = True Then
                .Cells(i, 6).value = "有"
            Else
                .Cells(i, 6).value = "無"
            End If
            
            '// 合算グループの入力
            If customerRs!combined_group <> "" And customerRs!combined_group <> 0 Then
                combinedRs.filter = "id = " & customerRs!combined_group
                .Cells(i, 7).value = customerRs!combined_group & ":" & combinedRs!Name
            End If
            
            '// 複数回入金グループの入力
            If customerRs!several_times_payment_group <> "" And customerRs!several_times_payment_group <> 0 Then
                severalTimesRs.filter = "id = " & customerRs!several_times_payment_group
                .Cells(i, 8).value = customerRs!several_times_payment_group & ":" & severalTimesRs!Name
            End If

            '// 削除チェックボックス追加
            chkController.add .Cells(i, 2), "chk" & customerRs!ID
        
        End With
        i = i + 1
        customerRs.MoveNext
    Loop
    
    Set chkController = Nothing
    
    With Sheets("取引先マスタ")
        '// 入金サイト列にドロップダウン設定
        .Range(.Cells(11, 5), .Cells(Rows.Count, 2).End(xlUp).Offset(, 3)).Validation.add _
            Type:=xlValidateList, Formula1:="翌月,翌々,翌翌々"
    
        '// 相殺の有無列にドロップダウン設定
        .Range(.Cells(11, 6), .Cells(Rows.Count, 2).End(xlUp).Offset(, 4)).Validation.add _
            Type:=xlValidateList, Formula1:="有,無"
    
        '// 合算グループ列にドロップダウン設定
        .Range(.Cells(11, 7), .Cells(Rows.Count, 2).End(xlUp).Offset(, 5)).Validation.add _
            Type:=xlValidateList, Formula1:=createDropDownList(combinedRs)
        
        '// 複数回入金グループ列にドロップダウン設定
        .Range(.Cells(11, 8), .Cells(Rows.Count, 2).End(xlUp).Offset(, 6)).Validation.add _
            Type:=xlValidateList, Formula1:=createDropDownList(severalTimesRs)
    End With
    
Break:
    With Sheets("取引先マスタ")
        ' //「取引先」のシートをテーブルに設定
        .ListObjects.add(xlSrcRange, .Range(.Cells(10, 2), .Cells(Rows.Count, 2).End(xlUp).Offset(, 6)), , xlYes, , "TableStyleLight1").Name = "customers"
        
        .Range(.Cells(11, 5), .Cells(Rows.Count, 2).End(xlUp).Offset(, 4)).HorizontalAlignment = xlCenter
        
        '// 取引先コード列のIMEモードを半角英数字に変更
        With .Range(.Cells(11, 2), .Cells(Rows.Count, 2).End(xlUp)).Validation
            .Delete
            .add Type:=xlValidateInputOnly
            .IMEMode = xlIMEModeAlpha
        End With
        
        '// 取引先名列・口座名義のIMEモードを日本語入力に変更(取引先コードのセルから移動したときに日本語入力になるように)
        With .Range(.Cells(11, 3), .Cells(Rows.Count, 4).End(xlUp)).Validation
            .Delete
            .add Type:=xlValidateInputOnly
            .IMEMode = xlIMEModeOn
        End With
        
        .Cells.Locked = True
        '// 検索欄のロック解除
        .Range(.Cells(6, 3), Cells(8, 3)).Locked = False
        
        .Range("customers").Font.Color = vbBlue
        .Cells.Font.Name = "Meiryo UI"
        
        '// 取引先コートが変更されたかを確認するA列とデータの値が変更されたかを確認するJ列を非表示
        .Columns(1).Hidden = True
        .Columns(10).Hidden = True
        
        '// 「リセット」・「編集」・「新規追加」ボタンを使用可能にする
        .Shapes("btnReset").Visible = True
        .Shapes("imgReset").Visible = True
        .Shapes("btnEdit").Visible = True
        .Shapes("imgEdit").Visible = True
        .Shapes("btnAdd").Visible = True
        .Shapes("imgAdd").Visible = True
        
        '// 「登録」・「削除」ボタンを.使用不可にする
        .Shapes("btnRegister").Visible = False
        .Shapes("imgRegister").Visible = False
        .Shapes("btnDelete").Visible = False
        .Shapes("imgDelete").Visible = False
        .Protect
        
    End With

    customerRs.Close
    combinedRs.Close
    severalTimesRs.Close
    
    Set customerRs = Nothing
    Set combinedRs = Nothing
    Set severalTimesRs = Nothing

    Set con = Nothing
    
 End Sub
 
'/**
 '* 合算グループ・複数回入金グループのドロップダウンリスト用の文字列作成
 '* 「id:グループ名」の形でドロップダウンに表示する
'**/
Private Function createDropDownList(ByVal rs As Recordset) As String

    rs.filter = adFilterNone
    rs.Sort = "id ASC"

    Dim dropDownList As String

    Do Until rs.EOF = True
        If dropDownList = "" Then
            dropDownList = rs!ID & ":" & rs!Name
        Else
            dropDownList = dropDownList & "," & rs!ID & ":" & rs!Name
        End If
        
        rs.MoveNext
    Loop
    
    createDropDownList = dropDownList

End Function

'/**
 '* 新規取引先追加のための行挿入
'**/
Public Sub insertRowForNewCustomer()

    Sheets("取引先マスタ").Unprotect
    
    If Sheets("取引先マスタ").Cells(11, 2).value = "" Then: Exit Sub
    
    '// 下の行の書式を引き継ぐ
    Sheets("取引先マスタ").Rows(11).Insert copyorigin:=xlFormatFromRightOrBelow
    
    With Sheets("取引先マスタ")
        .Range("customers").Font.Color = vbBlack
        .Cells(11, 2).Select
        .Cells(11, 10).value = "NEW"
        
        '// 「登録」ボタン表示
        .Shapes("btnRegister").Visible = True
        .Shapes("imgRegister").Visible = True
    End With
    
End Sub

'/**
 '* 編集のためにセルのロックを解除する
'**/
Public Sub unProtectToEditCustomer()

    With Sheets("取引先マスタ")
        .Unprotect
        
        On Error Resume Next
        .Range("customers").Font.Color = vbBlack
        On Error GoTo 0
        
        '// 「登録」・「削除」ボタン表示
        .Shapes("btnRegister").Visible = True
        .Shapes("imgRegister").Visible = True
        .Shapes("btnDelete").Visible = True
        .Shapes("imgDelete").Visible = True
    End With
    
End Sub

'/**
 '* データが変更された取引先のを更新・新規追加
'**/
Public Sub registerCustomers()

    Application.ScreenUpdating = False

    Dim customerRs As New Recordset
    
    customerRs.Open "SELECT * FROM [customers$] ORDER BY id", connectDb(ThisWorkbook.Path & "\database\customers.xlsx"), adOpenStatic, adLockOptimistic
    Dim i As Long
    
    '// 値が変更されたレコードのみ更新
    For i = 11 To Sheets("取引先マスタ").Cells(Rows.Count, 1).End(xlUp).Row
        If Sheets("取引先マスタ").Cells(i, 10).value <> True And Sheets("取引先マスタ").Cells(i, 10).value <> "NEW" Then
            GoTo Continue
        End If
        
        '// バリデーション
        If validate(i, customerRs) = False Then
            Exit Sub
        End If
        
        '// 新規追加
        If Sheets("取引先マスタ").Cells(i, 10).value = "NEW" Then
            Call addCustomer(i, customerRs)
        '// 更新
        Else
            Call updateCustomer(i, customerRs)
        End If
        
Continue:
    Next
    
    customerRs.Close
    Set customerRs = Nothing
    
    With Sheets("取引先マスタ")
        .Unprotect
        .Columns(10).ClearContents
    End With
    
    '// 文字列として保存されるデータを数値化
    Dim dbBook As Workbook: Set dbBook = Workbooks.Open(ThisWorkbook.Path & "\database\customers.xlsx")
    
    With dbBook.Sheets("customers")
        convertStr2Number .Range(.Cells(2, 1), .Cells(Rows.Count, 1).End(xlUp))
        convertStr2Number .Range(.Cells(2, 6), .Cells(Rows.Count, 1).End(xlUp).Offset(, 6))
    End With
    
    dbBook.Close True
    
    Set dbBook = Nothing
    
End Sub

'/**
 '* 入力した値のバリデーション
'**/
Private Function validate(ByVal targetRow As Long, ByVal customerRs As Recordset) As Boolean

    validate = False

    '// 取引先コードが入力されているか
    If Cells(targetRow, 2).value = "" Then
        MsgBox "取引先コードを入力してください。", vbExclamation, "取引先マスタ登録"
        Cells(targetRow, 2).Select
        Exit Function
    '// 取引先コードが数字か
    ElseIf IsNumeric(Cells(targetRow, 2).value) = False Then
        MsgBox "取引先コードには数字を入力してください。", vbExclamation, "取引先マスタ登録"
        Cells(targetRow, 2).Select
        Exit Function
    End If
    
    '// 取引先名が入力されているか
    If Cells(targetRow, 3).value = "" Then
        MsgBox "取引先名は必須項目です。", vbExclamation, "取引先マスタ登録"
        Cells(targetRow, 3).Select
        Exit Function
    End If
    
    '// 口座名義が入力されているか
    If Cells(targetRow, 4).value = "" Then
        If MsgBox("口座名義が入力されていませんが、よろしいですか?", vbQuestion + vbYesNo, "取引先マスタ登録") = vbNo Then
            Cells(targetRow, 4).Select
            Exit Function
        End If
    End If
    
    '// 口座名義が半角カナか
    Dim reg As New RegController
    If Cells(targetRow, 4).value <> "" And reg.pregMatch(Cells(targetRow, 4).value, "^[ｧ-ﾝﾞﾟ\-\(\)\（\）\.a-zA-Z]+$") = False Then
        MsgBox "口座名義は半角カタカナ、または半角アルファベットで入力してください。", vbExclamation + vbYesNo, "合算グループマスタ"
        Cells(targetRow, 4).Select
        Set reg = Nothing
        Exit Function
    End If
    
    '// 入金サイトが入力されているか
    If Cells(targetRow, 5).value = "" Then
        MsgBox "入金サイトは必須項目です。", vbExclamation, "取引先マスタ登録"
        Cells(targetRow, 5).Select
        Exit Function
    End If
    
    '// 取引先コードが変更された場合 → 変更後の取引先コードでフィルターをかけ、既に使用されている場合は処理を中断する
    If Cells(targetRow, 1).value <> Cells(targetRow, 2).value Then
        customerRs.filter = "id = " & Cells(targetRow, 2).value
    
        If customerRs.RecordCount > 0 Then
            MsgBox "取引先コード " & Cells(targetRow, 2).value & " は既に使用されています。", vbExclamation, "取引先マスタ登録"
            Cells(targetRow, 2).Select
            Exit Function
        End If
    End If

    validate = True
    
    Set reg = Nothing

End Function

'/**
 '* 新規取引先を追加する
'**/
Private Sub addCustomer(ByVal rowNumber As Long, ByVal customerRs As Recordset)

    customerRs.AddNew
    
    '// 比較用取引先コードをA列のセルに入力
    Sheets("取引先マスタ").Cells(rowNumber, 1).value = Sheets("取引先マスタ").Cells(rowNumber, 2).value
    
    '// 各項目の値を追加
    customerRs!ID = Sheets("取引先マスタ").Cells(rowNumber, 2).value
    customerRs!customer_name = Sheets("取引先マスタ").Cells(rowNumber, 3).value
    customerRs!customer_account = Sheets("取引先マスタ").Cells(rowNumber, 4).value
    customerRs!customer_site = Sheets("取引先マスタ").Cells(rowNumber, 5).value
    
    '// 合算グループが入力されている場合
    If Sheets("取引先マスタ").Cells(rowNumber, 7).value <> "" Then
        customerRs!combined_group = Split(Sheets("取引先マスタ").Cells(rowNumber, 7).value, ":")(0)
    End If
    
    '// 複数回入金グループが入力されている場合
    If Sheets("取引先マスタ").Cells(rowNumber, 8).value <> "" Then
        customerRs!several_times_payment_group = Split(Sheets("取引先マスタ").Cells(rowNumber, 2).value, ":")(0)
    End If
    
    '// 削除チェックボックスの追加
    Dim chkController As New checkBoxController
    chkController.add Cells(rowNumber, 2), "chk" & Cells(rowNumber, 2).value
    
    customerRs.Update
    
End Sub

'/**
 '* データベースの値を更新する
'**/
Public Sub updateCustomer(ByVal rowNumber As Long, ByVal customerRs As Recordset)
    
    customerRs.filter = "id = " & Sheets("取引先マスタ").Cells(rowNumber, 1).value
    
    customerRs!ID = Sheets("取引先マスタ").Cells(rowNumber, 2).value
    customerRs!customer_name = Sheets("取引先マスタ").Cells(rowNumber, 3).value
    customerRs!customer_account = Sheets("取引先マスタ").Cells(rowNumber, 4).value
    customerRs!customer_site = Sheets("取引先マスタ").Cells(rowNumber, 5).value
    customerRs!is_offset = Sheets("取引先マスタ").Cells(rowNumber, 6).value = "有"
    
    
    If Sheets("取引先マスタ").Cells(rowNumber, 7).value <> "" Then
        customerRs!combined_group = Split(Sheets("取引先マスタ").Cells(rowNumber, 7).value, ":")(0)
    Else
        customerRs!combined_group = 0
    End If
    
    If Sheets("取引先マスタ").Cells(rowNumber, 8).value <> "" Then
        customerRs!several_times_payment_group = Split(Sheets("取引先マスタ").Cells(rowNumber, 8).value, ":")(0)
    Else
        customerRs!several_times_payment_group = 0
    End If
    
    customerRs.Update
    
    '// チェックボックスの名前更新
    Sheets("取引先マスタ").CheckBoxes("chk" & Cells(rowNumber, 1).value).Name = "chk" & Cells(rowNumber, 2).value
    
    '// 取引先コードが変更されたかを確認するためのセルの値を変更
    Cells(rowNumber, 1).value = Cells(rowNumber, 2).value

End Sub

'/**
 '* チェックボックスにチェックが入っている取引先を削除
'**/
Public Sub deleteCustomers()

    If MsgBox("チェックボックスにチェックが入った取引先を削除しますがよろしいですか?", vbQuestion + vbYesNo, "取引先マスタ登録") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    '// DBとして使用しているエクセルブック(エクセルをDBとして使用するとレコードセットのDeleteメソッドを実行できないため、ブックを開き行を削除する)
    Dim dbBook As Workbook: Set dbBook = Workbooks.Open(ThisWorkbook.Path & "\database\customers.xlsx")
    
    ThisWorkbook.Sheets("取引先マスタ").Activate
    Dim i As Long
    Dim deleteRow As Long
    
    '// 処理を開始する前の最終行
    Dim lastRow As Long: lastRow = Cells(Rows.Count, 2).End(xlUp).Row
    
    '// チェックボックスにチェックが入っていたらデータ削除 & 行削除
    For i = 11 To Cells(Rows.Count, 2).End(xlUp).Row
        '// 行削除すると最終行の値が変更され、チェックボックスの値が取得できなくなるため、処理を開始する前の最終行 - 削除した行数をiが超えたらループを抜ける
        If i > lastRow Then
            Exit For
        End If
            
        '// 新規取引先は登録するまでチェックボックスがないのでとばす
        If Sheets("取引先マスタ").Cells(i, 1).value = "" Then
            GoTo Continue
        End If
        
        If Sheets("取引先マスタ").CheckBoxes("chk" & Cells(i, 1).value) = 1 Then
            
            '// DBのデータ削除
            deleteRow = WorksheetFunction.Match(Sheets("取引先マスタ").Cells(i, 1).value, dbBook.Sheets("customers").Columns(1), 0)
            dbBook.Sheets("customers").Rows(deleteRow).Delete
            
            '// チェックボックス削除 & シート「取引先マスタ」の行削除
            Sheets("取引先マスタ").CheckBoxes("chk" & Cells(i, 1).value).Delete
            Sheets("取引先マスタ").Rows(i).Delete
            
            i = i - 1
            lastRow = lastRow - 1
        End If

Continue:
    Next
        
    dbBook.Close True
    
    Set dbBook = Nothing

End Sub
