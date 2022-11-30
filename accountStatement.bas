Attribute VB_Name = "accountStatement"
'// 銀行明細関連(売掛金回収など)
Option Explicit

'// 入金明細の年を選択するフォーム起動
Public Sub openFormYear()

    Sheets("mode").Cells(1, 1).value = "IMPORT_STATEMENT"

    formYear.Show
    
End Sub

'// 三井住友の振込明細を加工してシート「銀行明細」へ貼り付け(メインプロシージャ)
Public Sub putBankStatement(ByVal targetYear As Long)
    
    Dim wsh As New WshShell
    
    '// ダイアログを表示して取り込むファイルを選択 selectFile [ダイアログタイトル], [初期表示フォルダ], [拡張子絞り込みメッセージ], [絞り込む拡張子]
    Dim fileName As String: fileName = selectFile("銀行明細取込", wsh.SpecialFolders(4), "CSVファイル", "*.csv")
    
    Set wsh = Nothing
    
    If fileName = "" Then: Exit Sub
    
    '// ヘッダーの設定
    Call createTableHeader
    
    '// 指定したファイルが正しい形式か確認
    If checkFile(fileName) = False Then
        Exit Sub
    End If
    
    Dim fso As New FileSystemObject
    fileName = fso.GetFileName(fileName)
    
    Set fso = Nothing
    
    Dim bankCsv As Workbook: Set bankCsv = Workbooks(fileName)

    With bankCsv.Sheets(1)
        
        Dim lastRow As Long: lastRow = .Cells(Rows.Count, 3).End(xlUp).Row - 1
    
        '// 日付を「00月00日」に変換
        .Cells(2, 3).Formula = "=MID(D2," & Len(.Cells(2, 4).value) - 3 & ",2) & ""月"" & TEXT(RIGHT(D2,2),0) & ""日"""
        .Cells(2, 3).AutoFill .Range(Cells(2, 3), Cells(lastRow, 3))
        
        '// 日付列をコピーして貼り付け
        .Range(Cells(2, 3), Cells(lastRow, 3)).Copy
        ThisWorkbook.Sheets("銀行明細").Cells(6, 2).PasteSpecial xlPasteValues
        
        '// 口座名義をコピーして貼り付け
        .Range(Cells(2, 8), Cells(lastRow, 8)).Copy
        ThisWorkbook.Sheets("銀行明細").Cells(6, 5).PasteSpecial xlPasteValues
        
        '// 金額をコピーして貼り付け
        .Range(Cells(2, 5), Cells(lastRow, 5)).Copy
        ThisWorkbook.Sheets("銀行明細").Cells(6, 6).PasteSpecial xlPasteValues
    End With
    
    bankCsv.Close False
    Set bankCsv = Nothing
        
    With ThisWorkbook.Sheets("銀行明細")
    
        ' //「銀行明細」のシートをテーブルに設定
        With .ListObjects.add(xlSrcRange, .Range(.Cells(5, 1), .Cells(.Cells(Rows.Count, 2).End(xlUp).Row, 9)), , xlYes, , "TableStyleLight14")
            .Name = "account_statements"
        End With
    End With
    
    '// 口座名義から取引先コードと取引先名を取得
    Call fetchCodeAndName
    
    '// 手数料の計算
    '// calculateHandlingCharge [入金年]
    Call calculateHandlingCharge(targetYear)
    
    '// 要確認フラグ・再計算フラグの設定
    Call setFlags
    
    With Sheets("銀行明細")
    
        '// 日付・取引先名・取引先コード・備考列の文字の大きさ設定
        .Range(.Cells(6, 2), .Cells(Rows.Count, 2).End(xlUp)).Font.Size = 10
        .Range(.Cells(6, 4), .Cells(Rows.Count, 5).End(xlUp)).Font.Size = 10
        .Range(.Cells(6, 8), .Cells(Rows.Count, 8).End(xlUp)).Font.Size = 9
        
        .Range(.Cells(6, 2), .Cells(Rows.Count, 2).End(xlUp)).HorizontalAlignment = xlCenter
        .Range(.Cells(6, 9), .Cells(Rows.Count, 1).End(xlUp).Offset(, 8)).HorizontalAlignment = xlCenter
        
        .Range(.Columns(2), .Columns(8)).EntireColumn.AutoFit
        
        .Columns(11).Hidden = True
        .Cells(6, 1).Select
    End With
    
    Application.CutCopyMode = False
    
End Sub

'/**
 '* ヘッダーの設定
'**/
Private Sub createTableHeader()

    With Sheets("銀行明細")
        .Range(.Cells(5, 1), .Cells(Rows.Count, 11)).Clear
        .Cells.Font.Name = "Meiryo UI"
        .Columns(1).HorizontalAlignment = xlCenter
        
        '// ヘッダーの内容
        .Cells(5, 1).value = "要確認"
        .Cells(5, 2).value = "日付"
        .Cells(5, 3).value = "取引先コード"
        .Cells(5, 4).value = "取引先名"
        .Cells(5, 5).value = "口座名義"
        .Cells(5, 6).value = "振込金額"
        .Cells(5, 7).value = "売掛金との差額"
        .Cells(5, 8).value = "備考"
        .Cells(5, 9).value = "再計算"
        '// 合算入金かどうかを示す列
        .Cells(5, 11).value = "combined"
        .Range(.Columns(6), .Columns(7)).NumberFormatLocal = "#,##0;[赤]-#,##0"
        
        '// ヘッダーの色など
        With .Range(.Cells(5, 1), .Cells(5, 9))
            .Interior.ColorIndex = 50
            .Font.Color = vbWhite
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        .Rows(5).RowHeight = 50
    
        '// ヘッダーの説明
        With .Cells(5, 1).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "こちらは確認が必要かを表します。" & vbLf & vbLf & "値が1のとき ！ を表示します。ダブルクリックで切り替え可能です。"
        End With
        
        With .Cells(5, 2).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "こちらは入金のあった日付を表します。"
        End With
        
        With .Cells(5, 3).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "こちらは入金のあった取引先コードを表します。" & vbLf & vbLf & "取引先が登録されていない場合、「取引先が見つかりません」と表示します。"
        End With
        
        With .Cells(5, 4).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "こちらは入金のあった取引先名を表します。" & vbLf & vbLf & "取引先が登録されていない場合、「取引先が見つかりません」と表示します。"
        End With
        
        With .Cells(5, 5).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "こちらは入金のあった口座名義を表します。"
        End With
        
        With .Cells(5, 6).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "こちらは振り込まれた金額を表します。"
        End With
        
        With .Cells(5, 7).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "こちらは売掛金と入金額の差額を表します。" & vbLf & vbLf & "複雑な入金には対応できていないため、手動で修正をお願いしますm(__)m"
        End With
        
        With .Cells(5, 8).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "こちらは備考を表します。" & vbLf & vbLf & "合算入金や複数回入金の場合などに使用されます。"
        End With
    
        With .Cells(5, 9).Validation
            .add Type:=xlValidateInputOnly
            .InputMessage = "こちらは再計算を行うかを表します。" & vbLf & vbLf & "ダブルクリックで切り替え可能です。"
        End With
        
    End With

End Sub

'/**
 '* 取込前にファイルの形式が正しいかを判断
'**/
Private Function checkFile(ByVal fileName As String) As Boolean

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(fileName)
    
    With targetFile.Sheets(1)
        If .Cells(1, 8).value <> "ﾐﾂｲｽﾐﾄﾓ" Then
            targetFile.Close False
            MsgBox "指定したファイルの形式が正しくありません。" & vbLf & "三井住友銀行の振込入金明細を指定してください。", vbExclamation, "入金明細取込"
            checkFile = False
            Exit Function
        End If
    End With
    
    checkFile = True

End Function

'// 口座名義をもとにデータベースから取引先コードを取引先名を取得
Public Sub fetchCodeAndName()

    Workbooks.Open ThisWorkbook.Path & "\database\customers.xlsx"
    
    With ThisWorkbook.Sheets("銀行明細")

        '// 取引先コード列・取引先名をXLOOKUPでデータベースファイルから取得
        .Cells(6, 3).Formula = "=XLOOKUP([@口座名義],[customers.xlsx]customers!$C:$C,[customers.xlsx]customers!$A:$A,""取引先が見つかりません"",0)"
        .Cells(6, 4).Formula = "=XLOOKUP([@口座名義],[customers.xlsx]customers!$C:$C,[customers.xlsx]customers!$B:$B,""取引先が見つかりません"",0)"
        
        .Cells(6, 3).AutoFill .Range(.Cells(6, 3), .Cells(Rows.Count, 2).End(xlUp).Offset(, 1)), xlFillValues
        .Cells(6, 4).AutoFill .Range(.Cells(6, 4), .Cells(Rows.Count, 2).End(xlUp).Offset(, 2)), xlFillValues
        
        '// 式を値に変換
        .Range(.Cells(6, 3), .Cells(Rows.Count, 3).End(xlUp)).Copy
        .Cells(6, 3).PasteSpecial xlPasteValues
        .Range(.Cells(6, 4), .Cells(Rows.Count, 4).End(xlUp)).Copy
        .Cells(6, 4).PasteSpecial xlPasteValues
    
    End With
    
    Application.CutCopyMode = False
    
    Workbooks("customers.xlsx").Close False

End Sub

'/**
 '* 要確認フラグ・再計算フラグの設定
'**/
Private Sub setFlags()

    With Sheets("銀行明細")
    
        '// 要確認列の内容設定(差額がマイナスの場合と1000円を超える場合に1にする)
        .Cells(6, 1).Formula = "=IFERROR(IF(OR([@売掛金との差額]<0,[@売掛金との差額]>1000,[@売掛金との差額] =""""),1,0),0)"
        .Range(.Cells(6, 1), .Cells(Rows.Count, 1).End(xlUp)).Copy
        .Cells(6, 1).PasteSpecial xlPasteValues
        
        '// 要確認フラグ列にドロップダウン設定(0か1)
        .Range(.Cells(6, 1), .Cells(Rows.Count, 1).End(xlUp)).Validation.add _
            Type:=xlValidateList, Formula1:="0,1"
        
        '// 要確認フラグのアイコン設定
        Dim confirmFlag As IconSetCondition
        Set confirmFlag = .Range(.Cells(6, 1), .Cells(Rows.Count, 1).End(xlUp)).FormatConditions.AddIconSetCondition
        
        With confirmFlag
            .IconSet = ThisWorkbook.IconSets(xl3Symbols2)
            .ShowIconOnly = True
        
            .IconCriteria(1).Icon = xlIconNoCellIcon
            
            '// 要確認フラグの値が1の場合は黄色いびっくりマークを表示する
            With .IconCriteria(2)
                .Type = xlConditionValueNumber
                .value = 1
                .Operator = xlGreaterEqual
            End With
            
            With .IconCriteria(3)
                .Icon = xlIconNoCellIcon
                .Type = xlConditionValueNumber
                .value = 1
                .Operator = xlGreater
            End With
            
        End With
        
        '// 要確認フラグが立っている行を薄い赤で色付けするようルールを設定
        With .Range(.Cells(6, 1), .Cells(Rows.Count, 1).Offset(, 8)).FormatConditions.add( _
            Type:=xlExpression, _
            Formula1:="=$A6 = 1" _
        )
            .Interior.ColorIndex = 40
            .StopIfTrue = False
        End With
    
        '// 再計算フラグ列にドロップダウン設定(0か1)
        .Range(.Cells(6, 9), .Cells(Rows.Count, 1).End(xlUp).Offset(, 8)).Validation.add _
            Type:=xlValidateList, Formula1:="0,1"
            
        '// 再計算フラグのアイコン設定
        Dim recalcFlag As IconSetCondition
        Set recalcFlag = .Range(.Cells(6, 9), .Cells(Rows.Count, 1).End(xlUp).Offset(, 8)).FormatConditions.AddIconSetCondition
        
        With recalcFlag
            .IconSet = ThisWorkbook.IconSets(xl3Flags)
            .ShowIconOnly = True
            
            .IconCriteria(1).Icon = xlIconNoCellIcon
            
            '// 再計算フラグの値が1の場合は緑の旗を表示
            With .IconCriteria(2)
                .Icon = xlIconGreenFlag
                .Type = xlConditionValueNumber
                .value = 1
                .Operator = xlGreaterEqual
            End With
            
            With .IconCriteria(3)
                .Icon = xlIconNoCellIcon
                .Type = xlConditionValueNumber
                .value = 1
                .Operator = xlGreater
            End With
        End With
        
    End With
    
    Set confirmFlag = Nothing
    Set recalcFlag = Nothing
    
End Sub

'/**
 '* 手数料再計算
'**/
Public Sub openFormToRecalculate()
    
    Sheets("mode").Cells(1, 1).value = "RECALCULATE"
    
    formYear.Show vbModeless

End Sub

'/**
 '* 手数料計算
 '*
'**/
Public Sub calculateHandlingCharge(ByVal paymentYear As Long, Optional ByVal isRecalc As Boolean = False)

    '// 取引先関連のデータベース接続
    Dim customerCon As ADODB.Connection: Set customerCon = connectDb(ThisWorkbook.Path & "\database\customers.xlsx")
    
    '// 売掛金のデータベース接続(今年と昨年)
    Dim salesConThisYear As ADODB.Connection: Set salesConThisYear = connectDb(ThisWorkbook.Path & "\database\sales\" & paymentYear & ".xlsx")
    Dim salesConLastYear As ADODB.Connection: Set salesConLastYear = connectDb(ThisWorkbook.Path & "\database\sales\" & paymentYear - 1 & ".xlsx")
    
    '// 取引先レコードセット
    Dim customerRs As New ADODB.Recordset
    customerRs.Open "SELECT * FROM [customers$]", customerCon, adOpenStatic, adLockOptimistic
    
    '// 売上レコードセット(今年と昨年)
    Dim salesRsThisyear As New ADODB.Recordset
    salesRsThisyear.Open "SELECT * FROM [sales$]", salesConThisYear, adOpenStatic, adLockOptimistic
    
    Dim salesRsLastYear As New ADODB.Recordset
    salesRsLastYear.Open "SELECT * FROM [sales$]", salesConLastYear, adOpenStatic, adLockOptimistic
    
    '// 合算グループレコードセット
    Dim groupRs As New Recordset
    groupRs.Open "SELECT * FROM [combined_groups$]", customerCon, adOpenStatic, adLockOptimistic
    
    '// 銀行明細レコードセット
    Dim bankRs As New Recordset
    bankRs.CursorLocation = adUseClient
    bankRs.Open "SELECT * FROM [銀行明細$A5:H]", connectDb(ThisWorkbook.FullName), adOpenStatic, adLockOptimistic
  
    '/**
     '* 売掛金と入金額の差額計算
    '**/
    Dim rowNumber As Long: rowNumber = 6
    Dim customerId As Long
    Dim paymentMonth As Long
      
    Do Until Sheets("銀行明細").Cells(rowNumber, 2).value = ""
        '// 取引先が見つからない場合、備考に複数回入金とあるものはとばす
        If Sheets("銀行明細").Cells(rowNumber, 3).value = "取引先が見つかりません" Or InStr(1, Sheets("銀行明細").Cells(rowNumber, 8).value, "複数回入金") > 0 Then
            GoTo Continue
        '// 再計算の場合は、再計算フラグが立っていない場合、合算の場合はとばす
        ElseIf isRecalc = True And Sheets("銀行明細").Cells(rowNumber, 9).value <> 1 Then
            GoTo Continue
        End If
        
        customerId = Sheets("銀行明細").Cells(rowNumber, 3).value
        customerRs.filter = "id = " & customerId
        
        '// 相殺ありの取引先の場合 → 備考に入力
        If customerRs!is_offset Then
            bankRs.filter = "取引先コード = " & customerId
            bankRs!備考 = "※「相殺有」の取引先です。"
            bankRs.Update
            bankRs.filter = adFilterNone
        End If
        
        paymentMonth = Split(Sheets("銀行明細").Cells(rowNumber, 2).value, "月")(0)
    
        '/**
         '* 合算の場合の処理
        '**/
        If customerRs!combined_group <> 0 Then
            '// 合算グループIDで取引先を絞る
            Dim combinedGroupId As Long: combinedGroupId = customerRs!combined_group
            customerRs.filter = "combined_group = " & combinedGroupId
            
            '// 売掛金と入金額の差額
            Dim combinedDiff As Long
            combinedDiff = calculateTotalSales(salesRsThisyear, salesRsLastYear, customerRs, paymentMonth) - Sheets("銀行明細").Cells(rowNumber, 6).value
                        
            '// 合算グループに登録されている口座名義が同じグループがある場合 → 全ての合算グループで売掛金と入金額の差額を求め、
            '// 最も差額が小さいグループを採用
            groupRs.filter = "account = " & "'" & customerRs!customer_account & "'"
            
            If groupRs.RecordCount >= 2 Then
                Dim tmpDiff As Long
                
                Do Until groupRs.EOF
                    customerRs.filter = "combined_group = " & groupRs!ID
                    tmpDiff = calculateTotalSales(salesRsThisyear, salesRsLastYear, customerRs, paymentMonth) - Sheets("銀行明細").Cells(rowNumber, 6).value
                    
                    '// 売掛金と入金額の差額が妥当な方を採用
                    combinedDiff = compareNumbersAsCommision(combinedDiff, tmpDiff)
                    
                    If combinedDiff = tmpDiff Then
                        combinedGroupId = groupRs!ID
                    End If
                    
                    groupRs.MoveNext
                Loop
            End If
        
            '// 合算処理で挿入された行数
            Dim insertedCount As Long
            
            '// calculateCombindedPayment [売掛金と入金額の差額], [入金月], [合算グループID],[売上レコードセット(今年)],[売上レコードセット(昨年)],[取引先レコードセット],[処理中の行番号]
            With Sheets("銀行明細")
                insertedCount = calculateCombinedPayment( _
                    combinedDiff, _
                    paymentMonth, _
                    combinedGroupId, _
                    salesRsThisyear, _
                    salesRsLastYear, _
                    customerRs, _
                    rowNumber _
                )
                rowNumber = rowNumber + insertedCount
            End With
            
            GoTo Continue
        End If
        
        '// 複数回入金のある取引先の場合
        If WorksheetFunction.CountIf(Sheets("銀行明細").Columns(4), Sheets("銀行明細").Cells(rowNumber, 4).value) > 1 Then
            Dim totalSales As Long
            Dim targetMonth As Long: targetMonth = getTargetMonth(customerRs!customer_site, paymentMonth)
            
            '// 対象月が昨年の場合 → 昨年の売掛金レコードセットを使用
            Dim salesRs As Recordset
            If targetMonth < 0 Then
                Set salesRs = salesRsLastYear
            Else
                Set salesRs = salesRsThisyear
            End If
            
            '// 同一口座から異なる取引先の売掛金が複数回に分かれて入金される場合
            If customerRs!several_times_payment_group <> 0 Then
                customerRs.filter = "several_times_payment_group = " & customerRs!several_times_payment_group
                totalSales = calculateTotalSales(salesRsThisyear, salesRsLastYear, customerRs, paymentMonth)
            
            '// 同一口座から1つの取引先の売掛金が複数回に分かれて入金される場合に対象月が昨年の場合
            Else
                salesRs.filter = "customer_id = " & customerId & " AND sales_month = " & Abs(targetMonth)
                totalSales = salesRs!sales
            End If
            
            '// calculateSeveralTimePayment [口座名義], [売掛金], [売掛金レコードセット], [取引先レコードセット], [銀行明細レコードセット], [入金月]
            Call calculateSeveralTimePayment( _
                customerRs!customer_account, totalSales, salesRs, customerRs, bankRs, paymentMonth _
            )
            GoTo Continue
        End If
        
        '// 上記以外の入金
        With Sheets("銀行明細")
            '// calculateDifference [取引先コード], [入金額], [入金月], [入金サイト], [売上レコードセット]
            .Cells(rowNumber, 7).value = _
                calculateDifference(customerId, .Cells(rowNumber, 6).value, paymentMonth, customerRs!customer_site, salesRsThisyear, salesRsLastYear)
        End With
        
        customerRs.filter = adFilterNone
Continue:
        rowNumber = rowNumber + 1
    Loop
    
    salesRsThisyear.Close
    salesRsLastYear.Close
    customerRs.Close
    groupRs.Close
    
    Set salesRsThisyear = Nothing
    Set salesRsLastYear = Nothing
    Set customerRs = Nothing
    Set groupRs = Nothing

End Sub

'/**
 '* 合算入金の場合の手数料などを求める
 '* @return 何行挿入したか
 '* @params amountDiff       売掛金と入金額の差額
 '* @params paymentMonth     入金月
 '* @params combinedGroupId  合算グループID
 '* @params salesRsThisYear  売上レコードセット(今年)
 '* @params salesRsLastYear  売上レコードセット(昨年)
 '* @pamras customerRs       取引先レコードセット
 '* @params sheetRow         シート「銀行明細」の現在処理中の行番号
'**/
Private Function calculateCombinedPayment(ByVal amountDiff As Long, ByVal paymentMonth As Long, ByVal combinedGroupId As Long, ByVal salesRsThisyear As Recordset, ByVal salesRsLastYear As Recordset, ByVal customerRs As Recordset, ByVal sheetRow As Long) As Long

    Dim totalPayment As Long: totalPayment = Sheets("銀行明細").Cells(sheetRow, 6).value

    customerRs.filter = "combined_group = " & combinedGroupId
        
    Sheets("銀行明細").Cells(sheetRow, 7).value = amountDiff
    customerRs.MoveFirst
    
    Dim insertedCount As Long
    Dim counter As Long: counter = -1
    Dim salesRs As ADODB.Recordset
    
    '// 合算グループに登録されている分の会社名などのデータを銀行明細に入力
    Do Until customerRs.EOF
                
        '// 取引先コードと対象月で売上を抽出 getTargetMonth [サイト],[入金月]
        Dim targetMonth As Long: targetMonth = getTargetMonth(customerRs!customer_site, paymentMonth)
        
        '// 対象月が昨年の場合(値がマイナスの場合) → 昨年の売掛金レコードセットを使用
        If targetMonth < 0 Then
            Set salesRs = salesRsLastYear
        Else
            Set salesRs = salesRsThisyear
        End If
        
        salesRs.filter = "customer_id = " & customerRs!ID & " AND sales_month = " & Abs(targetMonth)
        
        If salesRs.RecordCount = 0 Then: GoTo Continue
    
        counter = counter + 1
        
        '// 合算グループで売上が0円以上の最初の処理の場合 → 売上から手数料を引いた金額を入金額とする
        If counter = 0 Then
            Sheets("銀行明細").Cells(sheetRow + counter, 6).value = salesRs!sales - amountDiff
        '// 2件目以降 → 行挿入してデータ入力 売上高を入金額とし、手数料は0とし、combined列をTrueにする
        Else
            Sheets("銀行明細").Rows(sheetRow + counter).Insert xlDown
            insertedCount = insertedCount + 1
            Sheets("銀行明細").Cells(sheetRow + counter, 6).value = salesRs!sales
            Sheets("銀行明細").Cells(sheetRow + counter, 7).value = 0
            Sheets("銀行明細").Cells(sheetRow + counter, 11).value = True
        End If
        
        '// 日付・取引先コード・取引先名・口座名義入力(日付は値だけ入力するとユーザー定義になるため、コピーして貼り付ける)
        With Sheets("銀行明細")
            .Cells(sheetRow, 2).Copy Destination:=.Cells(sheetRow + counter, 2)
            .Cells(sheetRow + counter, 3).value = customerRs!ID
            .Cells(sheetRow + counter, 4).value = customerRs!customer_name
            .Cells(sheetRow + counter, 5).value = customerRs!customer_account
        End With
        
Continue:
        salesRs.filter = adFilterNone
        customerRs.MoveNext
    Loop
    
    '// 入金が複数の売上が合算されている場合(1つの売上に対する入金ではない場合) → 備考に合算入金額を表示、合算かどうかを示すcombined列の値をTrueにする
    If counter > 0 Then
        Sheets("銀行明細").Cells(sheetRow, 8).value = "合算入金額:" & Format(totalPayment, "#,##0") & "円 " & Sheets("銀行明細").Cells(sheetRow, 8).value
        Sheets("銀行明細").Cells(sheetRow, 11).value = True
        
        
        With Sheets("銀行明細").Range(Cells(sheetRow, 1), Cells(sheetRow + counter, 9))
            .Interior.Color = RGB(144, 238, 144)
            .Font.Color = vbBlack
        End With
    End If
            
    customerRs.filter = adFilterNone
    
    '// 何行挿入したかが返り値となる
    calculateCombinedPayment = insertedCount
    
End Function

'/**
 '* 月に複数回入金がある取引先の処理
 '* @params customerAccount 口座名義
 '* @params sales           売掛金
 '* @params customerRs      取引先レコードセット
 '* @params bankRs          銀行明細レコードセット
 '* @params paymentMonth    入金月
'**/
Private Sub calculateSeveralTimePayment(ByVal customerAccount As String, ByVal sales As Long, ByVal salesRs As Recordset, ByVal customerRs As Recordset, ByVal bankRs As Recordset, ByVal paymentMonth As Long)
    
    '// 入金額の合計
    Dim totalPayment As Long
    totalPayment = WorksheetFunction.SumIf(Sheets("銀行明細").Columns(5), customerAccount, Sheets("銀行明細").Columns(6))
    
   '/**
    '* 入金額と売掛金の差額を求め、1000円以上であれば差額を入金回数で割り、それぞれの手数料とする
    '* 1000円より小さい場合は入金額が最大のところの手数料とする
   '**/
    Dim amountDiff As Long: amountDiff = sales - totalPayment
    bankRs.filter = "口座名義 = " & "'" & customerAccount & "'"
    
    '// 入金回数
    Dim paymentCount As Long: paymentCount = WorksheetFunction.CountIf(Sheets("銀行明細").Columns(5), customerAccount)
    
    '// 売掛金と入金額合計の差額が0の場合 → 該当の取引先の手数料を全て0円にする
    If amountDiff = 0 Then
        Do While bankRs.EOF = False
            bankRs!売掛金との差額 = 0
            bankRs!備考 = "複数回入金 " & bankRs!備考
            bankRs.MoveNext
        Loop
    '// 差額が入金回数×880円以内の場合 → 手数料を入金回数で割ったものをそれぞれの手数料とする
    ElseIf 0 < amountDiff And amountDiff <= 880 * paymentCount Then
        Do While bankRs.EOF = False
            bankRs!売掛金との差額 = amountDiff / paymentCount
            bankRs!備考 = "複数回入金" & bankRs!備考
            bankRs.MoveNext
        Loop
    '// それ以外 → 差額を入金額が1番多いところの手数料とし、それ以外の手数料を0円にする
    Else
        bankRs.Sort = "振込金額 DESC"
        bankRs!売掛金との差額 = amountDiff
        bankRs!備考 = "複数回入金 " & bankRs!備考
        bankRs.MoveNext
        
        Do While bankRs.EOF = False
            bankRs!売掛金との差額 = 0
            bankRs!備考 = "複数回入金" & bankRs!備考
            bankRs.MoveNext
        Loop
    End If
    
    '// 同一口座から異なる取引先の売掛金が複数回に分けて入金される場合
    If customerRs.RecordCount > 1 Then
        bankRs.MoveFirst
            
        Dim tmpSales As Long
        Dim validCustomerId As Long
        
        '// 取引先レコードセットのフィルター:異なる取引先の売掛金が入金されてくる場合は, "several_times_payment_group = 0" のようなフィルターになっている
        Dim previousFilter As String: previousFilter = customerRs.filter
        
        '// 複数の取引先の売掛金と入金額の差額を比較して、最も妥当な売掛金の取引先のコードと名前を入力する
        Do While bankRs.EOF = False
            
            Do While customerRs.EOF = False
                salesRs.filter = "customer_id = " & customerRs!ID & " AND sales_month = " & getTargetMonth(customerRs!customer_site, paymentMonth)
                
                '// 入金額と2つの売掛金から、どちらが売掛金として妥当かを判断
                '// compareNumbersAsSales [入金額]. [売掛金1], [売掛金2]
                tmpSales = compareNumbersAsSales(bankRs!振込金額, tmpSales, salesRs!sales)
                        
                If tmpSales = salesRs!sales Then
                    validCustomerId = customerRs!ID
                End If
                
                customerRs.MoveNext
            Loop
            
            '// 妥当だと判断された取引先のコードと名前を明細に入力
            customerRs.MoveFirst
            customerRs.filter = "id = " & validCustomerId
            bankRs!取引先コード = customerRs!ID
            bankRs!取引先名 = customerRs!customer_name
        
            customerRs.filter = previousFilter
            bankRs.MoveNext
        Loop
    End If
                
    bankRs.MoveFirst
    bankRs.Update
    
    customerRs.filter = adFilterNone
    bankRs.filter = adFilterNone
    bankRs.Sort = ""
    
End Sub

'/**
 '* 複数売上の合計額を求める
 '* @params salesRsThisYear 売掛金レコードセット(今年)
 '* @params salesRsLastYear 売掛金レコードセット(昨年)
 '* @params customerRs      取引先レコードセット
 '* @params paymentMonth    入金月
 '* @return 売上の合計額
'**/
Private Function calculateTotalSales(ByVal salesRsThisyear As Recordset, ByVal salesRsLastYear As Recordset, customerRs As Recordset, paymentMonth As Long) As Long

    Dim targetMonth As Long
    Dim salesRs As ADODB.Recordset
    
    Do While customerRs.EOF = False
        '// 売上の対象月をサイトから求める(昨年の場合は対象月にマイナスがつく) 例) getTargetMonth [翌月], [1] → -12
        targetMonth = getTargetMonth(customerRs!customer_site, paymentMonth)
    
        '// 対象月が昨年の場合(値がマイナスの場合) → 昨年の売掛金レコードセットを使用
        If targetMonth < 0 Then
            Set salesRs = salesRsLastYear
        Else
            Set salesRs = salesRsThisyear
        End If
        
        salesRs.filter = "customer_id = " & customerRs!ID & " AND sales_month = " & Abs(targetMonth)
        
        If salesRs.RecordCount = 0 Then
            GoTo Continue
        End If
        
        calculateTotalSales = calculateTotalSales + salesRs!sales

Continue:
        salesRs.filter = adFilterNone
        customerRs.MoveNext
    Loop

    customerRs.MoveFirst

    Set salesRs = Nothing

End Function

'/**
 '* 入金日とサイトから何月分の売上が対象かを求める
'**/
Private Function getTargetMonth(ByVal site As String, ByVal paymentMonth As Long)

    If site = "翌月" Then
        getTargetMonth = passMonth(paymentMonth, -1)
    ElseIf site = "翌々" Then
        getTargetMonth = passMonth(paymentMonth, -2)
    ElseIf site = "翌翌々" Then
        getTargetMonth = passMonth(paymentMonth, -3)
    End If
    
End Function

'/**
 '* 売掛金と振込金額の差額を計算
 '*
'**/
Private Function calculateDifference(ByVal customerId As Long, ByVal amount As Long, ByVal paymentMonth As Long, ByVal site As String, ByVal salesRsThisyear As ADODB.Recordset, ByVal salesRsLastYear As ADODB.Recordset)

    Dim targetMonth As Long
    
    If site = "翌月" Then
        targetMonth = passMonth(paymentMonth, -1)
    ElseIf site = "翌々" Then
        targetMonth = passMonth(paymentMonth, -2)
    ElseIf site = "翌翌々" Then
        targetMonth = passMonth(paymentMonth, -3)
    End If
    
    Dim salesRs As ADODB.Recordset
    
    '// 入金日とサイトから売掛金の対象月を求めた結果がマイナスの場合(入金日の年より1年前が対象の場合)は1年前のレコードセットを使用する
    If targetMonth < 0 Then
        Set salesRs = salesRsLastYear
    Else
        Set salesRs = salesRsThisyear
    End If
    
    salesRs.filter = "customer_id = " & customerId & " AND sales_month = " & Abs(targetMonth)
    If salesRs.RecordCount = 0 Then
        salesRs.filter = adFilterNone
        Exit Function
    End If

    '/**
     '* 差額を計算し、差額が1000円を超える場合は前月の売上から3カ月分さかのぼり、それぞれ差額を計算する
     '* その中で差額が1000円以下のものが無ければサイトに登録されている売上から入金額を引いた金額を差額とする
    '**/
    
    '// DBに登録されているサイトの売上高と入金額の差額
    Dim amountDiffDefault As Long: amountDiffDefault = salesRs!sales - amount
    Dim amountDiff As Long
    
    '// サイトに登録されている月の売上高と入金額の差額が0円以上1000円以下の場合
    If 0 <= amountDiffDefault And amountDiffDefault <= 1000 Then
        amountDiff = amountDiffDefault
        GoTo Break
    End If
    
    Dim i As Long
    
    '// 差額が1000円以上の場合 → 入金月の前月売上から3カ月さかのぼり、それぞれの差額を計算し、最も妥当なものを採用する
    For i = 1 To 3
        '// 入金月からiを引いた値が0以下の場合 → 1年前の売掛金レコードセットを使用する
        If paymentMonth - i < 1 Then
            Set salesRs = salesRsLastYear
        Else
            Set salesRs = salesRsThisyear
        End If
        
        salesRs.filter = "customer_id = " & customerId & " AND sales_month = " & paymentMonth - i
        
        If salesRs.RecordCount = 0 Then
            salesRs.filter = adFilterNone
            GoTo Continue
        End If
        
        amountDiff = salesRs!sales - amount
        
        If 0 <= amountDiff And amountDiff <= 1000 Then
            GoTo Break
        End If
Continue:
    Next

Break:
    salesRs.filter = adFilterNone
    
    '// 入金月から3カ月さかのぼっても妥当な差額が見つからない場合は、サイト月の売上高と入金額の差額を使用する
    If amountDiff < 0 Or 1000 < amountDiff Then
        amountDiff = amountDiffDefault
    End If
    
    calculateDifference = amountDiff
    
    Set salesRs = Nothing
    
End Function

'// 管理帳へ貼り付け
Public Sub pasteToLedger()

    Application.ScreenUpdating = False

    If MsgBox("銀行明細を管理帳へ貼り付けますが、よろしいですか?", vbQuestion + vbYesNo, "山岸運送売掛金回収用ファイル") = vbNo Then
        Exit Sub
    End If
    
    Dim filePath As String: filePath = Sheets("設定").Cells(8, 3).value
    
    Dim fso As New FileSystemObject
    
    '// 貼り付け先が存在しなければ抜ける
    If fso.FileExists(filePath) = False Then
        MsgBox "貼り付け先として設定されているファイルが存在しません。", vbExclamation, "山岸運送売掛金回収用ファイル"
        Set fso = Nothing
        Exit Sub
    End If
    
    '/**
     '* 必要な箇所を貼り付け
    '**/
    Dim ledgerFile As Workbook: Set ledgerFile = Workbooks.Open(filePath)
    
    '// 管理帳に「ワーク2」のシートが無ければ抜ける
    If sheetExist(ledgerFile, "ワーク2") = False Then
        ledgerFile.Close False
        MsgBox "管理帳として指定されているファイルが適切ではありません。" & vbLf & "管理帳に「ワーク2」のシートがあることを確認してください。", vbExclamation, "山岸運送売掛金回収用ファイル"
        
        Set ledgerFile = Nothing
        Set fso = Nothing
        Exit Sub
    End If
    
    ledgerFile.Sheets("ワーク2").Cells.Clear
        
    With ThisWorkbook.Sheets("銀行明細")
        '// 日付貼り付け
        .Range(.Cells(5, 2), .Cells(Rows.Count, 2).End(xlUp)).Copy
        ledgerFile.Sheets("ワーク2").Cells(1, 1).PasteSpecial xlPasteValues
    
        '// 日付を「0月0日」から日付の数字だけに変更 例)10月3日 → 3
        With ledgerFile.Sheets("ワーク2")
            .Cells(2, 2).Formula = "=DAY(A2)"
            .Cells(2, 2).AutoFill .Range(.Cells(2, 2), .Cells(Rows.Count, 1).End(xlUp).Offset(, 1))
            .Columns(2).Copy
            .Columns(2).PasteSpecial xlPasteValues
            .Columns(1).Delete
            .Cells(1, 1).value = "日付"
        End With
        
        '// 取引先コード・取引先名貼り付け
        .Range(.Cells(5, 3), .Cells(Rows.Count, 4).End(xlUp)).Copy
        ledgerFile.Sheets("ワーク2").Cells(1, 2).PasteSpecial xlPasteValues
        
        '// 振込金額・差額貼り付け
        .Range(.Cells(5, 6), .Cells(Rows.Count, 6).End(xlUp).Offset(, 1)).Copy
        ledgerFile.Sheets("ワーク2").Cells(1, 4).PasteSpecial xlPasteValues
        
    End With
    
    Application.CutCopyMode = False
    
    ThisWorkbook.Sheets("銀行明細").Activate
    
    MsgBox "管理帳への貼り付けが完了しました。", vbInformation, "山岸運送売掛金回収用ファイル"

End Sub
