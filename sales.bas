Attribute VB_Name = "sales"
'// 売上関連
Option Explicit

'// 売上データ更新
Public Sub updateSales()
    
    Application.ScreenUpdating = False
    
    Dim wsh As New WshShell
    
    '// 取り込むデータを選択し、ファイルが選択されていない場合とファイルが適切出ない場合は抜ける
    Dim fileName As String: fileName = selectFile("売上データ取込", wsh.SpecialFolders(4), "CSVファイル", "*.csv")
    
    Set wsh = Nothing
    
    If fileName = "" Then
        Exit Sub
    ElseIf checkFile(fileName) = False Then
        Exit Sub
    End If
    
    Dim fso As New FileSystemObject
    
    Dim salesCsv As Workbook: Set salesCsv = Workbooks(fso.GetFileName(fileName))
    
    Set fso = Nothing
    
    '// 売上年
    Dim salesYear As Long: salesYear = Format(salesCsv.Sheets(1).Cells(2, 1).value, "yyyy")
    Dim salesMonth As Long: salesMonth = Format(salesCsv.Sheets(1).Cells(2, 1).value, "m")
    
    '// 売上の年でDBとして使用するExcelファイルを変更する
    Dim dbBookName As String: dbBookName = ThisWorkbook.Path & "\database\sales\" & salesYear & ".xlsx"
    Dim fc As New FileController
    fc.createFileIfNotExist (dbBookName)
    Set fc = Nothing
    
    '// DB接続
    Dim con As ADODB.Connection: Set con = connectDb(dbBookName)

    Dim rs As New ADODB.Recordset
    rs.Open "SELECT * FROM [sales$] ORDER BY sales_id", con, adOpenStatic, adLockOptimistic
    
    '// 売上のid(売上がまだDBに1つも登録されていなければidは1になる)
    Dim nextId As Long

    If rs.RecordCount = 0 Then
        nextId = 1
    Else
        rs.MoveLast
        nextId = rs![sales_id] + 1
    End If

    Dim i As Long

    '// 売上データの更新
    For i = 2 To salesCsv.Sheets(1).Cells(Rows.Count, 1).End(xlUp).Row

        '// 取引先コードと年月で検索して既にデータがあれば金額を更新し、無ければ新規追加する
        rs.filter = _
            "customer_id = " & salesCsv.Sheets(1).Cells(i, 9).value & _
            " AND sales_year = " & salesYear & _
            " AND sales_month = " & salesMonth

        '// 新規追加
        If rs.RecordCount = 0 Then
            rs.AddNew
            rs!sales_id = nextId
            rs!customer_id = salesCsv.Sheets(1).Cells(i, 9).value
            rs!sales_year = salesYear
            rs!sales_month = salesMonth
            nextId = nextId + 1
        End If

        rs!sales = salesCsv.Sheets(1).Cells(i, 12).value

        rs.Update

        rs.filter = adFilterNone
    Next

    rs.Close
    Set rs = Nothing
    con.Close
    Set con = Nothing
    
    salesCsv.Close False
    Set salesCsv = Nothing
    
    '// 売上を数値化
    Call convertStrSales2Number(dbBookName)
    
    MsgBox "取込が完了しました。", Title:="売上データ取込"
    
End Sub

'// 選択したファイルが適切か確認
Private Function checkFile(ByVal fileName As String) As Boolean

    Dim targetFile As Workbook: Set targetFile = Workbooks.Open(fileName)

    With targetFile.Sheets(1).Cells(1, 1)
        If .value = "開始日付" _
            And .Offset(, 1).value = "終了日付" _
            And .Offset(, 2).value = "コード" _
            And .Offset(, 3).value = "部門" _
            And .Offset(, 4).value = "コード" _
            And .Offset(, 5).value = "科目" _
            And .Offset(, 6).value = "コード" _
            And .Offset(, 7).value = "補助科目" _
            And .Offset(, 8).value = "コード" _
            And .Offset(, 9).value = "取引先名" _
            And .Offset(, 10).value = "繰越額" _
            And .Offset(, 11).value = "借方金額" _
            And .Offset(, 12).value = "貸方金額" _
            And .Offset(, 13).value = "残高" Then
        
            checkFile = True
        Else
            targetFile.Close False
            MsgBox "選択したファイルが適切ではありません。", vbExclamation, "売上データ取込"
            checkFile = False
        End If
    End With

End Function


'// 文字列として保存されている売上データを数値に変換
Private Sub convertStrSales2Number(ByVal dbBookName As String)
    
    Dim dbBook As Workbook: Set dbBook = Workbooks.Open(dbBookName)
    
    With dbBook.Sheets("sales")
        .Activate
        
        '// converStr2Number [数値化する範囲]
        Call convertStr2Number(.Range(.Cells(2, 1), .Cells(.Cells(Rows.Count, 1).End(xlUp).Row, 5)))
    End With
    
    dbBook.Close True
    Set dbBook = Nothing
    
End Sub

