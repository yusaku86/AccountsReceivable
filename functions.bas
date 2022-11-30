Attribute VB_Name = "functions"
Option Explicit

'// db接続
Public Function connectDb(ByVal dbBook As String) As ADODB.Connection

    Dim returnCon As New ADODB.Connection
    
    '// dbとして使用するファイルに接続
    With returnCon
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .Properties("Extended Properties") = "Excel 12.0"
        .Open dbBook
    End With

    Set connectDb = returnCon

End Function

'/**
 '* 文字列として保存されているデータを数値化
 '* @params targetRange データを数値化する範囲
'**/
Public Sub convertStr2Number(ByVal targetRange As Range)
    
    targetRange.value = Evaluate(targetRange.Address & "*1")
    
End Sub

'// ダイアログを表示し、ファイルを選択する
Public Function selectFile(ByVal dialogTitle As String, ByVal initialFile As String, ByVal targetDiscription As String, ByVal targetExtension As String) As String
    
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = initialFile
        .AllowMultiSelect = False
        .Title = dialogTitle
        
        '// 選択するファイルの拡張子設定
        .filters.Clear
        .filters.add targetDiscription, targetExtension
        
        If .Show Then
            selectFile = .SelectedItems(1)
        End If
    End With
    
End Function

'// ファイルに指定のシートが存在するか確認
Public Function sheetExist(ByVal activeFile As Workbook, ByVal sheetName As String) As Boolean

    Dim sheet As Worksheet
    
    For Each sheet In activeFile.Worksheets
        If sheet.Name = sheetName Then
            sheetExist = True
            Exit Function
        End If
    Next
    
    sheetExist = False

End Function

'/**
 '* 基準月から指定の月だけ経過(さかのぼった)月が何月かを求める
 '* 遡った月が昨年になる場合は返り値にマイナスを付ける
 '*
 '* 例) passMonth [1], [-3] → -10
'**/

Public Function passMonth(ByVal standardMonth As Long, ByVal passingMonths As Long)

    Dim returnMonth As Long: returnMonth = standardMonth + passingMonths

    If returnMonth <= 0 Then
        returnMonth = -(returnMonth + 12)
    ElseIf returnMonth > 12 Then
        returnMonth = returnMonth - 12
    End If
    
    passMonth = returnMonth

End Function

'// 2つの数字のうち、どちらが手数料として妥当かを判定
Public Function compareNumbersAsCommision(ByVal number1 As Long, ByVal number2 As Long) As Long

    '// 2つの数字が同じ場合 → 前者を返す
    If number1 = number2 Then
        compareNumbersAsCommision = number1
        Exit Function
    End If
    
    '// どちらかが0の場合は0の方が手数料として妥当
    If number1 = 0 Then
        compareNumbersAsCommision = number1
        Exit Function
    ElseIf number2 = 0 Then
        compareNumbersAsCommision = number2
        Exit Function
    End If
    
    '// どちらか１つのみが負の数字の場合→負の数字ではない方が妥当
    If 0 < number1 And number2 < 0 Then
        compareNumbersAsCommision = number1
        Exit Function
    ElseIf 0 < number2 And number1 < 0 Then
        compareNumbersAsCommision = number2
        Exit Function
    End If
    
    '// 絶対値が小さい方が妥当
    If Asc(number1) < Asc(number2) Then
        compareNumbersAsCommision = number1
    Else
        compareNumbersAsCommision = number2
    End If
    
End Function

'// 2つの数字のうち、入金額に対してどちらが売上として妥当かを判定
Public Function compareNumbersAsSales(ByVal payment As Long, ByVal sales1 As Long, ByVal sales2 As Long) As Long

    '// 2つの数字が同じ場合 → 前者を返す
    If sales1 = sales2 Then
        compareNumbersAsSales = sales1
        Exit Function
    End If
    
    '// 入金額とどちらか1つの数字の差額が0の場合 -> 差額が0になる方が妥当
    If sales1 - payment = 0 Then
        compareNumbersAsSales = sales1
        Exit Function
    ElseIf sales2 - payment = 0 Then
        compareNumbersAsSales = sales2
        Exit Function
    End If
    
    '// どちらかの入金額との差額が負の数字になる場合 → 負にならない方が妥当
    If 0 < sales1 - payment And sales2 - payment < 0 Then
        compareNumbersAsSales = sales1
        Exit Function
    ElseIf 0 < sales2 - payment And sales1 - payment < 0 Then
        compareNumbersAsSales = sales2
        Exit Function
    End If
    
    '// 上記以外 → 入金額との差額の絶対値が小さい方が妥当
    If Asc(sales1 - payment) < Asc(sales2 - payment) Then
        compareNumbersAsSales = sales1
    Else
        compareNumbersAsSales = sales2
    End If
    
End Function



