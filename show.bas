' ************************************************************************************************
' ***
' *** 概要：テーブルのデータを表示する
' ***
' ************************************************************************************************
Option Explicit

Dim database As String
'Dim dsn As String
Dim cnUser As String
Dim cnPass As String
Dim currentRow As Integer

'データ取得ボタン押下時の処理
Private Sub btnGetData_Click()
    On Error GoTo btnGetData_Error
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim sql As String
    Dim columnSize As Integer
    
    '入力値を取得
    database = Cells(1, 2)
    'dsn = Cells(1, 2)
    cnUser = Cells(1, 4)
    cnPass = Cells(2, 4)
    
    'シートのクリア
    Call clearResult

    currentRow = 3
    
    '*************************
    ' コネクションの生成
    '*************************
    cn.Open "Provider=OraOLEDB.Oracle;Data Source=" & database & ";User ID=" & cnUser & ";Password=" & cnPass & ";"
    'cn.Open "dsn=" & dsn & "; username=" & cnUser & "; password=" & cnPass & ";"
    
    '*************************
    ' Users 取得
    '*************************
    currentRow = currentRow + 1
    'テーブル名の表示
    Call displayTableName("Users")
    
    'Users を検索
    sql = ""
    sql = sql & "select * from Users"
    
    'レコードセットの生成
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
    
    currentRow = currentRow + 1
    'ヘッダ行の表示
    Call displayHeaderColumn(rs)
    
    currentRow = currentRow + 1
    'データ行の表示
    Call displayData(rs)
    
    'レコードセットのクローズ
    rs.Close
    Set rs = Nothing
    
    '*************************
    ' Books 取得
    '*************************
    currentRow = currentRow + 1
    'テーブル名の表示
    Call displayTableName("Books")
    
    'Books を検索
    sql = ""
    sql = sql & "select * from Books"
    
    'レコードセットの生成
    rs.Open sql, cn, adOpenStatic, adLockReadOnly
    
    currentRow = currentRow + 1
    'ヘッダ行の表示
    Call displayHeaderColumn(rs)
    
    currentRow = currentRow + 1
    'データ行の表示
    Call displayData(rs)
    
    'レコードセットのクローズ
    rs.Close
    Set rs = Nothing
    
    '*************************
    ' 列幅の調整
    '*************************
    Columns("A:AZ").EntireColumn.AutoFit
    Range("A1:A1").Select
    
    GoTo Terminate
    
btnGetData_Error:
    MsgBox Err.Description & ":" & Err.Source
    
    GoTo Terminate

Terminate:
    'クローズ処理
    If rs.State <> adStateClosed Then
        rs.Close
        Set rs = Nothing
    End If
    If cn.State <> adStateClosed Then
        cn.Close
        Set cn = Nothing
    End If
End Sub

'シートのクリア
Private Sub clearResult()
    Range("A4", "AZ1000").Select
    Selection.ClearContents

    With Rows("3:1000")
        .NumberFormat = "General"
        .Interior.ColorIndex = 2
        .Interior.Pattern = xlSolid
        .Borders.LineStyle = xlNone
        .Font.Name = "ＭＳ ゴシック"
    End With
End Sub

'テーブル名の表示
Private Sub displayTableName(tableName As String)
    Cells(currentRow, 1) = tableName
    Cells(currentRow, 1).Interior.ColorIndex = 37
End Sub

'ヘッダ行カラムの表示
Private Function displayHeaderColumn(rs As Recordset)
    Dim column As Integer

    'カラム名の表示
    For column = 1 To rs.Fields.Count
        Cells(currentRow, column) = rs(column - 1).Name
        
        'データ行の書式設定
        If rs(column - 1).Type = adVarWChar Then
            Range(Cells(currentRow, column), Cells(rs.RecordCount + currentRow, column)).NumberFormatLocal = "@"
        ElseIf rs(column - 1).Type = adDBTimeStamp Then
            Range(Cells(currentRow, column), Cells(rs.RecordCount + currentRow, column)).NumberFormatLocal = "yyyy/mm/dd hh:mm:ss"
        ElseIf rs(column - 1).Type = adVarNumeric Then
            Range(Cells(currentRow, column), Cells(rs.RecordCount + currentRow, column)).NumberFormatLocal = "0"
        End If
    Next
    
    'カラム行の書式設定
    With Range(Cells(currentRow, 1), Cells(currentRow, rs.Fields.Count))
        .Interior.ColorIndex = 15
        .Borders.Color = vbBlack
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    displayHeaderColumn = rs.Fields.Count
End Function

'データ行の表示
Private Sub displayData(rs As Recordset)
    Dim column As Integer
    Dim startRow As Integer
        
    startRow = currentRow
   
    'データの表示
    Do Until rs.EOF
        For column = 1 To rs.Fields.Count
            Cells(currentRow, column) = rs(column - 1).Value
        Next
        rs.MoveNext
        currentRow = currentRow + 1
    Loop

    'データ行の書式設定
    If startRow <> currentRow Then
        With Range(Cells(startRow, 1), Cells(currentRow - 1, rs.Fields.Count))
            .Borders.Color = vbBlack
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
    End If
End Sub
