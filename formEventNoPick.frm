VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEventNoPick 
   Caption         =   "大会選択"
   ClientHeight    =   6902
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   6804
   OleObjectBlob   =   "formEventNoPick.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formEventNoPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  formEventNoPick
'


Private Sub btnClose_Click()
    HyouShow.MyCon.Close
    HyouShow.MyCon = Nothing
    Unload Me
End Sub

Private Sub listEvent_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' エンターキーが押されたとき、 CommandButton1 をクリック
        Call btnOK_Click
    End If
End Sub

Sub add_list_item(row As Integer, item1 As String, item2 As String, item3 As String, item4 As String)
    formPrgNoPick.listPrg.AddItem ("")
    formPrgNoPick.listPrg.List(row, 0) = item1
    formPrgNoPick.listPrg.List(row, 1) = item2
    formPrgNoPick.listPrg.List(row, 2) = item3
    formPrgNoPick.listPrg.List(row, 3) = item4

End Sub





Sub CreateTableIfNotExists()
    Dim cmd As Object 'ADODB.Command
    Dim sql As String


    sql = "IF NOT EXISTS (" & _
          "SELECT 1 " & _
          "FROM INFORMATION_SCHEMA.TABLES " & _
          "WHERE TABLE_NAME = '印刷状況'" & _
          ") " & _
          "BEGIN " & _
          "CREATE TABLE 印刷状況 (" & _
          "大会番号 smallINT NOT NULL, " & _
          "競技番号 smallINT NOT NULL, " & _
          "印刷状況 smallint NOT NULL " & _
           " CONSTRAINT PK_印刷状況 PRIMARY KEY (大会番号, 競技番号)" & _
          "); " & _
          "END;"
    
    On Error GoTo ErrorHandler
    
    
    ' ADODB.Command オブジェクトを作成してSQLを実行
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = HyouShow.MyCon
    cmd.CommandText = sql
    cmd.Execute
    
    ' リソースを解放

    Set cmd = Nothing

    Exit Sub

ErrorHandler:
    ' エラー処理
    Debug.Print "エラーが発生しました: " & Err.Description

    Set cmd = Nothing

End Sub


Sub CopyToPrintStatusIfNotExists(target大会番号 As Integer)
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim connectionString As String
    Dim checkSql As String
    Dim insertSql As String
    
   
    
    ' 存在確認用SQL文
    checkSql = "SELECT 1 FROM 印刷状況 WHERE 大会番号 = " & target大会番号 & ";"
    
    ' 挿入用SQL文
    insertSql = "INSERT INTO 印刷状況 (大会番号, 競技番号, 印刷状況) " & _
                "SELECT 大会番号, 競技番号, 0 " & _
                "FROM プログラム " & _
                "WHERE 大会番号 = " & target大会番号 & ";"
    
    On Error GoTo ErrorHandler
    
    ' ADODB.Connection オブジェクトを作成

    
    ' 存在確認用SQLを実行
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = HyouShow.MyCon
    cmd.CommandText = checkSql
    
    Set rs = cmd.Execute
    If rs.EOF Then
        ' レコードが存在しない場合のみ挿入SQLを実行
        cmd.CommandText = insertSql
        cmd.Execute

    End If
    
    ' リソースを解放
    rs.Close

    Set rs = Nothing
    Set cmd = Nothing

    Exit Sub

ErrorHandler:
    ' エラー処理
    Debug.Print "エラーが発生しました: " & Err.Description
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set rs = Nothing
    Set cmd = Nothing

End Sub


Private Sub btnOK_Click()
    Dim Gender(5) As String
    Gender(1) = "男子"
    Gender(2) = "女子"
    Gender(3) = "混成"
    Gender(4) = "混合"
    Dim selectedItem As String
    Dim myRecordset As New ADODB.Recordset
    Dim myquery As String
    Dim row As Integer

    selectedItem = listEvent.Value
    HyouShow.EventNo = CInt(Left(selectedItem, 3))
    Call CreateTableIfNotExists
    CopyToPrintStatusIfNotExists (HyouShow.EventNo)
    If HyouShow.class_exist("") Then
        Call add_list_item(0, "#", "クラス", "種目", "st")
        row = 1
        myquery = "SELECT プログラム.表示用競技番号 as 競技番号, クラス.クラス名称 as クラス, " & _
              "プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, 種目.種目 as 種目 FROM プログラム" + _
              " INNER JOIN 種目 ON 種目.種目コード = プログラム.種目コード " + _
              " INNER JOIN クラス ON クラス.クラス番号=プログラム.クラス番号 " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & HyouShow.EventNo & " AND " + _
              " クラス.大会番号 = " & HyouShow.EventNo & _
              " order by 競技番号 asc;"
              
            myRecordset.Open myquery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordset.EOF

                Call add_list_item(row, Right("   " & myRecordset!競技番号, 3), myRecordset!クラス, _
                    Gender(myRecordset!性別) + myRecordset!距離 + myRecordset!種目, "")
                row = row + 1
                myRecordset.MoveNext
            Loop
    Else
        Call add_list_item(0, "#", "", "種目", "")
        row = 1
        myquery = "SELECT プログラム.競技番号 as 競技番号,  " & _
              "プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, 種目.種目 as 種目 FROM プログラム" + _
              " INNER JOIN 種目 ON 種目.種目コード = プログラム.種目コード " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & HyouShow.EventNo & ";"
            myRecordset.Open myquery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordset.EOF
                Call add_list_item(row, Right("   " & myRecordset!競技番号, 3), "", _
                    Gender(myRecordset!性別) + myRecordset!距離 + myRecordset!種目, "")

                row = row + 1
                myRecordset.MoveNext
            Loop
    End If
    formPrgNoPick.LastRow = row - 1
    

    myRecordset.Close
    Set myRecordset = Nothing
    Call HyouShow.init_senshu("")
    
    Unload Me
    formPrgNoPick.show vbModeless
End Sub


