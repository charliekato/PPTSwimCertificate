VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEventNoPick 
   Caption         =   "大会選択"
   ClientHeight    =   6902
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6804
   OleObjectBlob   =   "formEventNoPick.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formEventNoPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub listEvent_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' エンターキーが押されたとき、 CommandButton1 をクリック
        Call btnOK_Click
    End If
End Sub



Private Sub btnOK_Click()
    Dim Gender(4) As String
    Gender(1) = "男子"
    Gender(2) = "女子"
    Gender(3) = "混合"
    Dim selectedItem As String
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String

    selectedItem = listEvent.value
    HyouShow.EventNo = CInt(Left(selectedItem, 3))
    If HyouShow.class_exist(HyouShow.EventNo) Then
        myquery = "SELECT プログラム.表示用競技番号 as 競技番号, クラス.クラス名称 as クラス, " & _
              "プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, 種目.種目 as 種目 FROM プログラム" + _
              " INNER JOIN 種目 ON 種目.種目コード = プログラム.種目コード " + _
              " INNER JOIN クラス ON クラス.クラス番号=プログラム.クラス番号 " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & HyouShow.EventNo & " AND " + _
              " クラス.大会番号 = " & HyouShow.EventNo & _
              " order by 競技番号 asc;"
              
            myRecordSet.Open myquery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordSet.EOF
                
                formPrgNoPick.listPrg.AddItem Right("   " & myRecordSet!競技番号, 3) & "  " & _
                          Gender(if_not_null(myRecordSet!性別)) & " " & _
                          Right("               " + if_not_null_string(myRecordSet!クラス), 10) & " " & _
                          if_not_null_string(myRecordSet!距離) & " " & _
                          if_not_null_string(myRecordSet!種目)
                myRecordSet.MoveNext
            Loop
    Else
        myquery = "SELECT プログラム.競技番号 as 競技番号,  " & _
              "プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, 種目.種目 as 種目 FROM プログラム" + _
              " INNER JOIN 種目 ON 種目.種目コード = プログラム.種目コード " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & HyouShow.EventNo & ";"
            myRecordSet.Open myquery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordSet.EOF
                formPrgNoPick.listPrg.AddItem Right("   " & myRecordSet!競技番号, 3) & "  " & _
                          Gender(if_not_null(myRecordSet!性別)) & " " & _
                          if_not_null_string(myRecordSet!距離) & " " & _
                          if_not_null_string(myRecordSet!種目)
                myRecordSet.MoveNext
            Loop
    End If
    

    myRecordSet.Close
    Set myRecordSet = Nothing
    Call HyouShow.init_senshu(HyouShow.EventNo)
    
    Unload Me
    formPrgNoPick.show vbModeless
End Sub


