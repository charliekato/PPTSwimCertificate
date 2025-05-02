Attribute VB_Name = "HyouShow"
Option Explicit
Option Base 0

    Public MyCon As ADODB.Connection
    Public EventNo As Integer

    Public Gender(4) As String
    Public Shumoku(8) As String
    Public Swimmer() As String



Sub init_gender(dummy As String)
    Gender(1) = "男子"
    Gender(2) = "女子"
    Gender(3) = "混合"
End Sub
Sub init_shumoku(dummy As String)
    Shumoku(1) = "自由形"
    Shumoku(2) = "背泳ぎ"
    Shumoku(3) = "平泳ぎ"
    Shumoku(4) = "バタフライ"
    Shumoku(5) = "個人メドレー"
    Shumoku(6) = "フリーリレー"
    Shumoku(7) = "メドレーリレー"
End Sub

Sub 賞状作成()
    Dim ss As formServerSelect
    Call init_gender("")
    Call init_shumoku("")
    
    Set ss = New formServerSelect
    ss.show
End Sub




Function if_not_null(obj As Variant) As Integer
    If IsNull(obj) Then
        if_not_null = 0
    Else
        if_not_null = obj
    End If
End Function

Function if_not_null_string(obj As Variant) As String
    If IsNull(obj) Then
        if_not_null_string = ""
    Else
        if_not_null_string = obj
    End If
End Function




Public Function class_exist(dummy As String) As Boolean
    Dim myRecordset As New ADODB.Recordset
    Dim myquery As String
    Dim rc As Boolean
    myquery = "select * from クラス where 大会番号 = " & EventNo
    myRecordset.Open myquery, MyCon, adOpenStatic, adLockReadOnly
    If myRecordset.EOF Then
        rc = False
    Else
        rc = True
    End If
    myRecordset.Close
    Set myRecordset = Nothing
    class_exist = rc

End Function

Sub set_open_to_kengai()
    Dim sql1 As String
    Dim sql2 As String
    Dim cmd As Object
    
    sql1 = "UPDATE 記録 " & _
              " Set 記録.オープン = 1 " & _
              " From 記録 " & _
              " INNER JOIN プログラム on プログラム.競技番号 = 記録.競技番号 " & _
              " INNER JOIN 選手 ON 記録.選手番号 = 選手.選手番号 " & _
              " WHERE 記録.大会番号 = " & EventNo & " And 選手.大会番号 = " & EventNo & _
              " And プログラム.大会番号 = " & EventNo & _
              " and プログラム.種目コード < 6 " & _
              " and 選手.加盟団体番号 <> 25; "
              
    sql2 = "UPDATE 記録 " & _
           "SET 記録.オープン = 1 " & _
           "FROM 記録 " & _
           "INNER JOIN プログラム ON プログラム.競技番号 = 記録.競技番号 " & _
           "INNER JOIN リレーチーム ON 記録.選手番号 = リレーチーム.チーム番号 " & _
           "WHERE 記録.大会番号 = " & EventNo & " AND リレーチーム.大会番号 = " & _
            EventNo & " AND プログラム.大会番号 = " & EventNo & _
           "AND プログラム.種目コード > 5 " & _
           "AND リレーチーム.加盟団体番号 <> 25;"
           
    Set cmd = CreateObject("ADODB.Command")
'    On Error GoTo ErrorHandler
   
    cmd.ActiveConnection = MyCon
    cmd.CommandText = sql1
    cmd.CommandType = adCmdText
    cmd.Execute

    
    

    cmd.ActiveConnection = MyCon
    cmd.CommandText = sql2
    cmd.CommandType = adCmdText
    cmd.Execute
Cleanup:
    If Not cmd Is Nothing Then Set cmd = Nothing
    Exit Sub
ErrorHandler:
    MsgBox "error while setting open"
    Resume Cleanup
End Sub


Sub reset_open()
    Dim sql As String
    
    Dim cmd As Object
    
    sql = "UPDATE 記録" & _
   " Set 記録.オープン = 0" & _
   " From 記録" & _
   " INNER JOIN 選手 ON 記録.選手番号 = 選手.選手番号 " & _
   " WHERE 記録.大会番号 = " & EventNo & " AND 選手.大会番号=" & EventNo
              

           
    Set cmd = CreateObject("ADODB.Command")
    On Error GoTo ErrorHandler2
    With cmd
        .ActiveConnection = MyCon
        .CommandText = sql
        .CommandType = adCmdText
    End With
    cmd.Execute
    

Cleanup:
    If Not cmd Is Nothing Then Set cmd = Nothing
    Exit Sub
ErrorHandler2:
    MsgBox "error while setting open"
    Resume Cleanup
End Sub



Sub init_senshu(dummy As String)

    Dim myRecordset As New ADODB.Recordset
    Dim myquery As String
    Dim maxSwimmerNo As Integer
    myquery = "SELECT MAX(選手番号) as MAX from 選手 where 大会番号 = " & EventNo
    myRecordset.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    maxSwimmerNo = myRecordset!Max
    
    ReDim Swimmer(maxSwimmerNo)
    myRecordset.Close
    myquery = "SELECT 氏名, 選手番号 from 選手 where 大会番号 = " & EventNo
    myRecordset.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        Swimmer(myRecordset!選手番号) = myRecordset!氏名
        myRecordset.MoveNext
    Loop
    myRecordset.Close
    Set myRecordset = Nothing
End Sub

Public Function GetPrgNofromPrintPrgNo(printPrgNo As Integer) As Integer
    Dim myRecordset As New ADODB.Recordset
    Dim myquery As String
    myquery = "select 競技番号 from プログラム where 表示用競技番号=" & printPrgNo & _
              "and 大会番号= " & EventNo & ";"
    myRecordset.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    GetPrgNofromPrintPrgNo = if_not_null(myRecordset!競技番号)
    myRecordset.Close
    Set myRecordset = Nothing
End Function
Sub get_race_title(ByVal prgNo As Integer, ByRef Class As String, _
            ByRef genderStr As String, ByRef distance As String, ByRef styleNo As Integer)
    Dim myRecordset As New ADODB.Recordset
    Dim myquery As String
    Dim classExist As Boolean
    classExist = class_exist("")
    If classExist Then
        myquery = "SELECT クラス.クラス名称 as クラス, プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, プログラム.種目コード as 種目 FROM プログラム " + _
              " INNER JOIN クラス ON クラス.クラス番号=プログラム.クラス番号 " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & EventNo & " AND " + _
              " クラス.大会番号 = " & EventNo & " AND " & _
              " プログラム.競技番号 = " & prgNo & ";"
    Else
        myquery = "SELECT  プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, プログラム.種目コード as 種目 FROM プログラム " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & EventNo & " AND " + _
              " プログラム.競技番号 = " & prgNo & ";"
    End If
    myRecordset.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        If classExist Then
            Class = myRecordset!クラス
        Else
            Class = ""
        End If
        genderStr = Gender(myRecordset!性別)
        distance = myRecordset!距離
        styleNo = myRecordset!種目
        myRecordset.MoveNext
    Loop
    myRecordset.Close
    Set myRecordset = Nothing
                
              
    
End Sub






Function get_swimmer_by_rank(ByRef resultList As Collection, ByVal rank As Integer) As Collection
    Dim thisResult As result
    Dim swimmerList As Collection
    Set swimmerList = New Collection
    For Each thisResult In resultList
        If thisResult.順位 = rank Then
            swimmerList.Add thisResult
            Exit For
        End If
    Next thisResult
    Set get_swimmer_by_rank = swimmerList
End Function
Function is_relay(style As Integer) As Boolean
    is_relay = False
    If style > 5 Then is_relay = True
    
End Function

Function ConvertTimeFormat(timeString As String)
    Dim minutes As String
    Dim seconds As String
    Dim milliseconds As String
    Dim colonPos As Integer
    Dim dotPos As Integer
    
    ' コロンとドットの位置を探す
    colonPos = InStr(timeString, ":")
    dotPos = InStr(timeString, ".")
    milliseconds = Mid(timeString, dotPos + 1)
    
    ' 分, 秒, ミリ秒を抽出
    If colonPos > 0 Then
    minutes = Mid(timeString, 1, colonPos - 1)
    seconds = Mid(timeString, colonPos + 1, dotPos - colonPos - 1)
    
    ' 変換された時間を返す
    ConvertTimeFormat = minutes & "分" & seconds & "秒" & milliseconds
    Else
        seconds = Trim(Mid(timeString, 1, dotPos - 1))
        ConvertTimeFormat = seconds & "秒" & milliseconds
    End If
End Function
''eventNo, prgNo, className, genderName, distance, printenable

Sub fill_out_form_relay(prgNo As Integer, className As String, _
                genderName As String, distance As String, styleNo As Integer, printenable As Boolean)
    Dim myquery As String
    Dim junni As Integer
    Dim junnib As Integer
    Dim prevTime As String



    Dim myRecordset As New ADODB.Recordset
    Dim winnerName As String
    Dim myTime As String
    junni = 0
    junnib = 0
    prevTime = ""
    myquery = "SELECT リレーチーム.チーム名 as チーム名, 記録.ゴール as ゴール, " & _
            "記録.第１泳者, 記録.第２泳者, 記録.第３泳者, 記録.第４泳者, 記録.新記録印刷マーク " & _
            "FROM 記録 " & _
            " inner join リレーチーム on リレーチーム.チーム番号 = 記録.選手番号 " & _
            " where   記録.競技番号 = " & prgNo & _
            " and 記録.大会番号 = " & EventNo & " and 記録.事由入力ステータス=0 " & _
            " and リレーチーム.大会番号 = " & EventNo & _
            " and 記録.オープン = 0 " & _
            " order by ゴール asc;"

    myRecordset.Open myquery, MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordset.EOF
        junnib = junnib + 1
        If prevTime <> myRecordset!ゴール Then
            junni = junnib
            If junni > CInt(FormOption.tbxJunniLast) Then
                Exit Do
            End If
            If junni < CInt(FormOption.tbxJunniTop) Then
                GoTo DOLOOPEND
            End If
            prevTime = myRecordset!ゴール
        End If
        Call fill_name(myRecordset!チーム名)

        Call fill_shozoku(Swimmer(myRecordset!第1泳者) & "・" & Swimmer(myRecordset!第2泳者) & "・" & _
                     Swimmer(myRecordset!第3泳者) & "・" & Swimmer(myRecordset!第4泳者))
        
        Call fill_junni(junni)
         Call fill_class(className)
        Call fill_shumoku(genderName + distance + Shumoku(styleNo))
        Call fill_time(ConvertTimeFormat(myRecordset!ゴール) + " " + _
            if_not_null_string(myRecordset!新記録印刷マーク))
        If printenable Then
            Call print_it("")
        End If
DOLOOPEND:
        
        myRecordset.MoveNext
    Loop
                    ' クローズと解放
    myRecordset.Close
    'MyCon.Close
    Set myRecordset = Nothing
    'Set MyCon = Nothing
End Sub

Sub fill_out_form_kojin(prgNo As Integer, className As String, _
                    genderName As String, distance As String, styleNo As Integer, printenable As Boolean)
    Dim myquery As String
    Dim junni As Integer
    Dim junnib As Integer
    Dim prevTime As String


    Dim myRecordset As New ADODB.Recordset
    Dim winnerName As String
    Dim myTime As String
    junni = 0
    junnib = 0
    prevTime = ""
    
    myquery = "SELECT 選手.氏名 as 氏名, 記録.ゴール as ゴール, 選手.所属名称1, 記録.新記録印刷マーク " & _
        "FROM 記録 " & _
        " inner join 選手 on 選手.選手番号 = 記録.選手番号 " & _
        " where 選手.大会番号=" & EventNo & " and 記録.競技番号 = " & prgNo & _
        " and 記録.大会番号 = " & EventNo & " and 記録.事由入力ステータス=0 " & _
        " and 記録.オープン = 0 " & _
        " and 記録.選手番号>0 order by ゴール asc;"

    myRecordset.Open myquery, MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordset.EOF
        junnib = junnib + 1
        If IsNull(myRecordset!ゴール) Or myRecordset!ゴール = "" Then
            MsgBox ("該当データがありません。たぶんレースが終わっていないと思われます。")
            Exit Do
        End If
        If prevTime <> myRecordset!ゴール Then

            junni = junnib
            If junni > CInt(FormOption.tbxJunniLast) Then
                Exit Do
            End If
            If junni < CInt(FormOption.tbxJunniTop) Then
                GoTo LOOPEND2
            End If
            prevTime = myRecordset!ゴール
        End If
        Call fill_name(myRecordset!氏名)
        Call fill_shozoku(myRecordset!所属名称1)
        Call fill_class(className)
        Call fill_shumoku(genderName + distance + Shumoku(styleNo))
        Call fill_time(ConvertTimeFormat(myRecordset!ゴール) + " " + _
                 if_not_null_string(myRecordset!新記録印刷マーク))
        Call fill_junni(junni)
        If printenable Then
            Call print_it("")
        End If
LOOPEND2:
       
        myRecordset.MoveNext
    Loop
            ' クローズと解放
    myRecordset.Close
    'MyCon.Close
    Set myRecordset = Nothing
    'Set MyCon = Nothing
End Sub

Sub fill_time(myTime As String)
    If FormOption.cbxTime.Value Then
        Call show("タイム", myTime)
    Else
        Call show("タイム", "")
    End If
End Sub

Sub fill_class(className As String)
    If FormOption.cbxClass.Value Then
        Call show("クラス", className)
    Else
        Call show("クラス", "")
    End If
End Sub
Sub fill_shumoku(Shumoku As String)
    If FormOption.cbxStyle.Value Then
        Call show("種目", Shumoku)
    Else
        Call show("種目", "")
    End If
End Sub

Sub fill_junni(junni As Integer)
    If FormOption.cbxJunni.Value Then
        If FormOption.cbxJunniShowMethod1.Value Then
            Call show("順位", "" & junni)
        ElseIf FormOption.cbxJunniShowMethod2.Value Then
            Call show("順位", "第" & junni & "位")
        ElseIf FormOption.cbxJunniShowMethod3.Value Then
            If junni = 1 Then
                Call show("順位", "優勝")
            Else
                Call show("順位", "第" & junni & "位")
            End If
        End If
    Else
        Call show("順位", "")
    End If
End Sub

Sub fill_name(myName As String)
    If FormOption.cbxName.Value Then
        Call show("選手名", myName)
    Else
        Call show("選手名", "")
    End If
End Sub

Sub fill_shozoku(shozoku As String)
    If FormOption.cbxBelongsTo.Value Then
        Call show("所属", shozoku)
    Else
        Call show("所属", "")
    End If
End Sub

Sub fill_out_form(prgNo As Integer, printenable As Boolean)

    Dim myquery As String


    Dim className As String
    Dim genderName As String
    Dim distance As String
    Dim styleNo As Integer
    
    Call get_race_title(prgNo, className, genderName, distance, styleNo)

    '''------ 春季室内only 県外をopenにする---
    ' Call set_open_to_kengai
    '------------------------------------------
    If is_relay(styleNo) Then
        Call fill_out_form_relay(prgNo, className, genderName, distance, styleNo, printenable)
    Else
        Call fill_out_form_kojin(prgNo, className, genderName, distance, styleNo, printenable)
    End If
    '-------- 春季室内only
    'Call reset_open
    '---------------------------------------------
End Sub
Sub BackOff()
    ''ActivePresentation.Slides(1).FollowMasterBackground = msoFalse
End Sub
Sub BackOn()
    ActivePresentation.Slides(1).FollowMasterBackground = msoTrue
End Sub
Sub print_it(dummy As String)
    ActivePresentation.Slides(1).FollowMasterBackground = msoFalse
    ActivePresentation.PrintOut From:=1, To:=1, Copies:=1
    ActivePresentation.Slides(1).FollowMasterBackground = msoTrue
End Sub




Sub name_text_box(boxNo As Integer, myName As String)
    Dim slide As slide
    Set slide = ActivePresentation.Slides(1)
    slide.Shapes(boxNo).Name = myName
End Sub
Sub show(txtBoxName As String, dispText As String)

    ' スライドの取得
    Dim slide As slide
    Dim shp As Shape
    Dim shapeExists As Boolean
    
    Set slide = ActivePresentation.Slides(1) ' was slideIndex
    On Error Resume Next
     Set shp = slide.Shapes(txtBoxName)
     shapeExists = Not shp Is Nothing
    On Error GoTo 0
    If shapeExists Then
        slide.Shapes(txtBoxName).TextFrame.TextRange = dispText
    End If
End Sub


Sub InitTextBox()
    Call DisplayTextBoxName("選手名")
    Call DisplayTextBoxName("所属")
    Call DisplayTextBoxName("クラス")
    Call DisplayTextBoxName("種目")
    Call DisplayTextBoxName("順位")
    Call DisplayTextBoxName("タイム")
End Sub
Sub DisplayTextBoxName(txtBoxName As String)
    Dim slide As slide
    Dim shp As Shape
    Dim i As Integer
    Dim slideIndex As Integer
    Dim shapeExists As Boolean
    
    Set slide = ActivePresentation.Slides(1)
    On Error Resume Next
     Set shp = slide.Shapes(txtBoxName)
     shapeExists = Not shp Is Nothing
    On Error GoTo 0
            
    If shapeExists Then
                ' TextBoxの名前をTextRangeに設定
        shp.TextFrame.TextRange = txtBoxName
    End If
End Sub





