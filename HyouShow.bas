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
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    Dim rc As Boolean
    myquery = "select * from クラス where 大会番号 = " & EventNo
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockReadOnly
    If myRecordSet.EOF Then
        rc = False
    Else
        rc = True
    End If
    myRecordSet.Close
    Set myRecordSet = Nothing
    class_exist = rc

End Function

Sub init_senshu(dummy As String)

    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    Dim maxSwimmerNo As Integer
    myquery = "SELECT MAX(選手番号) as MAX from 選手 where 大会番号 = " & EventNo
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    maxSwimmerNo = myRecordSet!Max
    
    ReDim Swimmer(maxSwimmerNo)
    myRecordSet.Close
    myquery = "SELECT 氏名, 選手番号 from 選手 where 大会番号 = " & EventNo
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordSet.EOF
        Swimmer(myRecordSet!選手番号) = myRecordSet!氏名
        myRecordSet.MoveNext
    Loop
    myRecordSet.Close
    Set myRecordSet = Nothing
End Sub

Public Function get_prgNo(printPrgNo As Integer) As Integer
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    myquery = "select 競技番号 from プログラム where 表示用競技番号=" & printPrgNo & _
              "and 大会番号= " & EventNo & ";"
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    get_prgNo = if_not_null(myRecordSet!競技番号)
    myRecordSet.Close
    Set myRecordSet = Nothing
End Function
Sub get_race_title(ByVal PrgNo As Integer, ByRef Class As String, _
            ByRef genderStr As String, ByRef distance As String, ByRef styleNo As Integer)
    Dim myRecordSet As New ADODB.Recordset
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
              " プログラム.競技番号 = " & PrgNo & ";"
    Else
        myquery = "SELECT  プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, プログラム.種目コード as 種目 FROM プログラム " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & EventNo & " AND " + _
              " プログラム.競技番号 = " & PrgNo & ";"
    End If
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordSet.EOF
        If classExist Then
            Class = myRecordSet!クラス
        Else
            Class = ""
        End If
        genderStr = Gender(myRecordSet!性別)
        distance = myRecordSet!距離
        styleNo = myRecordSet!種目
        myRecordSet.MoveNext
    Loop
    myRecordSet.Close
    Set myRecordSet = Nothing
                
              
    
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

Sub fill_out_form_relay(PrgNo As Integer, className As String, _
                genderName As String, distance As String, styleNo As Integer, printenable As Boolean)
    Dim myquery As String
    Dim junni As Integer
    Dim junnib As Integer
    Dim prevTime As String



    Dim myRecordSet As New ADODB.Recordset
    Dim winnerName As String
    Dim myTime As String
    junni = 0
    junnib = 0
    prevTime = ""
    myquery = "SELECT リレーチーム.チーム名 as チーム名, 記録.ゴール as ゴール, " & _
            "記録.第１泳者, 記録.第２泳者, 記録.第３泳者, 記録.第４泳者, 記録.新記録印刷マーク " & _
            "FROM 記録 " & _
            " inner join リレーチーム on リレーチーム.チーム番号 = 記録.選手番号 " & _
            " where   記録.競技番号 = " & PrgNo & _
            " and 記録.大会番号 = " & EventNo & " and 記録.事由入力ステータス=0 " & _
            " and リレーチーム.大会番号 = " & EventNo & _
            " order by ゴール asc;"

    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordSet.EOF
        junnib = junnib + 1
        If prevTime <> myRecordSet!ゴール Then
            junni = junnib
            If junni > CInt(formPrgNoPick.tbxJunniLast) Then
                Exit Do
            End If
            If junni < CInt(formPrgNoPick.tbxJunniTop) Then
                GoTo DOLOOPEND
            End If
            prevTime = myRecordSet!ゴール
        End If
        Call fill_name(myRecordSet!チーム名)

        Call fill_shozoku(Swimmer(myRecordSet!第1泳者) & "・" & Swimmer(myRecordSet!第2泳者) & "・" & _
                     Swimmer(myRecordSet!第3泳者) & "・" & Swimmer(myRecordSet!第4泳者))
        
        Call fill_junni(junni)
         Call fill_class(className)
        Call fill_shumoku(genderName + distance + Shumoku(styleNo))
        Call fill_time(myRecordSet!ゴール + " " + _
            if_not_null_string(myRecordSet!新記録印刷マーク))
        If printenable Then
            Call print_it("")
        End If
DOLOOPEND:
        
        myRecordSet.MoveNext
    Loop
                    ' クローズと解放
    myRecordSet.Close
    'MyCon.Close
    Set myRecordSet = Nothing
    'Set MyCon = Nothing
End Sub

Sub fill_out_form_kojin(PrgNo As Integer, className As String, _
                    genderName As String, distance As String, styleNo As Integer, printenable As Boolean)
    Dim myquery As String
    Dim junni As Integer
    Dim junnib As Integer
    Dim prevTime As String


    Dim myRecordSet As New ADODB.Recordset
    Dim winnerName As String
    Dim myTime As String
    junni = 0
    junnib = 0
    prevTime = ""
    
    myquery = "SELECT 選手.氏名 as 氏名, 記録.ゴール as ゴール, 選手.所属名称1, 記録.新記録印刷マーク " & _
        "FROM 記録 " & _
        " inner join 選手 on 選手.選手番号 = 記録.選手番号 " & _
        " where 選手.大会番号=" & EventNo & " and 記録.競技番号 = " & PrgNo & _
        " and 記録.大会番号 = " & EventNo & " and 記録.事由入力ステータス=0 " & _
        " and 記録.選手番号>0 order by ゴール asc;"

    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordSet.EOF
        junnib = junnib + 1
        If IsNull(myRecordSet!ゴール) Or myRecordSet!ゴール = "" Then
            MsgBox ("該当データがありません。たぶんレースが終わっていないと思われます。")
            Exit Do
        End If
        If prevTime <> myRecordSet!ゴール Then

            junni = junnib
            If junni > CInt(formPrgNoPick.tbxJunniLast) Then
                Exit Do
            End If
            If junni < CInt(formPrgNoPick.tbxJunniTop) Then
                GoTo LOOPEND2
            End If
            prevTime = myRecordSet!ゴール
        End If
        Call fill_name(myRecordSet!氏名)
        Call fill_shozoku(myRecordSet!所属名称1)
        Call fill_class(className)
        Call fill_shumoku(genderName + distance + Shumoku(styleNo))
        Call fill_time(ConvertTimeFormat(myRecordSet!ゴール) + " " + _
                 if_not_null_string(myRecordSet!新記録印刷マーク))
        Call fill_junni(junni)
        If printenable Then
            Call print_it("")
        End If
LOOPEND2:
       
        myRecordSet.MoveNext
    Loop
            ' クローズと解放
    myRecordSet.Close
    'MyCon.Close
    Set myRecordSet = Nothing
    'Set MyCon = Nothing
End Sub

Sub fill_time(myTime As String)
    If formPrgNoPick.cbxTime.Value Then
        Call show("タイム", myTime)
    Else
        Call show("タイム", "")
    End If
End Sub

Sub fill_class(className As String)
    If formPrgNoPick.cbxClass.Value Then
        Call show("クラス", className)
    Else
        Call show("クラス", "")
    End If
End Sub
Sub fill_shumoku(Shumoku As String)
    If formPrgNoPick.cbxStyle.Value Then
        Call show("種目", Shumoku)
    Else
        Call show("種目", "")
    End If
End Sub

Sub fill_junni(junni As Integer)
    If formPrgNoPick.cbxJunni.Value Then
        If formPrgNoPick.cbxJunniShowMethod1.Value Then
            Call show("順位", "" & junni)
        ElseIf formPrgNoPick.cbxJunniShowMethod2.Value Then
            Call show("順位", "第" & junni & "位")
        ElseIf formPrgNoPick.cbxJunniShowMethod3.Value Then
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
    If formPrgNoPick.cbxName.Value Then
        Call show("選手名", myName)
    Else
        Call show("選手名", "")
    End If
End Sub

Sub fill_shozoku(shozoku As String)
    If formPrgNoPick.cbxBelongsTo.Value Then
        Call show("所属", shozoku)
    Else
        Call show("所属", "")
    End If
End Sub

Sub fill_out_form(PrgNo As Integer, printenable As Boolean)

    Dim myquery As String


    Dim className As String
    Dim genderName As String
    Dim distance As String
    Dim styleNo As Integer
    
    Call get_race_title(PrgNo, className, genderName, distance, styleNo)


    If is_relay(styleNo) Then
        Call fill_out_form_relay(PrgNo, className, genderName, distance, styleNo, printenable)
    Else
        Call fill_out_form_kojin(PrgNo, className, genderName, distance, styleNo, printenable)
    End If


End Sub
Sub BackOff()
    ActivePresentation.Slides(1).FollowMasterBackground = msoFalse
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

