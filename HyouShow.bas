Attribute VB_Name = "HyouShow"
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Option Explicit
Option Base 0
Const DefaultServerName = "localhost"
Const DebugMode As Boolean = False   ' false にしておくこと!!


    Public MyCon As ADODB.Connection
    Public EventNo As Integer

    Public Gender(4) As String
    Public Shumoku(8) As String
    Public Swimmer() As String
    Public MaxClassNo As Integer

    Public ClassTable() As String

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
    ss.txtBoxServerName = DefaultServerName
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


Public Function GetPrgNofromPrintPrgNo(printPrgNo As Integer) As Integer
    Dim myRecordset As New ADODB.Recordset
    Dim myQuery As String
    myQuery = "select 競技番号 from プログラム where 表示用競技番号=" & printPrgNo & _
              "and 大会番号= " & EventNo & ";"
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    GetPrgNofromPrintPrgNo = if_not_null(myRecordset!競技番号)
    myRecordset.Close
    Set myRecordset = Nothing
End Function



Sub get_race_title(ByVal prgNo As Integer, ByRef Class As String, _
            ByRef genderStr As String, ByRef distance As String, ByRef styleNo As Integer)
    Dim myRecordset As New ADODB.Recordset
    Dim myQuery As String
    Dim classBasedRace As Boolean
    classBasedRace = formEventNoPick.class_based_race("")
     
    
    If formEventNoPick.class_based_race("") Then
        myQuery = "SELECT クラス.クラス名称 as クラス, プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, プログラム.種目コード as 種目 FROM プログラム " + _
              " INNER JOIN クラス ON クラス.クラス番号=プログラム.クラス番号 " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & EventNo & " AND " + _
              " クラス.大会番号 = " & EventNo & " AND " & _
              " プログラム.競技番号 = " & prgNo & ";"
    Else
        myQuery = "SELECT  プログラム.性別コード as 性別, " & _
              "距離.距離 as 距離, プログラム.種目コード as 種目 FROM プログラム " + _
              " INNER JOIN 距離 ON 距離.距離コード = プログラム.距離コード " + _
              " WHERE プログラム.大会番号 = " & EventNo & " AND " + _
              " プログラム.競技番号 = " & prgNo & ";"
    End If
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        If classBasedRace Then
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

Function RelayDistance(distance As String) As String
    If distance = " 200m" Then
        RelayDistance = " 4×50m"
        Exit Function
    End If
    If distance = " 400m" Then
        RelayDistance = " 4×100m"
        Exit Function
    End If
    If distance = " 800m" Then
        RelayDistance = " 4×200m"
        Exit Function
    End If
End Function






Function kenmei(shozoku As String)
    kenmei = shozoku
    If shozoku = "大　阪" Then
        kenmei = kenmei + "　府"
        Exit Function
    End If
    If kenmei = "東　京" Then
        kenmei = kenmei + "　都"
        Exit Function
    End If
    If kenmei = "京　都" Then
        kenmei = kenmei + "　府"
        Exit Function
    End If
    If kenmei = "北海道" Then
        Exit Function
    End If
    If kenmei = "鹿児島" Then
        kenmei = kenmei + "県"
        Exit Function
    End If
    If kenmei = "神奈川" Then
        kenmei = kenmei + "県"
        Exit Function
    End If
    If kenmei = "和歌山" Then
        kenmei = kenmei + "県"
        Exit Function
    End If
    kenmei = kenmei + "　県"
End Function





Sub fill_time(myTime As String)
    If formOption.cbxTime.Value Then
        Call show("タイム", myTime)
    Else
        Call show("タイム", "")
    End If
End Sub

Sub fill_class(className As String)
    If formOption.cbxClass.Value Then
        Call show("クラス", className)
    Else
        Call show("クラス", "")
    End If
End Sub
Sub fill_shumoku(Shumoku As String)
    If formOption.cbxStyle.Value Then
        Call show("種目", Shumoku)
    Else
        Call show("種目", "")
    End If
End Sub

Sub fill_junni(junni As Integer)
    If formOption.cbxJunni.Value Then
        If formOption.cbxJunniShowMethod1.Value Then
            Call show("順位", "" & junni)
        ElseIf formOption.cbxJunniShowMethod2.Value Then
            Call show("順位", "第 " & junni & " 位")
        ElseIf formOption.cbxJunniShowMethod3.Value Then
            If junni = 1 Then
                Call show("順位", "優勝")
            Else
                Call show("順位", "第 " & junni & " 位")
            End If
        End If
    Else
        Call show("順位", "")
    End If
End Sub




Sub fill_name(myName As String)
    If formOption.cbxName.Value Then
        Call show("選手名", myName)
    Else
        Call show("選手名", "")
    End If
End Sub

Sub fill_shozoku(shozoku As String)
    If formOption.cbxBelongsTo.Value Then
        If formOption.cbxKenmeiMode Then
            Call show("所属", kenmei(shozoku))
        Else
            Call show("所属", shozoku)
        End If
    Else
        Call show("所属", "")
    End If
End Sub


Sub init_class(dummy As String)
    Dim myQuery As String


    

    Dim myRecordset As New ADODB.Recordset
    myQuery = "SELECT MAX(クラス番号) as MAX from クラス where 大会番号 = " & EventNo
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly

    MaxClassNo = myRecordset!Max
    
    ReDim ClassTable(MaxClassNo)
    myRecordset.Close
    Set myRecordset = Nothing
    
    myQuery = " select クラス番号,クラス名称 from クラス where 大会番号=" & EventNo
    myRecordset.Open myQuery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        ClassTable(CInt(myRecordset!クラス番号)) = myRecordset!クラス名称
                
        myRecordset.MoveNext
    Loop
    myRecordset.Close
    Set myRecordset = Nothing
End Sub
Function fill_out_form2(prgNo As Integer, printenable As Boolean) As Boolean
    Dim myQuery As String

    Dim myRecordset As New ADODB.Recordset
    Dim winnerName As String
    Dim myTime As String
    fill_out_form2 = True

    myQuery = _
"    IF EXISTS (select 1 from クラス where 大会番号=" & EventNo & ") " & _
"    BEGIN" & _
"        SELECT " & _
"           プログラム.種目コード," & _
"                   クラス.クラス名称,     " & _
"           case プログラム.性別コード  when 1 then '男子'  when 2 then '女子'   when 3 then '混成'  when 4 then '混合'" & _
"                   end as 性別, " & _
"           距離.距離," & _
"           種目.種目,  " & _
"           rank() over (partition by 記録.競技番号, 記録.新記録判定クラス " & _
"                           ORDER BY 記録.事由表示, 記録.ゴール ASC) as 順位," & _
"           case    WHEN プログラム.種目コード < 6 THEN 選手.氏名  ELSE 選手1.氏名  END AS 氏名1, " & _
"           case    WHEN プログラム.種目コード < 6 THEN ''         ELSE 選手2.氏名  END AS 氏名2, " & _
"           case    WHEN プログラム.種目コード < 6 THEN ''         ELSE 選手3.氏名  END AS 氏名3, " & _
"           case    WHEN プログラム.種目コード < 6 THEN ''         ELSE 選手4.氏名  END AS 氏名4, " & _
"           case    WHEN プログラム.種目コード < 6 THEN "
    myQuery = myQuery & _
"               case 選手.主所属 when 2 then 選手.所属名称2 " & _
"                                when 3 then 選手.所属名称3" & _
"                                else 選手.所属名称1 end" & _
"                   ELSE リレーチーム.チーム名     END AS 所属," & _
"           case  WHEN プログラム.種目コード < 6 THEN " & _
"                case 選手.主所属 when 2 then 所属2.所属名正式 " & _
"                            when 3 then 所属3.所属名正式 " & _
"                       else 所属1.所属名正式 end " & _
"                 else リレー所属.所属名正式 end as 所属名正式, "
    myQuery = myQuery & _
"           記録.ゴール, " & _
"           記録.新記録印刷マーク" & _
"       from 記録 " & _
"       LEFT JOIN 選手 ON 選手.選手番号 = 記録.選手番号  and 選手.大会番号=記録.大会番号" & _
"       left join リレーチーム on リレーチーム.チーム番号=記録.選手番号" & _
"             and リレーチーム.大会番号=記録.大会番号" & _
"       LEFT JOIN 選手 as 選手1 ON 選手1.選手番号 = 記録.第１泳者 and 選手1.大会番号=記録.大会番号" & _
"       LEFT join 選手 as 選手2 on 選手2.選手番号 = 記録.第２泳者 and 選手2.大会番号=記録.大会番号" & _
"       LEFT join 選手 as 選手3 on 選手3.選手番号 = 記録.第３泳者 and 選手3.大会番号=記録.大会番号" & _
"       LEFT join 選手 as 選手4 on 選手4.選手番号 = 記録.第４泳者 and 選手4.大会番号=記録.大会番号" & _
"       inner  join プログラム on プログラム.競技番号=記録.競技番号 " & _
"            and プログラム.大会番号=記録.大会番号" & _
"       inner join 距離 on 距離.距離コード=プログラム.距離コード" & _
"       inner join 種目 on 種目.種目コード=プログラム.種目コード" & _
"       inner join クラス on クラス.大会番号=記録.大会番号" & _
"                        and クラス.クラス番号=記録.新記録判定クラス" & _
"       left join 所属 as 所属1 on 所属1.所属番号=選手.所属番号1 and 所属1.大会番号=記録.大会番号 " & _
"       left join 所属 as 所属2 on 所属2.所属番号=選手.所属番号2 and 所属2.大会番号=記録.大会番号 " & _
"       left join 所属 as 所属3 on 所属3.所属番号=選手.所属番号3 and 所属3.大会番号=記録.大会番号 " & _
"       left join 所属 as リレー所属 on リレー所属.所属番号=リレーチーム.所属番号 and リレー所属.大会番号=記録.大会番号 " & _
"       WHERE  記録.大会番号= " & EventNo & _
"       　　and 記録.選手番号>0" & _
"           and プログラム.表示用競技番号=" & prgNo & _
"           and 記録.事由入力ステータス=0    and 記録.水路 < 11  end"
    myQuery = myQuery & _
"    else begin" & _
"        SELECT " & _
"           プログラム.種目コード," & _
"           '' as クラス名称 , " & _
"           case プログラム.性別コード " & _
"                 when 1 then '男子'" & _
"                 when 2 then '女子'" & _
"                 when 3 then '混成'" & _
"                 when 4 then '混合'" & _
"           　end as 性別, " & _
"           距離.距離," & _
"           種目.種目," & _
"           rank() over (partition by 記録.競技番号 " & _
"                    ORDER BY 記録.事由表示, 記録.ゴール ASC) as 順位," & _
"           case    WHEN プログラム.種目コード < 6 THEN 選手.氏名  ELSE 選手1.氏名  END AS 氏名1, " & _
"           case    WHEN プログラム.種目コード < 6 THEN ''         ELSE 選手2.氏名  END AS 氏名2, " & _
"           case    WHEN プログラム.種目コード < 6 THEN ''         ELSE 選手3.氏名  END AS 氏名3, " & _
"           case    WHEN プログラム.種目コード < 6 THEN ''         ELSE 選手4.氏名  END AS 氏名4, " & _
"           case    WHEN プログラム.種目コード < 6 THEN "
    myQuery = myQuery & _
"           case 選手.主所属 when 2 then 選手.所属名称2 " & _
"                            when 3 then 選手.所属名称3" & _
"                            else 選手.所属名称1 end" & _
"     ELSE リレーチーム.チーム名     END AS 所属," & _
"           case  WHEN プログラム.種目コード < 6 THEN " & _
"                case 選手.主所属 when 2 then 所属2.所属名正式 " & _
"                            when 3 then 所属3.所属名正式 " & _
"                       else 所属1.所属名正式 end " & _
"                 else リレー所属.所属名正式 end as 所属名正式, " & _
"           記録.ゴール, " & _
"           記録.新記録印刷マーク" & _
"       from 記録 "
    myQuery = myQuery & _
"       INNER JOIN 選手 ON 選手.選手番号 = 記録.選手番号 " & _
"                and 選手.大会番号=記録.大会番号" & _
"       LEFT JOIN リレーチーム on リレーチーム.チーム番号=記録.選手番号" & _
"                   and リレーチーム.大会番号=記録.大会番号" & _
"       LEFT JOIN 選手 as 選手1 ON 選手1.選手番号 = 記録.第１泳者 and 選手1.大会番号=記録.大会番号" & _
"       LEFT join 選手 as 選手2 on 選手2.選手番号 = 記録.第２泳者 and 選手2.大会番号=記録.大会番号" & _
"       LEFT join 選手 as 選手3 on 選手3.選手番号 = 記録.第３泳者 and 選手3.大会番号=記録.大会番号" & _
"       LEFT join 選手 as 選手4 on 選手4.選手番号 = 記録.第４泳者 and 選手4.大会番号=記録.大会番号" & _
"       inner join プログラム on プログラム.競技番号=記録.競技番号 " & _
"           and プログラム.大会番号=記録.大会番号" & _
"       inner join 距離 on 距離.距離コード=プログラム.距離コード" & _
"       inner join 種目 on 種目.種目コード=プログラム.種目コード" & _
"       left join 所属 as 所属1 on 所属1.所属番号=選手.所属番号1 and 所属1.大会番号=記録.大会番号 " & _
"       left join 所属 as 所属2 on 所属2.所属番号=選手.所属番号2 and 所属2.大会番号=記録.大会番号 " & _
"       left join 所属 as 所属3 on 所属3.所属番号=選手.所属番号3 and 所属3.大会番号=記録.大会番号 " & _
"       left join 所属 as リレー所属 on リレー所属.所属番号=リレーチーム.所属番号 and リレー所属.大会番号=記録.大会番号 " & _
"     WHERE  記録.大会番号= " & EventNo & _
"           and 記録.選手番号>0     " & _
"           and プログラム.表示用競技番号=" & prgNo & _
"           and 記録.事由入力ステータス=0" & _
"           and 記録.水路 < 11" & _
"    end;"
    Dim junni As Integer
    Dim relayMember As String
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordset.EOF

        If IsNull(myRecordset!ゴール) Or myRecordset!ゴール = "" Then
            MsgBox ("該当データがありません。たぶんレースが終わっていないと思われます。")
            fill_out_form2 = False
        Exit Do
        End If
        junni = CInt(myRecordset!順位)
        If junni > CInt(formPrgNoPick.tbxJunniLast) Then
            GoTo DOLOOPEND
        End If
        If junni < CInt(formPrgNoPick.tbxJunniTop) Then
            GoTo DOLOOPEND
        End If
        If CInt(myRecordset!種目コード) > 5 Then
            relayMember = myRecordset!氏名1 & "    " & myRecordset!氏名2 & vbCrLf & _
                          myRecordset!氏名3 & "    " & myRecordset!氏名4
            Call fill_name(relayMember)
        Else
            Call fill_name(myRecordset!氏名1)
        End If
        Call fill_shozoku(myRecordset!所属)
        
        Call fill_junni(myRecordset!順位)
        If formOption.cbxShumokuWithClass.Value Then
            If CInt(myRecordset!種目コード) > 5 Then
                Call fill_shumoku(myRecordset!クラス名称 + myRecordset!性別 + RelayDistance(myRecordset!距離) + myRecordset!種目)
            Else
                Call fill_shumoku(myRecordset!クラス名称 + myRecordset!性別 + myRecordset!距離 + myRecordset!種目)
            End If
        Else
            Call fill_class(myRecordset!クラス名称)
            If CInt(myRecordset!種目コード) > 5 Then
                Call fill_shumoku(myRecordset!性別 + RelayDistance(myRecordset!距離) + myRecordset!種目)
            Else
                Call fill_shumoku(myRecordset!性別 + myRecordset!距離 + myRecordset!種目)
            End If
        End If
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

    
End Function

Sub printMM()
    Call print_it("")
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





