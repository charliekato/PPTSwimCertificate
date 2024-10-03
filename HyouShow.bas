Attribute VB_Name = "HyouShow"
Option Explicit
Option Base 0

    Const JLIMIT = 1  '  何位まで賞状を出すか
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




Public Function class_exist(ByVal EventNo As Integer) As Boolean
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

Sub init_senshu(ByVal EventNo As Integer)

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
Sub get_race_title(ByVal EventNo As Integer, ByVal PrgNo As Integer, ByRef Class As String, _
            ByRef genderStr As String, ByRef distance As String, ByRef styleNo As Integer)
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    Dim classExist As Boolean
    classExist = class_exist(EventNo)
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

Sub fill_out_form(PrgNo As Integer, printEnable As Boolean)
    Dim rlist As Collection

    Dim junni As Integer
    Dim prevTime As String
'    Dim slideIndex As Integer
    Dim className As String
    Dim genderName As String
    Dim distance As String
    Dim styleNo As Integer
    Dim style As String

    Dim myRecordSet As New ADODB.Recordset
    Dim winnerName As String
    Dim myTime As String
    Dim myquery As String
    Call get_race_title(EventNo, PrgNo, className, genderName, distance, styleNo)
    junni = 0
    prevTime = ""

    If is_relay(styleNo) Then
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
            If prevTime <> myRecordSet!ゴール Then
                junni = junni + 1
                If junni > JLIMIT Then
                    Exit Do
                End If
                prevTime = myRecordSet!ゴール
            End If
            If formPrgNoPick.cbxName.Value Then
                Call show(junni, "選手名", myRecordSet!チーム名)
            Else
                Call show(junni, "選手名", "")
            End If
            If formPrgNoPick.cbxBelongsTo.Value Then
                Call show(junni, "所属", Swimmer(myRecordSet!第1泳者) & "・" & Swimmer(myRecordSet!第2泳者) & "・" & _
                     Swimmer(myRecordSet!第3泳者) & "・" & Swimmer(myRecordSet!第4泳者))
            Else
                Call show(junni, "所属", "")
            End If
            If formPrgNoPick.cbxJunni.Value Then
                Call show(junni, "順位", className)
            Else
                Call show(junni, "順位", "")
            End If
            If formPrgNoPick.cbxStyle.Value Then
                Call show(junni, "種目", genderName + distance + Shumoku(styleNo))
                Else
                Call show(junni, "種目", "")
            End If
            If formPrgNoPick.cbxTime.Value Then
                Call show(junni, "タイム", ConvertTimeFormat(myRecordSet!ゴール) + "  " + _
                      if_not_null_string(myRecordSet!新記録印刷マーク))
                Else
                Call show(junni, "タイム", "")
            End If
            If printEnable Then
                Call print_it("")
            End If
            myRecordSet.MoveNext
        Loop
    Else
        myquery = "SELECT 選手.氏名 as 氏名, 記録.ゴール as ゴール, 選手.所属名称1, 記録.新記録印刷マーク " & _
        "FROM 記録 " & _
        " inner join 選手 on 選手.選手番号 = 記録.選手番号 " & _
        " where 選手.大会番号=" & EventNo & " and 記録.競技番号 = " & PrgNo & _
        " and 記録.大会番号 = " & EventNo & " and 記録.事由入力ステータス=0 " & _
        " and 記録.選手番号>0 order by ゴール asc;"

        myRecordSet.Open myquery, MyCon, adOpenStatic, adLockReadOnly
        Do Until myRecordSet.EOF
            If prevTime <> myRecordSet!ゴール Then
                junni = junni + 1
                If junni > JLIMIT Then
                    Exit Do
                End If
            End If
            If formPrgNoPick.cbxName.Value Then
                Call show(junni, "選手名", myRecordSet!氏名)
            Else
                Call show(junni, "選手名", "")
            End If
            If formPrgNoPick.cbxBelongsTo.Value Then
                Call show(junni, "所属", myRecordSet!所属名称1)
            Else
                Call show(junni, "所属", "")
            End If
            If formPrgNoPick.cbxClass.Value Then
                Call show(junni, "クラス", className)
            Else
                Call show(junni, "クラス", "")
            End If
            If formPrgNoPick.cbxStyle.Value Then
                Call show(junni, "種目", genderName + distance + Shumoku(styleNo))
            Else
                Call show(junni, "種目", "")
            End If
            If formPrgNoPick.cbxTime.Value Then
                Call show(junni, "タイム", ConvertTimeFormat(myRecordSet!ゴール) + "  " + _
                    if_not_null_string(myRecordSet!新記録印刷マーク))
            Else
                Call show(junni, "タイム", "")
            End If
            If formPrgNoPick.cbxJunni.Value Then
                Call show(junni, "順位", junni & "位")
            Else
                Call show(junni, "順位", "")
            End If
            If printEnable Then
                Call print_it("")

            End If
            prevTime = myRecordSet!ゴール
            myRecordSet.MoveNext
        Loop
    End If
    ' クローズと解放
    myRecordSet.Close
    'MyCon.Close
    Set myRecordSet = Nothing
    'Set MyCon = Nothing

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
Sub show(slideIndex As Integer, txtBoxName As String, dispText As String)

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

Sub check_shape()
    Dim slide As slide
    Dim i As Integer
    Dim slideIndex As Integer
    ' スライドの取得 (テキストボックスが存在するスライド番号に設定)
    slideIndex = 1
    Set slide = ActivePresentation.Slides(slideIndex)

    ' すべてのシェイプをループ
    For i = 1 To slide.Shapes.Count

           ' slide.Shapes(i).Select
           ' MsgBox (" " & i & slide.Shapes(i).Name & ">" & slide.Shapes(i).TextFrame.TextRange)
           On Error Resume Next
        slide.Shapes(i).TextFrame.TextRange = ">>" & i & "<<"
    Next i
End Sub
