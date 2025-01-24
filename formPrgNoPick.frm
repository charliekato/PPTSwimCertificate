VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPrgNoPick 
   Caption         =   "競技選択"
   ClientHeight    =   7344
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   9288
   OleObjectBlob   =   "formPrgNoPick.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formPrgNoPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  formPrgNoPick
'
Public LastRow As Integer
Private NextRunTime As Single
Private Running As Boolean

Private Sub btnAutoExe_Click()
    If Running = False Then
        btnAutoExe.Caption = "停止"
        lblPrintStatus.Caption = "自動実行中..."
        Call StartProcess
    Else
        btnAutoExe.Caption = "自動実行"
        lblPrintStatus.Caption = ""
        Call StopProcess
    End If
End Sub
Sub StartProcess()
    ' 実行フラグを設定
    Running = True
    ' 現在の時刻に5秒を加算
    NextRunTime = Timer + 5
    ' タスクの実行を開始
    ExecuteTask
End Sub

Sub StopProcess()
    ' 実行フラグを解除
    Running = False
    ' Debug.Print "プロセスを停止しました。"
End Sub

Function AlreadyPrinted(printPrgNo As Integer) As Boolean
    Dim sqlStr As String
    Dim myRecordset As New ADODB.Recordset

    On Error GoTo ErrorHandler

    ' SQL文の作成
    sqlStr = "SELECT 印刷状況, 印刷状況.競技番号 " & _
             "FROM 印刷状況 " & _
             "INNER JOIN プログラム ON プログラム.競技番号 = 印刷状況.競技番号 " & _
             "WHERE プログラム.表示用競技番号 = " & printPrgNo & _
             " and プログラム.大会番号 = " & HyouShow.EventNo & _
             " and 印刷状況.大会番号 = " & HyouShow.EventNo

    ' レコードセットを開く
    myRecordset.Open sqlStr, HyouShow.MyCon, adOpenStatic, adLockReadOnly

    ' レコードが存在しない場合のチェック
    If Not myRecordset.EOF Then
        If myRecordset!印刷状況 = 1 Then
            AlreadyPrinted = True
        Else
            AlreadyPrinted = False
        End If
    Else
        ' レコードが存在しない場合、Falseを返す
        AlreadyPrinted = False
    End If

Cleanup:
    ' レコードセットをクローズして解放
    If Not myRecordset Is Nothing Then
        If myRecordset.State = adStateOpen Then myRecordset.Close
        Set myRecordset = Nothing
    End If
    Exit Function

ErrorHandler:
    Debug.Print "エラーが発生しました: " & Err.Description
    AlreadyPrinted = False
    Resume Cleanup
End Function
Sub ExecuteTask()
    ' 処理を実行
    ''Debug.Print "タスクを実行しました: " & Now
    Dim printPrgNo As Integer
    printPrgNo = GetLastPrintPrgNo
    If printPrgNo > 0 Then
        If AlreadyPrinted(printPrgNo) = False Then
            Call PrintGo(printPrgNo)
        End If
    End If
    ' 処理を継続する場合
    If Running Then
        ' 次の実行時間が来るまで待機
        Do While Timer < NextRunTime
            DoEvents ' 他の処理を妨げないようにする
        Loop
        ' 次の実行時間を設定
        NextRunTime = Timer + 5
        ' タスクを再帰的に呼び出す
        ExecuteTask
    End If
End Sub




Function GetLastPrintPrgNo() As Integer
    Dim sqlString As String
    Dim myRecordset As New ADODB.Recordset
    Dim lastPrintPrgNo As Integer
    Dim yoketsuCode As Integer
    
    lastPrintPrgNo = 0
    
    sqlString = "select 表示用競技番号 , 進行フラグ, 予決コード  from プログラム " & _
                " where 大会番号 = " & HyouShow.EventNo & _
                " Order by 表示用競技番号"
    myRecordset.Open sqlString, HyouShow.MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordset.EOF
        If myRecordset!進行フラグ = 1 Then
            GoTo closeExit
        End If
        yoketsuCode = myRecordset!予決コード
        If yoketsuCode < 7 And yoketsuCode > 2 Then
            lastPrintPrgNo = myRecordset!表示用競技番号
        Else
            lastPrintPrgNo = 0
        End If
        myRecordset.MoveNext
    Loop
closeExit:
    GetLastPrintPrgNo = lastPrintPrgNo
    myRecordset.Close
    Set myRecordset = Nothing

End Function


Private Sub btnClose_Click()
    Unload Me
End Sub
'---- error ---
Private Sub btnPreView_Click()
    Dim printPrgNo As Integer
    On Error GoTo subEnd

    printPrgNo = CInt(Left(listPrg.Value, 3))
  
    Call fill_out_form(HyouShow.GetPrgNofromPrintPrgNo(printPrgNo), False)
subEnd:
End Sub


Sub SetPrintedFlag(target競技番号 As Integer)
    Dim cmd As Object
    Dim sql As String

    On Error GoTo ErrorHandler


    sql = "UPDATE 印刷状況 " & _
          "SET 印刷状況 = 1 " & _
          "WHERE 大会番号 = " & HyouShow.EventNo & " " & _
          "AND 競技番号 = " & target競技番号 & ";"


    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = HyouShow.MyCon ' 既存の接続を使用
        .CommandText = sql
        .Execute
    End With


    ' リソースを解放
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    ' エラー処理
    Debug.Print "エラーが発生しました: " & Err.Description
    Set cmd = Nothing
End Sub


Private Sub PrintGo(printPrgNo As Integer)
    Dim prgNo As Integer
    prgNo = HyouShow.GetPrgNofromPrintPrgNo(printPrgNo)
    Call fill_out_form(prgNo, True) ' if last argument is true then print else just preview
    Call CheckPrinted(printPrgNo)
End Sub
Private Sub btnPrint_Click()
    Dim printPrgNo As Integer
    On Error GoTo MyExit
    printPrgNo = CInt(Left(listPrg.Value, 3))
    Call PrintGo(printPrgNo)
MyExit:
End Sub


Sub CheckPrinted(printPrgNo As Integer)
    Dim prgNo As Integer
    prgNo = HyouShow.GetPrgNofromPrintPrgNo(printPrgNo)
    Call SetPrintedFlag(prgNo)
    Call SetDoneFlagOnList(printPrgNo)
End Sub
Sub SetDoneFlagOnList(printPrgNo As Integer)
    Dim targetIndex As Integer

    ' 指定された競技番号のインデックスを取得
    targetIndex = GetLineNoFromPrintPrgNo(printPrgNo)

    ' インデックスが 0 の場合は該当項目なしと判断
    If targetIndex = 0 Then
        MsgBox "Error: SetDoneFlagOnList 該当する項目が見つかりません。"
        Exit Sub
    End If
    listPrg.List(targetIndex, 3) = "済"


 
End Sub
Sub SelectItemByPrgNo(printPrgNo As Integer)
    Dim targetIndex As Integer

    ' 指定された競技番号のインデックスを取得
    targetIndex = GetLineNoFromPrintPrgNo(printPrgNo)

    ' インデックスが 0 の場合は該当項目なしと判断
    If targetIndex = 0 Then
        MsgBox "該当する項目が見つかりません。"
        Exit Sub
    End If

    ' 全ての選択をクリア（複数選択モードの場合）
    Dim i As Integer
    For i = 0 To listPrg.ListCount - 1
        listPrg.Selected(i) = False
    Next i

    ' 該当する項目を選択状態にする
    listPrg.Selected(targetIndex) = True

    ' 選択された項目の値を表示（オプション）
    MsgBox "選択された項目: " & listPrg.List(targetIndex)
End Sub


Function GetLineNoFromPrintPrgNo(printPrgNo As Integer) As Integer
    Dim row As Integer
    Dim itemText As String

    ' ListCount - 1 までループ
    For row = 0 To listPrg.ListCount - 1
        itemText = listPrg.List(row, 0)
        
        ' 項目が3文字以上である場合のみ処理
        If Len(itemText) >= 3 Then
            If CInt(Left(itemText, 3)) = printPrgNo Then
                GetLineNoFromPrintPrgNo = row
                Exit Function
            End If
        End If
    Next row

    ' 該当項目がない場合は 0 を返す
    GetLineNoFromPrintPrgNo = 0
End Function

Private Sub cbxJunniShowMethod1_Click()
    If cbxJunniShowMethod1.Value = True Then
        cbxJunniShowMethod2.Value = False
        cbxJunniShowMethod3.Value = False
    End If
End Sub
Private Sub cbxJunniShowMethod2_Click()
    If cbxJunniShowMethod2.Value = True Then
        cbxJunniShowMethod1.Value = False
        cbxJunniShowMethod3.Value = False
    End If
End Sub
Private Sub cbxJunniShowMethod3_Click()
    If cbxJunniShowMethod3.Value = True Then
        cbxJunniShowMethod2.Value = False
        cbxJunniShowMethod1.Value = False
    End If
End Sub

Private Sub listPrg_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnPreView_Click
    End If
End Sub

