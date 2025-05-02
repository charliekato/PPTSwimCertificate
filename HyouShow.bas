Attribute VB_Name = "HyouShow"
Option Explicit
Option Base 0

    Public MyCon As ADODB.Connection
    Public EventNo As Integer

    Public Gender(4) As String
    Public Shumoku(8) As String
    Public Swimmer() As String



Sub init_gender(dummy As String)
    Gender(1) = "�j�q"
    Gender(2) = "���q"
    Gender(3) = "����"
End Sub
Sub init_shumoku(dummy As String)
    Shumoku(1) = "���R�`"
    Shumoku(2) = "�w�j��"
    Shumoku(3) = "���j��"
    Shumoku(4) = "�o�^�t���C"
    Shumoku(5) = "�l���h���["
    Shumoku(6) = "�t���[�����["
    Shumoku(7) = "���h���[�����["
End Sub

Sub �܏�쐬()
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
    myquery = "select * from �N���X where ���ԍ� = " & EventNo
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
    
    sql1 = "UPDATE �L�^ " & _
              " Set �L�^.�I�[�v�� = 1 " & _
              " From �L�^ " & _
              " INNER JOIN �v���O���� on �v���O����.���Z�ԍ� = �L�^.���Z�ԍ� " & _
              " INNER JOIN �I�� ON �L�^.�I��ԍ� = �I��.�I��ԍ� " & _
              " WHERE �L�^.���ԍ� = " & EventNo & " And �I��.���ԍ� = " & EventNo & _
              " And �v���O����.���ԍ� = " & EventNo & _
              " and �v���O����.��ڃR�[�h < 6 " & _
              " and �I��.�����c�̔ԍ� <> 25; "
              
    sql2 = "UPDATE �L�^ " & _
           "SET �L�^.�I�[�v�� = 1 " & _
           "FROM �L�^ " & _
           "INNER JOIN �v���O���� ON �v���O����.���Z�ԍ� = �L�^.���Z�ԍ� " & _
           "INNER JOIN �����[�`�[�� ON �L�^.�I��ԍ� = �����[�`�[��.�`�[���ԍ� " & _
           "WHERE �L�^.���ԍ� = " & EventNo & " AND �����[�`�[��.���ԍ� = " & _
            EventNo & " AND �v���O����.���ԍ� = " & EventNo & _
           "AND �v���O����.��ڃR�[�h > 5 " & _
           "AND �����[�`�[��.�����c�̔ԍ� <> 25;"
           
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
    
    sql = "UPDATE �L�^" & _
   " Set �L�^.�I�[�v�� = 0" & _
   " From �L�^" & _
   " INNER JOIN �I�� ON �L�^.�I��ԍ� = �I��.�I��ԍ� " & _
   " WHERE �L�^.���ԍ� = " & EventNo & " AND �I��.���ԍ�=" & EventNo
              

           
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
    myquery = "SELECT MAX(�I��ԍ�) as MAX from �I�� where ���ԍ� = " & EventNo
    myRecordset.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    maxSwimmerNo = myRecordset!Max
    
    ReDim Swimmer(maxSwimmerNo)
    myRecordset.Close
    myquery = "SELECT ����, �I��ԍ� from �I�� where ���ԍ� = " & EventNo
    myRecordset.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        Swimmer(myRecordset!�I��ԍ�) = myRecordset!����
        myRecordset.MoveNext
    Loop
    myRecordset.Close
    Set myRecordset = Nothing
End Sub

Public Function GetPrgNofromPrintPrgNo(printPrgNo As Integer) As Integer
    Dim myRecordset As New ADODB.Recordset
    Dim myquery As String
    myquery = "select ���Z�ԍ� from �v���O���� where �\���p���Z�ԍ�=" & printPrgNo & _
              "and ���ԍ�= " & EventNo & ";"
    myRecordset.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    GetPrgNofromPrintPrgNo = if_not_null(myRecordset!���Z�ԍ�)
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
        myquery = "SELECT �N���X.�N���X���� as �N���X, �v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, �v���O����.��ڃR�[�h as ��� FROM �v���O���� " + _
              " INNER JOIN �N���X ON �N���X.�N���X�ԍ�=�v���O����.�N���X�ԍ� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & EventNo & " AND " + _
              " �N���X.���ԍ� = " & EventNo & " AND " & _
              " �v���O����.���Z�ԍ� = " & prgNo & ";"
    Else
        myquery = "SELECT  �v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, �v���O����.��ڃR�[�h as ��� FROM �v���O���� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & EventNo & " AND " + _
              " �v���O����.���Z�ԍ� = " & prgNo & ";"
    End If
    myRecordset.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        If classExist Then
            Class = myRecordset!�N���X
        Else
            Class = ""
        End If
        genderStr = Gender(myRecordset!����)
        distance = myRecordset!����
        styleNo = myRecordset!���
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
        If thisResult.���� = rank Then
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
    
    ' �R�����ƃh�b�g�̈ʒu��T��
    colonPos = InStr(timeString, ":")
    dotPos = InStr(timeString, ".")
    milliseconds = Mid(timeString, dotPos + 1)
    
    ' ��, �b, �~���b�𒊏o
    If colonPos > 0 Then
    minutes = Mid(timeString, 1, colonPos - 1)
    seconds = Mid(timeString, colonPos + 1, dotPos - colonPos - 1)
    
    ' �ϊ����ꂽ���Ԃ�Ԃ�
    ConvertTimeFormat = minutes & "��" & seconds & "�b" & milliseconds
    Else
        seconds = Trim(Mid(timeString, 1, dotPos - 1))
        ConvertTimeFormat = seconds & "�b" & milliseconds
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
    myquery = "SELECT �����[�`�[��.�`�[���� as �`�[����, �L�^.�S�[�� as �S�[��, " & _
            "�L�^.��P�j��, �L�^.��Q�j��, �L�^.��R�j��, �L�^.��S�j��, �L�^.�V�L�^����}�[�N " & _
            "FROM �L�^ " & _
            " inner join �����[�`�[�� on �����[�`�[��.�`�[���ԍ� = �L�^.�I��ԍ� " & _
            " where   �L�^.���Z�ԍ� = " & prgNo & _
            " and �L�^.���ԍ� = " & EventNo & " and �L�^.���R���̓X�e�[�^�X=0 " & _
            " and �����[�`�[��.���ԍ� = " & EventNo & _
            " and �L�^.�I�[�v�� = 0 " & _
            " order by �S�[�� asc;"

    myRecordset.Open myquery, MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordset.EOF
        junnib = junnib + 1
        If prevTime <> myRecordset!�S�[�� Then
            junni = junnib
            If junni > CInt(FormOption.tbxJunniLast) Then
                Exit Do
            End If
            If junni < CInt(FormOption.tbxJunniTop) Then
                GoTo DOLOOPEND
            End If
            prevTime = myRecordset!�S�[��
        End If
        Call fill_name(myRecordset!�`�[����)

        Call fill_shozoku(Swimmer(myRecordset!��1�j��) & "�E" & Swimmer(myRecordset!��2�j��) & "�E" & _
                     Swimmer(myRecordset!��3�j��) & "�E" & Swimmer(myRecordset!��4�j��))
        
        Call fill_junni(junni)
         Call fill_class(className)
        Call fill_shumoku(genderName + distance + Shumoku(styleNo))
        Call fill_time(ConvertTimeFormat(myRecordset!�S�[��) + " " + _
            if_not_null_string(myRecordset!�V�L�^����}�[�N))
        If printenable Then
            Call print_it("")
        End If
DOLOOPEND:
        
        myRecordset.MoveNext
    Loop
                    ' �N���[�Y�Ɖ��
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
    
    myquery = "SELECT �I��.���� as ����, �L�^.�S�[�� as �S�[��, �I��.��������1, �L�^.�V�L�^����}�[�N " & _
        "FROM �L�^ " & _
        " inner join �I�� on �I��.�I��ԍ� = �L�^.�I��ԍ� " & _
        " where �I��.���ԍ�=" & EventNo & " and �L�^.���Z�ԍ� = " & prgNo & _
        " and �L�^.���ԍ� = " & EventNo & " and �L�^.���R���̓X�e�[�^�X=0 " & _
        " and �L�^.�I�[�v�� = 0 " & _
        " and �L�^.�I��ԍ�>0 order by �S�[�� asc;"

    myRecordset.Open myquery, MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordset.EOF
        junnib = junnib + 1
        If IsNull(myRecordset!�S�[��) Or myRecordset!�S�[�� = "" Then
            MsgBox ("�Y���f�[�^������܂���B���Ԃ񃌁[�X���I����Ă��Ȃ��Ǝv���܂��B")
            Exit Do
        End If
        If prevTime <> myRecordset!�S�[�� Then

            junni = junnib
            If junni > CInt(FormOption.tbxJunniLast) Then
                Exit Do
            End If
            If junni < CInt(FormOption.tbxJunniTop) Then
                GoTo LOOPEND2
            End If
            prevTime = myRecordset!�S�[��
        End If
        Call fill_name(myRecordset!����)
        Call fill_shozoku(myRecordset!��������1)
        Call fill_class(className)
        Call fill_shumoku(genderName + distance + Shumoku(styleNo))
        Call fill_time(ConvertTimeFormat(myRecordset!�S�[��) + " " + _
                 if_not_null_string(myRecordset!�V�L�^����}�[�N))
        Call fill_junni(junni)
        If printenable Then
            Call print_it("")
        End If
LOOPEND2:
       
        myRecordset.MoveNext
    Loop
            ' �N���[�Y�Ɖ��
    myRecordset.Close
    'MyCon.Close
    Set myRecordset = Nothing
    'Set MyCon = Nothing
End Sub

Sub fill_time(myTime As String)
    If FormOption.cbxTime.Value Then
        Call show("�^�C��", myTime)
    Else
        Call show("�^�C��", "")
    End If
End Sub

Sub fill_class(className As String)
    If FormOption.cbxClass.Value Then
        Call show("�N���X", className)
    Else
        Call show("�N���X", "")
    End If
End Sub
Sub fill_shumoku(Shumoku As String)
    If FormOption.cbxStyle.Value Then
        Call show("���", Shumoku)
    Else
        Call show("���", "")
    End If
End Sub

Sub fill_junni(junni As Integer)
    If FormOption.cbxJunni.Value Then
        If FormOption.cbxJunniShowMethod1.Value Then
            Call show("����", "" & junni)
        ElseIf FormOption.cbxJunniShowMethod2.Value Then
            Call show("����", "��" & junni & "��")
        ElseIf FormOption.cbxJunniShowMethod3.Value Then
            If junni = 1 Then
                Call show("����", "�D��")
            Else
                Call show("����", "��" & junni & "��")
            End If
        End If
    Else
        Call show("����", "")
    End If
End Sub

Sub fill_name(myName As String)
    If FormOption.cbxName.Value Then
        Call show("�I�薼", myName)
    Else
        Call show("�I�薼", "")
    End If
End Sub

Sub fill_shozoku(shozoku As String)
    If FormOption.cbxBelongsTo.Value Then
        Call show("����", shozoku)
    Else
        Call show("����", "")
    End If
End Sub

Sub fill_out_form(prgNo As Integer, printenable As Boolean)

    Dim myquery As String


    Dim className As String
    Dim genderName As String
    Dim distance As String
    Dim styleNo As Integer
    
    Call get_race_title(prgNo, className, genderName, distance, styleNo)

    '''------ �t�G����only ���O��open�ɂ���---
    ' Call set_open_to_kengai
    '------------------------------------------
    If is_relay(styleNo) Then
        Call fill_out_form_relay(prgNo, className, genderName, distance, styleNo, printenable)
    Else
        Call fill_out_form_kojin(prgNo, className, genderName, distance, styleNo, printenable)
    End If
    '-------- �t�G����only
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

    ' �X���C�h�̎擾
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
    Call DisplayTextBoxName("�I�薼")
    Call DisplayTextBoxName("����")
    Call DisplayTextBoxName("�N���X")
    Call DisplayTextBoxName("���")
    Call DisplayTextBoxName("����")
    Call DisplayTextBoxName("�^�C��")
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
                ' TextBox�̖��O��TextRange�ɐݒ�
        shp.TextFrame.TextRange = txtBoxName
    End If
End Sub





