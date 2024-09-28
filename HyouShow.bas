Attribute VB_Name = "HyouShow"
Option Explicit
Option Base 0

    Const JLIMIT = 1  '  ���ʂ܂ŏ܏���o����
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




Public Function class_exist(ByVal EventNo As Integer) As Boolean
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    Dim rc As Boolean
    myquery = "select * from �N���X where ���ԍ� = " & EventNo
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
    myquery = "SELECT MAX(�I��ԍ�) as MAX from �I�� where ���ԍ� = " & EventNo
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    maxSwimmerNo = myRecordSet!Max
    
    ReDim Swimmer(maxSwimmerNo)
    myRecordSet.Close
    myquery = "SELECT ����, �I��ԍ� from �I�� where ���ԍ� = " & EventNo
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordSet.EOF
        Swimmer(myRecordSet!�I��ԍ�) = myRecordSet!����
        myRecordSet.MoveNext
    Loop
    myRecordSet.Close
    Set myRecordSet = Nothing
End Sub

Public Function get_prgNo(printPrgNo As Integer) As Integer
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    myquery = "select ���Z�ԍ� from �v���O���� where �\���p���Z�ԍ�=" & printPrgNo & _
              "and ���ԍ�= " & EventNo & ";"
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    get_prgNo = if_not_null(myRecordSet!���Z�ԍ�)
    myRecordSet.Close
    Set myRecordSet = Nothing
End Function
Sub get_race_title(ByVal EventNo As Integer, ByVal PrgNo As Integer, ByRef Class As String, _
            ByRef genderStr As String, ByRef distance As String, ByRef styleNo As Integer)
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    myquery = "SELECT �N���X.�N���X���� as �N���X, �v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, �v���O����.��ڃR�[�h as ��� FROM �v���O���� " + _
              " INNER JOIN �N���X ON �N���X.�N���X�ԍ�=�v���O����.�N���X�ԍ� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & EventNo & " AND " + _
              " �N���X.���ԍ� = " & EventNo & " AND " & _
              " �v���O����.���Z�ԍ� = " & PrgNo & ";"
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordSet.EOF
        Class = myRecordSet!�N���X
        genderStr = Gender(myRecordSet!����)
        distance = myRecordSet!����
        styleNo = myRecordSet!���
        myRecordSet.MoveNext
    Loop
    myRecordSet.Close
    Set myRecordSet = Nothing
                
              
    
End Sub






Function get_swimmer_by_rank(ByRef resultList As Collection, ByVal rank As Integer) As Collection
    Dim thisResult As Result
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
        myquery = "SELECT �����[�`�[��.�`�[���� as �`�[����, �L�^.�S�[�� as �S�[��, " & _
            "�L�^.��P�j��, �L�^.��Q�j��, �L�^.��R�j��, �L�^.��S�j��, �L�^.�V�L�^����}�[�N " & _
            "FROM �L�^ " & _
            " inner join �����[�`�[�� on �����[�`�[��.�`�[���ԍ� = �L�^.�I��ԍ� " & _
            " where   �L�^.���Z�ԍ� = " & PrgNo & _
            " and �L�^.���ԍ� = " & EventNo & " and �L�^.���R���̓X�e�[�^�X=0 " & _
            " and �����[�`�[��.���ԍ� = " & EventNo & _
            " order by �S�[�� asc;"

        myRecordSet.Open myquery, MyCon, adOpenStatic, adLockReadOnly
        Do Until myRecordSet.EOF
            If prevTime <> myRecordSet!�S�[�� Then
                junni = junni + 1
                If junni > JLIMIT Then
                    Exit Do
                End If
                prevTime = myRecordSet!�S�[��
            End If
            Call show(1, 1, myRecordSet!�`�[����)
            Call show(1, 2, Swimmer(myRecordSet!��1�j��) & "�E" & Swimmer(myRecordSet!��2�j��) & "�E" & _
                     Swimmer(myRecordSet!��3�j��) & "�E" & Swimmer(myRecordSet!��4�j��))
            Call show(1, 3, className)
            Call show(1, 4, genderName + distance + Shumoku(styleNo))
            Call show(1, 5, myRecordSet!�S�[�� + "  " + if_not_null_string(myRecordSet!�V�L�^����}�[�N))
            If printEnable Then
                Call print_it("")
            End If
            myRecordSet.MoveNext
        Loop
    Else
        myquery = "SELECT �I��.���� as ����, �L�^.�S�[�� as �S�[��, �I��.��������1, �L�^.�V�L�^����}�[�N " & _
        "FROM �L�^ " & _
        " inner join �I�� on �I��.�I��ԍ� = �L�^.�I��ԍ� " & _
        " where �I��.���ԍ�=" & EventNo & " and �L�^.���Z�ԍ� = " & PrgNo & _
        " and �L�^.���ԍ� = " & EventNo & " and �L�^.���R���̓X�e�[�^�X=0 " & _
        " and �L�^.�I��ԍ�>0 order by �S�[�� asc;"

        myRecordSet.Open myquery, MyCon, adOpenStatic, adLockReadOnly
        Do Until myRecordSet.EOF
            If prevTime <> myRecordSet!�S�[�� Then
                junni = junni + 1
                If junni > JLIMIT Then
                    Exit Do
                End If
            End If
            Call show(1, 1, myRecordSet!����)
            Call show(1, 2, myRecordSet!��������1)
            Call show(1, 3, className)
            Call show(1, 4, genderName + distance + Shumoku(styleNo))
            Call show(1, 5, myRecordSet!�S�[�� + "  " + if_not_null_string(myRecordSet!�V�L�^����}�[�N))
            If printEnable Then
                Call print_it("")

            End If
            prevTime = myRecordSet!�S�[��
            myRecordSet.MoveNext
        Loop
    End If
    ' �N���[�Y�Ɖ��
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

Sub check_shape()
    Dim slide As slide
    Dim i As Integer
    Dim slideIndex As Integer
    ' �X���C�h�̎擾 (�e�L�X�g�{�b�N�X�����݂���X���C�h�ԍ��ɐݒ�)
    slideIndex = 1
    Set slide = ActivePresentation.Slides(slideIndex)

    ' ���ׂẴV�F�C�v�����[�v
    For i = 1 To slide.Shapes.Count

            slide.Shapes(i).Select
            MsgBox (" " & i & slide.Shapes(i).Name & ">" & slide.Shapes(i).TextFrame.TextRange)

    Next i
End Sub


Sub name_text_box(boxNo As Integer, myName As String)
    Dim slide As slide
    Set slide = ActivePresentation.Slides(1)
    slide.Shapes(boxNo).Name = myName
End Sub
Sub show(slideIndex As Integer, txtBoxIndex As Integer, dispText As String)

    ' �X���C�h�̎擾
    Dim slide As slide
    Set slide = ActivePresentation.Slides(slideIndex)

    slide.Shapes(txtBoxIndex).TextFrame.TextRange = dispText
End Sub

