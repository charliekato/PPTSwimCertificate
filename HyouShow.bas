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
    Dim classExist As Boolean
    classExist = class_exist(EventNo)
    If classExist Then
        myquery = "SELECT �N���X.�N���X���� as �N���X, �v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, �v���O����.��ڃR�[�h as ��� FROM �v���O���� " + _
              " INNER JOIN �N���X ON �N���X.�N���X�ԍ�=�v���O����.�N���X�ԍ� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & EventNo & " AND " + _
              " �N���X.���ԍ� = " & EventNo & " AND " & _
              " �v���O����.���Z�ԍ� = " & PrgNo & ";"
    Else
        myquery = "SELECT  �v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, �v���O����.��ڃR�[�h as ��� FROM �v���O���� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & EventNo & " AND " + _
              " �v���O����.���Z�ԍ� = " & PrgNo & ";"
    End If
    myRecordSet.Open myquery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordSet.EOF
        If classExist Then
            Class = myRecordSet!�N���X
        Else
            Class = ""
        End If
        genderStr = Gender(myRecordSet!����)
        distance = myRecordSet!����
        styleNo = myRecordSet!���
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
            If formPrgNoPick.cbxName.Value Then
                Call show(junni, "�I�薼", myRecordSet!�`�[����)
            Else
                Call show(junni, "�I�薼", "")
            End If
            If formPrgNoPick.cbxBelongsTo.Value Then
                Call show(junni, "����", Swimmer(myRecordSet!��1�j��) & "�E" & Swimmer(myRecordSet!��2�j��) & "�E" & _
                     Swimmer(myRecordSet!��3�j��) & "�E" & Swimmer(myRecordSet!��4�j��))
            Else
                Call show(junni, "����", "")
            End If
            If formPrgNoPick.cbxJunni.Value Then
                Call show(junni, "����", className)
            Else
                Call show(junni, "����", "")
            End If
            If formPrgNoPick.cbxStyle.Value Then
                Call show(junni, "���", genderName + distance + Shumoku(styleNo))
                Else
                Call show(junni, "���", "")
            End If
            If formPrgNoPick.cbxTime.Value Then
                Call show(junni, "�^�C��", ConvertTimeFormat(myRecordSet!�S�[��) + "  " + _
                      if_not_null_string(myRecordSet!�V�L�^����}�[�N))
                Else
                Call show(junni, "�^�C��", "")
            End If
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
            If formPrgNoPick.cbxName.Value Then
                Call show(junni, "�I�薼", myRecordSet!����)
            Else
                Call show(junni, "�I�薼", "")
            End If
            If formPrgNoPick.cbxBelongsTo.Value Then
                Call show(junni, "����", myRecordSet!��������1)
            Else
                Call show(junni, "����", "")
            End If
            If formPrgNoPick.cbxClass.Value Then
                Call show(junni, "�N���X", className)
            Else
                Call show(junni, "�N���X", "")
            End If
            If formPrgNoPick.cbxStyle.Value Then
                Call show(junni, "���", genderName + distance + Shumoku(styleNo))
            Else
                Call show(junni, "���", "")
            End If
            If formPrgNoPick.cbxTime.Value Then
                Call show(junni, "�^�C��", ConvertTimeFormat(myRecordSet!�S�[��) + "  " + _
                    if_not_null_string(myRecordSet!�V�L�^����}�[�N))
            Else
                Call show(junni, "�^�C��", "")
            End If
            If formPrgNoPick.cbxJunni.Value Then
                Call show(junni, "����", junni & "��")
            Else
                Call show(junni, "����", "")
            End If
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




Sub name_text_box(boxNo As Integer, myName As String)
    Dim slide As slide
    Set slide = ActivePresentation.Slides(1)
    slide.Shapes(boxNo).Name = myName
End Sub
Sub show(slideIndex As Integer, txtBoxName As String, dispText As String)

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

Sub check_shape()
    Dim slide As slide
    Dim i As Integer
    Dim slideIndex As Integer
    ' �X���C�h�̎擾 (�e�L�X�g�{�b�N�X�����݂���X���C�h�ԍ��ɐݒ�)
    slideIndex = 1
    Set slide = ActivePresentation.Slides(slideIndex)

    ' ���ׂẴV�F�C�v�����[�v
    For i = 1 To slide.Shapes.Count

           ' slide.Shapes(i).Select
           ' MsgBox (" " & i & slide.Shapes(i).Name & ">" & slide.Shapes(i).TextFrame.TextRange)
           On Error Resume Next
        slide.Shapes(i).TextFrame.TextRange = ">>" & i & "<<"
    Next i
End Sub
