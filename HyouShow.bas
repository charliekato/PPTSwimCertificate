Attribute VB_Name = "HyouShow"
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Option Explicit
Option Base 0
Const DefaultServerName = "localhost"
Const DebugMode As Boolean = False   ' false �ɂ��Ă�������!!


    Public MyCon As ADODB.Connection
    Public EventNo As Integer

    Public Gender(4) As String
    Public Shumoku(8) As String
    Public Swimmer() As String
    Public MaxClassNo As Integer

    Public ClassTable() As String

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


Public Function GetPrgNofromPrintPrgNo(printPrgNo As Integer) As Integer
    Dim myRecordset As New ADODB.Recordset
    Dim myQuery As String
    myQuery = "select ���Z�ԍ� from �v���O���� where �\���p���Z�ԍ�=" & printPrgNo & _
              "and ���ԍ�= " & EventNo & ";"
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    GetPrgNofromPrintPrgNo = if_not_null(myRecordset!���Z�ԍ�)
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
        myQuery = "SELECT �N���X.�N���X���� as �N���X, �v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, �v���O����.��ڃR�[�h as ��� FROM �v���O���� " + _
              " INNER JOIN �N���X ON �N���X.�N���X�ԍ�=�v���O����.�N���X�ԍ� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & EventNo & " AND " + _
              " �N���X.���ԍ� = " & EventNo & " AND " & _
              " �v���O����.���Z�ԍ� = " & prgNo & ";"
    Else
        myQuery = "SELECT  �v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, �v���O����.��ڃR�[�h as ��� FROM �v���O���� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & EventNo & " AND " + _
              " �v���O����.���Z�ԍ� = " & prgNo & ";"
    End If
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        If classBasedRace Then
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

Function RelayDistance(distance As String) As String
    If distance = " 200m" Then
        RelayDistance = " 4�~50m"
        Exit Function
    End If
    If distance = " 400m" Then
        RelayDistance = " 4�~100m"
        Exit Function
    End If
    If distance = " 800m" Then
        RelayDistance = " 4�~200m"
        Exit Function
    End If
End Function






Function kenmei(shozoku As String)
    kenmei = shozoku
    If shozoku = "��@��" Then
        kenmei = kenmei + "�@�{"
        Exit Function
    End If
    If kenmei = "���@��" Then
        kenmei = kenmei + "�@�s"
        Exit Function
    End If
    If kenmei = "���@�s" Then
        kenmei = kenmei + "�@�{"
        Exit Function
    End If
    If kenmei = "�k�C��" Then
        Exit Function
    End If
    If kenmei = "������" Then
        kenmei = kenmei + "��"
        Exit Function
    End If
    If kenmei = "�_�ސ�" Then
        kenmei = kenmei + "��"
        Exit Function
    End If
    If kenmei = "�a�̎R" Then
        kenmei = kenmei + "��"
        Exit Function
    End If
    kenmei = kenmei + "�@��"
End Function





Sub fill_time(myTime As String)
    If formOption.cbxTime.Value Then
        Call show("�^�C��", myTime)
    Else
        Call show("�^�C��", "")
    End If
End Sub

Sub fill_class(className As String)
    If formOption.cbxClass.Value Then
        Call show("�N���X", className)
    Else
        Call show("�N���X", "")
    End If
End Sub
Sub fill_shumoku(Shumoku As String)
    If formOption.cbxStyle.Value Then
        Call show("���", Shumoku)
    Else
        Call show("���", "")
    End If
End Sub

Sub fill_junni(junni As Integer)
    If formOption.cbxJunni.Value Then
        If formOption.cbxJunniShowMethod1.Value Then
            Call show("����", "" & junni)
        ElseIf formOption.cbxJunniShowMethod2.Value Then
            Call show("����", "�� " & junni & " ��")
        ElseIf formOption.cbxJunniShowMethod3.Value Then
            If junni = 1 Then
                Call show("����", "�D��")
            Else
                Call show("����", "�� " & junni & " ��")
            End If
        End If
    Else
        Call show("����", "")
    End If
End Sub




Sub fill_name(myName As String)
    If formOption.cbxName.Value Then
        Call show("�I�薼", myName)
    Else
        Call show("�I�薼", "")
    End If
End Sub

Sub fill_shozoku(shozoku As String)
    If formOption.cbxBelongsTo.Value Then
        If formOption.cbxKenmeiMode Then
            Call show("����", kenmei(shozoku))
        Else
            Call show("����", shozoku)
        End If
    Else
        Call show("����", "")
    End If
End Sub


Sub init_class(dummy As String)
    Dim myQuery As String


    

    Dim myRecordset As New ADODB.Recordset
    myQuery = "SELECT MAX(�N���X�ԍ�) as MAX from �N���X where ���ԍ� = " & EventNo
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly

    MaxClassNo = myRecordset!Max
    
    ReDim ClassTable(MaxClassNo)
    myRecordset.Close
    Set myRecordset = Nothing
    
    myQuery = " select �N���X�ԍ�,�N���X���� from �N���X where ���ԍ�=" & EventNo
    myRecordset.Open myQuery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordset.EOF
        ClassTable(CInt(myRecordset!�N���X�ԍ�)) = myRecordset!�N���X����
                
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
"    IF EXISTS (select 1 from �N���X where ���ԍ�=" & EventNo & ") " & _
"    BEGIN" & _
"        SELECT " & _
"           �v���O����.��ڃR�[�h," & _
"                   �N���X.�N���X����,     " & _
"           case �v���O����.���ʃR�[�h  when 1 then '�j�q'  when 2 then '���q'   when 3 then '����'  when 4 then '����'" & _
"                   end as ����, " & _
"           ����.����," & _
"           ���.���,  " & _
"           rank() over (partition by �L�^.���Z�ԍ�, �L�^.�V�L�^����N���X " & _
"                           ORDER BY �L�^.���R�\��, �L�^.�S�[�� ASC) as ����," & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN �I��.����  ELSE �I��1.����  END AS ����1, " & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN ''         ELSE �I��2.����  END AS ����2, " & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN ''         ELSE �I��3.����  END AS ����3, " & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN ''         ELSE �I��4.����  END AS ����4, " & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN "
    myQuery = myQuery & _
"               case �I��.�及�� when 2 then �I��.��������2 " & _
"                                when 3 then �I��.��������3" & _
"                                else �I��.��������1 end" & _
"                   ELSE �����[�`�[��.�`�[����     END AS ����," & _
"           case  WHEN �v���O����.��ڃR�[�h < 6 THEN " & _
"                case �I��.�及�� when 2 then ����2.���������� " & _
"                            when 3 then ����3.���������� " & _
"                       else ����1.���������� end " & _
"                 else �����[����.���������� end as ����������, "
    myQuery = myQuery & _
"           �L�^.�S�[��, " & _
"           �L�^.�V�L�^����}�[�N" & _
"       from �L�^ " & _
"       LEFT JOIN �I�� ON �I��.�I��ԍ� = �L�^.�I��ԍ�  and �I��.���ԍ�=�L�^.���ԍ�" & _
"       left join �����[�`�[�� on �����[�`�[��.�`�[���ԍ�=�L�^.�I��ԍ�" & _
"             and �����[�`�[��.���ԍ�=�L�^.���ԍ�" & _
"       LEFT JOIN �I�� as �I��1 ON �I��1.�I��ԍ� = �L�^.��P�j�� and �I��1.���ԍ�=�L�^.���ԍ�" & _
"       LEFT join �I�� as �I��2 on �I��2.�I��ԍ� = �L�^.��Q�j�� and �I��2.���ԍ�=�L�^.���ԍ�" & _
"       LEFT join �I�� as �I��3 on �I��3.�I��ԍ� = �L�^.��R�j�� and �I��3.���ԍ�=�L�^.���ԍ�" & _
"       LEFT join �I�� as �I��4 on �I��4.�I��ԍ� = �L�^.��S�j�� and �I��4.���ԍ�=�L�^.���ԍ�" & _
"       inner  join �v���O���� on �v���O����.���Z�ԍ�=�L�^.���Z�ԍ� " & _
"            and �v���O����.���ԍ�=�L�^.���ԍ�" & _
"       inner join ���� on ����.�����R�[�h=�v���O����.�����R�[�h" & _
"       inner join ��� on ���.��ڃR�[�h=�v���O����.��ڃR�[�h" & _
"       inner join �N���X on �N���X.���ԍ�=�L�^.���ԍ�" & _
"                        and �N���X.�N���X�ԍ�=�L�^.�V�L�^����N���X" & _
"       left join ���� as ����1 on ����1.�����ԍ�=�I��.�����ԍ�1 and ����1.���ԍ�=�L�^.���ԍ� " & _
"       left join ���� as ����2 on ����2.�����ԍ�=�I��.�����ԍ�2 and ����2.���ԍ�=�L�^.���ԍ� " & _
"       left join ���� as ����3 on ����3.�����ԍ�=�I��.�����ԍ�3 and ����3.���ԍ�=�L�^.���ԍ� " & _
"       left join ���� as �����[���� on �����[����.�����ԍ�=�����[�`�[��.�����ԍ� and �����[����.���ԍ�=�L�^.���ԍ� " & _
"       WHERE  �L�^.���ԍ�= " & EventNo & _
"       �@�@and �L�^.�I��ԍ�>0" & _
"           and �v���O����.�\���p���Z�ԍ�=" & prgNo & _
"           and �L�^.���R���̓X�e�[�^�X=0    and �L�^.���H < 11  end"
    myQuery = myQuery & _
"    else begin" & _
"        SELECT " & _
"           �v���O����.��ڃR�[�h," & _
"           '' as �N���X���� , " & _
"           case �v���O����.���ʃR�[�h " & _
"                 when 1 then '�j�q'" & _
"                 when 2 then '���q'" & _
"                 when 3 then '����'" & _
"                 when 4 then '����'" & _
"           �@end as ����, " & _
"           ����.����," & _
"           ���.���," & _
"           rank() over (partition by �L�^.���Z�ԍ� " & _
"                    ORDER BY �L�^.���R�\��, �L�^.�S�[�� ASC) as ����," & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN �I��.����  ELSE �I��1.����  END AS ����1, " & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN ''         ELSE �I��2.����  END AS ����2, " & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN ''         ELSE �I��3.����  END AS ����3, " & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN ''         ELSE �I��4.����  END AS ����4, " & _
"           case    WHEN �v���O����.��ڃR�[�h < 6 THEN "
    myQuery = myQuery & _
"           case �I��.�及�� when 2 then �I��.��������2 " & _
"                            when 3 then �I��.��������3" & _
"                            else �I��.��������1 end" & _
"     ELSE �����[�`�[��.�`�[����     END AS ����," & _
"           case  WHEN �v���O����.��ڃR�[�h < 6 THEN " & _
"                case �I��.�及�� when 2 then ����2.���������� " & _
"                            when 3 then ����3.���������� " & _
"                       else ����1.���������� end " & _
"                 else �����[����.���������� end as ����������, " & _
"           �L�^.�S�[��, " & _
"           �L�^.�V�L�^����}�[�N" & _
"       from �L�^ "
    myQuery = myQuery & _
"       INNER JOIN �I�� ON �I��.�I��ԍ� = �L�^.�I��ԍ� " & _
"                and �I��.���ԍ�=�L�^.���ԍ�" & _
"       LEFT JOIN �����[�`�[�� on �����[�`�[��.�`�[���ԍ�=�L�^.�I��ԍ�" & _
"                   and �����[�`�[��.���ԍ�=�L�^.���ԍ�" & _
"       LEFT JOIN �I�� as �I��1 ON �I��1.�I��ԍ� = �L�^.��P�j�� and �I��1.���ԍ�=�L�^.���ԍ�" & _
"       LEFT join �I�� as �I��2 on �I��2.�I��ԍ� = �L�^.��Q�j�� and �I��2.���ԍ�=�L�^.���ԍ�" & _
"       LEFT join �I�� as �I��3 on �I��3.�I��ԍ� = �L�^.��R�j�� and �I��3.���ԍ�=�L�^.���ԍ�" & _
"       LEFT join �I�� as �I��4 on �I��4.�I��ԍ� = �L�^.��S�j�� and �I��4.���ԍ�=�L�^.���ԍ�" & _
"       inner join �v���O���� on �v���O����.���Z�ԍ�=�L�^.���Z�ԍ� " & _
"           and �v���O����.���ԍ�=�L�^.���ԍ�" & _
"       inner join ���� on ����.�����R�[�h=�v���O����.�����R�[�h" & _
"       inner join ��� on ���.��ڃR�[�h=�v���O����.��ڃR�[�h" & _
"       left join ���� as ����1 on ����1.�����ԍ�=�I��.�����ԍ�1 and ����1.���ԍ�=�L�^.���ԍ� " & _
"       left join ���� as ����2 on ����2.�����ԍ�=�I��.�����ԍ�2 and ����2.���ԍ�=�L�^.���ԍ� " & _
"       left join ���� as ����3 on ����3.�����ԍ�=�I��.�����ԍ�3 and ����3.���ԍ�=�L�^.���ԍ� " & _
"       left join ���� as �����[���� on �����[����.�����ԍ�=�����[�`�[��.�����ԍ� and �����[����.���ԍ�=�L�^.���ԍ� " & _
"     WHERE  �L�^.���ԍ�= " & EventNo & _
"           and �L�^.�I��ԍ�>0     " & _
"           and �v���O����.�\���p���Z�ԍ�=" & prgNo & _
"           and �L�^.���R���̓X�e�[�^�X=0" & _
"           and �L�^.���H < 11" & _
"    end;"
    Dim junni As Integer
    Dim relayMember As String
    myRecordset.Open myQuery, MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordset.EOF

        If IsNull(myRecordset!�S�[��) Or myRecordset!�S�[�� = "" Then
            MsgBox ("�Y���f�[�^������܂���B���Ԃ񃌁[�X���I����Ă��Ȃ��Ǝv���܂��B")
            fill_out_form2 = False
        Exit Do
        End If
        junni = CInt(myRecordset!����)
        If junni > CInt(formPrgNoPick.tbxJunniLast) Then
            GoTo DOLOOPEND
        End If
        If junni < CInt(formPrgNoPick.tbxJunniTop) Then
            GoTo DOLOOPEND
        End If
        If CInt(myRecordset!��ڃR�[�h) > 5 Then
            relayMember = myRecordset!����1 & "    " & myRecordset!����2 & vbCrLf & _
                          myRecordset!����3 & "    " & myRecordset!����4
            Call fill_name(relayMember)
        Else
            Call fill_name(myRecordset!����1)
        End If
        Call fill_shozoku(myRecordset!����)
        
        Call fill_junni(myRecordset!����)
        If formOption.cbxShumokuWithClass.Value Then
            If CInt(myRecordset!��ڃR�[�h) > 5 Then
                Call fill_shumoku(myRecordset!�N���X���� + myRecordset!���� + RelayDistance(myRecordset!����) + myRecordset!���)
            Else
                Call fill_shumoku(myRecordset!�N���X���� + myRecordset!���� + myRecordset!���� + myRecordset!���)
            End If
        Else
            Call fill_class(myRecordset!�N���X����)
            If CInt(myRecordset!��ڃR�[�h) > 5 Then
                Call fill_shumoku(myRecordset!���� + RelayDistance(myRecordset!����) + myRecordset!���)
            Else
                Call fill_shumoku(myRecordset!���� + myRecordset!���� + myRecordset!���)
            End If
        End If
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





