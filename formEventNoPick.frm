VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEventNoPick 
   Caption         =   "���I��"
   ClientHeight    =   6902
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   6804
   OleObjectBlob   =   "formEventNoPick.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formEventNoPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'  formEventNoPick
'


Private Sub btnClose_Click()
    HyouShow.MyCon.Close
    HyouShow.MyCon = Nothing
    Unload Me
End Sub

Private Sub listEvent_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' �G���^�[�L�[�������ꂽ�Ƃ��A CommandButton1 ���N���b�N
        Call btnOK_Click
    End If
End Sub

Sub add_list_item(row As Integer, item1 As String, item2 As String, item3 As String, item4 As String)
    formPrgNoPick.listPrg.AddItem ("")
    formPrgNoPick.listPrg.List(row, 0) = item1
    formPrgNoPick.listPrg.List(row, 1) = item2
    formPrgNoPick.listPrg.List(row, 2) = item3
    formPrgNoPick.listPrg.List(row, 3) = item4

End Sub





Sub CreateTableIfNotExists()
    Dim cmd As Object 'ADODB.Command
    Dim sql As String


    sql = "IF NOT EXISTS (" & _
          "SELECT 1 " & _
          "FROM INFORMATION_SCHEMA.TABLES " & _
          "WHERE TABLE_NAME = '�����'" & _
          ") " & _
          "BEGIN " & _
          "CREATE TABLE ����� (" & _
          "���ԍ� smallINT NOT NULL, " & _
          "���Z�ԍ� smallINT NOT NULL, " & _
          "����� smallint NOT NULL " & _
           " CONSTRAINT PK_����� PRIMARY KEY (���ԍ�, ���Z�ԍ�)" & _
          "); " & _
          "END;"
    
    On Error GoTo ErrorHandler
    
    
    ' ADODB.Command �I�u�W�F�N�g���쐬����SQL�����s
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = HyouShow.MyCon
    cmd.CommandText = sql
    cmd.Execute
    
    ' ���\�[�X�����

    Set cmd = Nothing

    Exit Sub

ErrorHandler:
    ' �G���[����
    Debug.Print "�G���[���������܂���: " & Err.Description

    Set cmd = Nothing

End Sub


Sub CopyToPrintStatusIfNotExists(target���ԍ� As Integer)
    Dim conn As Object
    Dim cmd As Object
    Dim rs As Object
    Dim connectionString As String
    Dim checkSql As String
    Dim insertSql As String
    
   
    
    ' ���݊m�F�pSQL��
    checkSql = "SELECT 1 FROM ����� WHERE ���ԍ� = " & target���ԍ� & ";"
    
    ' �}���pSQL��
    insertSql = "INSERT INTO ����� (���ԍ�, ���Z�ԍ�, �����) " & _
                "SELECT ���ԍ�, ���Z�ԍ�, 0 " & _
                "FROM �v���O���� " & _
                "WHERE ���ԍ� = " & target���ԍ� & ";"
    
    On Error GoTo ErrorHandler
    
    ' ADODB.Connection �I�u�W�F�N�g���쐬

    
    ' ���݊m�F�pSQL�����s
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = HyouShow.MyCon
    cmd.CommandText = checkSql
    
    Set rs = cmd.Execute
    If rs.EOF Then
        ' ���R�[�h�����݂��Ȃ��ꍇ�̂ݑ}��SQL�����s
        cmd.CommandText = insertSql
        cmd.Execute

    End If
    
    ' ���\�[�X�����
    rs.Close

    Set rs = Nothing
    Set cmd = Nothing

    Exit Sub

ErrorHandler:
    ' �G���[����
    Debug.Print "�G���[���������܂���: " & Err.Description
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
    End If
    Set rs = Nothing
    Set cmd = Nothing

End Sub


Private Sub btnOK_Click()
    Dim Gender(5) As String
    Gender(1) = "�j�q"
    Gender(2) = "���q"
    Gender(3) = "����"
    Gender(4) = "����"
    Dim selectedItem As String
    Dim myRecordset As New ADODB.Recordset
    Dim myquery As String
    Dim row As Integer

    selectedItem = listEvent.Value
    HyouShow.EventNo = CInt(Left(selectedItem, 3))
    Call CreateTableIfNotExists
    CopyToPrintStatusIfNotExists (HyouShow.EventNo)
    If HyouShow.class_exist("") Then
        Call add_list_item(0, "#", "�N���X", "���", "st")
        row = 1
        myquery = "SELECT �v���O����.�\���p���Z�ԍ� as ���Z�ԍ�, �N���X.�N���X���� as �N���X, " & _
              "�v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, ���.��� as ��� FROM �v���O����" + _
              " INNER JOIN ��� ON ���.��ڃR�[�h = �v���O����.��ڃR�[�h " + _
              " INNER JOIN �N���X ON �N���X.�N���X�ԍ�=�v���O����.�N���X�ԍ� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & HyouShow.EventNo & " AND " + _
              " �N���X.���ԍ� = " & HyouShow.EventNo & _
              " order by ���Z�ԍ� asc;"
              
            myRecordset.Open myquery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordset.EOF

                Call add_list_item(row, Right("   " & myRecordset!���Z�ԍ�, 3), myRecordset!�N���X, _
                    Gender(myRecordset!����) + myRecordset!���� + myRecordset!���, "")
                row = row + 1
                myRecordset.MoveNext
            Loop
    Else
        Call add_list_item(0, "#", "", "���", "")
        row = 1
        myquery = "SELECT �v���O����.���Z�ԍ� as ���Z�ԍ�,  " & _
              "�v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, ���.��� as ��� FROM �v���O����" + _
              " INNER JOIN ��� ON ���.��ڃR�[�h = �v���O����.��ڃR�[�h " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & HyouShow.EventNo & ";"
            myRecordset.Open myquery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordset.EOF
                Call add_list_item(row, Right("   " & myRecordset!���Z�ԍ�, 3), "", _
                    Gender(myRecordset!����) + myRecordset!���� + myRecordset!���, "")

                row = row + 1
                myRecordset.MoveNext
            Loop
    End If
    formPrgNoPick.LastRow = row - 1
    

    myRecordset.Close
    Set myRecordset = Nothing
    Call HyouShow.init_senshu("")
    
    Unload Me
    formPrgNoPick.show vbModeless
End Sub


