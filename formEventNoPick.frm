VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formEventNoPick 
   Caption         =   "���I��"
   ClientHeight    =   6902
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   6804
   OleObjectBlob   =   "formEventNoPick.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formEventNoPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub listEvent_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        ' �G���^�[�L�[�������ꂽ�Ƃ��A CommandButton1 ���N���b�N
        Call btnOK_Click
    End If
End Sub



Private Sub btnOK_Click()
    Dim Gender(4) As String
    Gender(1) = "�j�q"
    Gender(2) = "���q"
    Gender(3) = "����"
    Dim selectedItem As String
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String

    selectedItem = listEvent.value
    HyouShow.EventNo = CInt(Left(selectedItem, 3))
    If HyouShow.class_exist(HyouShow.EventNo) Then
        myquery = "SELECT �v���O����.�\���p���Z�ԍ� as ���Z�ԍ�, �N���X.�N���X���� as �N���X, " & _
              "�v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, ���.��� as ��� FROM �v���O����" + _
              " INNER JOIN ��� ON ���.��ڃR�[�h = �v���O����.��ڃR�[�h " + _
              " INNER JOIN �N���X ON �N���X.�N���X�ԍ�=�v���O����.�N���X�ԍ� " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & HyouShow.EventNo & " AND " + _
              " �N���X.���ԍ� = " & HyouShow.EventNo & _
              " order by ���Z�ԍ� asc;"
              
            myRecordSet.Open myquery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordSet.EOF
                
                formPrgNoPick.listPrg.AddItem Right("   " & myRecordSet!���Z�ԍ�, 3) & "  " & _
                          Gender(if_not_null(myRecordSet!����)) & " " & _
                          Right("               " + if_not_null_string(myRecordSet!�N���X), 10) & " " & _
                          if_not_null_string(myRecordSet!����) & " " & _
                          if_not_null_string(myRecordSet!���)
                myRecordSet.MoveNext
            Loop
    Else
        myquery = "SELECT �v���O����.���Z�ԍ� as ���Z�ԍ�,  " & _
              "�v���O����.���ʃR�[�h as ����, " & _
              "����.���� as ����, ���.��� as ��� FROM �v���O����" + _
              " INNER JOIN ��� ON ���.��ڃR�[�h = �v���O����.��ڃR�[�h " + _
              " INNER JOIN ���� ON ����.�����R�[�h = �v���O����.�����R�[�h " + _
              " WHERE �v���O����.���ԍ� = " & HyouShow.EventNo & ";"
            myRecordSet.Open myquery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
            Do Until myRecordSet.EOF
                formPrgNoPick.listPrg.AddItem Right("   " & myRecordSet!���Z�ԍ�, 3) & "  " & _
                          Gender(if_not_null(myRecordSet!����)) & " " & _
                          if_not_null_string(myRecordSet!����) & " " & _
                          if_not_null_string(myRecordSet!���)
                myRecordSet.MoveNext
            Loop
    End If
    

    myRecordSet.Close
    Set myRecordSet = Nothing
    Call HyouShow.init_senshu(HyouShow.EventNo)
    
    Unload Me
    formPrgNoPick.show vbModeless
End Sub


