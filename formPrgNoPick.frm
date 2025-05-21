VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPrgNoPick 
   Caption         =   "���Z�I��"
   ClientHeight    =   7344
   ClientLeft      =   96
   ClientTop       =   408
   ClientWidth     =   9540
   OleObjectBlob   =   "formPrgNoPick.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
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
        btnAutoExe.Caption = "��~"
        lblPrintStatus.Caption = "�������s��..."
        Call StartProcess
    Else
        btnAutoExe.Caption = "�������s"
        lblPrintStatus.Caption = ""
        Call StopProcess
    End If
End Sub
Sub StartProcess()
    ' ���s�t���O��ݒ�
    Running = True
    ' ���݂̎�����5�b�����Z
    NextRunTime = Timer + 5
    ' �^�X�N�̎��s���J�n
    ExecuteTask
End Sub

Sub StopProcess()
    ' ���s�t���O������
    Running = False
    ' Debug.Print "�v���Z�X���~���܂����B"
End Sub
Function GetFirstPrintNeededPrtPrgNo() As Integer
    Dim sqlStr As String
    Dim myRecordset As New ADODB.Recordset
    
    On Error GoTo ErrorHandler
    
    sqlStr = "SELECT �����, �����.���Z�ԍ�, �v���O����.�\���p���Z�ԍ� as printPrgNo " & _
             " FROM ����� " & _
             " INNER JOIN �v���O���� ON �v���O����.���Z�ԍ� = �����.���Z�ԍ� " & _
             " WHERE  �v���O����.���ԍ� = " & HyouShow.EventNo & _
             " and �����.���ԍ� = " & HyouShow.EventNo & _
             " order by �v���O����.�\���p���Z�ԍ�"
    myRecordset.Open sqlStr, HyouShow.MyCon, adOpenStatic, adLockReadOnly
    Do Until myRecordset.EOF
        If myRecordset!����� = 0 Then
            GetFirstPrintNeededPrtPrgNo = myRecordset!printPrgNo
            If RaceDone(GetFirstPrintNeededPrtPrgNo) Then
                GoTo CloseExit
            Else
                GetFirstPrintNeededPrtPrgNo = 0
                GoTo CloseExit
            End If
        End If
        myRecordset.MoveNext
    Loop
    ' all races are printed.
    GetFirstPrintNeededPrtPrgNo = 0
CloseExit:
    If Not myRecordset Is Nothing Then
        If myRecordset.State = adStateOpen Then myRecordset.Close
        Set myRecordset = Nothing
    End If
    Exit Function
ErrorHandler:
    Debug.Print "Error in GetFirstPrintNeededPrtPrgNo : " & Err.Description
    GetFirstPrintNeededPrtPrgNo = 0
    Resume CloseExit
End Function
Function RaceDone(printPrgNo As Integer) As Boolean
    Dim sqlStr As String
    Dim myRecordset As New ADODB.Recordset

    On Error GoTo ErrorHandler

    ' SQL���̍쐬
    sqlStr = "SELECT �i�s�t���O from �v���O����" & _
             " WHERE �\���p���Z�ԍ� = " & printPrgNo & _
             " and ���ԍ� = " & HyouShow.EventNo

    ' ���R�[�h�Z�b�g���J��
    myRecordset.Open sqlStr, HyouShow.MyCon, adOpenStatic, adLockReadOnly

    ' ���R�[�h�����݂��Ȃ��ꍇ�̃`�F�b�N
    If Not myRecordset.EOF Then
        If myRecordset!�i�s�t���O = 2 Then
            RaceDone = True
        Else
            RaceDone = False
        End If
    Else
        Debug.Print "No record found while executing " & sqlStr
        RaceDone = False
    End If

Cleanup:
    ' ���R�[�h�Z�b�g���N���[�Y���ĉ��
    If Not myRecordset Is Nothing Then
        If myRecordset.State = adStateOpen Then myRecordset.Close
        Set myRecordset = Nothing
    End If
    Exit Function

ErrorHandler:
    Debug.Print "�G���[���������܂���: " & Err.Description
    RaceDone = False
    Resume Cleanup
End Function
Sub ExecuteTask()
    ' ���������s
    ''Debug.Print "�^�X�N�����s���܂���: " & Now
    Dim printPrgNo As Integer
    printPrgNo = GetFirstPrintNeededPrtPrgNo
    If printPrgNo > 0 Then

        Call PrintGo(printPrgNo)

    End If
    ' �������p������ꍇ
    If Running Then
        ' ���̎��s���Ԃ�����܂őҋ@
        Do While Timer < NextRunTime
            DoEvents ' ���̏�����W���Ȃ��悤�ɂ���
        Loop
        ' ���̎��s���Ԃ�ݒ�
        NextRunTime = Timer + 5
        ' �^�X�N���ċA�I�ɌĂяo��
        ExecuteTask
    End If
End Sub







Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnOption_Click()
    formOption.show (vbModeless)
End Sub

'---- error ---
Private Sub btnPreView_Click()
    Dim printPrgNo As Integer
 '   On Error GoTo subEnd

    printPrgNo = CInt(Left(listPrg.Value, 3))
  
    Call fill_out_form(HyouShow.GetPrgNofromPrintPrgNo(printPrgNo), False)
subEnd:
End Sub


Sub SetPrintedFlag(target���Z�ԍ� As Integer)
    Dim cmd As Object
    Dim sql As String

    On Error GoTo ErrorHandler


    sql = "UPDATE ����� " & _
          " SET ����� = 1 " & _
          " WHERE ���ԍ� = " & HyouShow.EventNo & _
          " AND ���Z�ԍ� = " & target���Z�ԍ� & ";"


    Set cmd = CreateObject("ADODB.Command")
    With cmd
        .ActiveConnection = HyouShow.MyCon ' �����̐ڑ����g�p
        .CommandText = sql
        .Execute
    End With


    ' ���\�[�X�����
    Set cmd = Nothing
    Exit Sub

ErrorHandler:
    ' �G���[����
    Debug.Print "�G���[���������܂���: " & Err.Description
    Set cmd = Nothing
End Sub


Private Sub PrintGo(printPrgNo As Integer)
    Dim prgNo As Integer
    prgNo = HyouShow.GetPrgNofromPrintPrgNo(printPrgNo)
    If fill_out_form(prgNo, True) Then
        Call CheckPrinted(printPrgNo)
    End If
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

    ' �w�肳�ꂽ���Z�ԍ��̃C���f�b�N�X���擾
    targetIndex = GetLineNoFromPrintPrgNo(printPrgNo)

    ' �C���f�b�N�X�� 0 �̏ꍇ�͊Y�����ڂȂ��Ɣ��f
    If targetIndex = 0 Then
        MsgBox "Error: SetDoneFlagOnList �Y�����鍀�ڂ�������܂���B"
        Exit Sub
    End If
    listPrg.List(targetIndex, 3) = "��"


 
End Sub
Sub SelectItemByPrgNo(printPrgNo As Integer)
    Dim targetIndex As Integer

    ' �w�肳�ꂽ���Z�ԍ��̃C���f�b�N�X���擾
    targetIndex = GetLineNoFromPrintPrgNo(printPrgNo)

    ' �C���f�b�N�X�� 0 �̏ꍇ�͊Y�����ڂȂ��Ɣ��f
    If targetIndex = 0 Then
        MsgBox "�Y�����鍀�ڂ�������܂���B"
        Exit Sub
    End If

    ' �S�Ă̑I�����N���A�i�����I�����[�h�̏ꍇ�j
    Dim i As Integer
    For i = 0 To listPrg.ListCount - 1
        listPrg.Selected(i) = False
    Next i

    ' �Y�����鍀�ڂ�I����Ԃɂ���
    listPrg.Selected(targetIndex) = True

    ' �I�����ꂽ���ڂ̒l��\���i�I�v�V�����j
    MsgBox "�I�����ꂽ����: " & listPrg.List(targetIndex)
End Sub


Function GetLineNoFromPrintPrgNo(printPrgNo As Integer) As Integer
    Dim row As Integer
    Dim itemText As String

    ' ListCount - 1 �܂Ń��[�v
    For row = 0 To listPrg.ListCount - 1
        itemText = listPrg.List(row, 0)
        
        ' ���ڂ�3�����ȏ�ł���ꍇ�̂ݏ���
        If Len(itemText) >= 3 Then
            If CInt(Left(itemText, 3)) = printPrgNo Then
                GetLineNoFromPrintPrgNo = row
                Exit Function
            End If
        End If
    Next row

    ' �Y�����ڂ��Ȃ��ꍇ�� 0 ��Ԃ�
    GetLineNoFromPrintPrgNo = 0
End Function



Private Sub listPrg_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnPreView_Click
    End If
End Sub

