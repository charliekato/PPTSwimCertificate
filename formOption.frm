VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formOption 
   Caption         =   "�I�v�V����"
   ClientHeight    =   7740
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8376
   OleObjectBlob   =   "formOption.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "formOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If VBA7 Then
    Private Declare PtrSafe Function ImmGetContext Lib "imm32.dll" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As LongPtr, ByVal Conversion As Long, ByVal Sentence As Long) As Long
    Private Declare PtrSafe Function ImmReleaseContext Lib "imm32.dll" (ByVal hWnd As LongPtr, ByVal hIMC As LongPtr) As Long
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
#Else
    Private Declare Function ImmGetContext Lib "imm32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function ImmSetConversionStatus Lib "imm32.dll" (ByVal hIMC As Long, ByVal Conversion As Long, ByVal Sentence As Long) As Long
    Private Declare Function ImmReleaseContext Lib "imm32.dll" (ByVal hWnd As Long, ByVal hIMC As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetForegroundWindow Lib "user32" () As Long
#End If

'''---------------------------------
''' �w�i�̐ݒ�
'''--------------------------------
    Public ImagePath As String

'�w�i�̉摜��path
'Const ImagePath = "C:\Users\user\OneDrive\MyPrograms\VBA\�܏����V�X�e��\���̘A.jpg"
'Const ImagePath = "C:\Users\user\OneDrive\MyPrograms\VBA\�܏����V�X�e��\�܏󎠉ꌧ.jpg"

'Const ImagePath = "C:\Users\user\OneDrive\MyPrograms\VBA\�܏����V�X�e��\���ꌧ�W���j�A.png"

'Const ImagePath = "C:\Users\user\OneDrive\MyPrograms\VBA\�܏����V�X�e��\������21OPEN.png"
    
Sub BackOn()
    Dim sld As slide

    If ImagePath = "" Then Exit Sub

    '
    ' �X���C�h1���擾
    Set sld = ActivePresentation.Slides(1)
    sld.FollowMasterBackground = msoFalse
    ' �w�i��ݒ�
    With sld.Background.Fill
        .Visible = msoTrue
        .UserPicture ImagePath
    End With
    sld.FollowMasterBackground = msoFalse
    
End Sub
Sub BackOff()
    ActivePresentation.Slides(1).FollowMasterBackground = msoTrue
End Sub
Private Sub cbxBackGround_Click()
    If cbxBackGround.Value = True Then
        Call BackOn
    Else
        Call BackOff
    End If
End Sub

Private Sub cmdBackGround_Click()
    Dim fd As FileDialog
    Dim selectedFile As String

    ' �t�@�C���_�C�A���O���t�@�C���I�����[�h�ō쐬
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "�w�i�摜��I�����Ă�������"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "�摜�t�@�C��", "*.jpg; *.jpeg; *.png; *.bmp; *.gif"
        .Filters.Add "���ׂẴt�@�C��", "*.*"

        ' �_�C�A���O��\�����A���[�U�[���t�@�C����I�񂾂��m�F
        If .show = -1 Then
            ImagePath = .SelectedItems(1)
            tbxBackGround.Text = ImagePath
        Else
            MsgBox "�L�����Z������܂���", vbInformation
        End If
    End With
    If cbxBackGround.Value Then
        Call BackOn
    End If
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub tbxBackGround_Change()

End Sub

Private Sub tbxBackGround_Enter()
    Dim hWnd As LongPtr
    Dim hIMC As LongPtr

    hWnd = GetForegroundWindow()
    hIMC = ImmGetContext(hWnd)

    If hIMC <> 0 Then
        ' Conversion: 0 = IME OFF
        ImmSetConversionStatus hIMC, 0, 0
        ImmReleaseContext hWnd, hIMC
    End If
End Sub
Private Sub UserForm_Click()

End Sub
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
