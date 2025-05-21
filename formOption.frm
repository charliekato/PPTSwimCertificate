VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formOption 
   Caption         =   "オプション"
   ClientHeight    =   7740
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8376
   OleObjectBlob   =   "formOption.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
''' 背景の設定
'''--------------------------------
    Public ImagePath As String

'背景の画像のpath
'Const ImagePath = "C:\Users\user\OneDrive\MyPrograms\VBA\賞状印刷システム\中体連.jpg"
'Const ImagePath = "C:\Users\user\OneDrive\MyPrograms\VBA\賞状印刷システム\賞状滋賀県.jpg"

'Const ImagePath = "C:\Users\user\OneDrive\MyPrograms\VBA\賞状印刷システム\滋賀県ジュニア.png"

'Const ImagePath = "C:\Users\user\OneDrive\MyPrograms\VBA\賞状印刷システム\いずみ21OPEN.png"
    
Sub BackOn()
    Dim sld As slide

    If ImagePath = "" Then Exit Sub

    '
    ' スライド1を取得
    Set sld = ActivePresentation.Slides(1)
    sld.FollowMasterBackground = msoFalse
    ' 背景を設定
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

    ' ファイルダイアログをファイル選択モードで作成
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "背景画像を選択してください"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "画像ファイル", "*.jpg; *.jpeg; *.png; *.bmp; *.gif"
        .Filters.Add "すべてのファイル", "*.*"

        ' ダイアログを表示し、ユーザーがファイルを選んだか確認
        If .show = -1 Then
            ImagePath = .SelectedItems(1)
            tbxBackGround.Text = ImagePath
        Else
            MsgBox "キャンセルされました", vbInformation
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
