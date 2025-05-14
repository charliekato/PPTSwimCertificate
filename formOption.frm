VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formOption 
   Caption         =   "オプション"
   ClientHeight    =   8364
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "formOption.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
    Me.Hide
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
