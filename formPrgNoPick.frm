VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formPrgNoPick 
   Caption         =   "競技選択"
   ClientHeight    =   7343
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   7301
   OleObjectBlob   =   "formPrgNoPick.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formPrgNoPick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnPreView_Click()
    Dim printPrgNo As Integer
    printPrgNo = CInt(Left(listPrg.value, 3))

    
    Call fill_out_form(HyouShow.get_prgNo(printPrgNo), False)
End Sub

Private Sub btnPrint_Click()
    Dim printPrgNo As Integer
    printPrgNo = CInt(Left(listPrg.value, 3))

    
    Call fill_out_form(HyouShow.get_prgNo(printPrgNo), True)
End Sub

Private Sub listPrg_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnOK_Click
    End If
End Sub

