VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formServerSelect 
   Caption         =   "サーバー選択"
   ClientHeight    =   2359
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   3150
   OleObjectBlob   =   "formServerSelect.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "formServerSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnOK_Click
    End If
End Sub

Private Sub txtBoxServerName_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call btnOK_Click
    End If
End Sub


Private Sub btnOK_Click()
    Dim serverName As String
    serverName = txtBoxServerName.Text
    'Dim MyCon As New ADODB.Connection
    Dim myRecordSet As New ADODB.Recordset
    Dim myquery As String
    Unload Me
'    On Error GoTo MyError
    Set HyouShow.MyCon = New ADODB.Connection
    HyouShow.MyCon.ConnectionString = "Provider=SQLOLEDB;Data Source=" & serverName & "\SQLEXPRESS;Initial Catalog=Sw;User ID=Sw;Password=;"
    HyouShow.MyCon.Open
    Dim eventPick As formEventNoPick
    Set eventPick = New formEventNoPick
    
    myquery = "SELECT 大会番号, 大会名1 FROM 大会設定"
    myRecordSet.Open myquery, HyouShow.MyCon, adOpenStatic, adLockOptimistic, adLockReadOnly
    Do Until myRecordSet.EOF
        formEventNoPick.listEvent.AddItem Right("   " & myRecordSet!大会番号, 3) & "   " & if_not_null_string(myRecordSet!大会名1)
        myRecordSet.MoveNext
    Loop
    myRecordSet.Close
    Set myRecordSet = Nothing
    formEventNoPick.show vbModeless
    Exit Sub
MyError:
    MsgBox ("cannot access server " & serverName)
End Sub


