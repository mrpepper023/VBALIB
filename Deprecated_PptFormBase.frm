VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PptFormBase 
   Caption         =   "UserForm1"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9975
   OleObjectBlob   =   "PptFormBase.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "PptFormBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private flag As Boolean
Private result As Object
Private argument As Object



Public Function UI(ByRef arg) As Object
    '引数にオブジェクトを用いる場合
    Set argument = arg
    
    '結果のプレースホルダ
    Set result = CreateObject("scripting.dictionary")
    flag = False
    
    Me.Show
    
    '結果の返却
    If Not flag Then Set result = Nothing
    Set UI = result
    Set result = Nothing
    Set argument = Nothing

End Function


Private Sub CommandButton1_Click()
'OK
    flag = True
    Me.Hide
End Sub

Private Sub CommandButton2_Click()
'CANCEL
    flag = False
    Me.Hide
End Sub
