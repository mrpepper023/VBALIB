Attribute VB_Name = "PptFormBaseLogic"
Public CF_str1

Sub �t�H�[���Ăяo��()
    
    Set temp = CreateObject("scripting.dictionary")
    Set result = MyFormBase.UI(temp)
    If result Is Nothing Then
        MsgBox "cancel"
    End If
    
    Unload MyFormBase
    
End Sub

