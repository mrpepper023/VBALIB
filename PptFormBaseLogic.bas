Attribute VB_Name = "PptFormBaseLogic"
Public CF_str1

Sub フォーム呼び出し()
    
    Set temp = CreateObject("scripting.dictionary")
    Set result = MyFormBase.UI(temp)
    If result Is Nothing Then
        MsgBox "cancel"
    End If
    
    Unload MyFormBase
    
End Sub

