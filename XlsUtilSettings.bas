Attribute VB_Name = "XlsUtilSettings"


Public Function GetSettings_FromSheet(ByRef sh)

    Set temp = CreateObject("scripting.dictionary")
    
    maxr = GetMaxRowSequence(sh, maxc)
    arr = RectRange(sh, 1, 1, maxr, maxc)
        
    For srcr = 1 To maxr
        '★Settings
    Next

    Set GetSettings_FromSheet = temp

End Function

Private Sub test_GetSettings_FromSheet()

    Set sh = ThisWorkbook.Sheets("Settings")
    
    Set dic = GetSettings_FromSheet(sh)

End Sub

'XlsUtilではなくUtilの方に移す
Public Function GetSettings_FromFile(ByRef sh)

End Function

