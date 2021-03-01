Attribute VB_Name = "模块1"
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Function GetSetting(SettingName) As String
On Error Resume Next
Dim WS As Worksheet
Set WS = ThisWorkbook.Sheets("设置参数")
For x = 2 To WS.UsedRange.Rows.Count
    If WS.Cells(x, 1).Text = SettingName Then
        GetSetting = WS.Cells(x, 2).Text
    End If
Next
End Function


Function RemoveChars(i As String) As String
Dim CharArr
chars = "/\*<>|?"":"
CharArr = Split(chars, "")
For Each c In CharArr
    i = Replace(i, c, "")
Next
'i = Replace(i, "*", "")
RemoveChars = i
End Function
