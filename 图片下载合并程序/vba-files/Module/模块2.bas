Attribute VB_Name = "ģ��2"
Sub ��1()
Attribute ��1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ��1 ��
'

'
    Range("A1:C3").Select
    Range("A2").Activate
    ActiveSheet.Range("$A$1:$C$3").RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
End Sub
