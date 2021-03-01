Attribute VB_Name = "친욥2"
Sub 브1()
Attribute 브1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 브1 브
'

'
    Range("A1:C3").Select
    Range("A2").Activate
    ActiveSheet.Range("$A$1:$C$3").RemoveDuplicates Columns:=Array(1, 2, 3), _
        Header:=xlYes
End Sub
