VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13635
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const CTitle = "id,tags,categories,date,modified,title,sentiment,link"

Sub CombineUrl()
On Error Resume Next
With Me
    baseurl = .T_Url
    token = .T_token
    Line = .T_lines
    Page = .T_pages
End With
Offset = Line * (Page - 1)
CBURL = baseurl & "?per_page=" & Line & "&token=" & token & "&offset=" & Offset
Me.L_CombineUrl.Caption = CBURL
End Sub


Function getData(ByVal url As String, sht As Worksheet, ByVal rowNum As Integer)
    Dim http As Object
    Set http = CreateObject("Microsoft.XMLHTTP")     ' ���� http �����Է�������
    http.Open "GET", url, False                      ' ���������ַ
    http.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"     '��������ͷ
    http.send    '��������
    If http.Status = 200 Then
        Dim json$                      '�����ַ��� json
        json = http.responseText       '��ȡ��Ӧ���
        '�������ǽ��� json
        Set objSC = CreateObject("ScriptControl")
        'Set objSC = CreateObjectx86("MSScriptControl.ScriptControl")   '��64λ��Excel�еĴ�����
        objSC.Language = "JScript"
        strJSON = "var json=" & json & ";"
        objSC.AddCode (strJSON)       '�� json ���ַ�������Ϊ����
        
        Dim j, k, l
        Dim arr()                               '����һ������������ json �е�����
                indexArr = Split(CTitle, ",")     '������ json �������������ݵ�����
        colNum = UBound(indexArr) + 1
        ReDim arr(1 To rowNum, 1 To colNum)     '��������� Excel ��Ԫ��������ݵ�Ч��

        'On Error GoTo err_handle                '������
        For j = 1 To rowNum
            For k = 1 To colNum
                Dim kk
                kk2 = "json.posts" + "[" + CStr(j - 1) + "]." + indexArr(k - 1)
                kk = "json.obj" + "[""posts""][" + CStr(j - 1) + "]." + indexArr(k - 1)
                arr(j, k) = objSC.eval(kk2)
            Next
            l = l + 1
        Next
      
err_handle:
    If l = "" Then
    Exit Function
Else
    startrow = sht.UsedRange.Rows.Count
    If startrow = 1 Then
        startrow = 1
    Else
        startrow = startrow + 1
    End If
    
    sht.Range(Cells(startrow, 1), Cells(startrow + l, colNum)).Value2 = arr   '���������� Excel ���
End If
    End If
End Function


Private Sub CB_Run_Click()
Pages = Me.T_pages
'Dim sheetname As String
sheetname = Me.T_TableName
Me.T_log = "����ʼ " & Format(Now, "h:mm:ss") & Chr(10)
For i = 1 To Pages
    DoEvents
    url = getpageurl(i)
    Me.T_log = "��������" & i & "ҳ " & Format(Now, "h:mm:ss") & Chr(10) & Me.T_log
    getData url, ThisWorkbook.Sheets(sheetname), 500
    Me.T_log = "��" & i & "ҳ��ȡ�ɹ� " & Format(Now, "h:mm:ss") & Chr(10) & Me.T_log
Next
Me.T_log = "������� " & Format(Now, "h:mm:ss") & Chr(10) & Me.T_log


End Sub
Function getpageurl(pageindex)
On Error Resume Next
With Me
    baseurl = .T_Url
    token = .T_token
    Line = .T_lines
    Page = .T_pages
End With
Offset = Line * (pageindex - 1)
getpageurl = baseurl & "?per_page=" & Line & "&token=" & token & "&offset=" & Offset
End Function


Private Sub T_lines_Change()
CombineUrl
End Sub

Private Sub T_pages_Change()
CombineUrl
End Sub

Private Sub T_token_Change()
CombineUrl
End Sub

Private Sub T_Url_Change()
CombineUrl
End Sub
