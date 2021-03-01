VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13635
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '所有者中心
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
    Set http = CreateObject("Microsoft.XMLHTTP")     ' 创建 http 对象以发送请求
    http.Open "GET", url, False                      ' 设置请求地址
    http.setRequestHeader "CONTENT-TYPE", "application/x-www-form-urlencoded"     '设置请求头
    http.send    '发送请求
    If http.Status = 200 Then
        Dim json$                      '定义字符串 json
        json = http.responseText       '获取相应结果
        '接下来是解析 json
        Set objSC = CreateObject("ScriptControl")
        'Set objSC = CreateObjectx86("MSScriptControl.ScriptControl")   '在64位版Excel中的处理方法
        objSC.Language = "JScript"
        strJSON = "var json=" & json & ";"
        objSC.AddCode (strJSON)       '将 json 由字符串解析为对象
        
        Dim j, k, l
        Dim arr()                               '定义一个数组来接收 json 中的数据
                indexArr = Split(CTitle, ",")     '用于在 json 对象中索引数据的数组
        colNum = UBound(indexArr) + 1
        ReDim arr(1 To rowNum, 1 To colNum)     '可以提高向 Excel 单元格填充数据的效率

        'On Error GoTo err_handle                '错误处理
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
    
    sht.Range(Cells(startrow, 1), Cells(startrow + l, colNum)).Value2 = arr   '将数组填入 Excel 表格
End If
    End If
End Function


Private Sub CB_Run_Click()
Pages = Me.T_pages
'Dim sheetname As String
sheetname = Me.T_TableName
Me.T_log = "任务开始 " & Format(Now, "h:mm:ss") & Chr(10)
For i = 1 To Pages
    DoEvents
    url = getpageurl(i)
    Me.T_log = "正在爬第" & i & "页 " & Format(Now, "h:mm:ss") & Chr(10) & Me.T_log
    getData url, ThisWorkbook.Sheets(sheetname), 500
    Me.T_log = "第" & i & "页爬取成功 " & Format(Now, "h:mm:ss") & Chr(10) & Me.T_log
Next
Me.T_log = "任务完成 " & Format(Now, "h:mm:ss") & Chr(10) & Me.T_log


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
