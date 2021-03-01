Attribute VB_Name = "Other"
Const errorcolor = &HFF&
Const normalcolor = &H80000006
Public targetTab As Worksheet
Public resultTab As Worksheet
Public attrTab As Worksheet

Public imgurl As String
Public queryurl As String
Public querytype As String
Public checktime As Long
Public querytime As Long
Public checkcount As Integer
Public querycount As Integer
Public checksuccess As Integer
Public querysuccess As Integer

'���幤����
Public Sub ready()
Set targetTab = ThisWorkbook.Worksheets("Ҫ��ѯ����Ϣ")
Set resultTab = ThisWorkbook.Worksheets("��ѯ���")
Set attrTab = ThisWorkbook.Worksheets("����")
End Sub

'��ȡ����
Function getAttr(AName As String) As String
With attrTab
    For X = 2 To .UsedRange.Rows.Count
        If (.Cells(X, 1).Text = AName) Then
            getAttr = .Cells(X, 3)
            Exit Function
        End If
    Next
End With

End Function

'���Ŀ��
Sub cleartarget()
For X = targetTab.UsedRange.Rows.Count To 3
    targetTab.Rows(3).Delete
Next
End Sub

'���ȫ��

Sub cleart1()
targetTab.Cells.Delete
End Sub

'��ս��
Sub clearresult()
resultTab.Cells.Delete
End Sub

'ѡ���ƶ�
Sub move(X As Integer, Y As Integer)
'Dim s As Range
targetTab.Activate
Set s = Selection
If s.Worksheet.Name <> targetTab.Name Then
    targetTab.Cells(1, 3).Select
    Exit Sub
End If
Cells(Selection.Row + X, Selection.Column + Y).Select
End Sub

'��ʾ����
Sub showform()
MainForm.Show (0)
End Sub

'�������Ϣ����
Function getCols() As Variant
    Dim by()
    t = ""
    With targetTab
        For Y = 1 To .UsedRange.Columns.Count
            If .Cells(1, Y).Text <> "" Then
                t = t & Y & ","
            End If
        Next
    End With
    If t = "" Then
        getCols = ""
    End If
    t = Left(t, Len(t) - 1)
    a = Split(t, ",")
    ReDim by(0 To UBound(a))
    For n = 0 To UBound(a)
        by(n) = CInt(a(n))
    Next
    getCols = by
End Function

'��ò���
Function getqueryhead() As Variant
    t = ""
    With targetTab
        For Y = 1 To .UsedRange.Columns.Count
            If .Cells(1, Y).Text <> "" Then
                t = t & .Cells(1, Y).Text & ","
            End If
        Next
    End With
    If t = "" Then
        getqueryhead = ""
    End If
    t = Left(t, Len(t) - 1)
    getqueryhead = Split(t, ",")
End Function

'�����ѡ�е���Ϣ
Sub getSelRow()
    Dim info
    Dim r As Integer
    r = Selection.Row
    If MainForm.Visible = False Then
        showform
    End If
    qh = getqueryhead
    cs = getCols
    'a = TypeName(cs)
    If TypeName(cs) <> "Variant()" Then
    'If cs = "" Then
        Exit Sub
    End If
    For Each c In cs
        If targetTab.Cells(2, c).Text <> "" Then
            info = info & targetTab.Cells(2, c) & ":" & targetTab.Cells(r, c) & ";"
        Else
            info = info & targetTab.Cells(1, c) & ":" & targetTab.Cells(r, c) & ";"
        End If
    Next
    MainForm.L_Target.Caption = info
    UpdateLocation
End Sub

'����ͼƬ
Sub updatecapimg()
MainForm.WB_img.Navigate2 getAttr("��֤����ַ") & "?v=" & CStr(Rnd())
'MainForm.img_cap.Picture = getAttr("��֤����ַ") & "?v=" & CStr(Rnd())
End Sub

'����״̬��
Sub updatestatus()
'ƽ��ʶ��ʱ��
'ƽ����ѯʱ��
'���������

MainForm.L_Status.Caption = "ʶ�������" & checkcount & Chr(9) & "ƽʱʶ���ʱ��" & Format(checktime / checkcount, "0.0") & Chr(9) & "�ɹ��ʣ�" & Format(checksuccess / checkcount, "0.0%") & Chr(10) & _
    "��ѯ������" & querycount & Chr(9) & "ƽ����ѯ��ʱ��" & Format(querytime / querycount, "0.0") & Chr(9) & "����ʣ�" & Format(querysuccess / querycount, "0.0%") & Chr(10) & _
    "���������" & resultTab.UsedRange.Rows.Count - 1



End Sub

'��������
Sub clearinput()
MainForm.T_Input.Value = ""

End Sub

'�����ѡ��
Sub getrow()
    
End Sub

'��ʾ����
Sub showerror(errorno As Integer)
t = ""
Select Case errorno
    Case 1
    t = "����������"
    Case 2
    t = "��ѯģʽ����"
    Case 3
    t = "��ѯ��ʱ"
End Select
MainForm.L_Status.Caption = t
End Sub

'������
Sub run()
    Dim resp As String '�ش���Ϣ
    Dim r As Integer
    Dim extime As Integer '��ʱʱ��
    Dim resjson As Object
    Dim datapos As String
    Dim fields As String
    Dim Successmark As String
    Dim reslist
    Dim resrow As String
    Dim Baseinfo As String
    Dim capfn As String
    
    querytype = getAttr("��ѯģʽ")
    queryurl = getAttr("��ѯ��ַ")
    datapos = getAttr("�б�����λ��")
    fields = getAttr("�ֶ��б�")
    Successmark = getAttr("�жϳɹ���־")
    capfn = getAttr("��֤���ֶ�")
    resp = ""
    extime = CInt(getAttr("��ѯ��ʱʱ��"))
    If querytype <> "POST" And querytype <> "GET" Then
        showerror 2
        Exit Sub
    End If
    If queryurl = "" Or extime = 0 Or fields = "" Or datapos = "" Or Successmark = "" Then
        showerror 1
        Exit Sub
    End If
    If Selection.Worksheet.Name <> targetTab.Name Then Exit Sub
    If targetTab.UsedRange.Rows.Count > 2 Then
        targetTab.Cells(3, 1).Select
    Else
        Exit Sub
    End If
    fieldsArr = Split(fields, ";")
    t1 = Timer
    qh = getqueryhead
    cs = getCols
    r = Selection.Row
    If TypeName(cs) <> "Variant()" Then Exit Sub
    '����������
    
    cap = MainForm.T_Input.Value
    querys = ""
    For Each c In cs
        querys = querys & targetTab.Cells(1, c).Text & "=" & UrlEncode(targetTab.Cells(r, c).Text) & "&"
    Next
    querys = querys & capfn & "=" & cap
    'querys = StrConv(querys, vbFromUnicode)
        
    Dim objxml As Object
    Set objxml = CreateObject("MSXML2.XmlHttp")
    If querytype = "POST" Then
        objxml.Open querytype, queryurl, False
        objxml.setrequestheader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
        objxml.Send querys
    Else
        objxml.Open querytype, queryurl & "?" & querys, False
        objxml.Send
    End If
    t2 = Timer
    Debug.Print t2
    Debug.Print t1
    Debug.Print objxml.readyState
    Do While (objxml.readyState <> 4 And (t2 - t1) / 1000 < extime)
        t2 = Timer
        DoEvents
    Loop
    resp = objxml.responsetext
    If resp = "" Then
        showerror 3
        Exit Sub
    End If
    checkcount = checkcount + 1
    querycount = querycount + 1
    t3 = Timer
    '��ѯʱ��ͳ��
    querytime = querytime + (t3 - t1)
    'resjson = jsjson(resp)
    If json2string(resp, Successmark) = "succeed" Then
        '�����֤�ɹ�
        checksuccess = checksuccess + 1
        reslist = json2obj(resp, datapos)
        'д�������Ϣ
        Baseinfo = "row:" & r & ";"
        For Each c In cs
            Baseinfo = Baseinfo & targetTab.Cells(1, c).Text & ":" & targetTab.Cells(r, c).Text & ";"
        Next
        For i = 0 To reslist.Count
            resrow = Baseinfo
            For Each f In fieldsArr
                resrow = resrow & f & ":" & json2string(resp, datapos & "[" & i & "]." & f)
            Next
            writeresult resrow
        Next

        clearinput
        updatecapimg
        '��ѯ�ɹ�������
        If MainForm.CB_AutoNext Then
            move 1, 0
        End If
    Else
        '�����֤ʧ��
        MainForm.T_Input.BorderColor = errorcolor
        
        
        
    End If
    updatestatus
    '���㷵��
    MainForm.T_Input.SetFocus
End Sub

'����jsonΪstring
Function json2string(str, pos) As String
    Dim X As Object
    Set X = CreateObject("ScriptControl"): X.Language = "JScript"
    k = X.eval("eval(" & str & ")" & pos)
    jsjson = X.eval("eval(" & str & ")" & pos)
End Function

'����jsonΪobj
Function json2obj(str, pos) As Object
    Dim X As Object
    Set X = CreateObject("ScriptControl"): X.Language = "JScript"
    k = X.eval("eval(" & str & ")" & pos)
    jsjson = X.eval("eval(" & str & ")" & pos)
End Function


'д����
Function writeresult(str As String)
maxr = resultTab.UsedRange.Rows.Count
startr = maxr + 1
If maxr = 1 And resultTab.Cells(1, 1).Text = "" Then
    startr = 1
End If
resultTab.Cells(startr, 1).Text = str
End Function

'����д��,����д������,����һ��
Function splitobj(obj As Object) As Integer
    



End Function

'����λ����Ϣ
Sub UpdateLocation()
    totalrow = 0
    thisrow = 0
    If targetTab.UsedRange.Rows.Count > 2 Then
        totalrow = targetTab.UsedRange.Rows.Count - 2
    End If
    If Selection.Row > 2 Then
        thisrow = Selection.Row - 2
    End If
    MainForm.L_Location.Caption = "��" & thisrow & "/" & totalrow & "��"
End Sub

'ת��
Public Function UrlEncode(ByRef szString As String) As String
       Dim szChar   As String
       Dim szTemp   As String
       Dim szCode   As String
       Dim szHex    As String
       Dim szBin    As String
       Dim iCount1  As Integer
       Dim iCount2  As Integer
       Dim iStrLen1 As Integer
       Dim iStrLen2 As Integer
       Dim lResult  As Long
       Dim lAscVal  As Long
       szString = Trim$(szString)
       iStrLen1 = Len(szString)
       For iCount1 = 1 To iStrLen1
           szChar = Mid$(szString, iCount1, 1)
           lAscVal = AscW(szChar)
           If lAscVal >= &H0 And lAscVal <= &HFF Then
              If (lAscVal >= &H30 And lAscVal <= &H39) Or _
                 (lAscVal >= &H41 And lAscVal <= &H5A) Or _
                 (lAscVal >= &H61 And lAscVal <= &H7A) Then
                 szCode = szCode & szChar
              Else
                 szCode = szCode & "%" & Hex(AscW(szChar))
              End If
           Else
              szHex = Hex(AscW(szChar))
              iStrLen2 = Len(szHex)
              For iCount2 = 1 To iStrLen2
                  szChar = Mid$(szHex, iCount2, 1)
                  Select Case szChar
                         Case Is = "0"
                              szBin = szBin & "0000"
                         Case Is = "1"
                              szBin = szBin & "0001"
                         Case Is = "2"
                              szBin = szBin & "0010"
                         Case Is = "3"
                              szBin = szBin & "0011"
                         Case Is = "4"
                              szBin = szBin & "0100"
                         Case Is = "5"
                        szBin = szBin & "0101"
                         Case Is = "6"
                              szBin = szBin & "0110"
                         Case Is = "7"
                              szBin = szBin & "0111"
                         Case Is = "8"
                              szBin = szBin & "1000"
                         Case Is = "9"
                              szBin = szBin & "1001"
                         Case Is = "A"
                              szBin = szBin & "1010"
                         Case Is = "B"
                              szBin = szBin & "1011"
                         Case Is = "C"
                              szBin = szBin & "1100"
                         Case Is = "D"
                              szBin = szBin & "1101"
                         Case Is = "E"
                              szBin = szBin & "1110"
                         Case Is = "F"
                              szBin = szBin & "1111"
                         Case Else
                  End Select
              Next iCount2
              szTemp = "1110" & Left$(szBin, 4) & "10" & Mid$(szBin, 5, 6) & "10" & Right$(szBin, 6)
              For iCount2 = 1 To 24
                  If Mid$(szTemp, iCount2, 1) = "1" Then
                     lResult = lResult + 1 * 2 ^ (24 - iCount2)
                  Else: lResult = lResult + 0 * 2 ^ (24 - iCount2)
                  End If
              Next iCount2
              szTemp = Hex(lResult)
                    szCode = szCode & "%" & Left$(szTemp, 2) & "%" & Mid$(szTemp, 3, 2) & "%" & Right$(szTemp, 2)
           End If
szBin = vbNullString
           lResult = 0
       Next iCount1
       UrlEncode = szCode
End Function


