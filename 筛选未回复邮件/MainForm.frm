VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "ɸѡδ�ظ��ʼ�"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6195
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code written By AntoniotheFuture
'antoniothefuture@qq.com


Public isstop As Boolean

'д��־
Sub Log(logstr)
Me.Label9.Caption = Now() & " " & logstr & Chr(10) & Me.Label9.Caption
End Sub

'�����־
Sub clearlog()
Me.Label9.Caption = ""
End Sub

Sub ScanMails()



End Sub



'���½�������λ��
Function SetProgress(PG As Double, n As Integer, M As Integer) As Integer
Me.L_Progress.Width = Me.L_ProgressBG.Width * PG
Me.Label3.Caption = "����(" & n & "/" & M & ")"
End Function

'-1: ����
Function GetMails() As Integer
Dim olMail As MailItem
Dim OLF As Folder
Dim MidFolder As Folder
Dim Emails
Dim DoLoad As Boolean
Dim Title As String
Dim Attms As String
Dim AttmArr
Dim MailC As Collection
Dim hit As Boolean
Dim Sh As Shape
Dim SC As Integer

Dim SMails As Collection
Dim RMails As Collection

'ʱ��
Dim ST As String
Dim ET As String
Dim STD As Date
Dim ETD As Date

Dim Folders
Dim RCFolders

Dim OLFolders As Folders

'��ȡ����
ST = Me.T_ST
ET = Me.T_ET
If ST <> "" Then
    If Not IsDate(ST) Then
        Log ("��ʼʱ���ʽ����")
        GetMails = -1
        Exit Function
    Else
        STD = CDate(ST)
    End If
End If
If ET <> "" Then
    If Not IsDate(ET) Then
        Log ("����ʱ���ʽ����")
        GetMails = -1
        Exit Function
    Else
        ETD = CDate(ET)
    End If
End If


Folders = Me.T_TargetFolders
RCFolders = Me.T_TargetRCFolders
If Folders = "" Or RCFolders = "" Then
    Log ("Ŀ���ļ���Ϊ�գ��޷�ִ��")
    GetMails = -1
    Exit Function
End If
If ST = "" And ET = "" Then Log ("δѡ��ʱ�䷶Χ����ɨ��Ŀ���ļ��е������ʼ�")

Folders = Split(Folders, ";")
RCFolders = Split(RCFolders, ";")


'�ж���ɸѡ�ļ����Ƿ����
On Error Resume Next
Set OLF = Application.GetNamespace("MAPI").Folders(1).Folders("��ɸѡ�ļ���")
If TypeName(OLF) = "Nothing" Then
    Log ("��ɸѡ�ļ��в����ڣ�������")
    Set OLFolders = Application.GetNamespace("MAPI").Folders
    
    Set OLF = OLFolders(1).Folders.Add("��ɸѡ�ļ���")
    If TypeName(OLF) = "Nothing" Then
        Log ("�ļ��д���ʧ��")
        GetMails = -1
        Exit Function
    End If
End If
'���ɸѡ�ļ���
'OLF.Items.Remove
For Each olMail In OLF.Items
    olMail.Delete
Next


Dim SFolders As Collection
Dim SSFolders As Collection
Dim RFolders As Collection
Set SFolders = New Collection
Set SSFolders = New Collection
Set RFolders = New Collection
'����ÿ�������ļ���
For Each F In Folders
    If F = "" Then GoTo nextF
    If Right(F, 2) = "/*" Then
        Set MidFolder = GetFolder(Left(F, Len(F) - 2))
        Set SSFolders = GetFolders(MidFolder)
        For Each F2 In SSFolders
            SFolders.Add F2
        Next
    Else
        SFolders.Add GetFolder(CStr(F))
    End If
nextF:
Next

'����ÿ���ռ��ļ���
For Each F In RCFolders
    If F = "" Then GoTo nextF2
    If Right(F, 2) = "/*" Then
        Set MidFolder = GetFolder(Left(F, Len(F) - 2))
        Set SSFolders = GetFolders(MidFolder)
        For Each F2 In SSFolders
            RFolders.Add F2
        Next
    Else
        RFolders.Add GetFolder(CStr(F))
    End If
nextF2:
Next

Log ("�����ļ���������" & SFolders.Count)
Log ("�ռ��ļ���������" & RFolders.Count)

Set SMails = New Collection
Set RMails = New Collection
For Each F In SFolders
    For Each olMail In F.Items
        SMails.Add olMail
    Next
Next
For Each F In RFolders
    For Each olMail In F.Items
        RMails.Add olMail
    Next
Next
    


'�ʼ�ȥ��
For I = SMails.Count - 1 To 0
    For ii = SMails.Count - 1 To 0
        If I <> ii And SMails(I).Subject = SMails(ii).Subject And SMails(I).ReceivedTime = SMails(ii).ReceivedTime Then
            SMails.Remove (I)
            Exit For
        End If
    Next
Next

'Dim olMailR As MailItem
Dim ReplyC As Integer
Dim mlcopy As MailItem
Dim n As Integer
'����ÿһ���ѷ����ʼ����ҵ���Ӧ�Ļظ��ʼ�������Ҳ����͸��Ƶ���ɸѡ�ļ�����
n = 0
hitn = 0
For Each olMail In SMails
    If isstop Then
        Log ("����ֹ")
        GetMails = hitn
        Exit Function
    End If
    ReplyC = 0
    hit = True
    SetProgress n / SMails.Count, n, SMails.Count
    DoEvents
    If ST <> "" Then If olMail.LastModificationTime < STD Then hit = False
    If ET <> "" And hit Then If olMail.LastModificationTime > ETD Then hit = False
    If hit Then
        Log ("������:" & olMail.Subject)
        For Each olMailR In RMails
            If olMail.LastModificationTime < olMailR.ReceivedTime Then
                If Me.O_Subject.Value Then
                    '����ƥ��ģʽ
                    If InStr(olMailR.Subject, olMail.Subject) > 0 Then
                        ReplyC = ReplyC + 1
                    End If
                Else
                    '�Ի�ƥ��ģʽ
                    If olMail.ConversationID = olMailR.ConversationID Then
                        ReplyC = ReplyC + 1
                    End If
                End If
            End If
        Next
        If ReplyC = 0 Then
            Set mlcopy = olMail.Copy
            mlcopy.Move OLF
            hitn = hitn + 1
        End If
    End If
    n = n + 1
Next
Log ("�������")
SetProgress 1, n, SMails.Count
GetMails = hitn
skip1:
If hitn > 0 Then
    Log ("ɸѡ��" & hitn & "���ʼ�")
Else
    Log ("û�з����������ʼ�")
End If
End Function

'�ݹ�ÿһ�����ļ���
Function GetFolders(F As Folder) As Collection
Dim C As Collection
Dim C2 As Collection
Dim FF As Folder
Set C = New Collection
Set C2 = New Collection
C.Add F
For Each FF In F.Folders
    Set C2 = GetFolders(FF)
    For Each FFF In C2
        C.Add FFF
    Next
Next
Set GetFolders = C
End Function

Private Sub BT_Run_Click()
isstop = False
clearlog
Log ("����ʼ")
GetMails


End Sub

Private Sub BT_SelST_Click()
With ALDTPicker
    .Show
    Me.T_ST = .DateTime
End With
End Sub

Private Sub BT_SetET_Click()
With ALDTPicker
    .Show
    Me.T_ET = .DateTime
End With
End Sub

Private Sub BT_stop_Click()
isstop = True
End Sub

Private Sub Label7_Click()

End Sub

Private Sub s12h_Click()
SelSTFromNow (12)
End Sub

Private Sub s24h_Click()
SelSTFromNow (24)
End Sub

Private Sub s2h_Click()
SelSTFromNow (2)
End Sub

Function SelSTFromNow(h As Integer) As Integer
Me.T_ST = DateAdd("h", h * -1, Now())
End Function

Private Sub s48h_Click()
SelSTFromNow (48)
End Sub

Private Sub s4h_Click()
SelSTFromNow (4)
End Sub

Private Sub s6h_Click()
SelSTFromNow (6)
End Sub

'����������ڴ�·����λ�ļ���
Public Function GetFolder(strFolderPath As String) As MAPIFolder
    ' strFolderPath needs to be something like
    '   "Public Folders\All Public Folders\Company\Sales" or
    '   "Personal Folders\Inbox\My Folder"
    
    Dim objApp 'As Outlook.Application
    Dim objNS As NameSpace
    Dim colFolders As Folders
    Dim objFolder 'As Outlook.MAPIFolder
    Dim arrFolders() As String
    Dim I As Long
    On Error Resume Next
    
    strFolderPath = Replace(strFolderPath, "/", "\")
    arrFolders() = Split(strFolderPath, "\")
    Set objApp = Application
    Set objNS = objApp.GetNamespace("MAPI")
    Set objFolder = objNS.Folders.Item(arrFolders(0))
    If Not objFolder Is Nothing Then
      For I = 1 To UBound(arrFolders)
        Set colFolders = objFolder.Folders
        Set objFolder = Nothing
        Set objFolder = colFolders.Item(arrFolders(I))
        If objFolder Is Nothing Then
          Exit For
        End If
      Next
    End If
    
    Set GetFolder = objFolder
    Set colFolders = Nothing
    Set objNS = Nothing
    Set objApp = Nothing
End Function
