VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ALDTPicker 
   Caption         =   "时间日期选择器"
   ClientHeight    =   6465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4710
   OleObjectBlob   =   "ALDTPicker.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "ALDTPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Code written By AntoniotheFuture
'antoniothefuture@qq.com


Public OldValue As String
Public DateFormat As String
Public DateTime As String
Public DateType As String
Dim Y As Integer
Dim M As Integer
Dim D As Integer
Dim HH As Integer
Dim MM As Integer
Dim SS As Integer

Private Sub BT_Cancel_Click()
DateTime = OldValue
Me.hide
End Sub


Private Sub BT_HAdd_Click()
If HH = 23 Then
    HH = 0
Else
    HH = HH + 1
End If
RefreshTextbox
End Sub

Private Sub BT_HCut_Click()
If HH = 0 Then
    HH = 23
Else
    HH = HH - 1
End If
RefreshTextbox
End Sub

Private Sub BT_HZero_Click()
HH = 0
RefreshTextbox
End Sub

Private Sub BT_MAdd_Click()
If MM = 59 Then
    MM = 0
Else
    MM = MM + 1
End If
RefreshTextbox
End Sub

Private Sub BT_MCut_Click()
If MM = 0 Then
    MM = 59
Else
    MM = MM - 1
End If
RefreshTextbox
End Sub

Private Sub BT_MZero_Click()
MM = 0
RefreshTextbox
End Sub

Private Sub BT_SAdd_Click()
If SS = 59 Then
    SS = 0
Else
    SS = SS + 1
End If
RefreshTextbox
End Sub

Private Sub BT_SCut_Click()
If SS = 0 Then
    SS = 59
Else
    SS = SS - 1
End If
RefreshTextbox
End Sub

Private Sub BT_SZero_Click()
SS = 0
RefreshTextbox
End Sub


Private Sub BT_OK_Click()
Me.hide
End Sub


Private Sub CommandButton1_Click()

End Sub

Private Sub BT_YearDown_Click()
Y = Y - 1
RefreshTextbox
End Sub

Private Sub BT_YearDownTen_Click()
Y = Y - 5
RefreshTextbox
End Sub

Private Sub BT_YearUp_Click()
Y = Y + 1
RefreshTextbox
End Sub

Private Sub BT_YearUpTen_Click()
Y = Y + 5
RefreshTextbox
End Sub


Private Sub CommandButton5_Click()

End Sub

Private Sub L_LMDs_Click()
D = L_LMDs.Value
RefreshTextbox
End Sub

Private Sub L_NM_Click()
Dim Mstr As String
Mstr = Me.L_NM.Value
M = Left(Mstr, Len(Mstr) - 1)
If Y = 0 Then Y = Year(Date)
Me.L_LMDs.Clear
EndM = DateAdd("M", 1, Format(Y & "-" & M & "-" & "01", "ddddd"))
Ds = Day(DateAdd("D", -1, EndM))
For I = 16 To Ds
    Me.L_LMDs.AddItem I
Next
RefreshTextbox

End Sub

Private Sub L_UMDs_Click()
D = L_UMDs.Value
RefreshTextbox
End Sub

Private Sub T_H_Exit(ByVal Cancel As MSForms.ReturnBoolean)
TBH = Me.T_H.Text
If IsNumeric(TBH) And TBH >= 0 And TBH < 24 Then
    HH = TBH
Else
    Me.T_H.Text = ""
End If
RefreshTextbox
End Sub


Private Sub T_M_Exit(ByVal Cancel As MSForms.ReturnBoolean)
TBH = Me.T_M.Text
If IsNumeric(TBH) And TBH >= 0 And TBH < 60 Then
    MM = TBH
Else
    Me.T_M.Text = ""
End If
RefreshTextbox
End Sub

Private Sub T_S_Exit(ByVal Cancel As MSForms.ReturnBoolean)
TBH = Me.T_S.Text
If IsNumeric(TBH) And TBH >= 0 And TBH < 60 Then
    SS = TBH
Else
    Me.T_S.Text = ""
End If
RefreshTextbox
End Sub



Private Sub T_Year_Exit(ByVal Cancel As MSForms.ReturnBoolean)
With Me
    If IsNumeric(.T_Year) Then
        If Not (.T_Year.Text > 0 And .T_Year = Int(.T_Year.Text)) Then
            .T_Year = Year(Date)
        End If
    Else
        .T_Year = Year(Date)
    End If
Y = .T_Year
RefreshTextbox
End With
    
End Sub


Sub RefreshTextbox()
Me.T_Year = Format(Y, "0000")
Me.T_H = Format(HH, "00")
Me.T_M = Format(MM, "00")
Me.T_S = Format(SS, "00")
Me.T_DateTime = Format(Format(Y, "0000") & "/" & Format(M, "00") & _
    "/" & Format(D, "00") & " " & Format(HH, "00") & ":" & Format(MM, "00") & ":" & Format(SS, "00"), DateFormat)
DateTime = Me.T_DateTime
End Sub

Sub SetDT(FDateType As String, Optional OldValue As String)
DateType = FDateType
If OldValue = "" Then
    Me.T_DateTime = Format(Date & " " & Time(), DateFormat)
    DateTime = Me.T_DateTime
Else
    DateTime = OldValue
End Sub

Private Sub UserForm_Activate()
If DateFormat = "" Then
    DateFormat = "YYYY/MM/DD hh:mm:ss"
End If
If DateType = "" Then
    DateType = "D&T"
End If
If DateType = "D" Then
    Me.F_Time.Visible = False
ElseIf DateType = "T" Then
    Me.F_Time.Visible = False
End If
For I = 1 To 12
    Me.L_NM.AddItem I & "月"
Next
For I = 1 To 15
    Me.L_UMDs.AddItem I
Next
If Not IsDate(DateTime) Or OldValue = "" Then
    DateTime = Date & " " & Time
End If
Y = Year(DateTime)
M = Month(DateTime)
D = Day(DateTime)
HH = Hour(DateTime)
MM = Minute(DateTime)
SS = Second(DateTime)
RefreshTextbox
End Sub

