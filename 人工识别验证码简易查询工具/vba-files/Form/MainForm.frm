VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "人工识别验证码简易查询工具-By AntoniotheFuture"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub img_cap_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
updatecapimg
End Sub

Private Sub T_Input_Change()
t = Me.T_Input.Value
clen = CInt(getAttr("验证码长度"))
If clen <= 0 Then Exit Sub
If Len(t) = clen Then
    Me.T_Input.BorderColor = normalcolor
    run
End If
End Sub

'打开时加载验证码
Private Sub UserForm_Initialize()
updatecapimg
End Sub

Private Sub WB_img_Enter()
updatecapimg
End Sub

