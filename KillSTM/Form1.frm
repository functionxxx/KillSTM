VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "去你的屏幕控制"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command7 
      Caption         =   "切换到免控频道"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "破解进程保护"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "直接关闭屏幕控制程序"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "获取管理员密码"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "一键破解屏幕控制程序"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "退出"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "干他一炮"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "倒计时关闭"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   750
         Left            =   1800
         Top             =   960
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "按取消键取消"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "min"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Min As Integer, Sec As Integer, background As Integer, sPID As Long

Private Sub Form_Initialize()
     InitCommonControls '启用视觉样式
     Me.Icon = LoadPicture("")
     App.Title = "KillSTM"
End Sub

Private Sub Form_Load()

     '判断是否已有实例正在运行
     
     If App.PrevInstance = False Then

         EnableShutDown '打开关机权限
      
         '初始化变量
 
         Min = 0

         Sec = 0
  
         background = 3
     
         Me.Icon = LoadPicture("") 'Hide
         App.TaskVisible = False
     
         DisableAbility Text1 '禁用复制粘贴
         
     ElseIf App.PrevInstance = True Then
     
         MsgBox "已有一个实例正在运行!", 48, "KillSTM"
         AppActivate App.Title '激活已有实例
         
         End
         
     End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

     SaveSetting "KillSTM", "Settings", "mehwnd", 0 '清空注册表HWnd记录

End Sub


'Fuck
Private Sub Command1_Click()
     
     If Text1.Text = "" Then
         Label2.Caption = "请输入数字!"
     
     ElseIf Text1.Text = 0 Then
         Label2.Caption = "请输入非零值!"
     
     Else
     
         sPID = GetPsPid("StudentMain.exe") '获取学生端进程PID
         
         If sPID = 0 Then
         
             MsgBox "屏幕控制程序未开启或已退出", 64, "无需结束"
             
         Else '进程存在
         
             Dim isPreventKill As Integer
             isPreventKill = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill") '获取进程保护状态
     
             If isPreventKill = 0 Then
     
                 Timer2.Enabled = True '进入预备状态
                 Label2.Caption = "即将转入后台待命!..." & background & "s"
                 Command1.Enabled = False
                 Command2.Caption = "取消"
                 Command3.Enabled = False
                 Command4.Enabled = False
                 Command5.Enabled = False
                 Command6.Enabled = False
                 Command7.Enabled = False
                 Text1.Enabled = False
                 Label3.Visible = True
                 
             ElseIf isPreventKill = 1 Then
             
                 MsgBox "进程保护已开启!" & vbCrLf & "请先破解进程保护", 48, "提示"
             
             End If
             
         End If
     
     End If
     
End Sub

Private Sub Command2_Click()
     
     If Timer2.Enabled = True Then '取消预备状态
  
         Timer2.Enabled = False
         Label3.Visible = False
         Command1.Enabled = True
         Command2.Caption = "退出"
         Command3.Enabled = True
         Command4.Enabled = True
         Command5.Enabled = True
         Command6.Enabled = True
         Command7.Enabled = True
         Text1.Enabled = True
         Label2.Caption = ""
         background = 3 '重置任务倒计时
         
     Else '非预备状态则退出
         
         Unload Me
         End
         
     End If
  
End Sub

Private Sub Command4_Click() '获取管理员密码

     On Error Resume Next
     
     Dim pwd As String
     pwd = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-learning Class Standard\1.00\UninstallPasswd")
     
     If pwd = "" Then
         
         MsgBox "获取失败!", 48, "提示"
         
     Else
         
         MsgBox "管理员密码为:" & vbCrLf & pwd, 64, "获取成功"
         
     End If
     
End Sub

Private Sub Command5_Click()
     
     On Error Resume Next
     
     sPID = GetPsPid("StudentMain.exe") '获取学生端进程PID
     
     If sPID = 0 Then
         
         MsgBox "屏幕控制程序未开启或已退出", 64, "无需结束"
         
     Else '进程存在
     
         Dim isPreventKill As Integer
         isPreventKill = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill") '获取进程保护状态
         
         If isPreventKill = 0 Then
     
             Shell "taskkill /f /im StudentMain.exe", vbHide '直接结束
         
             MsgBox "结束成功!", 64, "Killed"
             
         ElseIf isPreventKill = 1 Then
         
             MsgBox "进程保护已开启!" & vbCrLf & "请先破解进程保护", 48, "提示"
             
         End If
         
     End If
     
End Sub

Private Sub Command6_Click()

     Dim isPreventKill As Integer
     isPreventKill = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill")
         
     If isPreventKill = 1 Then
         
         CreateObject("WScript.Shell").RegWrite "HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill", 0, "REG_DWORD" '关闭进程保护
     
         MsgBox "破解成功!" & vbCrLf & "更改将在注销重启后生效", 64, "进程保护"
         MsgBox "系统即将注销以应用破解, 请保存好工作后按确定按钮以注销!", 48, "注销"
         
         ExitWindowsEx EWX_LOGOFF & EWX_FORCE, 0 '自动注销
         
         Unload Me
         End
         
     ElseIf isPreventKill = 0 Then
     
         MsgBox "进程保护已经关闭!", 64, "无需破解"
         
     End If

End Sub

Private Sub Command7_Click()
     
     Dim ChannelId As Integer
     ChannelId = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\ChannelId")
         
     If ChannelId = 1 Then

         CreateObject("WScript.Shell").RegWrite "HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\ChannelId", 5, "REG_DWORD" '更改学生频道为3
     
         MsgBox "更改成功!" & vbCrLf & "更改将在注销重新登录后生效", 64, "更改频道"
         MsgBox "系统即将注销以应用更改, 请保存好工作后按确定按钮以注销!", 48, "注销"
     
         ExitWindowsEx EWX_LOGOFF & EWX_FORCE, 0 '自动注销
         
         Unload Me
         End
         
     Else
     
         MsgBox "已切换到免控频道!", 64, "无需破解"
         
     End If

End Sub

Private Sub Command3_Click() '一键破解
     
     Dim ChannelId As Integer
     ChannelId = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\ChannelId")
    
     Dim isPreventKill As Integer
     isPreventKill = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill")
     
     If isPreventKill = 0 And ChannelId <> 1 Then
         
         MsgBox "已应用所有破解!", 64, "无需破解"
          
     Else

         CreateObject("WScript.Shell").RegWrite "HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\ChannelId", 5, "REG_DWORD"
         CreateObject("WScript.Shell").RegWrite "HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill", 0, "REG_DWORD"
    
         MsgBox "破解成功!" & vbCrLf & "更改将在注销重启后生效", 64, "一键破解"
         MsgBox "系统即将注销以应用更改, 请保存好工作后按确定按钮以注销!", 48, "注销"
    
         ExitWindowsEx EWX_LOGOFF & EWX_FORCE, 0 '自动注销
         
         Unload Me
         End
         
     End If

End Sub

Private Sub Text1_Change()
     
     If Text1.Text = "" Then
         Command1.Enabled = False
         
     Else
         Command1.Enabled = True
         
     End If

End Sub

'限制输入数字
Private Sub Text1_KeyPress(KeyAscii As Integer)
 
     Select Case KeyAscii
      
         Case 49 To 57 '数字键码
          
             Label2.Caption = "" '输入时清空提示
          
             Exit Sub
          
         Case Else
          
             If KeyAscii = 8 Then '允许退格键
              
                 Label2.Caption = ""
              
             ElseIf KeyAscii = 96 Then
                 
                 If Text1.Text = "" Then
                     Label2.Caption = "请输入非零值!"
                     KeyAscii = 0
                     
                 Else
                 
                 End If
              
             ElseIf KeyAscii = 48 Then
                 
                 If Text1.Text = "" Then
                     Label2.Caption = "请输入非零值!"
                     KeyAscii = 0
                     
                 Else
                 
                 End If
              
             Else
              
                 Label2.Caption = "请输入数字!"
                 KeyAscii = 0
              
             End If
      
     End Select
  
End Sub

Private Sub Timer1_Timer()
     
     On Error Resume Next
 
     Sec = Sec + 1
     
     If Sec = 60 Then
         
         Min = Min + 1
         
         If Min = Text1.Text Then '达到指定时间结束StudentMain.exe进程
         
             sPID = GetPsPid("StudentMain.exe") '获取学生端进程PID
             Shell "taskkill /f /im StudentMain.exe", vbHide
                 
             MsgBox "结束成功!", 64, "Killed"
             Timer1.Enabled = False
             
             '清零计时变量
             Min = 0
             Sec = 0
             background = 3
             Text1.Text = 0
             
             Me.Show '显示窗体
         
         Else
             
             Sec = 0
             
         End If
         
     Else
         
     End If
 
End Sub

Private Sub Timer2_Timer()
  
     background = background - 1
     
     If background = 0 Then
         Timer2.Enabled = False
         Timer1.Enabled = True
         Label2.Caption = ""
         Label3.Visible = False
         Text1.Enabled = True
         Command1.Enabled = True
         Command2.Caption = "退出"
         Command3.Enabled = True
         Command4.Enabled = True
         Command5.Enabled = True
         Command6.Enabled = True
         Command7.Enabled = True
         Me.Hide
     
     Else
         Label2.Caption = "即将转入后台待命!..." & background & "s"
     
     End If

End Sub

'Create By @functionxxx
'Create Date: Feb 17, 2017 15:40
'Update Date: Feb 26, 2017 14:30
'enjoy it
