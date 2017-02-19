VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "去你的屏幕控制"
   ClientHeight    =   2385
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
   ScaleHeight     =   2385
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "获取管理员密码"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "直接关闭屏幕控制程序"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "干他一炮"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "倒计时关闭"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   750
         Left            =   1080
         Top             =   840
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
         Left            =   360
         MaxLength       =   2
         TabIndex        =   1
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   360
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
         Left            =   1320
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
Private Min As Integer
Private Sec As Integer
Private background As Integer
Private sPID As Long

Private Sub Command2_Click()
  
     Unload Me
     End
  
End Sub

Private Sub Command3_Click()

     On Error Resume Next
     
     Shell "ntsd /p " & sPID, vbHide '直接结束
     
     If Err.Number = 0 Then
         
         MsgBox "结束成功!", 64, "Killed"
     
     Else
         
         MsgBox "找不到文件ntsd.exe,程序无法继续运行!", 16, "错误"
         Unload Me
         End
         
     End If

End Sub

Private Sub Command4_Click() '获取管理员密码
     
     On Error Resume Next
     
     Dim pwd As String
     Set ws = CreateObject("WScript.Shell")
     pwd = ws.RegRead("HKLM\SOFTWARE\TopDomain\e-learning Class Standard\1.00\UninstallPasswd")
     
     If pwd = "" Then
         
         MsgBox "获取失败!", 48, "提示"
         
     Else
         
         MsgBox pwd
         
     End If
     
End Sub

Private Sub Form_Initialize()
     InitCommonControls '启用视觉样式
     Me.Icon = LoadPicture("")
End Sub

'Fuck
Private Sub Command1_Click()
     
     If Text1.Text = "" Then
         Label2.Caption = "请输入数字!"
     
     ElseIf Text1.Text = 0 Then
         Label2.Caption = "请输入非零值!"
     
     Else
         Timer2.Enabled = True
         Label2.Caption = "已转入后台待命!..." & background & "s"
     
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

Private Sub Form_Load()
      
     '初始化变量
 
     Min = 0

     Sec = 0
  
     background = 3
     
     Me.Icon = LoadPicture("") 'Hide
     App.TaskVisible = False
     
     sPID = GetPsPid("StudentMain.exe") '获取学生端进程PID

End Sub

Private Sub Timer1_Timer()
     
     On Error Resume Next
 
     Sec = Sec + 1
     
     If Sec = 60 Then
         
         Min = Min + 1
         
         If Min = Text1.Text Then '达到指定时间结束StudentMain.exe进程
             
             Shell "ntsd /p " & sPID, vbHide
             
             If Err.Number = 0 Then
                 
                 MsgBox "结束成功!", 64, "Killed"
                 Timer1.Enabled = False
             
                 '清零计时变量
                 Min = 0
                 Sec = 0
                 Text1.Text = 0
             
                 Me.Show '显示窗体
            
             Else
                 
                 MsgBox "找不到文件ntsd.exe,程序无法继续运行!", 16, "错误"
                 Unload Me
                 End
                 
             End If
         
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
         Label1.Caption = ""
         Me.Hide
     
     Else
         Label2.Caption = "已转入后台待命!..." & background & "s"
     
     End If

End Sub

'Create By @functionxxx
'Create Date: Feb 17, 2017
'Update Date: Feb 19, 2017
'enjoy it