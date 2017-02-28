VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "输入密码"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "验证新密码:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "新密码:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
     Unload Me
End Sub

Private Sub Command2_Click()

     If Text1.Text = "" Then
     
         Label3.Caption = "请输入密码"

     ElseIf Text1.Text = Text2.Text Then
     
         CreateObject("WScript.Shell").RegWrite "HKLM\SOFTWARE\TopDomain\e-learning Class Standard\1.00\UninstallPasswd", Text1.Text, "REG_SZ"
         MsgBox "修改成功!", 64, "更改密码"
         Unload Me
         
     Else
     
         Label3.Caption = "两次输入密码不一致"
         
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
     Form1.Enabled = True
     Form1.Show
End Sub

Private Sub Text1_Change()
     Label3.Caption = ""
End Sub

Private Sub Text2_Change()
     Label3.Caption = ""
End Sub

'Create By @functionxxx
'Create Date: Feb 26, 2017 16:55
'Update Date: Feb 26, 2017 23:10
'enjoy it

