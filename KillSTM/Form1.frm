VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ȥ�����Ļ����"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command4 
      Caption         =   "��ȡ����Ա����"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ֱ�ӹر���Ļ���Ƴ���"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "�˳�"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����һ��"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ʱ�ر�"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   750
         Left            =   1800
         Top             =   960
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "����"
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
      Begin VB.Label Label3 
         Caption         =   "��ȡ����ȡ��"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "min"
         BeginProperty Font 
            Name            =   "����"
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

Private Sub Form_Initialize()
     InitCommonControls '�����Ӿ���ʽ
     Me.Icon = LoadPicture("")
End Sub

Private Sub Form_Load()
      
     '��ʼ������
 
     Min = 0

     Sec = 0
  
     background = 3
     
     Me.Icon = LoadPicture("") 'Hide
     App.TaskVisible = False
     
     DisableAbility Text1 '���ø���ճ��

End Sub

'Fuck
Private Sub Command1_Click()
     
     If Text1.Text = "" Then
         Label2.Caption = "����������!"
     
     ElseIf Text1.Text = 0 Then
         Label2.Caption = "���������ֵ!"
     
     Else
     
         sPID = GetPsPid("StudentMain.exe") '��ȡѧ���˽���PID
         
         If sPID = 0 Then
         
             MsgBox "��Ļ���Ƴ���δ���������˳�", 64, "�������"
             
         Else '���̴���
     
             Timer2.Enabled = True '����Ԥ��״̬
             Label2.Caption = "����ת���̨����!..." & background & "s"
             Command2.Caption = "ȡ��"
             Command1.Enabled = False
             Command3.Enabled = False
             Command4.Enabled = False
             Text1.Enabled = False
             Label3.Visible = True
             
         End If
     
     End If
     
End Sub

Private Sub Command2_Click()
     
     If Timer2.Enabled = True Then 'ȡ��Ԥ��״̬
  
         Timer2.Enabled = False
         Label3.Visible = False
         Command1.Enabled = True
         Command3.Enabled = True
         Command4.Enabled = True
         Text1.Enabled = True
         Command2.Caption = "�˳�"
         Label2.Caption = ""
         background = 3 '�������񵹼�ʱ
         
     Else '��Ԥ��״̬���˳�
         
         Unload Me
         End
         
     End If
  
End Sub

Private Sub Command3_Click()
     
     On Error Resume Next
     
     sPID = GetPsPid("StudentMain.exe") '��ȡѧ���˽���PID
     
     If sPID = 0 Then
         
         MsgBox "��Ļ���Ƴ���δ���������˳�", 64, "�������"
         
     Else '���̴���
     
         Shell "ntsd /p " & sPID, vbHide 'ֱ�ӽ���
     
         If Err.Number = 0 Then
         
             MsgBox "�����ɹ�!", 64, "Killed"
     
         Else
         
             MsgBox "�Ҳ����ļ�ntsd.exe,�����޷���������!", 16, "����"
             Unload Me
             End
         
         End If
         
     End If

End Sub

Private Sub Command4_Click() '��ȡ����Ա����
     
     On Error Resume Next
     
     Dim pwd As String
     Set ws = CreateObject("WScript.Shell")
     pwd = ws.RegRead("HKLM\SOFTWARE\TopDomain\e-learning Class Standard\1.00\UninstallPasswd")
     
     If pwd = "" Then
         
         MsgBox "��ȡʧ��!", 48, "��ʾ"
         
     Else
         
         MsgBox "����Ա����Ϊ:" & vbCrLf & pwd, 64, "��ȡ�ɹ�"
         
     End If
     
End Sub

Private Sub Text1_Change()
     
     If Text1.Text = "" Then
         Command1.Enabled = False
         
     Else
         Command1.Enabled = True
         
     End If

End Sub

'������������
Private Sub Text1_KeyPress(KeyAscii As Integer)
 
     Select Case KeyAscii
      
         Case 49 To 57 '���ּ���
          
             Label2.Caption = "" '����ʱ�����ʾ
          
             Exit Sub
          
         Case Else
          
             If KeyAscii = 8 Then '�����˸��
              
                 Label2.Caption = ""
              
             ElseIf KeyAscii = 96 Then
                 
                 If Text1.Text = "" Then
                     Label2.Caption = "���������ֵ!"
                     KeyAscii = 0
                     
                 Else
                 
                 End If
              
             ElseIf KeyAscii = 48 Then
                 
                 If Text1.Text = "" Then
                     Label2.Caption = "���������ֵ!"
                     KeyAscii = 0
                     
                 Else
                 
                 End If
              
             Else
              
                 Label2.Caption = "����������!"
                 KeyAscii = 0
              
             End If
      
     End Select
  
End Sub

Private Sub Timer1_Timer()
     
     On Error Resume Next
 
     Sec = Sec + 1
     
     If Sec = 60 Then
         
         Min = Min + 1
         
         If Min = Text1.Text Then '�ﵽָ��ʱ�����StudentMain.exe����
         
             sPID = GetPsPid("StudentMain.exe") '��ȡѧ���˽���PID
             Shell "ntsd /p " & sPID, vbHide
             
             If Err.Number = 0 Then
                 
                 MsgBox "�����ɹ�!", 64, "Killed"
                 Timer1.Enabled = False
             
                 '�����ʱ����
                 Min = 0
                 Sec = 0
                 background = 3
                 Text1.Text = 0
             
                 Me.Show '��ʾ����
            
             Else
                 
                 MsgBox "�Ҳ����ļ�ntsd.exe,�����޷���������!", 16, "����"
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
         Label2.Caption = ""
         Label3.Visible = False
         Text1.Enabled = True
         Command1.Enabled = True
         Command2.Caption = "�˳�"
         Command3.Enabled = True
         Command4.Enabled = True
         Me.Hide
     
     Else
         Label2.Caption = "����ת���̨����!..." & background & "s"
     
     End If

End Sub

'Create By @functionxxx
'Create Date: Feb 17, 2017 15:40
'Update Date: Feb 19, 2017 22:51
'enjoy it
