VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ȥ�����Ļ����"
   ClientHeight    =   3525
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
   ScaleHeight     =   3525
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command7 
      Caption         =   "�л������Ƶ��"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "�ƽ���̱���"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ֱ�ӹر���Ļ���Ƴ���"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "��ȡ����Ա����"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "һ���ƽ���Ļ���Ƴ���"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "΢���ź�"
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
      Caption         =   "�˳�"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����һ��"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "����ʱ�ر�"
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
         Left            =   240
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
            Name            =   "����"
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
     InitCommonControls '�����Ӿ���ʽ
     Me.Icon = LoadPicture("")
     App.Title = "KillSTM"
End Sub

Private Sub Form_Load()

     '�ж��Ƿ�����ʵ����������
     
     If App.PrevInstance = False Then

         EnableShutDown '�򿪹ػ�Ȩ��
      
         '��ʼ������
 
         Min = 0

         Sec = 0
  
         background = 3
     
         Me.Icon = LoadPicture("") 'Hide
         App.TaskVisible = False
     
         DisableAbility Text1 '���ø���ճ��
         
     ElseIf App.PrevInstance = True Then
     
         MsgBox "����һ��ʵ����������!", 48, "KillSTM"
         AppActivate App.Title '��������ʵ��
         
         End
         
     End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

     SaveSetting "KillSTM", "Settings", "mehwnd", 0 '���ע���HWnd��¼

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
         
             Dim isPreventKill As Integer
             isPreventKill = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill") '��ȡ���̱���״̬
     
             If isPreventKill = 0 Then
     
                 Timer2.Enabled = True '����Ԥ��״̬
                 Label2.Caption = "����ת���̨����!..." & background & "s"
                 Command1.Enabled = False
                 Command2.Caption = "ȡ��"
                 Command3.Enabled = False
                 Command4.Enabled = False
                 Command5.Enabled = False
                 Command6.Enabled = False
                 Command7.Enabled = False
                 Text1.Enabled = False
                 Label3.Visible = True
                 
             ElseIf isPreventKill = 1 Then
             
                 MsgBox "���̱����ѿ���!" & vbCrLf & "�����ƽ���̱���", 48, "��ʾ"
             
             End If
             
         End If
     
     End If
     
End Sub

Private Sub Command2_Click()
     
     If Timer2.Enabled = True Then 'ȡ��Ԥ��״̬
  
         Timer2.Enabled = False
         Label3.Visible = False
         Command1.Enabled = True
         Command2.Caption = "�˳�"
         Command3.Enabled = True
         Command4.Enabled = True
         Command5.Enabled = True
         Command6.Enabled = True
         Command7.Enabled = True
         Text1.Enabled = True
         Label2.Caption = ""
         background = 3 '�������񵹼�ʱ
         
     Else '��Ԥ��״̬���˳�
         
         Unload Me
         End
         
     End If
  
End Sub

Private Sub Command4_Click() '��ȡ����Ա����

     On Error Resume Next
     
     Dim pwd As String
     pwd = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-learning Class Standard\1.00\UninstallPasswd")
     
     If pwd = "" Then
         
         MsgBox "��ȡʧ��!", 48, "��ʾ"
         
     Else
         
         MsgBox "����Ա����Ϊ:" & vbCrLf & pwd, 64, "��ȡ�ɹ�"
         
     End If
     
End Sub

Private Sub Command5_Click()
     
     On Error Resume Next
     
     sPID = GetPsPid("StudentMain.exe") '��ȡѧ���˽���PID
     
     If sPID = 0 Then
         
         MsgBox "��Ļ���Ƴ���δ���������˳�", 64, "�������"
         
     Else '���̴���
     
         Dim isPreventKill As Integer
         isPreventKill = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill") '��ȡ���̱���״̬
         
         If isPreventKill = 0 Then
     
             Shell "taskkill /f /im StudentMain.exe", vbHide 'ֱ�ӽ���
         
             MsgBox "�����ɹ�!", 64, "Killed"
             
         ElseIf isPreventKill = 1 Then
         
             MsgBox "���̱����ѿ���!" & vbCrLf & "�����ƽ���̱���", 48, "��ʾ"
             
         End If
         
     End If
     
End Sub

Private Sub Command6_Click()

     Dim isPreventKill As Integer
     isPreventKill = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill")
         
     If isPreventKill = 1 Then
         
         CreateObject("WScript.Shell").RegWrite "HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill", 0, "REG_DWORD" '�رս��̱���
     
         MsgBox "�ƽ�ɹ�!" & vbCrLf & "���Ľ���ע����������Ч", 64, "���̱���"
         MsgBox "ϵͳ����ע����Ӧ���ƽ�, �뱣��ù�����ȷ����ť��ע��!", 48, "ע��"
         
         ExitWindowsEx EWX_LOGOFF & EWX_FORCE, 0 '�Զ�ע��
         
         Unload Me
         End
         
     ElseIf isPreventKill = 0 Then
     
         MsgBox "���̱����Ѿ��ر�!", 64, "�����ƽ�"
         
     End If

End Sub

Private Sub Command7_Click()
     
     Dim ChannelId As Integer
     ChannelId = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\ChannelId")
         
     If ChannelId = 1 Then

         CreateObject("WScript.Shell").RegWrite "HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\ChannelId", 5, "REG_DWORD" '����ѧ��Ƶ��Ϊ3
     
         MsgBox "���ĳɹ�!" & vbCrLf & "���Ľ���ע�����µ�¼����Ч", 64, "����Ƶ��"
         MsgBox "ϵͳ����ע����Ӧ�ø���, �뱣��ù�����ȷ����ť��ע��!", 48, "ע��"
     
         ExitWindowsEx EWX_LOGOFF & EWX_FORCE, 0 '�Զ�ע��
         
         Unload Me
         End
         
     Else
     
         MsgBox "���л������Ƶ��!", 64, "�����ƽ�"
         
     End If

End Sub

Private Sub Command3_Click() 'һ���ƽ�
     
     Dim ChannelId As Integer
     ChannelId = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\ChannelId")
    
     Dim isPreventKill As Integer
     isPreventKill = CreateObject("WScript.Shell").RegRead("HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill")
     
     If isPreventKill = 0 And ChannelId <> 1 Then
         
         MsgBox "��Ӧ�������ƽ�!", 64, "�����ƽ�"
          
     Else

         CreateObject("WScript.Shell").RegWrite "HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\ChannelId", 5, "REG_DWORD"
         CreateObject("WScript.Shell").RegWrite "HKLM\SOFTWARE\TopDomain\e-Learning Class\Student\PreventKill", 0, "REG_DWORD"
    
         MsgBox "�ƽ�ɹ�!" & vbCrLf & "���Ľ���ע����������Ч", 64, "һ���ƽ�"
         MsgBox "ϵͳ����ע����Ӧ�ø���, �뱣��ù�����ȷ����ť��ע��!", 48, "ע��"
    
         ExitWindowsEx EWX_LOGOFF & EWX_FORCE, 0 '�Զ�ע��
         
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
             Shell "taskkill /f /im StudentMain.exe", vbHide
                 
             MsgBox "�����ɹ�!", 64, "Killed"
             Timer1.Enabled = False
             
             '�����ʱ����
             Min = 0
             Sec = 0
             background = 3
             Text1.Text = 0
             
             Me.Show '��ʾ����
         
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
         Command5.Enabled = True
         Command6.Enabled = True
         Command7.Enabled = True
         Me.Hide
     
     Else
         Label2.Caption = "����ת���̨����!..." & background & "s"
     
     End If

End Sub

'Create By @functionxxx
'Create Date: Feb 17, 2017 15:40
'Update Date: Feb 26, 2017 14:30
'enjoy it
