Attribute VB_Name = "TextBoxDisableAbility"
Option Explicit
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Const WM_CUT = &H300 '-------------������Ϣ
Private Const WM_COPY As Long = &H301 '-------------������Ϣ
Private Const WM_PASTE As Long = &H302 '-------------ճ����Ϣ
Private Const WM_CLEAR = &H303 '-------------ɾ����Ϣ[�Ҽ��˵���ɾ��]
Private Const EM_UNDO = &HC7 '-------------������Ϣ
Private Const WM_CONTEXTMENU = &H7B '-------------�Ҽ��˵�
Private prevWndProc   As Long

Private Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case Msg
  Case WM_CUT, WM_COPY, WM_PASTE, WM_CLEAR, EM_UNDO, WM_CONTEXTMENU
    '���ﴦ���Զ�����¼�,���Ϊ��
  Case Else
    '�ص�ϵͳ��������
    WndProc = CallWindowProc(prevWndProc, hwnd, Msg, wParam, lParam)
  End Select
End Function

Public Sub DisableAbility(TargetTextBox As TextBox)
  '��������ʱ�������
  prevWndProc = GetWindowLong(TargetTextBox.hwnd, GWL_WNDPROC)
  SetWindowLong TargetTextBox.hwnd, GWL_WNDPROC, AddressOf WndProc
End Sub

'Create By @functionxxx
'Create Date: Feb 19, 2017 22:51
'Update Date: Feb 19, 2017 22:51
'enjoy it
