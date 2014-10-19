Attribute VB_Name = "menu"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2

Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
'download by http://www.codefans.net

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type


Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const GWL_WNDPROC = (-4)

Public Const WM_USER = &H400
Public Const WM_TRAYICON = WM_USER + 123 '托盘消息

Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205

Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

'===========================================

Public pWndProc As Long

Public Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_TRAYICON Then
        Select Case lParam
            Case WM_RBUTTONDOWN
                SetForegroundWindow hwnd '关键的一步
            Case WM_RBUTTONUP
                Form1.PopupMenu Form1.Mnu_Menu
        End Select
    End If
    
    WndProc = CallWindowProc(pWndProc, hwnd, Msg, wParam, lParam)
End Function
