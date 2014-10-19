Attribute VB_Name = "task1"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Const WM_CLOSE = &H10
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Dim lw&, lh&
'Public Const WM_LBUTTONDOWN = &H201
'Public Const WM_LBUTTONUP = &H202
Public Const BM_CLICK = &HF5
Public ret As String
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Option Explicit
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Dim IfPid As Long

'Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'运行指定程序
Public Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long



Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim Pid1 As Long
Dim wText As String * 255
    GetWindowThreadProcessId hwnd, Pid1
    If IfPid = Pid1 Then
        GetWindowText hwnd, wText, 100
        ftask.list.AddItem DQStr(Format(hwnd), 20, 1) & " " & wText
    End If
    EnumWindowsProc = True
End Function
Public Sub Find_Window(ByVal Pid As Long)
    IfPid = Pid
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub



Private Sub Main()
Dim i As Long
Dim AA As String
Dim rttitle As String
       
       WinWnd = FindWindow(TrayNotifyWnd, vbNullString)
       cnt = GetWindowText(WinWnd, rttitle, 255)
       AA = Left$(rttitle, cnt)
       i = PostMessage(WinWnd, WM_CLOSE, &O0, &O0)           '结束任务管理器

 
End Sub
Public Function DQStr(Str As String, Slen As Long, Fs As Long) As String
Dim Sp As Long
Dim i As Long
Dim LenCou As Long
Dim Str1 As String
Dim Str2 As String
Dim SPP As Long
Dim SPP2 As Long
LenCou = 0
Sp = Len(Str)
For i = 1 To Sp
    Str1 = Mid(Str, i, 1)
    If (Str1 < "A" Or Str1 > "z") And (Str1 < "0" Or Str1 > "9") And Str1 <> "." Then
       LenCou = LenCou + 2
    Else
       LenCou = LenCou + 1
    End If
Next i
    If Slen > LenCou Then
       If ((Slen - LenCou) / 2) <> (Fix((Slen - LenCou) / 2)) Then
           SPP = (Fix((Slen - LenCou) / 2))
           SPP2 = (Fix((Slen - LenCou) / 2)) + 1
       Else
          SPP = (Fix((Slen - LenCou) / 2))
          SPP2 = (Fix((Slen - LenCou) / 2))
       End If
       Str1 = Space(Slen - LenCou)
       If Fs = 1 Then Str2 = Str + Str1
       If Fs = 2 Then Str2 = Space(SPP) + Str + Space(SPP2)
       If Fs = 3 Then Str2 = Str1 + Str
       DQStr = Str2
    Else
       DQStr = Str
    End If
End Function
Public Sub Delay(delaytime As Long)

Dim Savetime As Double

Savetime = timeGetTime '记下开始时的时间
While timeGetTime < Savetime + delaytime '循环等待
DoEvents '转让控制权，以便让操作系统处理其它的事件
Wend



End Sub


