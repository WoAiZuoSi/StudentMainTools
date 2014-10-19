VERSION 5.00
Begin VB.Form fmain 
   Appearance      =   0  'Flat
   BackColor       =   &H00E8A200&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Student Tools Simple"
   ClientHeight    =   756
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   1296
   FillColor       =   &H00E8A200&
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E8A200&
   Icon            =   "fmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   756
   ScaleWidth      =   1296
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tscr 
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer tkill 
      Left            =   240
      Top             =   240
   End
   Begin VB.Menu Mnu_Menu 
      Caption         =   "0"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu using 
         Caption         =   "启用免打扰模式"
      End
      Begin VB.Menu scshot 
         Caption         =   "截屏"
      End
      Begin VB.Menu Mnu_SubMenu2 
         Caption         =   "-"
      End
      Begin VB.Menu goweb 
         Caption         =   "访问我们的官网"
      End
      Begin VB.Menu Mnu_SubMenu 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu exit 
         Caption         =   "退出"
      End
   End
End
Attribute VB_Name = "fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Type PointAPI
    X As Long
    Y As Long
End Type

Dim ScreenPoint As PointAPI

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub exit_Click()
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Single
lMsg = X / Screen.TwipsPerPixelX
If canuse = 1 Then
    Select Case lMsg
    Case WM_LBUTTONUP
    'ShowWindow fmain.hwnd, SW_RESTORE
    'Me.Hide
    'fmain.Show
    'Me.SetFocus
    '单击左键，显示窗体
    
    '下面两句的目的是把窗口显示在窗口最顶层
    'Me.Show
    Case WM_RBUTTONUP
    PopupMenu Mnu_Menu
    'Case WM_MOUSEMOVE
    'Case WM_LBUTTONDOWN
    'Case WM_LBUTTONDBLCLK
    'Case WM_RBUTTONDOWN
    'Case WM_RBUTTONDBLCLK
    'Case Else
    End Select
Else
    Select Case lMsg
    Case WM_LBUTTONUP
    ShowWindow flogin.hwnd, SW_RESTORE
    'Me.Hide
    flogin.Show
    'Me.SetFocus
    'MsgBox "请用鼠标右键点击图标!", vbInformation, "实时播音专家"
    '单击左键，显示窗体
    
    '下面两句的目的是把窗口显示在窗口最顶层
    'Me.Show
    'Case WM_RBUTTONUP
    'PopupMenu Mnu_Menu
    'Case WM_MOUSEMOVE
    'Case WM_LBUTTONDOWN
    'Case WM_LBUTTONDBLCLK
    'Case WM_RBUTTONDOWN
    'Case WM_RBUTTONDBLCLK
    'Case Else
    End Select
End If
End Sub

Private Sub Form_QueryUnload(cancel As Integer, UnloadMode As Integer)
Call Shell_NotifyIcon(NIM_DELETE, nfIconData)
End Sub



Private Sub Form_Load()
Me.Hide
'laset.Caption = killother
'labout.Caption = processname
With nfIconData
.hwnd = Me.hwnd
.uID = Me.Icon
.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
.uCallbackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon.Handle

.szTip = App.Title & vbNullChar
.cbSize = Len(nfIconData)
End With

Call Shell_NotifyIcon(NIM_ADD, nfIconData)
If FC = 0 Then
    'If Dir("C:\ProgramData\WoAiZuoSi\Student Tools Simple\Can use\Can use.wazs", vbHidden) = "" Or Dir("C:\Users\Public\Documents\WoAiZuoSi\Student Tools Simple\Can use\Can use.wazs", vbHidden) = "" Then
    '    Me.Hide
    '    flogin.Show
    'Else
    '    canuse = 1
    'End If
    canuse = 1
    usescr = 1
    scrend = ".JPG"
    wta = 1000
    filepath = "C:\ScreenShot"
    fsmode = 1
    
    usekiller = 1
    killother = 0
    processname = "StudentMain.exe"
    wtb = 1000
    
    FC = 1
    
    mode = 0
    using.Caption = "启用免打扰模式"
    
    wins = 1
End If
On Error Resume Next
MkDir filepath
SetAttr filepath, vbNormal
Me.AutoRedraw = True

tscr.Interval = wta
tkill.Interval = wtb
End Sub


Private Sub goweb_Click()
ShellExecute Me.hwnd, "open", "http://www.woaizuosi.icoc.cc/", "", "", 5
End Sub

Private Sub scshot_Click()
    If canuse = 1 And mode = 0 And usescr = 1 Then
        'fmain.Hide
        
        BitBlt fmain.hDC, 0, 0, Screen.Width, Screen.Height, _
            GetDC(GetActiveWindow), 0, 0, vbSrcCopy
        
        Dim sFile As String
        sFile = filepath & "\" & Format(Now, "yyyy_mm_dd - hh_mm_ss") & scrend
        
        
        SavePicture fmain.Image, sFile
        'fmain.Picture = LoadPicture("")
        'fmain.Show
        
        'Me.Picture = LoadPicture("")
        'Me.Show
        'MsgBox "1"
        'flash.Show
        'MsgBox "2"
        'Unload Me
        'Form3.Show
    End If
End Sub

Private Sub tkill_Timer()
toola (5)
End Sub

Private Sub tscr_Timer()
toolb (5)
End Sub

Private Sub using_Click()
If mode = 0 Then
    mode = 1
    using.Caption = "停用免打扰模式"
Else
    mode = 0
    using.Caption = "启用免打扰模式"
End If
End Sub
