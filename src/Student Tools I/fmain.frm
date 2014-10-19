VERSION 5.00
Begin VB.Form fmain 
   BackColor       =   &H00E8A200&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Student Tools I"
   ClientHeight    =   2064
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4668
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
   ScaleHeight     =   2064
   ScaleWidth      =   4668
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tscr 
      Left            =   4320
      Top             =   0
   End
   Begin VB.Timer tkill 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lmini 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   2052
   End
   Begin VB.Label lhide 
      Alignment       =   2  'Center
      BackColor       =   &H00855D03&
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2052
   End
   Begin VB.Label labout 
      Alignment       =   2  'Center
      BackColor       =   &H00855D03&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   2052
   End
   Begin VB.Label laset 
      Alignment       =   2  'Center
      BackColor       =   &H00855D03&
      Caption         =   "Set"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2052
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student Tools I"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   21.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4692
   End
   Begin VB.Menu Mnu_Menu 
      Caption         =   "0"
      Visible         =   0   'False
      Begin VB.Menu using 
         Caption         =   "启用免打扰模式"
      End
      Begin VB.Menu Mnu_SubMenu 
         Caption         =   "-"
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
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long

Private Type PointAPI
    X As Long
    Y As Long
End Type

Dim ScreenPoint As PointAPI

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub kset_Click()

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub labout_Click()
Unload Me
fabout.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lMsg As Single
lMsg = X / Screen.TwipsPerPixelX
If canuse = 1 Then
    Select Case lMsg
    Case WM_LBUTTONUP
    ShowWindow fmain.hWnd, SW_RESTORE
    'Me.Hide
    fmain.Show
    'Me.SetFocus
    'MsgBox "请用鼠标右键点击图标!", vbInformation, "实时播音专家"
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
    ShowWindow flogin.hWnd, SW_RESTORE
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
'laset.Caption = killother
'labout.Caption = processname
With nfIconData
.hWnd = Me.hWnd
.uID = Me.Icon
.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
.uCallbackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon.Handle

.szTip = App.Title & vbNullChar
.cbSize = Len(nfIconData)
End With

Call Shell_NotifyIcon(NIM_ADD, nfIconData)

If FC = 0 Then
    'If Dir("C:\ProgramData\WoAiZuoSi\Student Tools I\Can use\Can use.wazs", vbHidden) = "" Or Dir("C:\Users\Public\Documents\WoAiZuoSi\Student Tools I\Can use\Can use.wazs", vbHidden) = "" Then
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
    
    usekiller = 1
    killother = 0
    processname = "StudentMain.exe"
    wtb = 1000
    
    FC = 1
    
    mode = 0
    using.Caption = "启用免打扰模式"
End If

On Error Resume Next
MkDir filepath
SetAttr filepath, vbNormal
Me.AutoRedraw = True

tscr.Interval = wta
tkill.Interval = wtb
End Sub


Private Sub lhide_Click()
Me.Hide
End Sub

Private Sub lmini_Click()
End
End Sub

Private Sub laset_Click()
Unload Me
fset.Show
End Sub

Private Sub tkill_Timer()
GetCursorPos ScreenPoint
If canuse = 1 And mode = 0 And usekiller = 1 And ScreenPoint.X = 0 And ScreenPoint.Y = 0 Then
    If killother = 1 Then
        Shell App.Path + "\ntsd.exe -c q -pn " + processname, vbNormalFocus
    Else
        Shell App.Path + "\ntsd.exe -c q -pn StudentMain.exe", vbNormalFocus
    End If
End If
End Sub

Private Sub tscr_Timer()
GetCursorPos ScreenPoint
xxw = Screen.Width \ Screen.TwipsPerPixelX - 2
If canuse = 1 And mode = 0 And usescr = 1 And ScreenPoint.X > xxw And ScreenPoint.Y = 0 Then
    Me.Hide
    BitBlt Me.hDC, 0, 0, Screen.Width, Screen.Height, _
        GetDC(GetActiveWindow), 0, 0, vbSrcCopy
    Dim sFile As String
    sFile = filepath & "\" & Format(Now, "yyyy_mm_dd - hh_mm_ss") & scrend
    SavePicture Me.Image, sFile
    Me.Picture = LoadPicture("")
    Me.Show
        'MsgBox "1"
        'flash.Show
        'MsgBox "2"
        'Unload Me
        'Form3.Show
End If
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
