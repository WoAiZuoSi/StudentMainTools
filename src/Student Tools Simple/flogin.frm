VERSION 5.00
Begin VB.Form flogin 
   Appearance      =   0  'Flat
   BackColor       =   &H00E8A200&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3252
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3252
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox mini 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2170
      Picture         =   "flogin.frx":0000
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   6
      Top             =   0
      Width           =   540
   End
   Begin VB.PictureBox close 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2710
      Picture         =   "flogin.frx":0B6A
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   5
      Top             =   0
      Width           =   540
   End
   Begin VB.PictureBox puname 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   360
      Picture         =   "flogin.frx":16D4
      ScaleHeight     =   480
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   960
      Width           =   240
   End
   Begin VB.PictureBox pupass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   360
      Picture         =   "flogin.frx":2076
      ScaleHeight     =   480
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   240
   End
   Begin VB.TextBox upass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2292
   End
   Begin VB.TextBox uname 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   480
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   2292
   End
   Begin VB.Label go 
      Alignment       =   2  'Center
      BackColor       =   &H00855A00&
      Caption         =   "→ Login"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1560
      TabIndex        =   7
      Top             =   2400
      Width           =   1332
   End
   Begin VB.Label l1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   732
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   1932
   End
End
Attribute VB_Name = "flogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
End
End Sub

Private Sub go_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
go.BackColor = &H644600
End Sub

Private Sub go_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
go.BackColor = &H855A00
If username = "CW" And userpass = "David" Then
    canuse = 1
End If
If username = "LC" And userpass = "404" Then
    canuse = 1
End If
If username = "WoAiZuoSi" And userpass = "ZuoSiBa" Then
    canuse = 1
End If
If username = "DBR" And userpass = "404" Then
    canuse = 1
End If
If username = "32768" And userpass = "" Then
    canuse = 1
End If
If canuse = 0 Then
    MsgBox "用户名或密码错误！"
    upass.Text = ""
    userpass = ""
    flogin.Show
Else
    canuse = 1
    
    On Error Resume Next
    MkDir "C:\ProgramData\WoAiZuoSi"
    SetAttr "C:\ProgramData\WoAiZuoSi", vbHidden
    MkDir "C:\ProgramData\WoAiZuoSi\Student Tools Simple"
    SetAttr "C:\ProgramData\WoAiZuoSi\Student Tools Simple", vbHidden
    MkDir "C:\ProgramData\WoAiZuoSi\Student Tools Simple\Can use"
    SetAttr "C:\ProgramData\WoAiZuoSi\Student Tools Simple\Can use", vbHidden
    
    MkDir "C:\Users\Public\Documents\WoAiZuoSi"
    SetAttr "C:\Users\Public\Documents\WoAiZuoSi", vbHidden
    MkDir "C:\Users\Public\Documents\WoAiZuoSi\Student Tools Simple"
    SetAttr "C:\Users\Public\Documents\WoAiZuoSi\Student Tools Simple", vbHidden
    MkDir "C:\Users\Public\Documents\WoAiZuoSi\Student Tools Simple\Can use"
    SetAttr "C:\Users\Public\Documents\WoAiZuoSi\Student Tools Simple\Can use", vbHidden

    Open "C:\ProgramData\WoAiZuoSi\Student Tools Simple\Can use\Can use.wazs" For Output As #1
    Open "C:\Users\Public\Documents\WoAiZuoSi\Student Tools Simple\Can use\Can use.wazs" For Output As #2
    
    Print #1, "1"
    Print #2, "1"
    
    Unload Me
    'fmain.Show
End If
End Sub

Private Sub mini_Click()
flogin.Hide
End Sub

Private Sub uname_Change()
username = uname.Text
End Sub

Private Sub upass_Change()
userpass = upass.Text
End Sub

