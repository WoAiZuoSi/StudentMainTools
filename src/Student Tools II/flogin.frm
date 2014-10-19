VERSION 5.00
Begin VB.Form flogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3612
   Icon            =   "flogin.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "flogin.frx":8582
   ScaleHeight     =   3483.871
   ScaleMode       =   0  'User
   ScaleWidth      =   3576.238
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox mini 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   2535
      Picture         =   "flogin.frx":4B4EC
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   4
      Top             =   0
      Width           =   540
   End
   Begin VB.TextBox upass 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00686868&
      Height          =   629
      Left            =   840
      TabIndex        =   2
      Top             =   1850
      Width           =   2323
   End
   Begin VB.TextBox uname 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   22.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00686868&
      Height          =   629
      Left            =   840
      TabIndex        =   1
      Top             =   871
      Width           =   2323
   End
   Begin VB.PictureBox close 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   3071
      Picture         =   "flogin.frx":4C056
      ScaleHeight     =   252
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   0
      Width           =   540
   End
   Begin VB.Label go 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   571
      Left            =   1949
      TabIndex        =   3
      Top             =   2790
      Width           =   1283
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

Private Sub go_Click()
If username = "CW" And userpass = "David" Then
    canuse = 1
End If
If username = "LC" And userpass = "404" Then
    canuse = 1
End If
If username = "WoAiZuoSi" And userpass = "ZuoSiBa!" Then
    canuse = 1
End If
If username = "DBR" And userpass = "404" Then
    canuse = 1
End If
If username = "odui" And userpass = "theniceboy" Then
    canuse = 1
End If
If username = "32768" And userpass = "" Then
    canuse = 1
End If
If canuse = 0 Then
    MsgBox "用户名或密码错误！"
    upass.Text = ""
    userpass = ""
Else
    canuse = 1
    
    MkDir "C:\ProgramData\WoAiZuoSi"
    SetAttr "C:\ProgramData\WoAiZuoSi", vbHidden
    MkDir "C:\ProgramData\WoAiZuoSi\Student Tools I"
    SetAttr "C:\ProgramData\WoAiZuoSi\Student Tools I", vbHidden
    MkDir "C:\ProgramData\WoAiZuoSi\Student Tools I\Can use"
    SetAttr "C:\ProgramData\WoAiZuoSi\Student Tools I\Can use", vbHidden
    
    MkDir "C:\Users\Public\Documents\WoAiZuoSi"
    SetAttr "C:\Users\Public\Documents\WoAiZuoSi", vbHidden
    MkDir "C:\Users\Public\Documents\WoAiZuoSi\Student Tools I"
    SetAttr "C:\Users\Public\Documents\WoAiZuoSi\Student Tools I", vbHidden
    MkDir "C:\Users\Public\Documents\WoAiZuoSi\Student Tools I\Can use"
    SetAttr "C:\Users\Public\Documents\WoAiZuoSi\Student Tools I\Can use", vbHidden

    Open "C:\ProgramData\WoAiZuoSi\Student Tools I\Can use\Can use.wazs" For Output As #1
    Open "C:\Users\Public\Documents\WoAiZuoSi\Student Tools I\Can use\Can use.wazs" For Output As #2
    
    Print #1, "1"
    Print #2, "1"
    
    Unload Me
    fmain.Show
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
