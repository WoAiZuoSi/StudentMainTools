VERSION 5.00
Begin VB.Form ffile 
   Appearance      =   0  'Flat
   BackColor       =   &H00E8A200&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Tools I - File view"
   ClientHeight    =   6264
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4836
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ffile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6264
   ScaleWidth      =   4836
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox flpath 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3972
   End
   Begin VB.DirListBox filelook 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      ForeColor       =   &H00FFFFFF&
      Height          =   4632
      IMEMode         =   5  'DBCS KATAKANA
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4572
   End
   Begin VB.Label new 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   492
   End
   Begin VB.Label ok 
      Alignment       =   2  'Center
      BackColor       =   &H00855A03&
      Caption         =   "OK"
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
      Left            =   360
      TabIndex        =   2
      Top             =   5520
      Width           =   1812
   End
   Begin VB.Label cancel 
      Alignment       =   2  'Center
      BackColor       =   &H00855A03&
      Caption         =   "Cancel"
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
      Left            =   2640
      TabIndex        =   1
      Top             =   5520
      Width           =   1812
   End
End
Attribute VB_Name = "ffile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
fset.Show
Unload Me
End Sub

Private Sub filelook_Change()
flpath.Text = filelook.Path
End Sub

Private Sub Form_Load()
filelook.Path = filepath
flpath.Text = filelook.Path
End Sub

Private Sub new_Click()
Dim nowfp As String
nowfp = InputBox("请输入新文件夹名称")
If nowfp <> "" Then
    filepath = filelook.Path
    On Error Resume Next
    MkDir filepath & "\" & nowfp
    SetAttr filepath & "\" & nowfp, vbNormal
    filelook.Refresh
Else
End If
End Sub

Private Sub ok_Click()
filepath = filelook.Path
fset.Show
Unload Me
End Sub
