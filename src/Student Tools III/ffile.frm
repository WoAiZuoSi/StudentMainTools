VERSION 5.00
Begin VB.Form ffile 
   Appearance      =   0  'Flat
   BackColor       =   &H00E8A200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Tools III - File view"
   ClientHeight    =   6132
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4812
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
   ScaleHeight     =   6132
   ScaleWidth      =   4812
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tkill 
      Left            =   0
      Top             =   0
   End
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
      DragIcon        =   "ffile.frx":9CCA
      ForeColor       =   &H00FFFFFF&
      Height          =   4860
      IMEMode         =   5  'DBCS KATAKANA
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4572
   End
   Begin VB.Label bnew 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
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
      BackColor       =   &H0000FF00&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   0
      TabIndex        =   2
      Top             =   5640
      Width           =   2412
   End
   Begin VB.Label cancel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   16.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   2400
      TabIndex        =   1
      Top             =   5640
      Width           =   2412
   End
End
Attribute VB_Name = "ffile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub givecolor(a As Long)

If a = 1 Then
    ffile.BackColor = &HE8A200
    filelook.BackColor = &H855A00
End If

If a = 2 Then
    ffile.BackColor = &H855A00
    filelook.BackColor = &HE8A200
End If

If a = 3 Then
    ffile.BackColor = &H323232
    filelook.BackColor = &H80000010
End If

If a = 4 Then
    ffile.BackColor = &H80000010
    filelook.BackColor = &H323232
End If

End Sub

Private Sub bnew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If wins = 1 Then bnew.BackColor = &H855A00
If wins = 2 Then bnew.BackColor = &HE8A200
If wins = 3 Then bnew.BackColor = &H80000010
If wins = 4 Then bnew.BackColor = &H323232
Dim nowfp As String
nowfp = InputBox("请输入新文件夹名称")
If nowfp <> "" Then
    On Error Resume Next
    filepath = filelook.Path
    MkDir filepath & "\" & nowfp
    SetAttr filepath & "\" & nowfp, vbNormal
    filelook.Refresh
Else
    MsgBox "新文件夹的名称不得为空"
End If
End Sub


Private Sub cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BackColor = &HC0&
End Sub

Private Sub cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
fset.Show
Unload Me
End Sub

Private Sub filelook_Change()
flpath.Text = filelook.Path
End Sub

Private Sub Form_Load()
If wins = 1 Then bnew.BackColor = &H855A00
If wins = 2 Then bnew.BackColor = &HE8A200
If wins = 3 Then bnew.BackColor = &H80000010
If wins = 4 Then bnew.BackColor = &H323232
tkill.Interval = wta
givecolor (wins)
filelook.Path = filepath
flpath.Text = filelook.Path
End Sub



Private Sub bnew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If wins = 1 Then bnew.BackColor = &H644600
If wins = 2 Then bnew.BackColor = &HB47800
If wins = 3 Then bnew.BackColor = &H80000011
If wins = 4 Then bnew.BackColor = &H0&
End Sub

Private Sub ok_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ok.BackColor = &H80000010
End Sub

Private Sub ok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
filepath = filelook.Path
fset.Show
Unload Me
End Sub


Private Sub tkill_Timer()
toola (3)
End Sub
