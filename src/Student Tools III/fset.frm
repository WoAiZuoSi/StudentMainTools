VERSION 5.00
Begin VB.Form fset 
   BackColor       =   &H00E8A200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Tools III - Set"
   ClientHeight    =   7812
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   7.8
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7812
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tkill 
      Left            =   0
      Top             =   0
   End
   Begin VB.OptionButton m1 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "年_月_日-时_分_秒"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   25
      Top             =   3240
      Width           =   1812
   End
   Begin VB.OptionButton m4 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "当前文件夹的文件数"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2160
      TabIndex        =   22
      Top             =   3480
      Width           =   2172
   End
   Begin VB.OptionButton m3 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "时_分_秒"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   21
      Top             =   3480
      Width           =   1812
   End
   Begin VB.OptionButton m2 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "年月日 - 时分秒"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2160
      TabIndex        =   20
      Top             =   3240
      Width           =   1812
   End
   Begin VB.TextBox scwt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   2172
   End
   Begin VB.CheckBox usesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "Use ScreenShot Tools"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   384
      Left            =   360
      TabIndex        =   6
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2892
   End
   Begin VB.TextBox scend 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   2532
   End
   Begin VB.CheckBox useki 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "Use StudentMainKiller"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   384
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Value           =   1  'Checked
      Width           =   2892
   End
   Begin VB.CheckBox kother 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A03&
      Caption         =   "Kill Other Process"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   4680
      Width           =   2892
   End
   Begin VB.TextBox kiwt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1560
      TabIndex        =   2
      Top             =   5400
      Width           =   2172
   End
   Begin VB.TextBox filep 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   840
      TabIndex        =   1
      Top             =   2520
      Width           =   2772
   End
   Begin VB.TextBox tkother 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   840
      TabIndex        =   0
      Top             =   5040
      Width           =   3372
   End
   Begin VB.Label lwin 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   840
      TabIndex        =   31
      Top             =   6600
      Width           =   2052
   End
   Begin VB.Label choose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Choose"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   3360
      TabIndex        =   30
      Top             =   6600
      Width           =   852
   End
   Begin VB.Label l10 
      BackStyle       =   0  'Transparent
      Caption         =   "Window Style :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   360
      TabIndex        =   28
      Top             =   6300
      Width           =   2052
   End
   Begin VB.Label l9 
      BackStyle       =   0  'Transparent
      Caption         =   "Other Settings"
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
      Height          =   612
      Left            =   360
      TabIndex        =   27
      Top             =   5880
      Width           =   2892
   End
   Begin VB.Label ll1 
      BackStyle       =   0  'Transparent
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   3840
      TabIndex        =   29
      Top             =   1848
      Width           =   372
   End
   Begin VB.Label look 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "浏览"
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   3720
      TabIndex        =   26
      Top             =   2520
      Width           =   492
   End
   Begin VB.Label l2 
      BackStyle       =   0  'Transparent
      Caption         =   "ScreenShot Tool - Set"
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
      Left            =   360
      TabIndex        =   24
      Top             =   720
      Width           =   3852
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
      TabIndex        =   23
      Top             =   7320
      Width           =   2292
   End
   Begin VB.Label l6 
      BackStyle       =   0  'Transparent
      Caption         =   "File Save Mode :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   360
      TabIndex        =   19
      Top             =   2800
      Width           =   2052
   End
   Begin VB.Label l1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Set"
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
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   4572
   End
   Begin VB.Label l4 
      BackStyle       =   0  'Transparent
      Caption         =   "Wait Time :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   360
      TabIndex        =   15
      Top             =   1848
      Width           =   1332
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
      Left            =   2280
      TabIndex        =   14
      Top             =   7320
      Width           =   2292
   End
   Begin VB.Label l3 
      BackStyle       =   0  'Transparent
      Caption         =   "File Format :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   360
      TabIndex        =   13
      Top             =   1475
      Width           =   1452
   End
   Begin VB.Label l7 
      BackStyle       =   0  'Transparent
      Caption         =   "StudentMainKiller - Set"
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
      Left            =   360
      TabIndex        =   11
      Top             =   3960
      Width           =   3612
   End
   Begin VB.Label l8 
      BackStyle       =   0  'Transparent
      Caption         =   "Wait Time :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   360
      TabIndex        =   10
      Top             =   5316
      Width           =   1332
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "Segoe UI Light 8"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   3840
      TabIndex        =   9
      Top             =   5316
      Width           =   372
   End
   Begin VB.Label l5 
      BackStyle       =   0  'Transparent
      Caption         =   "File Path :"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   360
      TabIndex        =   8
      Top             =   2196
      Width           =   1332
   End
   Begin VB.Label fro1 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3132
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   4092
   End
   Begin VB.Label fro2 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A03&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1812
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Width           =   4092
   End
   Begin VB.Label fro3 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A03&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1212
      Left            =   240
      TabIndex        =   18
      Top             =   5880
      Width           =   4092
   End
End
Attribute VB_Name = "fset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fsm As Long
Public win As Long

Private Sub givecolor(a As Long)

If a = 1 Then
fset.BackColor = &HE8A200
fro1.BackColor = &H855A00
fro2.BackColor = &H855A00
fro3.BackColor = &H855A00
usesc.BackColor = &H855A00
m1.BackColor = &H855A00
m2.BackColor = &H855A00
m3.BackColor = &H855A00
m4.BackColor = &H855A00
useki.BackColor = &H855A00
kother.BackColor = &H855A00
End If

If a = 2 Then
fset.BackColor = &H855A00
fro1.BackColor = &HE8A200
fro2.BackColor = &HE8A200
fro3.BackColor = &HE8A200
usesc.BackColor = &HE8A200
m1.BackColor = &HE8A200
m2.BackColor = &HE8A200
m3.BackColor = &HE8A200
m4.BackColor = &HE8A200
useki.BackColor = &HE8A200
kother.BackColor = &HE8A200
End If

If a = 3 Then
fset.BackColor = &H323232
fro1.BackColor = &H80000010
fro2.BackColor = &H80000010
fro3.BackColor = &H80000010
usesc.BackColor = &H80000010
m1.BackColor = &H80000010
m2.BackColor = &H80000010
m3.BackColor = &H80000010
m4.BackColor = &H80000010
useki.BackColor = &H80000010
kother.BackColor = &H80000010
End If

If a = 4 Then
fset.BackColor = &H80000010
fro1.BackColor = &H323232
fro2.BackColor = &H323232
fro3.BackColor = &H323232
usesc.BackColor = &H323232
m1.BackColor = &H323232
m2.BackColor = &H323232
m3.BackColor = &H323232
m4.BackColor = &H323232
useki.BackColor = &H323232
kother.BackColor = &H323232
End If
End Sub
Private Sub cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BackColor = &HC0&
End Sub

Private Sub cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
fmain.Show
End Sub

Private Sub choose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
choose.BackColor = &H80000011
End Sub

Private Sub choose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
fchoose.Show
End Sub

Private Sub Form_Load()

tkill.Interval = wta
givecolor (wins)
'MsgBox wtb
'back.Caption = usescr
If usesc.Value = 1 Then
    scend.Enabled = True
    scwt.Enabled = True
    filep.Enabled = True
    m1.Enabled = True
    m2.Enabled = True
    m3.Enabled = True
    m4.Enabled = True
    look.Enabled = True
    
Else
    scend.Enabled = False
    scwt.Enabled = Falses
    filep.Enabled = False
    m1.Enabled = False
    m2.Enabled = False
    m3.Enabled = False
    m4.Enabled = False
    look.Enabled = False
End If
usesc.Value = usescr
scend.Text = scrend
scwt = wta
filep.Text = filepath
If fsmode = 1 Then m1.Value = 1
If fsmode = 2 Then m2.Value = 1
If fsmode = 3 Then m3.Value = 1
If fsmode = 4 Then m4.Value = 1

useki.Value = usekiller
kother.Value = killother
tkother.Text = processname
kiwt.Text = wtb

If killother = 1 Then
    tkother.Enabled = True
Else
    tkother.Enabled = False
End If

If wins = 1 Then lwin.Caption = "Win8 (Default)"
If wins = 2 Then lwin.Caption = "Win8 (Opposite)"
If wins = 3 Then lwin.Caption = "Gray Style"
If wins = 4 Then lwin.Caption = "Gray (Opposite)"
End Sub


Private Sub kother_Click()
If kother.Value = 1 Then
    tkother.Enabled = True
Else
    tkother.Enabled = False
End If
End Sub


Private Sub look_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
look.BackColor = &H80000011
End Sub

Private Sub look_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ffile.Show
Unload Me
End Sub

Private Sub m1_Click()
fsm = 1
End Sub

Private Sub m2_Click()
fsm = 2
End Sub

Private Sub m3_Click()
fsm = 3
End Sub

Private Sub m4_Click()
fsm = 4
End Sub

Private Sub ok_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ok.BackColor = &HC000&
End Sub

Private Sub ok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
usescr = usesc.Value
scrend = scend.Text
wta = scwt.Text
filepath = filep.Text

fsmode = fsm

usekiller = useki.Value
killother = kother.Value
processname = tkother.Text
wtb = kiwt.Text

'MsgBox wtb

Unload Me
fmain.Show
End Sub


Private Sub tkill_Timer()
toola (6)
End Sub



Private Sub useki_Click()
If useki.Value = 1 Then
    kother.Enabled = True
    If kother.Value = 1 Then
        tkother.Enabled = True
    Else
        tkother.Enabled = False
    End If
    kiwt.Enabled = True
Else
    kother.Enabled = False
    tkother.Enabled = False
    kiwt.Enabled = False
End If
End Sub


Private Sub usesc_Click()
If usesc.Value = 1 Then
    scend.Enabled = True
    scwt.Enabled = True
    filep.Enabled = True
    m1.Enabled = True
    m2.Enabled = True
    m3.Enabled = True
    m4.Enabled = True
    look.Enabled = True
    
Else
    scend.Enabled = False
    scwt.Enabled = False
    filep.Enabled = False
    m1.Enabled = False
    m2.Enabled = False
    m3.Enabled = False
    m4.Enabled = False
    look.Enabled = False
End If
End Sub

