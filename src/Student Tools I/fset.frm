VERSION 5.00
Begin VB.Form fset 
   BackColor       =   &H00E8A200&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Tools I - Set"
   ClientHeight    =   5640
   ClientLeft      =   36
   ClientTop       =   360
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
   Icon            =   "fset.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   3252
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton ffilep 
      BackColor       =   &H00FFFFFF&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   4.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   20
      Top             =   2532
      Width           =   252
   End
   Begin VB.TextBox scwt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   1212
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
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2892
   End
   Begin VB.TextBox scend 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   1572
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
      Left            =   240
      TabIndex        =   4
      Top             =   3480
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
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   2892
   End
   Begin VB.TextBox kiwt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1440
      TabIndex        =   2
      Top             =   4560
      Width           =   1212
   End
   Begin VB.TextBox filep 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   2172
   End
   Begin VB.TextBox tkother 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   480
      TabIndex        =   0
      Top             =   4200
      Width           =   2532
   End
   Begin VB.Label Label2 
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
      Height          =   492
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   3252
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ms"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   2760
      TabIndex        =   17
      Top             =   1920
      Width           =   252
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   16
      Top             =   1848
      Width           =   1332
   End
   Begin VB.Label back 
      Alignment       =   2  'Center
      BackColor       =   &H00855A03&
      Caption         =   "Back"
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
      Left            =   0
      TabIndex        =   15
      Top             =   5160
      Width           =   3252
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File Name :"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1500
      Width           =   1452
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ScreenShot - Set"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   2892
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Killer - Set"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   2892
   End
   Begin VB.Label Label10 
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
      Left            =   240
      TabIndex        =   10
      Top             =   4476
      Width           =   1332
   End
   Begin VB.Label Label11 
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
      Left            =   2760
      TabIndex        =   9
      Top             =   4476
      Width           =   372
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   8
      Top             =   2196
      Width           =   1332
   End
   Begin VB.Label Label7 
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
      Height          =   2172
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   3012
   End
   Begin VB.Label Label8 
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
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   3012
   End
End
Attribute VB_Name = "fset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ffilep_Click()
ffile.Show
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
'MsgBox wtb
'back.Caption = usescr

usesc.Value = usescr
scend.Text = scrend
scwt = wta
filep.Text = filepath

useki.Value = usekiller
kother.Value = killother
tkother.Text = processname
kiwt.Text = wtb

If killother = 1 Then
    tkother.Enabled = True
Else
    tkother.Enabled = False
End If
End Sub

Private Sub back_Click()
usescr = usesc.Value
scrend = scend.Text
wta = scwt.Text
filepath = filep.Text

usekiller = useki.Value
killother = kother.Value
processname = tkother.Text
wtb = kiwt.Text

'MsgBox wtb



Unload Me
fmain.Show
End Sub


Private Sub kother_Click()
If kother.Value = 1 Then
    tkother.Enabled = True
Else
    tkother.Enabled = False
End If
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
Else
    scend.Enabled = False
    scwt.Enabled = False
    filep.Enabled = False
End If
End Sub

