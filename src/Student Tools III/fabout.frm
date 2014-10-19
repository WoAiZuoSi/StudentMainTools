VERSION 5.00
Begin VB.Form fabout 
   BackColor       =   &H00E8A200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Tools III - About"
   ClientHeight    =   2880
   ClientLeft      =   48
   ClientTop       =   408
   ClientWidth     =   3372
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fabout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3372
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tkill 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label goweb 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "→ 去我们的官网"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1920
      Width           =   2412
   End
   Begin VB.Label back 
      Alignment       =   2  'Center
      BackColor       =   &H00855A00&
      Caption         =   "Back"
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
      Top             =   2400
      Width           =   3372
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "About"
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
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label l1 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   $"fabout.frx":9CCA
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
      Height          =   1092
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3132
   End
End
Attribute VB_Name = "fabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Sub givecolor(a As Long)

If a = 1 Then
    fabout.BackColor = &HE8A200
    l1.BackColor = &H855A00
    back.BackColor = &H855A00
End If

If a = 2 Then
    fabout.BackColor = &H855A00
    l1.BackColor = &HE8A200
    back.BackColor = &HE8A200
End If

If a = 3 Then
    fabout.BackColor = &H323232
    l1.BackColor = &H80000010
    back.BackColor = &H80000010
End If

If a = 4 Then
    fabout.BackColor = &H80000010
    l1.BackColor = &H323232
    back.BackColor = &H323232
End If

End Sub

Private Sub back_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If wins = 1 Then back.BackColor = &H644600
If wins = 2 Then back.BackColor = &HB47800
If wins = 3 Then back.BackColor = &H80000011
If wins = 4 Then back.BackColor = &H0&
End Sub

Private Sub back_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
fmain.Show
End Sub

Private Sub Form_Load()

tkill.Interval = wta

givecolor (wins)
End Sub

Private Sub goweb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
goweb.BackColor = &H40C0&
End Sub

Private Sub goweb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ShellExecute Me.hwnd, "open", "http://www.woaizuosi.icoc.cc/", "", "", 5
goweb.BackColor = &H80FF&
End Sub

Private Sub tkill_Timer()
toola (1)
End Sub
