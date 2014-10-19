VERSION 5.00
Begin VB.Form fabout 
   BackColor       =   &H00E8A200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Tools II - About"
   ClientHeight    =   2520
   ClientLeft      =   48
   ClientTop       =   408
   ClientWidth     =   3360
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
   ScaleHeight     =   2520
   ScaleWidth      =   3360
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Label back 
      Alignment       =   2  'Center
      BackColor       =   &H00855A00&
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
      TabIndex        =   2
      Top             =   2040
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
         Size            =   21.6
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
      Caption         =   $"fabout.frx":8582
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
givecolor (wins)
End Sub
