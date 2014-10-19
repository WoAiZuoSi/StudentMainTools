VERSION 5.00
Begin VB.Form fchoose 
   Appearance      =   0  'Flat
   BackColor       =   &H00E8A200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Tools III - Style Seting"
   ClientHeight    =   3732
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4080
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Style Seting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3732
   ScaleWidth      =   4080
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer tkill 
      Left            =   0
      Top             =   0
   End
   Begin VB.OptionButton w1 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "Win8 (Default)"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   1572
   End
   Begin VB.OptionButton w2 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "Win8 (Opposite)"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   1812
   End
   Begin VB.OptionButton w4 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "Gray Style"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1812
   End
   Begin VB.OptionButton w3 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "Gray (Opposite)"
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1812
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
      Left            =   2040
      TabIndex        =   8
      Top             =   3240
      Width           =   2052
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
      TabIndex        =   7
      Top             =   3240
      Width           =   2052
   End
   Begin VB.Label fro2 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   $"Style Seting.frx":9CCA
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
      Height          =   1332
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3852
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Style Choose"
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
      Height          =   612
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   4092
   End
   Begin VB.Label fro1 
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
      Height          =   732
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3852
   End
End
Attribute VB_Name = "fchoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lswin As Long

Private Sub givecolor(a As Long)

'If b = 1 Then
If a = 1 Then
    fchoose.BackColor = &HE8A200
    w1.BackColor = &H855A00
    w2.BackColor = &H855A00
    w3.BackColor = &H855A00
    w4.BackColor = &H855A00
    fro1.BackColor = &H855A00
    fro2.BackColor = &H855A00
End If

If a = 2 Then
    fchoose.BackColor = &H855A00
    w1.BackColor = &HE8A200
    w2.BackColor = &HE8A200
    w3.BackColor = &HE8A200
    w4.BackColor = &HE8A200
    fro1.BackColor = &HE8A200
    fro2.BackColor = &HE8A200
End If

If a = 3 Then
    fchoose.BackColor = &H323232
    w1.BackColor = &H80000010
    w2.BackColor = &H80000010
    w3.BackColor = &H80000010
    w4.BackColor = &H80000010
    fro1.BackColor = &H80000010
    fro2.BackColor = &H80000010
End If

If a = 4 Then
    fchoose.BackColor = &H80000010
    w1.BackColor = &H323232
    w2.BackColor = &H323232
    w3.BackColor = &H323232
    w4.BackColor = &H323232
    fro1.BackColor = &H323232
    fro2.BackColor = &H323232
End If

'End If

End Sub

Private Sub back_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If wins = 1 Then back.BackColor = &H644600
If wins = 2 Then back.BackColor = &HB47800
If wins = 3 Then back.BackColor = &H80000011
If wins = 4 Then back.BackColor = &H0&
End Sub

Private Sub back_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
fset.Show
End Sub

Private Sub cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cancel.BackColor = &HC0&
End Sub

Private Sub cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
wins = lswin
Unload Me
fset.Show
End Sub

Private Sub tkill_Timer()
toola (2)
End Sub

Private Sub Form_Load()

tkill.Interval = wta
lswin = wins
givecolor (wins)
If wins = 1 Then
    w1.Value = 1
End If

If wins = 2 Then
    w2.Value = 1
End If

If wins = 3 Then
    w3.Value = 1
End If

If wins = 4 Then
    w4.Value = 1
End If

End Sub


Private Sub ok_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ok.BackColor = &HC000&
End Sub

Private Sub ok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Unload Me
fset.Show
End Sub

Private Sub w1_Click()
givecolor (1)
wins = 1
End Sub

Private Sub w2_Click()
givecolor (2)
wins = 2
End Sub

Private Sub w3_Click()
givecolor (3)
wins = 3
End Sub

Private Sub w4_Click()
givecolor (4)
wins = 4
End Sub
