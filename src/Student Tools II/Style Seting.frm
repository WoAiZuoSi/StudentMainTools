VERSION 5.00
Begin VB.Form fchoose 
   Appearance      =   0  'Flat
   BackColor       =   &H00E8A200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Tools II - Style Seting"
   ClientHeight    =   3732
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4068
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
   ScaleWidth      =   4068
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.OptionButton w1 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "Win8 (Default)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1572
   End
   Begin VB.OptionButton w2 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "Win8 (Opposite)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   1812
   End
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
      TabIndex        =   6
      Top             =   3240
      Width           =   4092
   End
   Begin VB.Label fro2 
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   $"Style Seting.frx":8582
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
    back.BackColor = &H855A00
End If

If a = 2 Then
    fchoose.BackColor = &H855A00
    w1.BackColor = &HE8A200
    w2.BackColor = &HE8A200
    w3.BackColor = &HE8A200
    w4.BackColor = &HE8A200
    fro1.BackColor = &HE8A200
    fro2.BackColor = &HE8A200
    back.BackColor = &HE8A200
End If

If a = 3 Then
    fchoose.BackColor = &H323232
    w1.BackColor = &H80000010
    w2.BackColor = &H80000010
    w3.BackColor = &H80000010
    w4.BackColor = &H80000010
    fro1.BackColor = &H80000010
    fro2.BackColor = &H80000010
    back.BackColor = &H80000010
End If

If a = 4 Then
    fchoose.BackColor = &H80000010
    w1.BackColor = &H323232
    w2.BackColor = &H323232
    w3.BackColor = &H323232
    w4.BackColor = &H323232
    fro1.BackColor = &H323232
    fro2.BackColor = &H323232
    back.BackColor = &H323232
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

Private Sub Form_Load()
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
