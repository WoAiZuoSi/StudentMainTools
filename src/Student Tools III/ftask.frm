VERSION 5.00
Begin VB.Form ftask 
   Appearance      =   0  'Flat
   BackColor       =   &H00E8A200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Tools III - Task List"
   ClientHeight    =   5760
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ftask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4800
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer tkill 
      Left            =   0
      Top             =   0
   End
   Begin VB.ListBox list 
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
      Height          =   4392
      ItemData        =   "ftask.frx":9CCA
      Left            =   120
      List            =   "ftask.frx":9CCC
      TabIndex        =   0
      Top             =   120
      Width           =   4572
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
      TabIndex        =   3
      Top             =   5280
      Width           =   4812
   End
   Begin VB.Label ref 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00855A00&
      Caption         =   "Refresh"
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
      TabIndex        =   2
      Top             =   4680
      Width           =   2412
   End
   Begin VB.Label endtask 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Kill the Process"
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
      TabIndex        =   1
      Top             =   4680
      Width           =   2412
   End
End
Attribute VB_Name = "ftask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub givecolor(a As Long)

If a = 1 Then
ftask.BackColor = &HE8A200
list.BackColor = &H855A00
ref.BackColor = &H855A00
back.BackColor = &H855A00
End If

If a = 2 Then
ftask.BackColor = &H855A00
list.BackColor = &HE8A200
ref.BackColor = &HE8A200
back.BackColor = &HE8A200
End If

If a = 3 Then
ftask.BackColor = &H323232
list.BackColor = &H80000010
ref.BackColor = &H80000010
back.BackColor = &H80000010
End If

If a = 4 Then
ftask.BackColor = &H80000010
list.BackColor = &H323232
ref.BackColor = &H323232
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

Private Sub endtask_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
endtask.BackColor = &HC0&
End Sub

Private Sub endtask_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
endtask.BackColor = &HFF&
Dim Str1Len As Long
Dim Pid As Long
Dim PFS As String
Dim l1 As Long
Dim i As Long
If list.Text = "" Then Exit Sub
l1 = list.ListIndex
For i = Len(list.Text) To 1 Step -1
    Str1Len = InStr(i, list.Text, " ")
    If Str1Len > 0 Then Exit For
Next i
Pid = Val(Trim(Mid(list.Text, Str1Len, Len(list.Text) - Str1Len + 1)))
Dim mProcID As Long
mProcID = OpenProcess(1&, -1&, Pid)
i = TerminateProcess(mProcID, 0&)
If i = 0 Then MsgBox "²Ù×÷Ê§°Ü  ", 0 + 48, ""
Delay 500
list.Clear
'List1.Clear
CloseProess ("")
If list.ListCount - 1 < l1 Then l1 = list.ListCount - 1
list.ListIndex = l1
End Sub

Private Sub Form_Load()
givecolor (wins)
tkill.Interval = wta
Dim l1 As Long
'Dim l2 As Long
l1 = list.ListIndex
'l2 = List1.ListIndex
list.Clear
CloseProess ("")
If list.ListCount - 1 < l1 Then l1 = list.ListCount - 1
    list.ListIndex = l1
'If List1.ListCount - 1 < l2 Then l2 = List1.ListCount - 1
'List1.ListIndex = l2
'Load Dialog
End Sub


Private Sub Label1_Click()

End Sub

Private Sub list_Click()
Dim Str1Len As Long
Dim Str2Long As Long
Dim i As Long
'List1.Clear
'For i = Len(list.Text) To 1 Step -1
  '     Str1Len = InStr(i, list.Text, " ")
  '     If Str1Len > 0 Then Exit For
  'Next i
    
    'Str2Long = Val(Trim(Mid(list.Text, Str1Len, Len(list.Text) - Str1Len + 1)))
'    yy = GetProcessIdFromProcessName()
    'Find_Window Str2Long
End Sub

Private Sub ref_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If wins = 1 Then ref.BackColor = &H644600
If wins = 2 Then ref.BackColor = &HB47800
If wins = 3 Then ref.BackColor = &H80000011
If wins = 4 Then ref.BackColor = &H0&
End Sub

Private Sub ref_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If wins = 1 Then ref.BackColor = &H855A00
If wins = 2 Then ref.BackColor = &HE8A200
If wins = 3 Then ref.BackColor = &H80000010
If wins = 4 Then ref.BackColor = &H323232
Dim l1 As Long
'Dim l2 As Long
l1 = list.ListIndex
'l2 = List1.ListIndex
list.Clear
Call CloseProess("")
If list.ListCount - 1 < l1 Then l1 = List2.ListCount - 1
list.ListIndex = l1
'If List1.ListCount - 1 < l2 Then l2 = List1.ListCount - 1
'List1.ListIndex = l2
'Load Dialog
End Sub

Private Sub tkill_Timer()
toola (7)
End Sub


