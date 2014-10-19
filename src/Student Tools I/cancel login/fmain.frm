VERSION 5.00
Begin VB.Form fmain 
   Appearance      =   0  'Flat
   BackColor       =   &H00E8A200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "一键取消对 Student Tools I 的注册"
   ClientHeight    =   1560
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   3852
   Icon            =   "fmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3852
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton go 
      Caption         =   "一键取消对 Student Tools I 的注册"
      Height          =   1092
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3612
   End
End
Attribute VB_Name = "fmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum FO_Operation
    FO_MOVE = 1
    FO_COPY = 2
    FO_DELETE = 3
    FO_RENAME = 4
End Enum

Public Enum FOFlags
    FOF_MULTIDESTFILES = &H1 'Destination specifies multiple files
    FOF_SILENT = &H4 'Don't display progress dialog
    FOF_RENAMEONCOLLISION = &H8 'Rename if destination already exists
    FOF_NOCONFIRMATION = &H10 'Don't prompt user
    FOF_WANTMAPPINGHANDLE = &H20 'Fill in hNameMappings member
    FOF_ALLOWUNDO = &H40 'Store undo information if possible
    FOF_FILESONLY = &H80 'On *.*, don't copy directories
    FOF_SIMPLEPROGRESS = &H100 'Don't show name of each file
    FOF_NOCONFIRMMKDIR = &H200 'Don't confirm making any needed dirs
End Enum

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long ' only used if FOF_SIMPLEPROGRESS
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private op As SHFILEOPSTRUCT

Public Sub DeleteFolder(sDeleteFolder As String, Optional Interface As Boolean = False)
    
    SetAttr sDeleteFolder, vbNormal
    With op
        .wFunc = FO_DELETE
        .pFrom = sDeleteFolder
        .fFlags = IIf(Interface = False, FOF_NOCONFIRMATION, FOF_NOCONFIRMATION And FOF_SILENT)
    End With
    SHFileOperation op
    
End Sub

Private Sub go_Click()
On Error Resume Next
DeleteFolder ("C:\ProgramData\WoAiZuoSi")
DeleteFolder ("C:\Users\Public\Documents\WoAiZuoSi")
End Sub

