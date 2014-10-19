Attribute VB_Name = "Gib"
Private Declare Function AddFontResource Lib "gdi32" Alias "AddFontResourceA" (ByVal lpFileName As String) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Type PointAPI
    X As Long
    Y As Long
End Type

Dim ScreenPoint As PointAPI

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long


Public usescr As Long
Public scrend As String
Public wta As Long
Public filepath As String
Public fsmode As String

Public usekiller As Long
Public killother As Long
Public processname As String
Public wtb As Long

Public FC As Long

Public mode As Long
'0 Õý³£
'1 Ãâ´òÈÅ

Public canuse As Long

Public username As String
Public userpass As String

Public wins As Long


'1 fabout
'2 fchoose
'3 ffile
'4 flogin
'5 fmain
'6 fset
'7 ftask

Public Sub toola(a As Long)

    GetCursorPos ScreenPoint
    If canuse = 1 And mode = 0 And usekiller = 1 And ScreenPoint.X = 0 And ScreenPoint.Y = 0 Then
        If killother = 1 Then
            Shell App.Path + "\ntsd.exe -c q -pn " + processname, vbNormalFocus
        Else
            Shell App.Path + "\ntsd.exe -c q -pn StudentMain.exe", vbNormalFocus
        End If
    End If
    
End Sub
Public Sub toolb(a As Long)

    GetCursorPos ScreenPoint
    xxw = Screen.Width \ Screen.TwipsPerPixelX - 2
    If canuse = 1 And mode = 0 And usescr = 1 And ScreenPoint.X > xxw And ScreenPoint.Y = 0 Then
        If a = 1 Then fabout.Hide
        If a = 2 Then fchoose.Hide
        If a = 3 Then ffile.Hide
        If a = 4 Then flogin.Hide
        If a = 5 Then fmain.Hide
        If a = 6 Then fset.Hide
        If a = 7 Then ftask.Hide
        
        If a = 1 Then
            BitBlt fabout.hDC, 0, 0, Screen.Width, Screen.Height, _
            GetDC(GetActiveWindow), 0, 0, vbSrcCopy
        End If
        If a = 2 Then
            BitBlt fchoose.hDC, 0, 0, Screen.Width, Screen.Height, _
            GetDC(GetActiveWindow), 0, 0, vbSrcCopy
        End If
        If a = 3 Then
            BitBlt ffile.hDC, 0, 0, Screen.Width, Screen.Height, _
            GetDC(GetActiveWindow), 0, 0, vbSrcCopy
        End If
        If a = 4 Then
            BitBlt flogin.hDC, 0, 0, Screen.Width, Screen.Height, _
            GetDC(GetActiveWindow), 0, 0, vbSrcCopy
        End If
        If a = 5 Then
            BitBlt fmain.hDC, 0, 0, Screen.Width, Screen.Height, _
            GetDC(GetActiveWindow), 0, 0, vbSrcCopy
        End If
        If a = 6 Then
            BitBlt fset.hDC, 0, 0, Screen.Width, Screen.Height, _
            GetDC(GetActiveWindow), 0, 0, vbSrcCopy
        End If
        If a = 7 Then
            BitBlt ftask.hDC, 0, 0, Screen.Width, Screen.Height, _
            GetDC(GetActiveWindow), 0, 0, vbSrcCopy
        End If
        
        
        Dim sFile As String
        If fsmode = 1 Then sFile = filepath & "\" & Format(Now, "yyyy_mm_dd - hh_mm_ss") & scrend
        If fsmode = 2 Then sFile = filepath & "\" & Format(Now, "yyyymmdd - hhmmss") & scrend
        If fsmode = 3 Then sFile = filepath & "\" & Format(Now, "hh_mm_ss") & scrend
        If fsmode = 4 Then
            Dim lss As String, howm As Long
            lss = Dir(filepath & "\*.*")
            Do Until lss = ""
                howm = howm + 1
                lss = Dir
            Loop
            sFile = filepath & "\" & howm & scrend
        'sFile = filepath & "\" & Format(Now, "yyyy_mm_dd - hh_mm_ss") & scrend
        End If
        
        If a = 1 Then SavePicture fabout.Image, sFile
        If a = 2 Then SavePicture fchoose.Image, sFile
        If a = 3 Then SavePicture ffile.Image, sFile
        If a = 4 Then SavePicture flogin.Image, sFile
        If a = 5 Then SavePicture fmain.Image, sFile
        If a = 6 Then SavePicture fset.Image, sFile
        If a = 7 Then SavePicture ftask.Image, sFile
        
        If a = 1 Then fabout.Picture = LoadPicture("")
        If a = 2 Then fchoose.Picture = LoadPicture("")
        If a = 3 Then ffile.Picture = LoadPicture("")
        If a = 4 Then flogin.Picture = LoadPicture("")
        If a = 5 Then fmain.Picture = LoadPicture("")
        If a = 6 Then fset.Picture = LoadPicture("")
        If a = 7 Then ftask.Picture = LoadPicture("")
        
        
        If a = 1 Then fabout.Show
        If a = 2 Then fchoose.Show
        If a = 3 Then ffile.Show
        If a = 4 Then flogin.Show
        If a = 5 Then fmain.Show
        If a = 6 Then fset.Show
        If a = 7 Then ftask.Show
        
        'Me.Picture = LoadPicture("")
        'Me.Show
        'MsgBox "1"
        'flash.Show
        'MsgBox "2"
        'Unload Me
        'Form3.Show
    End If

End Sub


