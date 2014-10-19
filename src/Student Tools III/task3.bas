Attribute VB_Name = "task3"
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long

Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 1024
End Type

Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPPROCESS = &H2
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Dim Pid As Long
Dim pname As String
Const sEndProess As String = "explorer.exe" '注意必须小写，是关闭的进程名称


Public Sub CloseProess(ProessFile As String)
Dim my As PROCESSENTRY32
Dim l As Long
Dim l1 As Long
Dim flag As Boolean
Dim mName As String
Dim i As Integer

l = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) '进程快照
If l Then
   my.dwSize = 1060
   If (Process32First(l, my)) Then '遍历第一个进程
  Do
   i = InStr(1, my.szExeFile, Chr(0))
   mName = LCase(Left(my.szExeFile, i - 1))
   ftask.list.AddItem DQStr(mName, 35, 1) + "" + Format(my.th32ProcessID)
   If mName = LCase(ProessFile) Then
      Pid = my.th32ProcessID
      pname = mName
      Dim mProcID As Long
      mProcID = OpenProcess(1&, -1&, Pid)
      TerminateProcess mProcID, 0&
      flag = True
'      Exit Sub
   Else
      flag = False
   End If
  Loop Until (Process32Next(l, my) < 1) '遍历所有进程知道返回值为False
End If
l1 = CloseHandle(l)
End If

End Sub

