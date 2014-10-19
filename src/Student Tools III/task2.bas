Attribute VB_Name = "task2"
Private Declare Function CloseHandle Lib "Kernel32.dll " (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll " (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcesses Lib "PSAPI.DLL " (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL " (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL " (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long


'返回符合进程名称的所有进程PID
'如果为没有，则返回空   (Empty)
Public Function GetProcessIdFromProcessName(ByVal strExeName As String) As Variant
        On Error Resume Next
        Const clMaxNumProcesses     As Long = 5000
        Const MAX_PATH = 260
        Const PROCESS_QUERY_INFORMATION = 1024
        Const PROCESS_VM_READ = 16
        Dim strModuleName     As String * MAX_PATH
        Dim strProcessNamePath     As String
        Dim strProcessName     As String
        Dim allMatchingProcessIDs()     As Long
        Dim alModules(1 To 400)         As Long
        Dim lBytesReturned     As Long
        Dim lNumMatching     As Long
        Dim lNumProcesses     As Long
        Dim lBytesNeeded     As Long
        Dim alProcIDs()     As Long
        Dim lHwndProcess     As Long
        Dim lThisProcess     As Long
        Dim lRet     As Long
        On Error GoTo Z
        strExeName = UCase$(Trim$(strExeName))
        ReDim alProcIDs(clMaxNumProcesses * 4) As Long
        lRet = EnumProcesses(alProcIDs(1), clMaxNumProcesses * 4, lBytesReturned)
        lNumProcesses = lBytesReturned / 4
        ReDim Preserve alProcIDs(lNumProcesses)
        ReDim allMatchingProcessIDs(1 To lNumProcesses)
        For lThisProcess = 1 To lNumProcesses
                If lHwndProcess > 0 Then lRet = CloseHandle(lHwndProcess)
                lHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, alProcIDs(lThisProcess))
                If lHwndProcess <> 0 Then
                      lRet = EnumProcessModules(lHwndProcess, alModules(1), 200&, lBytesNeeded)
                      If lRet <> 0 Then
                            lRet = GetModuleFileNameExA(lHwndProcess, alModules(1), strModuleName, MAX_PATH)
                            strProcessNamePath = Trim$(UCase$(Left$(strModuleName, lRet)))
                            strProcessName = Mid$(strProcessNamePath, InStrRev(strProcessNamePath, "\ ") + 1)
                            If strProcessName = strExeName Then
                                  lNumMatching = lNumMatching + 1
                                  allMatchingProcessIDs(lNumMatching) = alProcIDs(lThisProcess)
                            End If
                      End If
                      If lHwndProcess > 0 Then lRet = CloseHandle(lHwndProcess)
                End If
        Next
        If lNumMatching Then
              ReDim Preserve allMatchingProcessIDs(1 To lNumMatching)
              GetProcessIdFromProcessName = allMatchingProcessIDs
        Else
              GetProcessIdFromProcessName = Empty
        End If
        Exit Function
Z:
        GetProcessIdFromProcessName = Empty
End Function

