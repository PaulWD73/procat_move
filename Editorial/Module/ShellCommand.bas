Attribute VB_Name = "ShellCommand"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' File name:    ShellCommand.bas
' System:       manageBackup
'
' Description:  Provides a single public function that returns whether the
'               Execution of a shell command happend without error, and the
'               Return code of that error
'
' Amendment history:
'   d1  dis    10Nov99 Created
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' Data structures that the win32 api needs
Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

' win32 api functions that need to be declared to be seen
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal _
    hHandle As Long, ByRef LPDWORD As Long) As Boolean

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Public Function ExecCmd(cmdline$) As Boolean
'   Runs a given command synchronously
'   Returns false if an error occured in this function
'   iReturnStatus is the return code of the shell command
'
'   d1  dis    02Nov99 Created
    
    Dim sSource As String
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim dWaitReturn As Double
    Dim dCloseReturn As Double
    Dim bExitCode As Boolean
    Dim vReturnStatus As Variant
    Dim vProcReturn As Variant
    
    sSource = "ShellCommand:ExecCmd"
    ExecCmd = False
    
    On Error GoTo errorhandler
    
    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)
    
    ' Start the shelled application:
    vProcReturn = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, _
       NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
    
    ' Wait for the shelled application to finish:
    dWaitReturn = WaitForSingleObject(proc.hProcess, INFINITE)
    
    bExitCode = GetExitCodeProcess(proc.hProcess, vReturnStatus)
    dCloseReturn = CloseHandle(proc.hProcess)
    
    
    ExecCmd = True
   
    Exit Function
    
errorhandler:
    ConvertError sSource

End Function
