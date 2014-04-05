Attribute VB_Name = "vbcCommand"
Option Explicit

'=========================================================
' mathimagics@yahoo.co.uk
'=========================================================
'
' MVBLC Link Control Tool:   Module "vbcCommand"
'
'   EXECUTE(commandstring, modalflag)
'      Executes a Windows command line.  If the modal flag
'      is non-zero the command is executed modally (we
'      wait for the command execution to finish). We need
'      this function to invoke the real VB6 linker, as we
'      can't use the VB Shell function (it has no modal
'      option).
'
'=========================================================

Type STARTUPINFO  ' structure used with CreateProcess API
   cb As Long
   lpReserved As String
   lpDesktop As Long
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

Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
   End Type


Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
         hHandle As Long, ByVal dwMilliseconds As Long) As Long

Declare Function CreateProcessA Lib "kernel32" (ByVal _
         lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
         lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
         ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
         ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
         lpStartupInfo As STARTUPINFO, lpProcessInformation As _
         PROCESS_INFORMATION) As Long


Sub Execute(WinCommand As String, ByVal Modal As Long)
   
   Const NORMAL_PRIORITY_CLASS = &H20&
   Dim ProcInfo   As PROCESS_INFORMATION
   Dim StartInfo  As STARTUPINFO
   Dim NullString As Long
   
   StartInfo.cb = Len(StartInfo)
   StartInfo.lpDesktop = VarPtr(NullString)
   StartInfo.dwFlags = 1
   
   Call CreateProcessA(0&, WinCommand, 0&, 0&, 1&, _
            NORMAL_PRIORITY_CLASS, 0&, 0&, _
            StartInfo, ProcInfo)
            
   If Modal Then Call WaitForSingleObject(ProcInfo.hProcess, -1)

End Sub


