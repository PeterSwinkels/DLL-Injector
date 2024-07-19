Attribute VB_Name = "DLL_InjectorModule"
'This module contains this program's core procedures.
Option Explicit

'Defines the Microsoft Windows API constants and functions used by this program.
Private Const ERROR_SUCCESS As Long = 0
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200&
Private Const MEM_COMMIT As Long = &H1000&
Private Const MEM_DECOMMIT As Long = &H4000&
Private Const PAGE_READWRITE As Long = &H4&
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Const WAIT_TIMEOUT As Long = &H102&

Private Declare Function CloseHandle Lib "Kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateRemoteThread Lib "Kernel32.dll" (ByVal ProcessHandle As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Any, ByVal lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function FormatMessageA Lib "Kernel32.dll" (ByVal dwFlags As Long, lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function GetExitCodeThread Lib "Kernel32.dll" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function GetModuleHandleA Lib "Kernel32.dll" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "Kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "Kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function VirtualAllocEx Lib "Kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal fAllocType As Long, FlProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "Kernel32.dll" (ByVal hProcess As Long, ByVal lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function WaitForSingleObject Lib "Kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function WriteProcessMemory Lib "Kernel32.dll" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

'The constants used by this program.
Private Const MAX_STRING As Long = 65535   'Defines the maximum number of characters used for a string buffer.
Private Const NO_HANDLE As Long = 0        'Defines "no handle".
Private Const NO_PID As Long = 0           'Defines "no process id".

'This procedure checks whether an error has occurred during the most recent Windows API call.
Private Function CheckForError(ReturnValue As Long) As Long
Dim Description As String
Dim ErrorCode As Long
Dim Length As Long

   ErrorCode = Err.LastDllError
   Err.Clear
   
   On Error GoTo ErrorTrap
   
   If Not ErrorCode = ERROR_SUCCESS Then
       Description = String$(MAX_STRING, vbNullChar)
       Length = FormatMessageA(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, CLng(0), ErrorCode, CLng(0), Description, Len(Description), CLng(0))
       If Length = 0 Then
          Description = "No description."
       ElseIf Length > 0 Then
          Description = Left$(Description, Length - 1)
       End If
      
       Description = "API error: " & CStr(ErrorCode) & vbCr & Description & vbCr
       Description = Description & "Return value: " & CStr(ReturnValue)
       MsgBox Description, vbExclamation
   End If
   
EndProcedure:
   CheckForError = ReturnValue
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure ejects the specified DLL from the specified process.
Private Function EjectDLL(InjectedDLLH As Long, ProcessH As Long) As Long
On Error GoTo ErrorTrap
Dim ExitCode As Long
Dim FreeLibraryAddress As Long
Dim ModuleH As Long
Dim ReturnValue As Long
Dim ThreadH As Long

   ExitCode = ERROR_SUCCESS
   ModuleH = CheckForError(GetModuleHandleA("Kernel32.dll"))
   If Not ModuleH = NO_HANDLE Then
      FreeLibraryAddress = CheckForError(GetProcAddress(ModuleH, "FreeLibrary"))
      If Not FreeLibraryAddress = 0 Then
         ThreadH = CheckForError(CreateRemoteThread(ProcessH, CLng(0), CLng(0), FreeLibraryAddress, InjectedDLLH, CLng(0), CLng(0)))
         If Not ThreadH = NO_HANDLE Then
            ReturnValue = CheckForError(WaitForSingleObject(ThreadH, CLng(1000)))
            If Not ReturnValue = WAIT_TIMEOUT Then
               ReturnValue = CheckForError(GetExitCodeThread(ThreadH, ExitCode))
               If Not ReturnValue = ERROR_SUCCESS Then
                  ReturnValue = CheckForError(CloseHandle(ThreadH))
               End If
            End If
         End If
      End If
   End If
   
EndProcedure:
   EjectDLL = ExitCode
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure handles any errors that occur.
Private Function HandleError(Optional ReturnPreviousChoice As Boolean = False) As Long
Dim Description As String
Dim ErrorCode As Long
Static Choice As Long

   Description = Err.Description
   ErrorCode = Err.Number
   On Error Resume Next
   If Not ReturnPreviousChoice Then
      Choice = MsgBox(Description & "." & vbCr & "Error code: " & CStr(ErrorCode), vbAbortRetryIgnore Or vbDefaultButton2 Or vbExclamation)
   End If
   
   If Choice = vbAbort Then End
   
   HandleError = Choice
End Function
'This procedure injects the specified DLL into the specified process.
Private Function InjectDLL(DLLPath As String, ProcessH As Long) As Long
On Error GoTo ErrorTrap
Dim BaseAddress As Long
Dim ExitCode As Long
Dim InjectedDLLH As Long
Dim ModuleH As Long
Dim ProcedureAddress As Long
Dim ReturnValue As Long
Dim ThreadH As Long

   InjectedDLLH = NO_HANDLE
   BaseAddress = CheckForError(VirtualAllocEx(ProcessH, CLng(0), Len(DLLPath), MEM_COMMIT, ByVal PAGE_READWRITE))
   If Not BaseAddress = 0 Then
      ReturnValue = CheckForError(WriteProcessMemory(ProcessH, BaseAddress, ByVal DLLPath, Len(DLLPath), CLng(0)))
      If Not ReturnValue = 0 Then
         ModuleH = CheckForError(GetModuleHandleA("Kernel32.dll"))
         If Not ModuleH = NO_HANDLE Then
            ProcedureAddress = CheckForError(GetProcAddress(ModuleH, "LoadLibraryA"))
            If Not ProcedureAddress = 0 Then
               ThreadH = CheckForError(CreateRemoteThread(ProcessH, CLng(0), CLng(0), ProcedureAddress, BaseAddress, CLng(0), CLng(0)))
               If Not ThreadH = NO_HANDLE Then
                  ReturnValue = CheckForError(WaitForSingleObject(ThreadH, CLng(1000)))
                  If Not ReturnValue = WAIT_TIMEOUT Then
                     ReturnValue = CheckForError(GetExitCodeThread(ThreadH, ExitCode))
                     If Not ReturnValue = 0 Then
                        CheckForError CloseHandle(ThreadH)
                        InjectedDLLH = ExitCode
                     End If
                  End If
               End If
            End If
         End If
      End If
      ReturnValue = CheckForError(VirtualFreeEx(ProcessH, BaseAddress, Len(DLLPath), MEM_DECOMMIT))
   End If
   
EndProcedure:
   InjectDLL = InjectedDLLH
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function


'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
Dim DLLPath As String
Dim ExitCode As Long
Dim InjectedDLLH As Long
Dim ProcessH As Long
Dim ProcessId As Long
Dim TargetPath As String

   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   TargetPath = InputBox$("Enter a program's path:", ProgramInformation())
   If TargetPath = vbNullString Then Exit Sub
   DLLPath = InputBox$("Enter a *.dll file's path:", ProgramInformation())
   If DLLPath = vbNullString Then Exit Sub
   
   If Dir$(DLLPath, vbArchive Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = vbNullString Then
      MsgBox "Could not find the specified DLL.", vbExclamation
   Else
      ProcessId = Shell(TargetPath, vbNormal)
      If Not ProcessId = NO_PID Then
         ProcessH = CheckForError(OpenProcess(PROCESS_ALL_ACCESS, CLng(True), ProcessId))
         If Not ProcessH = NO_HANDLE Then
            InjectedDLLH = InjectDLL(DLLPath, ProcessH)
            
            If InjectedDLLH = NO_HANDLE Then
               MsgBox "Could not inject DLL.", vbExclamation
            Else
               MsgBox "DLL has been injected. Handle: " & CStr(InjectedDLLH) & vbCr & "Click ""Ok"" to eject and terminate the process.", vbInformation
               ExitCode = EjectDLL(InjectedDLLH, ProcessH)
               MsgBox "DLL has been ejected. Exit code: " & CStr(ExitCode), vbInformation
            End If
         
            CheckForError TerminateProcess(ProcessH, ExitCode)
            CheckForError CloseHandle(ProcessH)
         End If
      End If
   End If
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub


'This procedure returns information about this program.
Private Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String

   With App
      Information = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With

EndProcedure:
   ProgramInformation = Information
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function
