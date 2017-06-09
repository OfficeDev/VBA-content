---
title: Determine When a Shelled Process Ends
ms.prod: access
ms.assetid: 16a6fb03-0ff5-76a9-8efb-9348d5a6beef
ms.date: 06/08/2017
---


# Determine When a Shelled Process Ends

When you run the  **[Shell](http://msdn.microsoft.com/library/033bffb0-540f-2c17-2aed-d25d10bedd8c%28Office.15%29.aspx)** function in a Visual Basic for Applications (VBA) procedure, it starts an executable program asynchronously and returns control to the procedure. This shelled program continues to run independently of your procedure until you close it.

If your procedure needs to wait for the shelled process to end, you can use the Windows API to poll the status of the application, but this is not very efficient. This topic explains a more efficient method. 

The Windows API has integrated functionality that enables your application to wait until a shelled process has completed. To use these functions, you need to have a handle to the shelled process. To accomplish this, use the  **CreateProcess** function instead of the **Shell** function to begin your shelled program.


## Creating the Shelled Process

To create an addressable process, use the  **CreateProcess** function to start your shelled application. The **CreateProcess** function gives your program the process handle of the shelled process via one of its passed parameters.


## Waiting for the Shelled Process to End

After you use the  **CreateProcess** function to get a process handle, you can pass that handle to the **WaitForSingleObject** function. This causes your VBA procedure to suspend execution until the shelled process ends.

The following steps are necessary to build a VBA procedure that uses the  **CreateProcess** function to run the Windows Notepad application. This code shows how to use the Windows API **CreateProcess** and **WaitForSingleObject** functions to wait until a shelled process ends before resuming execution.

The syntax of the  **CreateProcess** function is complex, so in the example code, it is encapsulated into a function called **ExecCmd**. **ExecCmd** takes one parameter, the command line of the application to execute.

Create a standard module and paste the following lines in the Declarations section: 




```vb
Option Explicit 
 
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
 
Private Const NORMAL_PRIORITY_CLASS = &;H20&; 
Private Const INFINITE = -1&; 

```

Paste the following code into the module:




```vb
Public Sub ExecCmd(cmdline As String) 
 Dim proc As PROCESS_INFORMATION 
 Dim start As STARTUPINFO 
 Dim ReturnValue As Integer 
 
 ' Initialize the STARTUPINFO structure: 
 start.cb = Len(start) 
 
 ' Start the shelled application: 
 ReturnValue = CreateProcessA(0&;, cmdline$, 0&;, 0&;, 1&;, _ 
 NORMAL_PRIORITY_CLASS, 0&;, 0&;, start, proc) 
 
 ' Wait for the shelled application to finish: 
 Do 
 ReturnValue = WaitForSingleObject(proc.hProcess, 0) 
 DoEvents 
 Loop Until ReturnValue <> 258 
 
 ReturnValue = CloseHandle(proc.hProcess) 
End Sub 
 
Sub Testing() 
 ExecCmd "NOTEPAD.EXE" 
 MsgBox "Process Finished" 
End Sub
```

Type  **Testing** in the **Immediate** window. Notepad will start. After you close Notepad, a message box will notify you that the process has completed.


