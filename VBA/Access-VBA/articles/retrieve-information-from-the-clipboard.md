---
title: Retrieve Information from the Clipboard
ms.prod: access
ms.assetid: 593d3047-c6c8-ab22-cdeb-aadc8b56ca81
ms.date: 06/08/2017
---


# Retrieve Information from the Clipboard

## Using the RunCommand Method

You can use the  **[RunCommand](docmd-runcommand-method-access.md)** method with the **acCmdPaste** constant to paste the contents of the Clipboard into the active control on a form or report. The following example illustrates how to paste the contents of the Clipboard into a text box named txtNotes.


```vb
Private Sub cmdPaste_Click() 
   Me!txtNotes.SetFocus 
   DoCmd.RunCommand acCmdPaste 
End Sub
```


## Using the Windows API

To use API calls to retrieve information from the Clipboard, paste the following code into the Declarations section of a standard module.


```vb
Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) _ 
   As Long 
Declare Function CloseClipboard Lib "User32" () As Long 
Declare Function GetClipboardData Lib "User32" (ByVal wFormat As _ 
   Long) As Long 
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags&;, ByVal _ 
   dwBytes As Long) As Long 
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) _ 
   As Long 
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) _ 
   As Long 
Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) _ 
   As Long 
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _ 
   ByVal lpString2 As Any) As Long 
 
Public Const GHND = &;H42 
Public Const CF_TEXT = 1 
Public Const MAXSIZE = 4096
```

Paste the following code into a standard module.




```vb
Function ClipBoard_GetData() 
   Dim hClipMemory As Long 
   Dim lpClipMemory As Long 
   Dim MyString As String 
   Dim RetVal As Long 
 
   If OpenClipboard(0&;) = 0 Then 
      MsgBox "Cannot open Clipboard. Another app. may have it open" 
      Exit Function 
   End If 
          
   ' Obtain the handle to the global memory 
   ' block that is referencing the text. 
   hClipMemory = GetClipboardData(CF_TEXT) 
   If IsNull(hClipMemory) Then 
      MsgBox "Could not allocate memory" 
      GoTo OutOfHere 
   End If 
 
   ' Lock Clipboard memory so we can reference 
   ' the actual data string. 
   lpClipMemory = GlobalLock(hClipMemory) 
 
   If Not IsNull(lpClipMemory) Then 
      MyString = Space$(MAXSIZE) 
      RetVal = lstrcpy(MyString, lpClipMemory) 
      RetVal = GlobalUnlock(hClipMemory) 
       
      ' Peel off the null terminating character. 
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1) 
   Else 
      MsgBox "Could not lock memory to copy string from." 
   End If 
 
OutOfHere: 
 
   RetVal = CloseClipboard() 
   ClipBoard_GetData = MyString 
 
End Function
```


