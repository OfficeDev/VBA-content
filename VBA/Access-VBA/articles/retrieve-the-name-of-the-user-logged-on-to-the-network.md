---
title: Retrieve the Name of the User Logged On To the Network
ms.prod: access
ms.assetid: 3bf335a1-08d0-c8d5-8d89-36f0c29d47d0
ms.date: 06/08/2017
---


# Retrieve the Name of the User Logged On To the Network

This topic contians a user-defined function, GetLogonName, that returns the current user name. The GetLogonName function utilizes the GetUserNameA Windows API to retrieve the current user name. 


```vb
' Access the GetUserNameA function in advapi32.dll and 
' call the function GetUserName. 
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _ 
 (ByVal lpBuffer As String, nSize As Long) As Long 
 
' Main routine to retrieve user name. 
Function GetLogonName() As String 
 
 ' Dimension variables 
 Dim lpBuff As String * 255 
 Dim ret As Long 
 
 ' Get the user name minus any trailing spaces found in the name. 
 ret = GetUserName(lpBuff, 255) 
 
 If ret > 0 Then 
 GetLogonName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1) 
 Else 
 GetLogonName = vbNullString 
 End If 
End Function
```


