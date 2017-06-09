---
title: Calling the Windows API (Differences in String Function Operations)
ms.prod: access
ms.assetid: ee882d00-46f5-2bfc-09fc-ce2941302c5e
ms.date: 06/08/2017
---


# Calling the Windows API (Differences in String Function Operations)

  

**Applies to:** Access 2013 | Access 2016

The memory storage formats for text differ between Visual Basic for Applications (VBA) code and Access Basic code. (Access Basic was used in early versions of Microsoft Access.) Text is stored in ANSI format within Access Basic code and in Unicode format in Visual Basic. This topic discusses one potential issue when handling strings in the current version of Microsoft Access. For more information, see [Differences in String Function Operations](http://msdn.microsoft.com/library/40ce2b9a-cac6-589e-2b5e-d63be37efeee%28Office.15%29.aspx).

In several Windows API functions, the byte length of a string has a special meaning. For example, the following program returns a folder set up in Windows. In Microsoft Access,  **LeftB** (Buffer, ret) does not return the correct string. This is because, in spite of the fact that it shows the byte length of an ANSI string, the **LeftB** function processes Unicode strings. In this case, use the **InStr** function so that only the character string, without nulls, is returned.



```vb
Private Declare Function GetWindowsDirectory Lib "kernel32" _ 
 Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _ 
 ByVal nSize As Long) As Long 
 
Private Sub Command1_Click() 
 Buffer$ = Space(255) 
 ret = GetWindowsDirectory(Buffer$, 255) 
 ' WinDir = LeftB(Buffer, ret) '<--- Incorrect code" 
 
 WinDir = Left(Buffer$, InStr(Buffer$, Chr(0)) - 1) 
 '<--Correct code" 
 Print WinDir 
End Sub
```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

