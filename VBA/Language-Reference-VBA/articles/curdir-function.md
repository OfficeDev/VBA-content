---
title: CurDir Function
keywords: vblr6.chm1008881
f1_keywords:
- vblr6.chm1008881
ms.prod: office
ms.assetid: 5abad447-9c6b-8e9c-d6bb-f43f23dc45ad
ms.date: 06/08/2017
---


# CurDir Function



Returns a  **Variant** ( **String** ) representing the current path.
 **Syntax**
 **CurDir** [ **(**_drive_**)** ]
The optional  _drive_[argument](vbe-glossary.md) is a[string expression](vbe-glossary.md) that specifies an existing drive. If no drive is specified or if _drive_ is a zero-length string (""), **CurDir** returns the path for the current drive. On the Macintosh, **CurDir** ignores any _drive_ specified and simply returns the path for the current drive.

## Example

This example uses the  **CurDir** function to return the current path. On the Macintosh, _drive_ specifications given to **CurDir** are ignored. The default drive name is "HD" and portions of the pathname are separated by colons instead of backslashes. Similarly, you would specify Macintosh folders instead of \Windows.


```vb
' Assume current path on C drive is "C:\WINDOWS\SYSTEM" (on Microsoft Windows).
' Assume current path on D drive is "D:\EXCEL".
' Assume C is the current drive.
Dim MyPath
MyPath = CurDir    ' Returns "C:\WINDOWS\SYSTEM".
MyPath = CurDir("C")    ' Returns "C:\WINDOWS\SYSTEM".
MyPath = CurDir("D")    ' Returns "D:\EXCEL".

```


