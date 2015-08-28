
# CurDir Function

 **Last modified:** July 28, 2015


Returns a  **Variant** ( **String**) representing the current path.
 **Syntax**
 **CurDir**[ **(**_drive_**)**]
The optional  _drive_ [argument](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) is a [string expression](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) that specifies an existing drive. If no drive is specified or if _drive_ is a zero-length string (""), **CurDir** returns the path for the current drive. On the Macintosh, **CurDir** ignores any _drive_ specified and simply returns the path for the current drive.

## Example

This example uses the  **CurDir** function to return the current path. On the Macintosh, _drive_ specifications given to **CurDir** are ignored. The default drive name is "HD" and portions of the pathname are separated by colons instead of backslashes. Similarly, you would specify Macintosh folders instead of \Windows.


```
' Assume current path on C drive is "C:\WINDOWS\SYSTEM" (on Microsoft Windows).
' Assume current path on D drive is "D:\EXCEL".
' Assume C is the current drive.
Dim MyPath
MyPath = CurDir    ' Returns "C:\WINDOWS\SYSTEM".
MyPath = CurDir("C")    ' Returns "C:\WINDOWS\SYSTEM".
MyPath = CurDir("D")    ' Returns "D:\EXCEL".

```

