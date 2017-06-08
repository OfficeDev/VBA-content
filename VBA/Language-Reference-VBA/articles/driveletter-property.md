---
title: DriveLetter Property
keywords: vblr6.chm2181955
f1_keywords:
- vblr6.chm2181955
ms.prod: office
api_name:
- Office.DriveLetter
ms.assetid: 29bf179a-8bf7-56de-3cf5-53fd0d2151e0
ms.date: 06/08/2017
---


# DriveLetter Property



 **Description**
Returns the drive letter of a physical local drive or a network share. Read-only.
 **Syntax**
 _object_. **DriveLetter**
The  _object_ is always a **Drive** object.
 **Remarks**
The  **DriveLetter** property returns a zero-length string ("") if the specified drive is not associated with a drive letter, for example, a network share that has not been mapped to a drive letter.
The following code illustrates the use of the  **DriveLetter** property:



```vb
Sub ShowDriveLetter(drvPath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    s = "Drive " &; d.DriveLetter &; ": - " 
    s = s &; d.VolumeName  &; vbCrLf
    s = s &; "Free Space: " &; FormatNumber(d.FreeSpace/1024, 0) 
    s = s &; " Kbytes"
    MsgBox s
End Sub
```


