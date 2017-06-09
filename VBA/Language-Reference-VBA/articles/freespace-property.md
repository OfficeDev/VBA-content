---
title: FreeSpace Property
keywords: vblr6.chm2181958
f1_keywords:
- vblr6.chm2181958
ms.prod: office
api_name:
- Office.FreeSpace
ms.assetid: 6bb743ca-d196-a5ab-2cff-3dd6ecff6e0b
ms.date: 06/08/2017
---


# FreeSpace Property



 **Description**
Returns the amount of free space available to a user on the specified drive or network share. Read-only.
 **Syntax**
 _object_. **FreeSpace**
The  _object_ is always a **Drive** object.
 **Remarks**
The value returned by the  **FreeSpace** property is typically the same as that returned by the **AvailableSpace** property. Differences may occur between the two for computer systems that support quotas.
The following code illustrates the use of the  **FreeSpace** property:



```vb
Sub ShowFreeSpace(drvPath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    s = "Drive " &; UCase(drvPath) &; " - " 
    s = s &; d.VolumeName  &; vbCrLf
    s = s &; "Free Space: " &; FormatNumber(d.FreeSpace/1024, 0) 
    s = s &; " Kbytes"
    MsgBox s
End Sub
```


