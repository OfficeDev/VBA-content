---
title: AvailableSpace Property
keywords: vblr6.chm2181954
f1_keywords:
- vblr6.chm2181954
ms.prod: office
api_name:
- Office.AvailableSpace
ms.assetid: c7a2a011-1b90-7091-4dcb-0149c75a6ee6
ms.date: 06/08/2017
---


# AvailableSpace Property



 **Description**
Returns the amount of space available to a user on the specified drive or network share.
 **Syntax**
 _object_. **AvailableSpace**
The  _object_ is always a **Drive** object.
 **Remarks**
The value returned by the  **AvailableSpace** property is typically the same as that returned by the **FreeSpace** property. Differences may occur between the two values for computer systems that support quotas.
The following code illustrates the use of the  **AvailableSpace** property:



```vb
Sub ShowAvailableSpace(drvPath)
    Dim fs, d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    s = "Drive " &; UCase(drvPath) &; " - " 
    s = s &; d.VolumeName  &; vbCrLf
    s = s &; "Available Space: " &; FormatNumber(d.AvailableSpace/1024, 0) 
    s = s &; " Kbytes"
    MsgBox s
End Sub
```


