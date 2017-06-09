---
title: DriveType Property
keywords: vblr6.chm2181956
f1_keywords:
- vblr6.chm2181956
ms.prod: office
api_name:
- Office.DriveType
ms.assetid: 398dbcdb-9b39-1694-cdd0-499bc0d34704
ms.date: 06/08/2017
---


# DriveType Property



 **Description**
Returns a value indicating the type of a specified drive.
 **Syntax**
 _object_. **DriveType**
The  _object_ is always a **Drive** object.
 **Remarks**
The following code illustrates the use of the  **DriveType** property:



```vb
Sub ShowDriveType(drvpath)
    Dim fs, d, s, t
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(drvpath)
    Select Case d.DriveType
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    s = "Drive " &; d.DriveLetter &; ": - " &; t
    MsgBox s
End Sub
```


