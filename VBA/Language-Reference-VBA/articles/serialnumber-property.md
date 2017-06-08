---
title: SerialNumber Property
keywords: vblr6.chm2181962
f1_keywords:
- vblr6.chm2181962
ms.prod: office
api_name:
- Office.SerialNumber
ms.assetid: fdeb1410-3772-7f41-9a48-3bb7d2bd107a
ms.date: 06/08/2017
---


# SerialNumber Property



 **Description**
Returns the decimal serial number used to uniquely identify a disk volume.
 **Syntax**
 _object_. **SerialNumber**
The  _object_ is always a **Drive** object.
 **Remarks**
You can use the  **SerialNumber** property to ensure that the correct disk is inserted in a drive with removable media.
The following code illustrates the use of the  **SerialNumber** property:



```vb
Sub ShowDriveInfo(drvpath)
    Dim fs, d, s, t
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(drvpath)))
    Select Case d.DriveType
        Case 0: t = "Unknown"
        Case 1: t = "Removable"
        Case 2: t = "Fixed"
        Case 3: t = "Network"
        Case 4: t = "CD-ROM"
        Case 5: t = "RAM Disk"
    End Select
    s = "Drive " &; d.DriveLetter &; ": - " &; t
    s = s &; vbCrLf &; "SN: " &; d.SerialNumber
    MsgBox s
End Sub
```


