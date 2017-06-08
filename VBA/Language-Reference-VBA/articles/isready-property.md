---
title: IsReady Property
keywords: vblr6.chm2181959
f1_keywords:
- vblr6.chm2181959
ms.prod: office
api_name:
- Office.IsReady
ms.assetid: e4c0771b-ea30-1431-2106-ca53a13543f2
ms.date: 06/08/2017
---


# IsReady Property



 **Description**
Returns  **True** if the specified drive is ready; **False** if it is not.
 **Syntax**
object. **IsReady**
The object is always a  **Drive** object.
 **Remarks**
For removable-media drives and CD-ROM drives,  **IsReady** returns **True** only when the appropriate media is inserted and ready for access.
The following code illustrates the use of the  **IsReady** property:



```vb
Sub ShowDriveInfo(drvpath)
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
    If d.IsReady Then 
        s = s &; vbCrLf &; "Drive is Ready."
    Else
        s = s &; vbCrLf &; "Drive is not Ready."
    End If
    MsgBox s
End Sub
```


