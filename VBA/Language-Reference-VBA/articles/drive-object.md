---
title: Drive Object
keywords: vblr6.chm2181923
f1_keywords:
- vblr6.chm2181923
ms.prod: office
api_name:
- Office.Drive
ms.assetid: 95229345-790b-d77d-c3b4-6b4998aa0336
ms.date: 06/08/2017
---


# Drive Object



 **Description**
Provides access to the properties of a particular disk drive or network share.
 **Remarks**
The following code illustrates the use of the  **Drive** object to access drive properties:



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


