---
title: ShareName Property
keywords: vblr6.chm2181963
f1_keywords:
- vblr6.chm2181963
ms.prod: office
api_name:
- Office.ShareName
ms.assetid: 913ae336-102c-9c1c-4995-9b37aae79b3e
ms.date: 06/08/2017
---


# ShareName Property



 **Description**
Returns the network share name for a specified drive.
 **Syntax**
 _object_. **ShareName**
The  _object_ is always a **Drive** object.
 **Remarks**
If  _object_ is not a network drive, the **ShareName** property returns a zero-length string ("").
The following code illustrates the use of the  **ShareName** property:



```vb
Sub ShowDriveInfo(drvpath)
    Dim fs, d, s 
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(drvpath)))
    s = "Drive " &; d.DriveLetter &; ": - " &; d.ShareName
    MsgBox s
End Sub
```


