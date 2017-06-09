---
title: Drives Collection
keywords: vblr6.chm2181924
f1_keywords:
- vblr6.chm2181924
ms.prod: office
api_name:
- Office.Drives
ms.assetid: 729c2d39-5b4e-44f2-a9ed-4f06ba7ac1b7
ms.date: 06/08/2017
---


# Drives Collection



 **Description**
Read-only collection of all available drives.
 **Remarks**
Removable-media drives need not have media inserted for them to appear in the  **Drives** collection.
The following code illustrates how to get the  **Drives** collection and iterate the collection using the **For Each...Next** statement:



```vb
Sub ShowDriveList
    Dim fs, d, dc, s, n
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set dc = fs.Drives
    For Each d in dc
        s = s &; d.DriveLetter &; " - " 
        If d.DriveType = Remote Then
            n = d.ShareName
        Else
            n = d.VolumeName
        End If
        s = s &; n &; vbCrLf
    Next
    MsgBox s
End Sub
```


