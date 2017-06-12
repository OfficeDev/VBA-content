---
title: FileSystem Property
keywords: vblr6.chm2181957
f1_keywords:
- vblr6.chm2181957
ms.prod: office
api_name:
- Office.FileSystem
ms.assetid: 123ba29e-0b94-0afe-5f3d-323e903dd38e
ms.date: 06/08/2017
---


# FileSystem Property



 **Description**
Returns the type of file system in use for the specified drive.
 **Syntax**
 _object_. **FileSystem**
The  _object_ is always a **Drive** object.
 **Remarks**
Available return types include FAT, NTFS, and CDFS.
The following code illustrates the use of the  **FileSystem** property:



```vb
Sub ShowFileSystemType
    Dim fs,d, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive("e:")
    s = d.FileSystem
    MsgBox s
End Sub
```


