---
title: Drive Property
keywords: vblr6.chm2181976
f1_keywords:
- vblr6.chm2181976
ms.prod: office
api_name:
- Office.Drive
ms.assetid: 34512359-067f-f625-5f19-db7b0faa0138
ms.date: 06/08/2017
---


# Drive Property



 **Description**
Returns the drive letter of the drive on which the specified file or folder resides. Read-only.
 **Syntax**
 _object_. **Drive**
The  _object_ is always a **File** or **Folder** object.
 **Remarks**
The following code illustrates the use of the  **Drive** property:



```vb
Sub ShowFileAccessInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.Name &; " on Drive " &; UCase(f.Drive) &; vbCrLf
    s = s &; "Created: " &; f.DateCreated &; vbCrLf
    s = s &; "Last Accessed: " &; f.DateLastAccessed &; vbCrLf
    s = s &; "Last Modified: " &; f.DateLastModified  
    MsgBox s, 0, "File Access Info"
End Sub
```


