---
title: Path Property (FileSystemObject object)
keywords: vblr6.chm2181960
f1_keywords:
- vblr6.chm2181960
ms.prod: office
ms.assetid: 15eed13b-9252-e195-0c54-9e3c82ce987f
ms.date: 06/08/2017
---


# Path Property (FileSystemObject object)



 **Description**
Returns the path for a specified file, folder, or drive.
 **Syntax**
 _object_. **Path**
The  _object_ is always a **File**, **Folder**, or **Drive** object.
 **Remarks**
For drive letters, the root drive is not included. For example, the path for the C drive is C:, not C:\.
The following code illustrates the use of the  **Path** property with a **File** object:



```vb
Sub ShowFileAccessInfo(filespec)
    Dim fs, d, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = UCase(f.Path) &; vbCrLf
    s = s &; "Created: " &; f.DateCreated &; vbCrLf
    s = s &; "Last Accessed: " &; f.DateLastAccessed &; vbCrLf
    s = s &; "Last Modified: " &; f.DateLastModified  
    MsgBox s, 0, "File Access Info"
End Sub
```


