---
title: DateLastAccessed Property
keywords: vblr6.chm2181974
f1_keywords:
- vblr6.chm2181974
ms.prod: office
api_name:
- Office.DateLastAccessed
ms.assetid: 33ab9638-8c76-98ca-4d48-b9f39ad71025
ms.date: 06/08/2017
---


# DateLastAccessed Property



 **Description**
Returns the date and time that the specified file or folder was last accessed. Read-only.
 **Syntax**
 _object_. **DateLastAccessed**
The  _object_ is always a **File** or **Folder** object.
 **Remarks**
The following code illustrates the use of the  **DateLastAccessed** property with a file:



```vb
Sub ShowFileAccessInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = UCase(filespec) &; vbCrLf
    s = s &; "Created: " &; f.DateCreated &; vbCrLf
    s = s &; "Last Accessed: " &; f.DateLastAccessed &; vbCrLf
    s = s &; "Last Modified: " &; f.DateLastModified  
    MsgBox s, 0, "File Access Info"
End Sub
```



 **Important**  This method depends on the underlying operating system for its behavior. If the operating system does not support providing time information, none will be returned.



