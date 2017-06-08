---
title: DateLastModified Property
keywords: vblr6.chm2181975
f1_keywords:
- vblr6.chm2181975
ms.prod: office
api_name:
- Office.DateLastModified
ms.assetid: 5b8c6ee5-e514-a374-8725-ece0658b579a
ms.date: 06/08/2017
---


# DateLastModified Property



 **Description**
Returns the date and time that the specified file or folder was last modified. Read-only.
 **Syntax**
 _object_. **DateLastModified**
The  _object_ is always a **File** or **Folder** object.
 **Remarks**
The following code illustrates the use of the  **DateLastModified** property with a file:



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


