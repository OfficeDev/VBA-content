---
title: Type Property (FileSystemObject object)
keywords: vblr6.chm2182001
f1_keywords:
- vblr6.chm2182001
ms.prod: office
ms.assetid: b2e9bd7b-b9ea-1fe0-bd00-1f734d165e37
ms.date: 06/08/2017
---


# Type Property (FileSystemObject object)



 **Description**
Returns information about the type of a file or folder. For example, for files ending in .TXT, "Text Document" is returned.
 **Syntax**
 _object_. **Type**
The  _object_ is always a **File** or **Folder** object.
 **Remarks**
The following code illustrates the use of the  **Type** property to return a folder type. In this example, try providing the path of the Recycle Bin or other unique folder to the procedure.



```vb
Sub ShowFileSize(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(filespec)
    s = UCase(f.Name) &; " is a " &; f.Type 
    MsgBox s, 0, "File Size Info"
End Sub
```


