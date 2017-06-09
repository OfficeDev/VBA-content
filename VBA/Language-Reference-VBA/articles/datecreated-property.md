---
title: DateCreated Property
keywords: vblr6.chm2181973
f1_keywords:
- vblr6.chm2181973
ms.prod: office
api_name:
- Office.DateCreated
ms.assetid: 2b176d36-d598-f922-ceba-989411368253
ms.date: 06/08/2017
---


# DateCreated Property



 **Description**
Returns the date and time that the specified file or folder was created. Read-only.
 **Syntax**
 _object_. **DateCreated**
The  _object_ is always a **File** or **Folder** object.
 **Remarks**
The following code illustrates the use of the  **DateCreated** property with a file:



```vb
Sub ShowFileInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = "Created: " &; f.DateCreated
    MsgBox s
End Sub
```


