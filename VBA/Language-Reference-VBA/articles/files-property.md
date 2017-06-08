---
title: Files Property
keywords: vblr6.chm2182095
f1_keywords:
- vblr6.chm2182095
ms.prod: office
api_name:
- Office.Files
ms.assetid: 80ee842f-759f-a018-c69c-4233d9714938
ms.date: 06/08/2017
---


# Files Property



 **Description**
Returns a  **Files** collection consisting of all **File** objects contained in the specified folder, including those with hidden and system file attributes set.
 **Syntax**
 _object_. **Files**
The  _object_ is always a **Folder** object.
 **Remarks**
The following code illustrates the use of the  **Files** property:



```vb
Sub ShowFileList(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 in fc
        s = s &; f1.name 
        s = s &;  vbCrLf
    Next
    MsgBox s
End Sub
```


