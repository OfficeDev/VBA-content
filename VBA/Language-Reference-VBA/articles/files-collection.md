---
title: Files Collection
keywords: vblr6.chm2181926
f1_keywords:
- vblr6.chm2181926
ms.prod: office
api_name:
- Office.Files
ms.assetid: 1c69f6df-debc-448a-6f22-a2a41d069dc4
ms.date: 06/08/2017
---


# Files Collection



 **Description**
Collection of all  **File** objects within a folder.
 **Remarks**
The following code illustrates how to get a  **Files** collection and iterate the collection using the **For Each...Next** statement:



```vb
Sub ShowFolderList(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 in fc
        s = s &; f1.name 
        s = s &; vbCrLf
    Next
    MsgBox s
End Sub
```


