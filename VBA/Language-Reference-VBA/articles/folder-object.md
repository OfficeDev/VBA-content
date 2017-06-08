---
title: Folder Object
keywords: vblr6.chm2181928
f1_keywords:
- vblr6.chm2181928
ms.prod: office
api_name:
- Office.Folder
ms.assetid: 877e81a5-5a34-9ef9-2375-3c60d35d3255
ms.date: 06/08/2017
---


# Folder Object



 **Description**
Provides access to all the properties of a folder.
 **Remarks**
The following code illustrates how to obtain a  **Folder** object and how to return one of its properties:



```vb
Sub ShowFolderInfo(folderspec)
    Dim fs, f, s,
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    s = f.DateCreated
    MsgBox s
End Sub
```


