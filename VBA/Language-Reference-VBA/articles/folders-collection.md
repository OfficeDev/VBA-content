---
title: Folders Collection
keywords: vblr6.chm2181929
f1_keywords:
- vblr6.chm2181929
ms.prod: office
api_name:
- Office.Folders
ms.assetid: 84c95d58-9183-4820-bd45-817164497234
ms.date: 06/08/2017
---


# Folders Collection



 **Description**
Collection of all  **Folder** objects contained within a **Folder** object.
 **Remarks**
The following code illustrates how to get a  **Folders** collection and how to iterate the collection using the **For Each...Next** statement:



```vb
Sub ShowFolderList(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.SubFolders
    For Each f1 in fc
        s = s &; f1.name 
        s = s &;  vbCrLf
    Next
    MsgBox s
End Sub
```


