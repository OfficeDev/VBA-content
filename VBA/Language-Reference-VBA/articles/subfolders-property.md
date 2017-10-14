---
title: SubFolders Property
keywords: vblr6.chm2182070
f1_keywords:
- vblr6.chm2182070
ms.prod: office
api_name:
- Office.SubFolders
ms.assetid: 60bc795f-22f9-6846-00d3-05229f062099
ms.date: 06/08/2017
---


# SubFolders Property



 **Description**
Returns a  **Folders** collection consisting of all folders contained in a specified folder, including those with Hidden and System file attributes set.
 **Syntax**
 _object_. **SubFolders**
The  _object_ is always a **Folder** object.
 **Remarks**
The following code illustrates the use of the  **SubFolders** property:



```vb
Sub ShowFolderList(folderspec)
    Dim fs, f, f1, s, sf
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set sf = f.SubFolders
    For Each f1 in sf
        s = s &; f1.name 
        s = s &;  vbCrLf
    Next
    MsgBox s
End Sub
```


