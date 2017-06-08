---
title: IsRootFolder Property
keywords: vblr6.chm2182069
f1_keywords:
- vblr6.chm2182069
ms.prod: office
api_name:
- Office.IsRootFolder
ms.assetid: 4d47b8c1-9ca0-a6d4-996d-584d55033cc1
ms.date: 06/08/2017
---


# IsRootFolder Property



 **Description**
Returns  **True** if the specified folder is the root folder; **False** if it is not.
 **Syntax**
 _object_. **IsRootFolder**
The  _object_ is always a **Folder** object.
 **Remarks**
The following code illustrates the use of the  **IsRootFolder** property:



```vb
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
Sub DisplayLevelDepth(pathspec)
    Dim f, n
    Set f = fs.GetFolder(pathspec)
    If f.IsRootFolder Then
        MsgBox "The specified folder is the root folder."
    Else
        Do Until f.IsRootFolder
            Set f = f.ParentFolder
            n = n + 1
        Loop
        MsgBox "The specified folder is nested " &; n &; " levels deep."
    End If
End Sub
```


