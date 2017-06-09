---
title: ShortPath Property
keywords: vblr6.chm2181998
f1_keywords:
- vblr6.chm2181998
ms.prod: office
api_name:
- Office.ShortPath
ms.assetid: 9d473ea7-d555-0d79-9dfc-4822aa99ccd8
ms.date: 06/08/2017
---


# ShortPath Property



 **Description**
Returns the short path used by programs that require the earlier 8.3 file naming convention.
 **Syntax**
 _object_. **ShortPath**
The  _object_ is always a **File** or **Folder** object.
 **Remarks**
The following code illustrates the use of the  **ShortName** property with a **File** object:



```vb
Sub ShowShortPath(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = "The short path for " &; "" &; UCase(f.Name)
    s = s &; "" &; vbCrLf
    s = s &; "is: " &; "" &; f.ShortPath &; ""
    MsgBox s, 0, "Short Path Info"
End Sub
```


