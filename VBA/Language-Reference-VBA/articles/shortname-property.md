---
title: ShortName Property
keywords: vblr6.chm2181997
f1_keywords:
- vblr6.chm2181997
ms.prod: office
api_name:
- Office.ShortName
ms.assetid: 62d95787-61c7-777d-56d0-d17d4d8e0f18
ms.date: 06/08/2017
---


# ShortName Property



 **Description**
Returns the short name used by programs that require the earlier 8.3 naming convention.
 **Syntax**
 _object_. **ShortName**
The  _object_ is always a **File** or **Folder** object.
 **Remarks**
The following code illustrates the use of the  **ShortName** property with a **File** object:



```vb
Sub ShowShortName(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = "The short name for " &; "" &; UCase(f.Name)
    s = s &; "" &; vbCrLf
    s = s &; "is: " &; "" &; f.ShortName &; ""
    MsgBox s, 0, "Short Name Info"
End Sub
```


