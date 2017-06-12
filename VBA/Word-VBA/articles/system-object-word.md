---
title: System Object (Word)
keywords: vbawd10.chm2357
f1_keywords:
- vbawd10.chm2357
ms.prod: word
api_name:
- Word.System
ms.assetid: db15d780-3bbc-9515-a988-ea798777496f
ms.date: 06/08/2017
---


# System Object (Word)

Contains information about the computer system.


## Remarks

Use the  **System** property to return the **System** object. If the operating system is Windows, the following example makes a network connection to \\Project\Info.


```vb
If System.OperatingSystem = "Windows" Then 
 System.Connect Path:="\\Project\Info" 
End If
```

The following example displays the current screen resolution (for example, "1024 x 768").




```
horz = System.HorizontalResolution 
vert = System.VerticalResolution 
MsgBox "Resolution = " &; horz &; " x " &; vert
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


