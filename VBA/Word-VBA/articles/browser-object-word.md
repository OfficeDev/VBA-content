---
title: Browser Object (Word)
keywords: vbawd10.chm2350
f1_keywords:
- vbawd10.chm2350
ms.prod: word
api_name:
- Word.Browser
ms.assetid: 447bcab6-cfb2-77b0-9bbd-35e774417a60
ms.date: 06/08/2017
---


# Browser Object (Word)

Represents the browser tool used to move the insertion point to objects in a document. This tool is composed of the three buttons at the bottom of the vertical scroll bar.


## Remarks

Use the  **[Browser](application-browser-property-word.md)** property to return the **Browser** object. The following example moves the insertion point just before the next field in the active document.


```vb
With Application.Browser 
 .Target = wdBrowseField 
 .Next 
End With
```

The following example moves the insertion point to the previous table and selects it.




```vb
With Application.Browser 
 .Target = wdBrowseTable 
 .Previous 
End With 
If Selection.Information(wdWithInTable) = True Then 
 Selection.Tables(1).Select 
End If
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

