---
title: Application.PrintPreview Property (Word)
keywords: vbawd10.chm158335003
f1_keywords:
- vbawd10.chm158335003
ms.prod: word
api_name:
- Word.Application.PrintPreview
ms.assetid: 6f522dc1-60ad-d5c1-029b-961fce1992e5
ms.date: 06/08/2017
---


# Application.PrintPreview Property (Word)

 **True** if print preview is the current view. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintPreview**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example switches the view to print preview.


```vb
PrintPreview = True
```

This example switches the active window from print preview to normal view.




```vb
PrintPreview = False 
ActiveDocument.ActiveWindow.View.Type = wdNormalView
```


## See also


#### Concepts


[Application Object](application-object-word.md)

