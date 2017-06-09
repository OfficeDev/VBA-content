---
title: Zoom.Percentage Property (Word)
keywords: vbawd10.chm161873920
f1_keywords:
- vbawd10.chm161873920
ms.prod: word
api_name:
- Word.Zoom.Percentage
ms.assetid: 4d49583f-6991-3c6d-fcf4-535e6663c3b7
ms.date: 06/08/2017
---


# Zoom.Percentage Property (Word)

Returns or sets the magnification for a window as a percentage. Read/write  **Long** .


## Syntax

 _expression_ . **Percentage**

 _expression_ An expression that returns a **[Zoom](zoom-object-word.md)** object.


## Example

This example switches the active window to normal view and sets the magnification to 80 percent.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdNormalView 
 .Zoom.Percentage = 80 
End With
```

This example increases the magnification of the active window by 10 percent.




```vb
Set myZoom = ActiveDocument.ActiveWindow.View.Zoom 
myZoom.Percentage = myZoom.Percentage + 10
```


## See also


#### Concepts


[Zoom Object](zoom-object-word.md)

