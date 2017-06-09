---
title: Window.StyleAreaWidth Property (Word)
keywords: vbawd10.chm157417493
f1_keywords:
- vbawd10.chm157417493
ms.prod: word
api_name:
- Word.Window.StyleAreaWidth
ms.assetid: 2256deb8-1682-3c09-ac64-0557185c3d39
ms.date: 06/08/2017
---


# Window.StyleAreaWidth Property (Word)

Returns or sets the width of the style area in points. Read/write  **Single** .


## Syntax

 _expression_ . **StyleAreaWidth**

 _expression_ An expression that returns a **[Window](window-object-word.md)** object.


## Remarks

When the  **StyleAreaWidth** property is greater than 0 (zero), style names are displayed to the left of the text. The style area isn't visible in print layout or Web layout view.


## Example

This example switches the active window to normal view and sets the width of the style area to 1 inch.


```vb
With ActiveDocument.ActiveWindow 
 .View.Type = wdNormalView 
 .StyleAreaWidth = InchesToPoints(1) 
End With
```


## See also


#### Concepts


[Window Object](window-object-word.md)

