---
title: LineFormat.Transparency Property (Word)
keywords: vbawd10.chm164233327
f1_keywords:
- vbawd10.chm164233327
ms.prod: word
api_name:
- Word.LineFormat.Transparency
ms.assetid: c9b3adcb-c884-cfd1-6500-f430fdf86423
ms.date: 06/08/2017
---


# LineFormat.Transparency Property (Word)

Returns or sets the degree of transparency of line. Read/write  **Single** .


## Syntax

 _expression_ . **Transparency**

 _expression_ Required. A variable that represents a **[LineFormat](lineformat-object-word.md)** object.


## Remarks

 The value of this property can be expressed as a value between 0.0 (opaque) and 1.0 (clear). This property affects the appearance of solid-colored lines only; it has no effect on the appearance of patterned lines.


## Example

This example sets the shadow of shape three in the active document to semitransparent red. If the shape doesn't already have a shadow, this example adds one to it.


```vb
With ActiveDocument.Shapes(3).Shadow 
 .Visible = True 
 .ForeColor.RGB = RGB(255, 0, 0) 
 .Transparency = 0.5 
End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-word.md)

