---
title: LineFormat.Weight Property (Word)
keywords: vbawd10.chm164233329
f1_keywords:
- vbawd10.chm164233329
ms.prod: word
api_name:
- Word.LineFormat.Weight
ms.assetid: 81439a12-175e-9ea6-7fd8-ee4207a23752
ms.date: 06/08/2017
---


# LineFormat.Weight Property (Word)

Returns or sets the thickness of the specified line in points. Read/write  **Single** .


## Syntax

 _expression_ . **Weight**

 _expression_ An expression that returns a **[LineFormat](lineformat-object-word.md)** object.


## Example

This example adds a green dashed line two points thick to the active document.


```vb
With ActiveDocument.Shapes.AddLine(10, 10, 250, 250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(0, 255, 255) 
 .Weight = 2 
End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-word.md)

