---
title: ShapeRange.Cut Method (PowerPoint)
keywords: vbapp10.chm548050
f1_keywords:
- vbapp10.chm548050
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Cut
ms.assetid: 0e86d67c-7d52-4f3a-4cdd-6363667600a1
ms.date: 06/08/2017
---


# ShapeRange.Cut Method (PowerPoint)

Deletes the specified object and places it on the Clipboard.


## Syntax

 _expression_. **Cut**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example deletes shapes one and two from slide one in the active presentation, places copies of them on the Clipboard, and then pastes the copies onto slide two.


```vb
With ActivePresentation

    .Slides(1).Shapes.Range(Array(1, 2)).Cut

    .Slides(2).Shapes.Paste

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

