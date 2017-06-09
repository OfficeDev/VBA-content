---
title: Shape.Cut Method (PowerPoint)
keywords: vbapp10.chm547050
f1_keywords:
- vbapp10.chm547050
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Cut
ms.assetid: 908c998d-a15f-5075-33e0-de6c124a0ec7
ms.date: 06/08/2017
---


# Shape.Cut Method (PowerPoint)

Deletes the specified object and places it on the Clipboard.


## Syntax

 _expression_. **Cut**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-powerpoint.md)

