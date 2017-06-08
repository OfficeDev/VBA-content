---
title: ShapeRange.Copy Method (PowerPoint)
keywords: vbapp10.chm548051
f1_keywords:
- vbapp10.chm548051
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Copy
ms.assetid: ddc0dad9-6647-e2f4-393a-347c273656dd
ms.date: 06/08/2017
---


# ShapeRange.Copy Method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

 _expression_. **Copy**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks

Use the  **Paste** method to paste the contents of the Clipboard.


## Example

This example copies shapes one and two on slide one in the active presentation to the Clipboard and then pastes the copies onto slide two.


```vb
With ActivePresentation

    .Slides(1).Shapes.Range(Array(1, 2)).Copy

    .Slides(2).Shapes.Paste

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

