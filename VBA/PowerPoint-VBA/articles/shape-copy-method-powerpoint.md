---
title: Shape.Copy Method (PowerPoint)
keywords: vbapp10.chm547051
f1_keywords:
- vbapp10.chm547051
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Copy
ms.assetid: 41c82fd1-9ee7-c937-0a75-77b84c33c972
ms.date: 06/08/2017
---


# Shape.Copy Method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

 _expression_. **Copy**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-powerpoint.md)

