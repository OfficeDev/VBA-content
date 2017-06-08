---
title: Shape.Delete Method (PowerPoint)
keywords: vbapp10.chm547003
f1_keywords:
- vbapp10.chm547003
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Delete
ms.assetid: 998a345f-31e3-1270-7826-17d84d60634b
ms.date: 06/08/2017
---


# Shape.Delete Method (PowerPoint)

Deletes the specified  **Shape** object.


## Syntax

 _expression_. **Delete**

 _expression_ A variable that represents a **Shape** object.


## Example

This example deletes all freeform shapes from slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes

    For intShape = .Count To 1 Step -1

        With .Item(intShape)

            If .Type = msoFreeform Then .Delete

        End With

    Next

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

