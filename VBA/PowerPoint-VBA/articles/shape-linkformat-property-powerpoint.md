---
title: Shape.LinkFormat Property (PowerPoint)
keywords: vbapp10.chm547045
f1_keywords:
- vbapp10.chm547045
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.LinkFormat
ms.assetid: b742d78a-2fd3-1eb9-76d1-f2a2263cc68a
ms.date: 06/08/2017
---


# Shape.LinkFormat Property (PowerPoint)

Returns a  **[LinkFormat](linkformat-object-powerpoint.md)** object that contains the properties that are unique to linked OLE objects. Read-only.


## Syntax

 _expression_. **LinkFormat**

 _expression_ A variable that represents a **Shape** object.


### Return Value

LinkFormat


## Example

This example updates the links between any OLE objects on slide one in the active presentation and their source files.


```vb
For Each sh In ActivePresentation.Slides(1).Shapes

    If sh.Type = msoLinkedOLEObject Then

        With sh.LinkFormat

            .Update

        End With

    End If

Next
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

