---
title: ThreeDFormat.Perspective Property (Publisher)
keywords: vbapb10.chm3801347
f1_keywords:
- vbapb10.chm3801347
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.Perspective
ms.assetid: 5a85f7fa-2c72-e9b0-75f0-e6d6680ecd99
ms.date: 06/08/2017
---


# ThreeDFormat.Perspective Property (Publisher)

 **msoTrue** if the extrusion appears in perspective — that is, if the walls of the extrusion narrow toward a vanishing point. **msoFalse** if the extrusion is a parallel, or orthographic, projection — that is, if the walls don't narrow toward a vanishing point. Read/write.


## Syntax

 _expression_. **Perspective**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

MsoTriState


## Example

This example sets the extrusion depth for shape one on the first page to 100 points and specifies that the extrusion be parallel, or orthographic. For this example to work, the specified shape must be a 3-D shape.


```vb
Sub ChangePerspective() 
 With ActiveDocument.Pages(1).Shapes(1).ThreeD 
 .Visible = True 
 .Depth = 100 
 .Perspective = msoFalse 
 End With 
End Sub
```


