---
title: ThreeDFormat.Perspective Property (PowerPoint)
keywords: vbapp10.chm557010
f1_keywords:
- vbapp10.chm557010
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.Perspective
ms.assetid: 1da4fd78-c1ae-16c8-0232-71cc0b2273e2
ms.date: 06/08/2017
---


# ThreeDFormat.Perspective Property (PowerPoint)

Determines whether the extrusion appears in perspective. Read/write.


## Syntax

 _expression_. **Perspective**

 _expression_ A variable that represents a **ThreeDFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Perspective** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The extrusion is a parallel, or orthographic, projection?that is, if the walls don't narrow toward a vanishing point. |
|**msoTrue**| The extrusion appears in perspective?that is, if the walls of the extrusion narrow toward a vanishing point.|

## Example

This example sets the extrusion depth for shape one on  `myDocument` to 100 points and specifies that the extrusion be parallel, or orthographic.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    .Visible = True

    .Depth = 100

    .Perspective = msoFalse

End With
```


