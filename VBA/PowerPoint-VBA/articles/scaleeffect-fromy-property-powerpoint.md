---
title: ScaleEffect.FromY Property (PowerPoint)
keywords: vbapp10.chm660006
f1_keywords:
- vbapp10.chm660006
ms.prod: powerpoint
api_name:
- PowerPoint.ScaleEffect.FromY
ms.assetid: a63e5ec1-35c6-bb1e-58d2-57e2c7299f6e
ms.date: 06/08/2017
---


# ScaleEffect.FromY Property (PowerPoint)

Returns or sets a  **Single** that represents the starting height of a **[ScaleEffect](scaleeffect-object-powerpoint.md)** object, specified as a percentage of the screen width. Read/write.


## Syntax

 _expression_. **FromY**

 _expression_ A variable that represents a **ScaleEffect** object.


### Return Value

Single


## Remarks

The default value of this property is  **Empty**, in which case the current position of the object is used.

Use this property in conjunction with the  **ToY** property to resize or jump from one position to another.

Do not confuse this property with the  **From** property of the **[ColorEffect](coloreffect-object-powerpoint.md)**, **[RotationEffect](rotationeffect-object-powerpoint.md)**, or **[PropertyEffect](propertyeffect-object-powerpoint.md)** objects, which is used to set or change colors, rotations, or other properties of an animation behavior, respectively.


## See also


#### Concepts


[ScaleEffect Object](scaleeffect-object-powerpoint.md)

