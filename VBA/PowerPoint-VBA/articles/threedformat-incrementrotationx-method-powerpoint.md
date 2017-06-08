---
title: ThreeDFormat.IncrementRotationX Method (PowerPoint)
keywords: vbapp10.chm557002
f1_keywords:
- vbapp10.chm557002
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.IncrementRotationX
ms.assetid: 39f421e7-93d6-0744-9065-ac4d95738879
ms.date: 06/08/2017
---


# ThreeDFormat.IncrementRotationX Method (PowerPoint)

Changes the rotation of the specified shape around the x-axis by the specified number of degrees. 


## Syntax

 _expression_. **IncrementRotationX**( **_Increment_** )

 _expression_ A variable that represents an **ThreeDFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required|**Single**|Specifies how much (in degrees) the rotation of the shape around the x-axis is to be changed. Can be a value from ? 90 through 90. A positive value tilts the shape up; a negative value tilts it down.|

## Remarks

Use the  **[RotationX](threedformat-rotationx-property-powerpoint.md)** property to set the absolute rotation of the shape around the x-axis.

You cannot adjust the rotation around the x-axis of the specified shape past the upper or lower limit for the  **RotationX** property (90 degrees to ? 90 degrees). For example, if the **RotationX** property is initially set to 80 and you specify 40 for the Increment argument, the resulting rotation will be 90 (the upper limit for the **RotationX** property) instead of 120.

To change the rotation of a shape around the y-axis, use the  **[IncrementRotationY](threedformat-incrementrotationy-method-powerpoint.md)** method. To change the rotation around the z-axis, use the **[IncrementRotation](shape-incrementrotation-method-powerpoint.md)** method.


## Example

This example tilts shape one on  `myDocument` up 10 degrees. Shape one must be an extruded shape for you to see the effect of this code.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).ThreeD.IncrementRotationX 10
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-powerpoint.md)

