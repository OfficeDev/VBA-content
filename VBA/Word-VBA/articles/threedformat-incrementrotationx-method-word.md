---
title: ThreeDFormat.IncrementRotationX Method (Word)
keywords: vbawd10.chm164626442
f1_keywords:
- vbawd10.chm164626442
ms.prod: word
api_name:
- Word.ThreeDFormat.IncrementRotationX
ms.assetid: 5892375e-b6a5-ae2c-c85c-bc0798558407
ms.date: 06/08/2017
---


# ThreeDFormat.IncrementRotationX Method (Word)

Changes the rotation of the specified shape around the x-axis by the specified number of degrees.


## Syntax

 _expression_ . **IncrementRotationX**( **_Increment_** )

 _expression_ Required. A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much (in degrees) the rotation of the shape around the x-axis is to be changed. Can be a value from ?90 through 90. A positive value tilts the shape up; a negative value tilts it down.|

## Remarks

Use the  **[RotationX](threedformat-rotationx-property-word.md)** property to set the absolute rotation of the shape around the x-axis.

You cannot adjust the rotation around the x-axis of the specified shape past the upper or lower limit for the  **RotationX** property (90 degrees to ?90 degrees). For example, if the **RotationX** property is initially set to 80 and you specify 40 for the Increment argument, the resulting rotation will be 90 (the upper limit for the **RotationX** property) instead of 120.

To change the rotation of a shape around the y-axis, use the  **IncrementRotationY** method. To change the rotation around the z-axis, use the **IncrementRotation** method.


## Example

This example tilts the first shape on the active document up 10 degrees. The first shape must be an extruded shape for you to see the effect of this code.


```vb
ActiveDocument.Shapes(1).ThreeD.IncrementRotationX 10
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)

