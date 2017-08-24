---
title: ThreeDFormat.IncrementRotationX Method (Publisher)
keywords: vbapb10.chm3801104
f1_keywords:
- vbapb10.chm3801104
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.IncrementRotationX
ms.assetid: d64204d6-ff4e-aa25-7795-858006ba2cf2
ms.date: 06/08/2017
---


# ThreeDFormat.IncrementRotationX Method (Publisher)

Changes the rotation of the specified shape around the x-axis (horizontal) by the specified number of degrees.


## Syntax

 _expression_. **IncrementRotationX**( **_Increment_**)

 _expression_A variable that represents a  **ThreeDFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Increment|Required| **Single**|Specifies by how many degrees to rotate the shape around the x-axis. Can be a value from - 90 through 90. A positive value tilts the shape up; a negative value tilts it down.|

## Remarks

Use the  **[RotationX](threedformat-rotationx-property-publisher.md)** property to set the absolute rotation of the shape around the x-axis.

You cannot adjust the rotation around the x-axis of the specified shape past the upper or lower limit for the  **RotationX** property (90 degrees to - 90 degrees). For example, if the **RotationX** property is initially set to 80 and you specify 40 for the **_Increment_** argument, the resulting rotation will be 90 (the upper limit for the **RotationX** property) instead of 120.

To change the rotation of a shape around the y-axis (vertical), use the  **[IncrementRotationY](threedformat-incrementrotationy-method-publisher.md)** method. To change the rotation around the z-axis (extends outward from the plane of the publication), use the  **[IncrementRotation](shape-incrementrotation-method-publisher.md)** method.


## Example

This example tilts the first shape in the active publication up 10 degrees. The shape must be an extruded shape for you to see the effect of this code.


```vb
ActiveDocument.Pages(1).Shapes(1).ThreeD _ 
 .IncrementRotationX Increment:=10
```


