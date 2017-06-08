---
title: ThreeDFormat.IncrementRotationY Method (Publisher)
keywords: vbapb10.chm3801105
f1_keywords:
- vbapb10.chm3801105
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.IncrementRotationY
ms.assetid: 54260253-c914-6600-60ef-17bdde12be59
ms.date: 06/08/2017
---


# ThreeDFormat.IncrementRotationY Method (Publisher)

Changes the rotation of the specified shape around the y-axis (vertical) by the specified number of degrees.


## Syntax

 _expression_. **IncrementRotationY**( **_Increment_**)

 _expression_A variable that represents a  **ThreeDFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Increment|Required| **Single**|Specifies by how many degrees to rotate the shape around the y-axis. Can be a value from - 90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.|

## Remarks

Use the  **[RotationY](threedformat-rotationy-property-publisher.md)** property to set the absolute rotation of the shape around the y-axis.

You cannot adjust the rotation around the y-axis of the specified shape past the upper or lower limit for the  **RotationY** property (90 degrees to - 90 degrees). For example, if the **RotationY** property is initially set to 80 and you specify 40 for the **_Increment_** argument, the resulting rotation will be 90 (the upper limit for the **RotationY** property) instead of 120.

To change the rotation of a shape around the x-axis (horizontal), use the  **[IncrementRotationX](threedformat-incrementrotationx-method-publisher.md)** method. To change the rotation around the z-axis (extends outward from the plane of the publication), use the  **[IncrementRotation](shape-incrementrotation-method-publisher.md)** method.


## Example

This example tilts the first shape in the active publication 10 degrees to the right. The shape must be an extruded shape for you to see the effect of this code.


```vb
ActiveDocument.Pages(1).Shapes(1).ThreeD _ 
 .IncrementRotationY Increment:=-10
```


