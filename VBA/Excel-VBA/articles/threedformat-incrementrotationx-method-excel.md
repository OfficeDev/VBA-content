---
title: ThreeDFormat.IncrementRotationX Method (Excel)
keywords: vbaxl10.chm119020
f1_keywords:
- vbaxl10.chm119020
ms.prod: excel
api_name:
- Excel.ThreeDFormat.IncrementRotationX
ms.assetid: 599f96ca-6a8a-3b9e-5d76-c053f3236522
ms.date: 06/08/2017
---


# ThreeDFormat.IncrementRotationX Method (Excel)

Changes the rotation of the specified shape around the x-axis by the specified number of degrees. Use the  **[RotationX](threedformat-rotationx-property-excel.md)** property to set the absolute rotation of the shape around the x-axis.


## Syntax

 _expression_ . **IncrementRotationX**( **_Increment_** )

 _expression_ A variable that represents a **ThreeDFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much (in degrees) the rotation of the shape around the x-axis is to be changed. Can be a value from ? 90 through 90. A positive value tilts the shape up; a negative value tilts it down.|

## Remarks

You cannot adjust the specified shape's rotation around the x-axis past the upper or lower limit for the  **RotationX** property (90 degrees to ? 90 degrees). For example, if the **RotationX** property is initially set to 80 and you specify 40 for the _Increment_ argument, the resulting rotation will be 90 (the upper limit for the **RotationX** property) instead of 120.

To change the rotation of a shape around the y-axis, use the  **[IncrementRotationY](threedformat-incrementrotationy-method-excel.md)** method. To change the rotation around the z-axis, use the **[IncrementRotationZ](threedformat-incrementrotationz-method-excel.md)** method.


## Example

This example tilts shape one on  `myDocument` up 10 degrees. Shape one must be an extruded shape for you to see the effect of this code.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).ThreeD.IncrementRotationX 10
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)

