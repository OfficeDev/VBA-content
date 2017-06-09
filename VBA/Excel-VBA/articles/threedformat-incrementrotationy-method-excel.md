---
title: ThreeDFormat.IncrementRotationY Method (Excel)
keywords: vbaxl10.chm119021
f1_keywords:
- vbaxl10.chm119021
ms.prod: excel
api_name:
- Excel.ThreeDFormat.IncrementRotationY
ms.assetid: 56dde624-a56d-41f1-3192-f4c5c28e0a66
ms.date: 06/08/2017
---


# ThreeDFormat.IncrementRotationY Method (Excel)

Changes the rotation of the specified shape around the y-axis by the specified number of degrees. Use the  **[RotationY](threedformat-rotationy-property-excel.md)** property to set the absolute rotation of the shape around the y-axis.


## Syntax

 _expression_ . **IncrementRotationY**( **_Increment_** )

 _expression_ A variable that represents a **ThreeDFormat** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much (in degrees) the rotation of the shape around the y-axis is to be changed. Can be a value from ? 90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.|

## Remarks

To change the rotation of a shape around the x-axis, use the  **[IncrementRotationX](threedformat-incrementrotationx-method-excel.md)** method. To change the rotation around the z-axis, use the **[IncrementRotationZ](threedformat-incrementrotationz-method-excel.md)** method.

You cannot adjust the specified shape's rotation around the y-axis shape past the upper or lower limit for the  **RotationY** property (90 degrees to ? 90 degrees). For example, if the **RotationY** property is initially set to 80 and you specify 40 for the _Increment_ argument, the resulting rotation will be 90 (the upper limit for the **RotationY** property) instead of 120.


## Example

This example tilts shape one on  `myDocument` 10 degrees to the right. Shape one must be an extruded shape for you to see the effect of this code.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes(1).ThreeD.IncrementRotationY -10
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-excel.md)

