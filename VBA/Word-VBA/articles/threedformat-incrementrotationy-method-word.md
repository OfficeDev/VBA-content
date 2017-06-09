---
title: ThreeDFormat.IncrementRotationY Method (Word)
keywords: vbawd10.chm164626443
f1_keywords:
- vbawd10.chm164626443
ms.prod: word
api_name:
- Word.ThreeDFormat.IncrementRotationY
ms.assetid: 924559dd-1e64-d5f4-c462-5d5fb931cd57
ms.date: 06/08/2017
---


# ThreeDFormat.IncrementRotationY Method (Word)

Changes the rotation of the specified shape around the y-axis by the specified number of degrees.


## Syntax

 _expression_ . **IncrementRotationY**( **_Increment_** )

 _expression_ Required. A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how much (in degrees) the rotation of the shape around the y-axis is to be changed. Can be a value from ? 90 through 90. A positive value tilts the shape to the left; a negative value tilts it to the right.|

## Remarks

Use the  **[RotationY](threedformat-rotationy-property-word.md)** property to set the absolute rotation of the shape around the y-axis.

To change the rotation of a shape around the x-axis, use the  **IncrementRotationX** method. To change the rotation around the z-axis, use the **IncrementRotation** method.

You cannot adjust the rotation around the y-axis of the specified shape past the upper or lower limit for the  **RotationY** property (90 degrees to ? 90 degrees). For example, if the **RotationY** property is initially set to 80 and you specify 40 for the Increment argument, the resulting rotation will be 90 (the upper limit for the **RotationY** property) instead of 120.


## Example

This example tilts the first shape on the active document 10 degrees to the right. The first shape must be an extruded shape for you to see the effect of this code.


```vb
ActiveDocument.Shapes(1).ThreeD.IncrementRotationY -10
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)

