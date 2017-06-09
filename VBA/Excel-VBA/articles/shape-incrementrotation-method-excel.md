---
title: Shape.IncrementRotation Method (Excel)
keywords: vbaxl10.chm636079
f1_keywords:
- vbaxl10.chm636079
ms.prod: excel
api_name:
- Excel.Shape.IncrementRotation
ms.assetid: 3b9f1ae0-da53-b0e7-6569-dc3cd4595b12
ms.date: 06/08/2017
---


# Shape.IncrementRotation Method (Excel)

Changes the rotation of the specified shape around the z-axis by the specified number of degrees. Use the  **[Rotation](shape-rotation-property-excel.md)** property to set the absolute rotation of the shape.


## Syntax

 _expression_ . **IncrementRotation**( **_Increment_** )

 _expression_ A variable that represents a **Shape** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shape is to be rotated horizontally, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.|

## Remarks

To rotate a three-dimensional shape around the x-axis or the y-axis, use the  **[IncrementRotationX](threedformat-incrementrotationx-method-excel.md)** method or the **[IncrementRotationY](threedformat-incrementrotationy-method-excel.md)** method.


## Example

This example duplicates shape one on  `myDocument`, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(1).Duplicate 
 .Fill.PresetTextured msoTextureGranite 
 .IncrementLeft 70 
 .IncrementTop -50 
 .IncrementRotation 30 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-excel.md)

