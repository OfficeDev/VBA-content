---
title: ShapeRange.IncrementRotation Method (Word)
keywords: vbawd10.chm162856977
f1_keywords:
- vbawd10.chm162856977
ms.prod: word
api_name:
- Word.ShapeRange.IncrementRotation
ms.assetid: bf77da5d-7043-fa09-1b78-410d2514cde1
ms.date: 06/08/2017
---


# ShapeRange.IncrementRotation Method (Word)

Changes the rotation of the specified shape around the z-axis by the specified number of degrees. .


## Syntax

 _expression_ . **IncrementRotation**( **_Increment_** )

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Increment_|Required| **Single**|Specifies how far the shape is to be rotated horizontally, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise.|

## Remarks

Use the  **Rotation** property to set the absolute rotation of the shape. To rotate a three-dimensional shape around the x-axis or the y-axis, use the **[IncrementRotationX](threedformat-incrementrotationx-method-word.md)** or **[IncrementRotationY](threedformat-incrementrotationy-method-word.md)** method of the **[ThreeDFormat](threedformat-object-word.md)** .


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

