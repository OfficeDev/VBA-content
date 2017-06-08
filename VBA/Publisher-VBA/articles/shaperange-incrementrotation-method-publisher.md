---
title: ShapeRange.IncrementRotation Method (Publisher)
keywords: vbapb10.chm2293793
f1_keywords:
- vbapb10.chm2293793
ms.prod: publisher
api_name:
- Publisher.ShapeRange.IncrementRotation
ms.assetid: c58cdc12-948a-d6f8-2ddd-113008c7201b
ms.date: 06/08/2017
---


# ShapeRange.IncrementRotation Method (Publisher)

Changes the rotation of the specified shape around the z-axis (extends outward from the plane of the publication) by the specified number of degrees.


## Syntax

 _expression_. **IncrementRotation**( **_Increment_**)

 _expression_A variable that represents a  **ShapeRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Increment|Required| **Single**|Specifies how far the shape is to be rotated around the z-axis, in degrees. A positive value rotates the shape clockwise; a negative value rotates it counterclockwise. Valid values are between - 360 and 360.|

## Remarks

Use the  **[Rotation](shaperange-rotation-property-publisher.md)** property to set the absolute rotation of the shape.

To rotate a three-dimensional shape around the x-axis (horizontal) or the y-axis (vertical), use the  **[IncrementRotationX](threedformat-incrementrotationx-method-publisher.md)** method or the  **[IncrementRotationY](threedformat-incrementrotationy-method-publisher.md)** method, respectively.


## Example

This example duplicates the first shape on the active publication, sets the fill for the duplicate, moves it 70 points to the right and 50 points up, and rotates it 30 degrees clockwise.


```vb
With ActiveDocument.Pages(1).Shapes(1).Duplicate 
 .Fill.PresetTextured PresetTexture:=msoTextureGranite 
 .IncrementLeft Increment:=70 
 .IncrementTop Increment:=-50 
 .IncrementRotation Increment:=30 
End With
```


