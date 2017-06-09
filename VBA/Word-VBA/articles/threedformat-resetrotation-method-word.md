---
title: ThreeDFormat.ResetRotation Method (Word)
keywords: vbawd10.chm164626444
f1_keywords:
- vbawd10.chm164626444
ms.prod: word
api_name:
- Word.ThreeDFormat.ResetRotation
ms.assetid: ab8b1bb6-2d39-2488-5db9-8405f8494014
ms.date: 06/08/2017
---


# ThreeDFormat.ResetRotation Method (Word)

Resets the extrusion rotation around the x-axis and the y-axis to 0 (zero) so that the front of the extrusion faces forward.


## Syntax

 _expression_ . **ResetRotation**

 _expression_ Required. A variable that represents a **[ThreeDFormat](threedformat-object-word.md)** object.


## Remarks

To set the extrusion rotation around the x-axis and the y-axis to anything other than 0 (zero), use the  **RotationX** and **RotationY** properties of the **ThreeDFormat** object. To set the extrusion rotation around the z-axis, use the **Rotation** property of the **Shape** object that represents the extruded shape.


 **Note**  This method does not reset the rotation around the z-axis.


## Example

This example resets the rotation around the x-axis and the y-axis to 0 (zero) for the extrusion of the first shape on the active document.


```vb
ActiveDocument.Shapes(1).ThreeD.ResetRotation
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-word.md)

