---
title: ShapeRange.Duplicate Method (Word)
keywords: vbawd10.chm162856974
f1_keywords:
- vbawd10.chm162856974
ms.prod: word
api_name:
- Word.ShapeRange.Duplicate
ms.assetid: 98efa3b3-3405-152a-b629-d4bb654c8029
ms.date: 06/08/2017
---


# ShapeRange.Duplicate Method (Word)

Creates a duplicate of the specified  **ShapeRange** object, adds the new range of shapes to the **Shapes** collection at a standard offset from the original shapes, and then returns a **Shape** object.


## Syntax

 _expression_ . **Duplicate**

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Example

This example creates a duplicate of shape one on the active document and then changes the fill for the new shape.


```vb
Set newShape = ActiveDocument.Shapes(1).Duplicate 
With newShape 
 .Fill.PresetGradient msoGradientVertical, 1, msoGradientGold 
End With
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

