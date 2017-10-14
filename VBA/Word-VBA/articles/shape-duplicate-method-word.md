---
title: Shape.Duplicate Method (Word)
keywords: vbawd10.chm161480716
f1_keywords:
- vbawd10.chm161480716
ms.prod: word
api_name:
- Word.Shape.Duplicate
ms.assetid: 8734d0f7-62fa-01eb-7aa8-9f959d36d195
ms.date: 06/08/2017
---


# Shape.Duplicate Method (Word)

Creates a duplicate of the specified  **Shape** object, adds the new shape to the **Shapes** collection at a standard offset from the original shapes, and then returns the new **Shape** object.


## Syntax

 _expression_ . **Duplicate**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


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


[Shape Object](shape-object-word.md)

