---
title: ShapeRange.PickUp Method (Word)
keywords: vbawd10.chm162856980
f1_keywords:
- vbawd10.chm162856980
ms.prod: word
api_name:
- Word.ShapeRange.PickUp
ms.assetid: 6074168d-5cb2-2f86-fca4-c609dd2333f8
ms.date: 06/08/2017
---


# ShapeRange.PickUp Method (Word)

Copies the formatting of the specified shape.


## Syntax

 _expression_ . **PickUp**

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

Use the  **[Apply](shaperange-apply-method-word.md)** method to apply the copied formatting to another shape.


## Example

This example copies the formatting of shape one on  _myDocument_ and then applies the copied formatting to shape two.


```vb
Set myDocument = ActiveDocument 
With myDocument 
 .Shapes(1).PickUp 
 .Shapes(2).Apply 
End With
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

