---
title: ShapeRange.Group Method (Word)
keywords: vbawd10.chm162856979
f1_keywords:
- vbawd10.chm162856979
ms.prod: word
api_name:
- Word.ShapeRange.Group
ms.assetid: 2220e1d9-24aa-d2ba-f086-130e1139b346
ms.date: 06/08/2017
---


# ShapeRange.Group Method (Word)

Groups the shapes in the specified range, and returns the grouped shapes as a single  **Shape** object.


## Syntax

 _expression_ . **Group**

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

Because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the  **Shapes** collection and changes the index numbers of items that come after the affected items in the collection.


## Example

This example adds two shapes to myDocument, groups the two new shapes, sets the fill for the group, rotates the group, and sends the group to the back of the drawing layer.


```vb
Set myDocument = ActiveDocument 
With myDocument.Shapes 
 .AddShape(msoShapeCan, 50, 10, 100, 200).Name = "shpOne" 
 .AddShape(msoShapeCube, 150, 250, 100, 200).Name = "shpTwo" 
 With .Range(Array("shpOne", "shpTwo")).Group 
 .Fill.PresetTextured msoTextureBlueTissuePaper 
 .Rotation = 45 
 .ZOrder msoSendToBack 
 End With 
End With
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

