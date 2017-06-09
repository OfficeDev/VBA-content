---
title: Selection.ChildShapeRange Property (Publisher)
keywords: vbapb10.chm851973
f1_keywords:
- vbapb10.chm851973
ms.prod: publisher
api_name:
- Publisher.Selection.ChildShapeRange
ms.assetid: 8ef96e85-2f25-7b3a-4465-7e22fdbbaa9a
ms.date: 06/08/2017
---


# Selection.ChildShapeRange Property (Publisher)

Returns a  **[ShapeRange](shaperange-object-publisher.md)** object representing the child shapes of a selection.


## Syntax

 _expression_. **ChildShapeRange**

 _expression_A variable that represents a  **Selection** object.


### Return Value

ShapeRange


## Example

This example creates a new page in the active publication, populates the page with shapes, and selects and groups the shapes. Then, after canceling the selection of two of the group shapes, it changes the AutoShape type for one of the shapes.


```vb
Sub ChangeFillToChildShape() 
 
 With ThisDocument.Pages(1) 
 With .Shapes 
 .AddShape msoShape4pointStar, 10, 10, 175, 175 
 .AddShape msoShapeOval, 100, 100, 175, 75 
 .AddShape msoShapeOval, 150, 150, 175, 75 
 .Range.Group 
 .SelectAll 
 End With 
 .Shapes(1).GroupItems(1).Select msoFalse 
 .Shapes(1).GroupItems(2).Select msoFalse 
 End With 
 
 Selection.ChildShapeRange(3).AutoShapeType = msoShapeDiamond 
 
End Sub
```


