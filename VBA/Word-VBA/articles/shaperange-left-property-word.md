---
title: ShapeRange.Left Property (Word)
keywords: vbawd10.chm162857071
f1_keywords:
- vbawd10.chm162857071
ms.prod: word
api_name:
- Word.ShapeRange.Left
ms.assetid: 18ef49c4-d3b9-d65a-c992-9939479b607d
ms.date: 06/08/2017
---


# ShapeRange.Left Property (Word)

Returns or sets a  **Single** that represents the horizontal position, measured in points, of the specified range of shapes. Can also be any valid **[WdShapePosition](wdshapeposition-enumeration-word.md)** constant. Read/write.


## Syntax

 _expression_ . **Left**

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

The position of a shape is measured from the upper-left corner of the shape's bounding box to the shape's anchor. The  **RelativeHorizontalPosition** property controls whether the anchor is positioned alongside a character, column, margin, or the edge of the page.

For a  **ShapeRange** object that contains more than one shape, the **Left** property sets the horizontal position of each shape.


## Example

This example sets the horizontal position of the first and second shapes in the active document to 1 inch from the left edge of the column.


```vb
With ActiveDocument.Shapes.Range(Array(1, 2)) 
 .RelativeHorizontalPosition = _ 
 wdRelativeHorizontalPositionColumn 
 .Left = InchesToPoints(1) 
End With
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

