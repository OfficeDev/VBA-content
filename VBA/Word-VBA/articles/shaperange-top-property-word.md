---
title: ShapeRange.Top Property (Word)
keywords: vbawd10.chm162857083
f1_keywords:
- vbawd10.chm162857083
ms.prod: word
api_name:
- Word.ShapeRange.Top
ms.assetid: 2bfa4057-2b4e-6ea6-6d0f-3efd6eb3c63d
ms.date: 06/08/2017
---


# ShapeRange.Top Property (Word)

Returns or sets the vertical position of the specified shape or shape range in points. Read/write  **Single** .


## Syntax

 _expression_ . **Top**

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

The position of a shape is measured from the upper-left corner of the shape's bounding box to the shape's anchor. The  **RelativeVerticalPosition** property controls whether the shape's anchor is positioned alongside the line, the paragraph, the margin, or the edge of the page.

For a  **ShapeRange** object that contains more than one shape, the **Top** property sets the vertical position of each shape.


## Example

This example sets the vertical position of the first and second shapes in the active document to 1 inch from the top of the page.


```vb
With ActiveDocument.Shapes.Range(Array(1, 2)) 
 .RelativeVerticalPosition = wdRelativeVerticalPositionPage 
 .Top = InchesToPoints(1) 
End With
```


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

