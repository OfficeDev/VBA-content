---
title: Shape.Left Property (Word)
keywords: vbawd10.chm161480815
f1_keywords:
- vbawd10.chm161480815
ms.prod: word
api_name:
- Word.Shape.Left
ms.assetid: 9c14ebc2-70fa-027b-63f0-6e44e60f8eed
ms.date: 06/08/2017
---


# Shape.Left Property (Word)

Returns or sets a  **Single** that represents the horizontal position, measured in points, of the specified shape or shape range. Can also be any valid **[WdShapePosition](wdshapeposition-enumeration-word.md)** constant. Read/write.


## Syntax

 _expression_ . **Left**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

The position of a shape is measured from the upper-left corner of the shape's bounding box to the shape's anchor. The  **RelativeHorizontalPosition** property controls whether the anchor is positioned alongside a character, column, margin, or the edge of the page.


## Example

This example sets the horizontal position of the first shape in the active document to 1 inch from the left edge of the page.


```vb
With ActiveDocument.Shapes(1) 
 .RelativeHorizontalPosition = _ 
 wdRelativeHorizontalPositionPage 
 .Left = InchesToPoints(1) 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

