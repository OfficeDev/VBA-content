---
title: Shape.Top Property (Word)
keywords: vbawd10.chm161480827
f1_keywords:
- vbawd10.chm161480827
ms.prod: word
api_name:
- Word.Shape.Top
ms.assetid: 59000c91-58c0-7849-2945-48b9fb8d8b17
ms.date: 06/08/2017
---


# Shape.Top Property (Word)

Returns or sets the vertical position of the specified shape or shape range in points. Read/write  **Single** .


## Syntax

 _expression_ . **Top**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

The position of a shape is measured from the upper-left corner of the shape's bounding box to the shape's anchor. The  **RelativeVerticalPosition** property controls whether the shape's anchor is positioned alongside the line, the paragraph, the margin, or the edge of the page.

For a  **ShapeRange** object that contains more than one shape, the **Top** property sets the vertical position of each shape.


## Example

This example sets the vertical position of the first shape in the active document to 1 inch from the top of the page.


```vb
With ActiveDocument.Shapes(1) 
 .RelativeVerticalPosition = wdRelativeVerticalPositionPage 
 .Top = InchesToPoints(1) 
End With
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

