---
title: WrapFormat.Side Property (Word)
keywords: vbawd10.chm163774565
f1_keywords:
- vbawd10.chm163774565
ms.prod: word
api_name:
- Word.WrapFormat.Side
ms.assetid: eb4aec92-a51b-df53-1643-bd5dca45c9b5
ms.date: 06/08/2017
---


# WrapFormat.Side Property (Word)

Returns or sets a value that indicates whether the document text should wrap on both sides of the specified shape, on either the left or right side only, or on the side of the shape that's farthest from the page margin.Read/write  **WdWrapSideType** .


## Syntax

 _expression_ . **Side**

 _expression_ Required. A variable that represents a **[WrapFormat](wrapformat-object-word.md)** object.


## Remarks

 If the text wraps on only one side of the shape, there is a text-free area between the other side of the shape and the page margin.


## Example

This example adds an oval to the active document and specifies that the document text wrap around the left and right sides of the square that circumscribes the oval. The example sets a 0.1-inch margin between the document text and the top, bottom, left side, and right side of the square.


```vb
Set myOval = ActiveDocument.Shapes.AddShape(msoShapeOval, _ 
 0, 0, 200, 50) 
With myEll.WrapFormat 
 .Type = wdWrapSquare 
 .Side = wdWrapBoth 
 .DistanceTop = InchesToPoints(0.1) 
 .DistanceBottom = InchesToPoints(0.1) 
 .DistanceLeft = InchesToPoints(0.1) 
 .DistanceRight = InchesToPoints(0.1) 
End With
```


## See also


#### Concepts


[WrapFormat Object](wrapformat-object-word.md)

