---
title: Shape.WrapFormat Property (Word)
keywords: vbawd10.chm161481007
f1_keywords:
- vbawd10.chm161481007
ms.prod: word
api_name:
- Word.Shape.WrapFormat
ms.assetid: 7ed0561f-7dcd-a9bd-3524-880237ebf1cb
ms.date: 06/08/2017
---


# Shape.WrapFormat Property (Word)

Returns a  **WrapFormat** object that contains the properties for wrapping text around the specified shape. Read-only.


## Syntax

 _expression_ . **WrapFormat**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Example

This example adds an oval to the active document and specifies that the document text wrap around the left and right sides of the square that circumscribes the oval. The example sets a 0.1-inch margin between the document text and the top, bottom, left side, and right side of the square.


```vb
Set myOval = _ 
 ActiveDocument.Shapes.AddShape(msoShapeOval, 36, 36, 90, 50) 
With myOval.WrapFormat 
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


[Shape Object](shape-object-word.md)

