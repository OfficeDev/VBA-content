---
title: WrapFormat.DistanceRight Property (Word)
keywords: vbawd10.chm163774569
f1_keywords:
- vbawd10.chm163774569
ms.prod: word
api_name:
- Word.WrapFormat.DistanceRight
ms.assetid: 37aae328-abbc-e755-a402-5649a2f81e18
ms.date: 06/08/2017
---


# WrapFormat.DistanceRight Property (Word)

 Returns or sets the distance (in points) between the document text and the right edge of the text-free area surrounding the specified shape. Read/write **Single** .


## Syntax

 _expression_ . **DistanceRight**

 _expression_ A variable that represents a **[WrapFormat](wrapformat-object-word.md)** object.


## Remarks

The size and shape of the specified shape, together with the values of the  **Type** and **Side** properties of the **WrapFormat** object, determine the size and shape of this text-free area.


## Example

This example sets text to wrap around the first table in the active document and sets the distance for wrapped text to 20 points on all sides of the table.


```vb
With ActiveDocument.Tables(1).Rows 
 .WrapAroundText = True 
 .DistanceLeft = 20 
 .DistanceRight = 20 
 .DistanceTop = 20 
 .DistanceBottom = 20 
End With
```

This example adds an oval to the active document and specifies that the document text wrap around the left and right sides of the square that circumscribes the oval. The example sets a 0.1-inch margin between the document text and the top, bottom, left side, and right side of the square.




```vb
Dim shapeOval As Shape 
 
Set shapeOval = ActiveDocument.Shapes.AddShape(msoShapeOval, _ 
 0, 0, 200, 50) 
With shapeOval.WrapFormat 
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

