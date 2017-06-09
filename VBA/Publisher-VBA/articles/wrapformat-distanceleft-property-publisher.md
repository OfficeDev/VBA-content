---
title: WrapFormat.DistanceLeft Property (Publisher)
keywords: vbapb10.chm786439
f1_keywords:
- vbapb10.chm786439
ms.prod: publisher
api_name:
- Publisher.WrapFormat.DistanceLeft
ms.assetid: 4d05ac86-f4a2-8a5e-bc7c-e303fee67e18
ms.date: 06/08/2017
---


# WrapFormat.DistanceLeft Property (Publisher)

When the  **[Type](wrapformat-type-property-publisher.md)** property of the **[WrapFormat](wrapformat-object-publisher.md)** object is set to **pbWrapTypeSquare**, returns or sets a  **Variant** that represents the distance (in points) between the document text and the left edge of the specified shape. Read/write.


## Syntax

 _expression_. **DistanceLeft**

 _expression_A variable that represents a  **WrapFormat** object.


## Example

This example adds an oval to the active document and specifies that the document text wrap around the left and right sides of the square that circumscribes the oval. The example sets a 0.1-inch margin between the document text and the top, bottom, left side, and right side of the square.


```vb
Sub AddNewShape() 
 Dim shpOval As Shape 
 
 Set shpOval = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, Left:=36, _ 
 Top:=36, Width:=100, Height:=35) 
 With shpOval.TextWrap 
 .Type = pbWrapTypeSquare 
 .Side = pbWrapSideBoth 
 .DistanceAuto = msoFalse 
 .DistanceTop = InchesToPoints(0.1) 
 .DistanceBottom = InchesToPoints(0.1) 
 .DistanceLeft = InchesToPoints(0.1) 
 .DistanceRight = InchesToPoints(0.1) 
 End With 
End Sub
```


