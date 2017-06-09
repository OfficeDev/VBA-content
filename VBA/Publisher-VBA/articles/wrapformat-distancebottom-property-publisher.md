---
title: WrapFormat.DistanceBottom Property (Publisher)
keywords: vbapb10.chm786440
f1_keywords:
- vbapb10.chm786440
ms.prod: publisher
api_name:
- Publisher.WrapFormat.DistanceBottom
ms.assetid: b8aa651a-4fe1-8b0b-f0d1-672e3f060809
ms.date: 06/08/2017
---


# WrapFormat.DistanceBottom Property (Publisher)

When the  **[Type](wrapformat-type-property-publisher.md)** property of the **[WrapFormat](wrapformat-object-publisher.md)** object is set to **pbWrapTypeSquare**, returns or sets a  **Variant** that represents the distance (in points) between the document text and the bottom edge of the specified shape. Read/write.


## Syntax

 _expression_. **DistanceBottom**

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


