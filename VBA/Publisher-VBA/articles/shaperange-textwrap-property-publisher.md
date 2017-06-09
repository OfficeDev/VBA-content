---
title: ShapeRange.TextWrap Property (Publisher)
keywords: vbapb10.chm2294020
f1_keywords:
- vbapb10.chm2294020
ms.prod: publisher
api_name:
- Publisher.ShapeRange.TextWrap
ms.assetid: 40fbc7aa-0a1b-7835-76bf-1815d7ccffc4
ms.date: 06/08/2017
---


# ShapeRange.TextWrap Property (Publisher)

Returns a  **[WrapFormat](wrapformat-object-publisher.md)** object that represents the properties for wrapping text around a shape or shape range.


## Syntax

 _expression_. **TextWrap**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

The following example adds an oval to the active publication and specifies that publication text wrap around the left and right sides of the square that circumscribes the oval. There will be a 0.1-inch margin between the publication text and the top, bottom, left side, and right side of the square.


```vb
Sub SetTextWrapFormatProperties() 
 Dim shpOval As Shape 
 
 Set shpOval = ActiveDocument.Pages(1).Shapes.AddShape(Type:=msoShapeOval, _ 
 Left:=36, Top:=36, Width:=100, Height:=35) 
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


