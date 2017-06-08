---
title: ShapeRange.Line Property (Publisher)
keywords: vbapb10.chm2293826
f1_keywords:
- vbapb10.chm2293826
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Line
ms.assetid: e9a6e8a0-f57a-63af-3040-5c43f8aba423
ms.date: 06/08/2017
---


# ShapeRange.Line Property (Publisher)

Returns a  **[LineFormat](lineformat-object-publisher.md)** object that contains line formatting properties for the specified shape. (For a line, the  **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.).


## Syntax

 _expression_. **Line**

 _expression_A variable that represents a  **ShapeRange** object.


## Example

This example adds a blue dashed line to the active publication.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=10, BeginY:=10, _ 
 EndX:=250, EndY:=250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With
```

This example adds a cross to the first page and then sets its border to be 8 points thick and red.




```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeCross, _ 
 Left:=10, Top:=10, Width:=50, Height:=70).Line 
 .Weight = 8 
 .ForeColor.RGB = RGB(255, 0, 0) 
End With
```


