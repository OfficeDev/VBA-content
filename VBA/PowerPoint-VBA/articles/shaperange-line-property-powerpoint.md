---
title: ShapeRange.Line Property (PowerPoint)
keywords: vbapp10.chm548027
f1_keywords:
- vbapp10.chm548027
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.Line
ms.assetid: 27f648e0-d7eb-27a9-312b-8aa1784e7001
ms.date: 06/08/2017
---


# ShapeRange.Line Property (PowerPoint)

Returns a  **[LineFormat](lineformat-object-powerpoint.md)** object that contains line formatting properties for the specified shape. (For a line, the **LineFormat** object represents the line itself; for a shape with a border, the **LineFormat** object represents the border.) Read-only.


## Syntax

 _expression_. **Line**

 _expression_ A variable that represents a **ShapeRange** object.


### Return Value

LineFormat


## Example

This example adds a blue dashed line to  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(10, 10, 250, 250).Line

    .DashStyle = msoLineDashDotDot

    .ForeColor.RGB = RGB(50, 0, 128)

End With
```

This example adds a cross to the first slide and then sets its border to be 8 points thick and red.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeCross, 10, 10, 50, 70).Line

    .Weight = 8

    .ForeColor.RGB = RGB(255, 0, 0)

End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-powerpoint.md)

