---
title: LineFormat.Weight Property (PowerPoint)
keywords: vbapp10.chm553015
f1_keywords:
- vbapp10.chm553015
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.Weight
ms.assetid: 5141d66f-4706-060d-fb4c-f244f9ac6437
ms.date: 06/08/2017
---


# LineFormat.Weight Property (PowerPoint)

Returns or sets the thickness of the specified line, in points. Read/write.


## Syntax

 _expression_. **Weight**

 _expression_ A variable that represents a **LineFormat** object.


### Return Value

Single


## Example

This example adds a green dashed line two points thick to  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(10, 10, 250, 250).Line

    .DashStyle = msoLineDashDotDot

    .ForeColor.RGB = RGB(0, 255, 255)

    .Weight = 2

End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-powerpoint.md)

