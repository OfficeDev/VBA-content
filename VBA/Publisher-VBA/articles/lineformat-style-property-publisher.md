---
title: LineFormat.Style Property (Publisher)
keywords: vbapb10.chm3408144
f1_keywords:
- vbapb10.chm3408144
ms.prod: publisher
api_name:
- Publisher.LineFormat.Style
ms.assetid: 3826eb43-b90e-e24b-31d5-8d9eddd3ed4e
ms.date: 06/08/2017
---


# LineFormat.Style Property (Publisher)

Returns or sets an  **MsoLineStyle** constant that represents the style of line to apply to a shape or border. Read/write.


## Syntax

 _expression_. **Style**

 _expression_A variable that represents a  **LineFormat** object.


### Return Value

MsoLineStyle


## Remarks

The  **Style** property value can be one of the **MsoLineStyle** constants declared in the Microsoft Office type library and shown in the following table.



| **msoLineSingle**|
| **msoLineStyleMixed**|
| **msoLineThickBetweenThin**|
| **msoLineThickThin**|
| **msoLineThinThick**|
| **msoLineThinThin**|

## Example

This example adds a new shape and sets the line properties for the shape.


```vb
Sub SetLineStyle() 
 With ActiveDocument.Pages(1).Shapes.AddShape(msoShapeRectangle, _ 
 Left:=72, Top:=140, Width:=200, Height:=100) 
 .Rotation = 120 
 With .Line 
 .Weight = 5 
 .DashStyle = msoLineDashDotDot 
 .Style = msoLineThickBetweenThin 
 End With 
 End With 
End Sub
```


