---
title: PlotArea.InsideHeight Property (Excel)
keywords: vbaxl10.chm618091
f1_keywords:
- vbaxl10.chm618091
ms.prod: excel
api_name:
- Excel.PlotArea.InsideHeight
ms.assetid: a9b2e591-afc5-331e-86b5-bbeb47696c3d
ms.date: 06/08/2017
---


# PlotArea.InsideHeight Property (Excel)

Returns the inside height of the plot area, in points. Read-write  **Double** .


## Syntax

 _expression_ . **InsideHeight**

 _expression_ A variable that represents a **PlotArea** object.


## Remarks

The plot area used for this measurement doesn't include the axis labels. The  **Height** property for the plot area uses the bounding rectangle that includes the axis labels.


## Example

This example draws a dotted rectangle around the inside of the plot area in Chart1.


```vb
With Charts("chart1") 
 Set pa = .PlotArea 
 With .Shapes.AddShape(msoShapeRectangle, _ 
 pa.InsideLeft, pa.InsideTop, _ 
 pa.InsideWidth, pa.InsideHeight) 
 .Fill.Transparency = 1 
 .Line.DashStyle = msoLineDashDot 
 End With 
End With
```


## See also


#### Concepts


[PlotArea Object](plotarea-object-excel.md)

