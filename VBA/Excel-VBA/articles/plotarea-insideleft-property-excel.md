---
title: PlotArea.InsideLeft Property (Excel)
keywords: vbaxl10.chm618088
f1_keywords:
- vbaxl10.chm618088
ms.prod: excel
api_name:
- Excel.PlotArea.InsideLeft
ms.assetid: 193934e2-c3ca-c3cf-fb90-2dd45e17f9b8
ms.date: 06/08/2017
---


# PlotArea.InsideLeft Property (Excel)

Returns the distance from the chart edge to the inside left edge of the plot area, in points. Read-write  **Double** .


## Syntax

 _expression_ . **InsideLeft**

 _expression_ A variable that represents a **PlotArea** object.


## Remarks

The plot area used for this measurement doesn't include the axis labels. The  **Left** property for the plot area uses the bounding rectangle that includes the axis labels.


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

