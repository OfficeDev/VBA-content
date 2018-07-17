---
title: PlotArea.InsideWidth Property (Excel)
keywords: vbaxl10.chm618090
f1_keywords:
- vbaxl10.chm618090
ms.prod: excel
api_name:
- Excel.PlotArea.InsideWidth
ms.assetid: 2ebad523-2f25-28c1-5d6e-56517e2690b7
ms.date: 06/08/2017
---


# PlotArea.InsideWidth Property (Excel)

Returns the inside width of the plot area, in points. Read-write  **Double** .


## Syntax

 _expression_ . **InsideWidth**

 _expression_ A variable that represents a **PlotArea** object.


## Remarks

The plot area used for this measurement doesn't include the axis labels. The  **Width** property for the plot area uses the bounding rectangle that includes the axis labels.


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

