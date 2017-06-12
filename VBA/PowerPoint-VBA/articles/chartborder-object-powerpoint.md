---
title: ChartBorder Object (PowerPoint)
keywords: vbapp10.chm685000
f1_keywords:
- vbapp10.chm685000
ms.prod: powerpoint
api_name:
- PowerPoint.ChartBorder
ms.assetid: fd651a9a-4068-9a9b-f605-9228da5e6183
ms.date: 06/08/2017
---


# ChartBorder Object (PowerPoint)

Represents the border of an object.


## Remarks

Most bordered objects have a border that is treated as a single entity, regardless of how many sides it has. The entire border must be returned as a unit. To return a  **Border** object, use the **Border** property for the particular bordered object (for example, the **[Border](trendline-border-property-powerpoint.md)** property of a **[TrendLine](trendline-object-powerpoint.md)** object).


## Example




 **Note**  Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

 The following example changes the type and line style of a trendline on the active chart.




```vb
With ActiveDocument.InlineShapes(1).Chart.SeriesCollection(1).Trendlines(1)

    .Type = xlLinear

    .Border.LineStyle = xlDash

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

