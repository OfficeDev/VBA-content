---
title: ChartSeries.FillColor Property (Access)
keywords: vbaac10.chm14785
f1_keywords:
- vbaac10.chm14785
ms.prod: access
api_name:
- Access.ChartSeries.FillColor
ms.date: 05/02/2018
---


# ChartSeries.FillColor Property (Access)

Returns or sets the fill color of a series visualization. Read/write **String** .

You can use a **[system color constant](../../language-reference-vba/articles/system-color-constants.md)** or the RGB function as shown in the example below.


## Syntax

 _expression_ . **FillColor**

 _expression_ A variable that represents a **ChartSeries** object.


## Example

The following example sets the border and fill color of the first series in a collection.

```vb
With myChart.ChartSeriesCollection.Item(0)
 .BorderColor = RGB(0, 0, 0)
 .FillColor = RGB(210, 250, 210)
End With
```

## See also


#### Concepts


[ChartSeries Object](chartseries-object-access.md)

[Chart Object](chart-object-access.md)