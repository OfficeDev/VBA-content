---
title: ChartSeries.ComboChartType Property (Access)
keywords: vbaac10.chm14781
f1_keywords:
- vbaac10.chm14781
ms.prod: access
api_name:
- Access.ChartSeries.ComboChartType
ms.date: 05/02/2018
---


# ChartSeries.ComboChartType Property (Access)

Returns or sets the chart type for the specified series. Read/write **[AcChartType](accharttype-enumeration-access.md)** .

This setting is only applicable when the **[ChartType](chart-charttype-property-access.md)** of the parent **[Chart](chart-object-access.md)** is set to **acChartCombo**.


## Syntax

 _expression_ . **ComboChartType**

 _expression_ A variable that represents a **ChartSeries** object.


## Example

This example checks if a chart is a combo chart, and if so, sets the **ComboChartType** of the first series to **acChartLine**.

```vb
With myChart
 If .ChartType = acChartCombo Then
  .ChartSeriesCollection.Item(0).ComboChartType = acChartLine
 End If
End With
```

## See also


#### Concepts


[AcChartType Enumeration](accharttype-enumeration-access.md)

[ChartSeries Object](chartseries-object-access.md)

[Chart Object](chart-object-access.md)