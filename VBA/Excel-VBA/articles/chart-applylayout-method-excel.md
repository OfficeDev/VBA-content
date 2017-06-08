---
title: Chart.ApplyLayout Method (Excel)
keywords: vbaxl10.chm149174
f1_keywords:
- vbaxl10.chm149174
ms.prod: excel
api_name:
- Excel.Chart.ApplyLayout
ms.assetid: 0e07936d-c179-9b38-a6d4-1d71d1c5af3b
ms.date: 06/08/2017
---


# Chart.ApplyLayout Method (Excel)

Applies the layouts shown in the ribbon.


## Syntax

 _expression_ . **ApplyLayout**( **_Layout_** , **_ChartType_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Layout_|Required| **Long**|Specifies the type of layout. The type of layout is denoted by a number from 1 to 10.|
| _ChartType_|Optional| **[XlChartType](xlcharttype-enumeration-excel.md)**|The type of chart.|

## Remarks

When you use a layout on the current chart type, a number from 1 to 10 is applied to the chart type. You can also apply the layout of one chart type on another chart type. For example, you can apply the layouts that are available from a line chart to a column chart. The layout only adds chart elements that are available for that particular chart type.


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

