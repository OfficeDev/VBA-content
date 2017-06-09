---
title: Chart.ApplyLayout Method (PowerPoint)
keywords: vbapp10.chm68004
f1_keywords:
- vbapp10.chm68004
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.ApplyLayout
ms.assetid: 94368674-4ecb-a3aa-da91-8e92980d3cc8
ms.date: 06/08/2017
---


# Chart.ApplyLayout Method (PowerPoint)

Applies the layouts shown in the Ribbon.


## Syntax

 _expression_. **ApplyLayout**( **_Layout_**, **_ChartType_** )

 _expression_ A variable that represents a **[Chart](chart-object-powerpoint.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Layout_|Required|**Long**|The type of layout. The type of layout is denoted by a number from 1 to 10.|
| _ChartType_|Optional|**Variant**|An  **[XlChartType](http://msdn.microsoft.com/library/bba4ee89-ee91-f55a-d2e0-59a73e5bfabe%28Office.15%29.aspx)** constant that represents the type of chart.|

## Remarks

When you use a layout on the current chart type, a number from 1 to 10 is applied to the chart type. You can also apply the layout of one chart type on another chart type. For example, you can apply the layouts that are available from a line chart to a column chart. The layout adds only chart elements that are available for that particular chart type.


## See also


#### Concepts


[Chart Object](chart-object-powerpoint.md)

