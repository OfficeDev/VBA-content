---
title: Shapes.AddChart2 Method (Excel)
keywords: vbaxl10.chm638096
f1_keywords:
- vbaxl10.chm638096
ms.prod: excel
ms.assetid: 2d4569df-2f77-40d5-5f81-859e13e0abb7
ms.date: 06/08/2017
---


# Shapes.AddChart2 Method (Excel)

Adds a chart to the document. Returns a  **Shape** object that represents a chart and adds it to the specified collection.


## Syntax

 _expression_ . **AddChart2**_(Style,_ _XlChartType,_ _Left,_ _Top,_ _Width,_ _Height,_ _NewLayout)_

 _expression_ A variable that represents a **Shapes** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Style_|Optional|VARIANT|The chart style. Use "-1" to get the default style for the chart type specified in  **XlChartType**. |
| _XlChartType_|Optional|VARIANT|The type of chart.|
| _Left_|Optional|VARIANT|The position, in points, of the left edge of the chart, relative to the anchor.|
| _Top_|Optional|VARIANT|The position, in points, of the top edge of the chart, relative to the anchor.|
| _Width_|Optional|VARIANT|The width, in points, of the chart.|
| _Height_|Optional|VARIANT|The height, in points, of the chart.|
| _NewLayout_|Optional|VARIANT|If  **NewLayout** is **True** , the chart is inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|

### Return value

 **SHAPE**


## See also


#### Concepts


[Shapes Object](shapes-object-excel.md)

