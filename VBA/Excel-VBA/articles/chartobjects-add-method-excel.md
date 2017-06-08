---
title: ChartObjects.Add Method (Excel)
keywords: vbaxl10.chm497103
f1_keywords:
- vbaxl10.chm497103
ms.prod: excel
api_name:
- Excel.ChartObjects.Add
ms.assetid: 46f28b34-83a5-b3d9-c19b-a1dc8e05dff7
ms.date: 06/08/2017
---


# ChartObjects.Add Method (Excel)

Creates a new embedded chart.


## Syntax

 _expression_ . **Add**( **_Left_** , **_Top_** , **_Width_** , **_Height_** )

 _expression_ A variable that represents a **ChartObjects** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Left_|Required| **Double**|The initial coordinates of the new object (in points), relative to the upper-left corner of cell A1 on a worksheet or to the upper-left corner of a chart.|
| _Width_|Required| **Double**|The initial size of the new object, in points.|

### Return Value

A  **[ChartObject](chartobject-object-excel.md)** object that represents the new embedded chart.


## Example

This example creates a new embedded chart..


```vb
Set co = Sheets("Sheet1").ChartObjects.Add(50, 40, 200, 100) 
co.Chart.ChartWizard Source:=Worksheets("Sheet1").Range("A1:B2"), _ 
 Gallery:=xlColumn, Format:=6, PlotBy:=xlColumns, _ 
 CategoryLabels:=1, SeriesLabels:=0, HasLegend:=1
```


## See also


#### Concepts


[ChartObjects Object](chartobjects-object-excel.md)

