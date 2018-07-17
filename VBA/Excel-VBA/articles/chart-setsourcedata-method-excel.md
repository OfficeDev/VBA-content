---
title: Chart.SetSourceData Method (Excel)
keywords: vbaxl10.chm149162
f1_keywords:
- vbaxl10.chm149162
ms.prod: excel
api_name:
- Excel.Chart.SetSourceData
ms.assetid: fc41cc05-087a-f53c-2f54-fd6307de51d6
ms.date: 06/08/2017
---


# Chart.SetSourceData Method (Excel)

Sets the source data range for the chart.


## Syntax

 _expression_ . **SetSourceData**( **_Source_** , **_PlotBy_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **Range**|The range that contains the source data.|
| _PlotBy_|Optional| **Variant**|Specifies the way the data is to be plotted. Can be either of the following  **[XlRowCol](xlrowcol-enumeration-excel.md)** constants: **xlColumns** or **xlRows** .|

## Example

This example sets the source data range for chart one.


```vb
Charts(1).SetSourceData Source:=Sheets(1).Range("a1:a10"), _ 
 PlotBy:=xlColumns
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

