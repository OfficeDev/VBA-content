---
title: Chart.Location Method (Excel)
keywords: vbaxl10.chm149125
f1_keywords:
- vbaxl10.chm149125
ms.prod: excel
api_name:
- Excel.Chart.Location
ms.assetid: 3744f7f3-f7df-3ac2-48b7-b57ce3a8c812
ms.date: 06/08/2017
---


# Chart.Location Method (Excel)

Moves the chart to a new location.


## Syntax

 _expression_ . **Location**( **_Where_** , **_Name_** )

 _expression_ An expression that returns a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Where_|Required| **[XlChartLocation](xlchartlocation-enumeration-excel.md)**|Where to move the chart.|
| _Name_|Optional| **Variant**|Required if  _Where_ is **xlLocationAsObject** . The name of the sheet where the chart will be embedded if _Where_ is **xlLocationAsObject** or the name of the new sheet if _Where_ is **xlLocationAsNewSheet** .|

### Return Value

Chart


## Example

This example moves the embedded chart to a new chart sheet named "Monthly Sales." 


```vb
Worksheets(1).ChartObjects(1).Chart _ 
 .Location xlLocationAsNewSheet, "Monthly Sales"
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

