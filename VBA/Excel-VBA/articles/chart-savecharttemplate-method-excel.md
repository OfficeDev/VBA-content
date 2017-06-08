---
title: Chart.SaveChartTemplate Method (Excel)
keywords: vbaxl10.chm149181
f1_keywords:
- vbaxl10.chm149181
ms.prod: excel
api_name:
- Excel.Chart.SaveChartTemplate
ms.assetid: d9e36023-b5bb-aaf4-5b34-9a22df468ced
ms.date: 06/08/2017
---


# Chart.SaveChartTemplate Method (Excel)

Saves a custom chart template to the list of available chart templates.


## Syntax

 _expression_ . **SaveChartTemplate**( **_Filename_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|The name of the chart template.|

## Remarks

By default, this method saves the active chart to the user's chart template directory. If a UNC or URL is specified, the chart will be saved to the specified location instead. 


## Example

This example adds a new chart template based on the active chart.


```vb
ActiveChart.SaveChartTemplate _ 
 Filename:="Presentation Chart" 

```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

