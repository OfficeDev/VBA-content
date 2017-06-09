---
title: Worksheet.ChartObjects Method (Excel)
keywords: vbaxl10.chm175081
f1_keywords:
- vbaxl10.chm175081
ms.prod: excel
api_name:
- Excel.Worksheet.ChartObjects
ms.assetid: 234cab0e-a8a2-2174-8881-39b5fb37c743
ms.date: 06/08/2017
---


# Worksheet.ChartObjects Method (Excel)

Returns an object that represents either a single embedded chart (a  **[ChartObject](chartobject-object-excel.md)** object) or a collection of all the embedded charts (a **[ChartObjects](chartobjects-object-excel.md)** object) on the sheet.


## Syntax

 _expression_ . **ChartObjects**( **_Index_** )

 _expression_ A variable that represents a **Worksheet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The name or number of the chart. This argument can be an array, to specify more than one chart.|

### Return Value

Object


## Remarks

This method isn't equivalent to the  **[Charts](workbook-charts-property-excel.md)** property. This method returns embedded charts; the **Charts** property returns chart sheets. Use the **[Chart](chartobject-chart-property-excel.md)** property to return the **Chart** object for an embedded chart.


## Example

This example adds a title to embedded chart one on Sheet1.


```vb
With Worksheets("Sheet1").ChartObjects(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "1995 Rainfall Totals by Month" 
End With
```

This example creates a new series in embedded chart one on Sheet1. The data source for the new series is the range B1:B10 on Sheet1.




```vb
Worksheets("Sheet1").ChartObjects(1).Activate 
ActiveChart.SeriesCollection.Add _ 
 source:=Worksheets("Sheet1").Range("B1:B10")
```

This example clears the formatting of embedded chart one on Sheet1.




```vb
Worksheets("Sheet1").ChartObjects(1).Chart.ChartArea.ClearFormats
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

