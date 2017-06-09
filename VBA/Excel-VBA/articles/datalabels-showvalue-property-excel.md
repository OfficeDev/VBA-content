---
title: DataLabels.ShowValue Property (Excel)
keywords: vbaxl10.chm584101
f1_keywords:
- vbaxl10.chm584101
ms.prod: excel
api_name:
- Excel.DataLabels.ShowValue
ms.assetid: e078ade5-d3d0-5b5c-8b40-667e69e38793
ms.date: 06/08/2017
---


# DataLabels.ShowValue Property (Excel)

Returns or sets a  **Boolean** corresponding to a specified chart's data label values display behavior. **True** displays the values. **False** to hide. Read/write.


## Syntax

 _expression_ . **ShowValue**

 _expression_ A variable that represents a **DataLabels** object.


## Remarks

The specified chart must first be active before you can access the data labels programmatically or a run-time error will occur.


## Example

This example enables the value to be shown for the data labels of the first series, on the first chart. This example assumes a chart exists on the active worksheet.


```vb
Sub UseValue() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowValue = True 
 
End Sub
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-excel.md)

