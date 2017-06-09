---
title: DataLabel.ShowValue Property (Excel)
keywords: vbaxl10.chm582101
f1_keywords:
- vbaxl10.chm582101
ms.prod: excel
api_name:
- Excel.DataLabel.ShowValue
ms.assetid: 83d4ead9-3539-d420-d4bd-2b474e174e10
ms.date: 06/08/2017
---


# DataLabel.ShowValue Property (Excel)

Returns or sets a  **Boolean** corresponding to a specified chart's data label values display behavior. **True** displays the values. **False** to hide. Read/write.


## Syntax

 _expression_ . **ShowValue**

 _expression_ A variable that represents a **DataLabel** object.


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


[DataLabel Object](datalabel-object-excel.md)

