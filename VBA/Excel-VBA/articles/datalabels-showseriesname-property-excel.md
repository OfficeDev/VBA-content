---
title: DataLabels.ShowSeriesName Property (Excel)
keywords: vbaxl10.chm584099
f1_keywords:
- vbaxl10.chm584099
ms.prod: excel
api_name:
- Excel.DataLabels.ShowSeriesName
ms.assetid: 19fcea65-a796-3c02-f162-33b5cb03aad3
ms.date: 06/08/2017
---


# DataLabels.ShowSeriesName Property (Excel)

Returns or sets a  **Boolean** to indicate the series name display behavior for the data labels on a chart. **True** to show the series name. **False** to hide. Read/write.


## Syntax

 _expression_ . **ShowSeriesName**

 _expression_ A variable that represents a **DataLabels** object.


## Remarks

The chart must first be active before you can access the data labels programmatically or a run-time error will occur.


## Example

This example enables the series name to be shown for the data labels of the first series on the first chart. This example assumes a chart exists on the active worksheet.


```vb
Sub UseSeriesName() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowSeriesName = True 
 
End Sub
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-excel.md)

