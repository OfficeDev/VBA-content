---
title: DataLabel.ShowPercentage Property (Excel)
keywords: vbaxl10.chm582102
f1_keywords:
- vbaxl10.chm582102
ms.prod: excel
api_name:
- Excel.DataLabel.ShowPercentage
ms.assetid: 9d084502-545d-7a9a-1b6d-e12d4e2b34e6
ms.date: 06/08/2017
---


# DataLabel.ShowPercentage Property (Excel)

 **True** to display the percentage value for the data labels on a chart. **False** to hide. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowPercentage**

 _expression_ A variable that represents a **DataLabel** object.


## Remarks

The chart must first be active before you can access the data labels programmatically or a run-time error will occur.


## Example

This example enables the percentage value to be shown for the data labels of the first series on the first chart. This example assumes a chart exists on the active worksheet.


```vb
Sub UsePercentage() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowPercentage = True 
 
End Sub
```


## See also


#### Concepts


[DataLabel Object](datalabel-object-excel.md)

