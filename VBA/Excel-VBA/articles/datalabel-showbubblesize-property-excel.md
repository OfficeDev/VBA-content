---
title: DataLabel.ShowBubbleSize Property (Excel)
keywords: vbaxl10.chm582103
f1_keywords:
- vbaxl10.chm582103
ms.prod: excel
api_name:
- Excel.DataLabel.ShowBubbleSize
ms.assetid: e2768811-a45a-40cb-5327-64e3985095f0
ms.date: 06/08/2017
---


# DataLabel.ShowBubbleSize Property (Excel)

 **True** to show the bubble size for the data labels on a chart. **False** to hide. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowBubbleSize**

 _expression_ An expression that returns a **DataLabel** object.


## Remarks

The chart must first be active before you can access the data labels programmatically or a run-time error will occur.


## Example

This example shows the bubble size for the data labels of the first series on the first chart. This example assumes a chart exists on the active worksheet.


```vb
Sub UseBubbleSize() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowBubbleSize = True 
 
End Sub
```


## See also


#### Concepts


[DataLabel Object](datalabel-object-excel.md)

