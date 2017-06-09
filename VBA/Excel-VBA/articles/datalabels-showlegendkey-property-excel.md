---
title: DataLabels.ShowLegendKey Property (Excel)
keywords: vbaxl10.chm584096
f1_keywords:
- vbaxl10.chm584096
ms.prod: excel
api_name:
- Excel.DataLabels.ShowLegendKey
ms.assetid: 7bd5c103-b704-448a-35e0-38bd8f120cac
ms.date: 06/08/2017
---


# DataLabels.ShowLegendKey Property (Excel)

 **True** if the data label legend key is visible. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowLegendKey**

 _expression_ A variable that represents a **DataLabels** object.


## Example

This example sets the data labels for series one in Chart1 to show values and the legend key.


```vb
With Charts("Chart1").SeriesCollection(1).DataLabels 
 .ShowLegendKey = True 
 .Type = xlShowValue 
End With
```


## See also


#### Concepts


[DataLabels Object](datalabels-object-excel.md)

