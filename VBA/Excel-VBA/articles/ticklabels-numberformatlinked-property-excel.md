---
title: TickLabels.NumberFormatLinked Property (Excel)
keywords: vbaxl10.chm616078
f1_keywords:
- vbaxl10.chm616078
ms.prod: excel
api_name:
- Excel.TickLabels.NumberFormatLinked
ms.assetid: 8ca8dc6c-b061-503e-f874-cd506242ea07
ms.date: 06/08/2017
---


# TickLabels.NumberFormatLinked Property (Excel)

 **True** if the number format is linked to the cells (so that the number format changes in the labels when it changes in the cells). Read/write **Boolean** .


## Syntax

 _expression_ . **NumberFormatLinked**

 _expression_ A variable that represents a **TickLabels** object.


## Example

This example links the number format for tick-mark labels to its cells for the value axis in Chart1.


```vb
Charts("Chart1").Axes(xlValue).TickLabels.NumberFormatLinked = True
```


## See also


#### Concepts


[TickLabels Object](ticklabels-object-excel.md)

