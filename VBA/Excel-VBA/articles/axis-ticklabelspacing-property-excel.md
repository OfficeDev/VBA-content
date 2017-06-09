---
title: Axis.TickLabelSpacing Property (Excel)
keywords: vbaxl10.chm561101
f1_keywords:
- vbaxl10.chm561101
ms.prod: excel
api_name:
- Excel.Axis.TickLabelSpacing
ms.assetid: 69e74146-31db-356a-3c00-e5aa35367dc3
ms.date: 06/08/2017
---


# Axis.TickLabelSpacing Property (Excel)

Returns or sets the number of categories or series between tick-mark labels. Applies only to category and series axes. Can be a value from 1 through 31999. Read/write  **Long** .


## Syntax

 _expression_ . **TickLabelSpacing**

 _expression_ A variable that represents an **Axis** object.


## Remarks

Tick-mark label spacing on the value axis is always calculated by Microsoft Excel.


## Example

This example sets the number of categories between tick-mark labels on the category axis in Chart1.


```vb
Charts("Chart1").Axes(xlCategory).TickLabelSpacing = 10 

```


## See also


#### Concepts


[Axis Object](axis-object-excel.md)

