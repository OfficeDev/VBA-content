---
title: PivotField.DrilledDown Property (Excel)
keywords: vbaxl10.chm240125
f1_keywords:
- vbaxl10.chm240125
ms.prod: excel
api_name:
- Excel.PivotField.DrilledDown
ms.assetid: 6fb6ae8b-ce41-9343-316c-d26bb1ae9630
ms.date: 06/08/2017
---


# PivotField.DrilledDown Property (Excel)

 **True** if the flag for the specified PivotTable field or PivotTable item is set to "drilled" (expanded, or visible). Read/write **Boolean** .


## Syntax

 _expression_ . **DrilledDown**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

You can use this property only for OLAP data sources.

You cannot set this property if the field or item is hidden.


## Example

This example sets the flags to "not drilled" for all items in the state field in the third PivotTable report on the active worksheet.


```vb
ActiveSheet.PivotTables("PivotTable3") _ 
 .PivotFields("state").DrilledDown = False
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

