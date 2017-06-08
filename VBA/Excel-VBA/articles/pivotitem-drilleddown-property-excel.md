---
title: PivotItem.DrilledDown Property (Excel)
keywords: vbaxl10.chm246091
f1_keywords:
- vbaxl10.chm246091
ms.prod: excel
api_name:
- Excel.PivotItem.DrilledDown
ms.assetid: 863909c6-7d2c-4b54-7fb9-de79a6487e4d
ms.date: 06/08/2017
---


# PivotItem.DrilledDown Property (Excel)

 **True** if the flag for the specified PivotTable field or PivotTable item is set to "drilled" (expanded, or visible). Read/write **Boolean** .


## Syntax

 _expression_ . **DrilledDown**

 _expression_ A variable that represents a **PivotItem** object.


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


[PivotItem Object](pivotitem-object-excel.md)

