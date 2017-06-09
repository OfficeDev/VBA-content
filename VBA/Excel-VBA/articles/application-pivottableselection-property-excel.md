---
title: Application.PivotTableSelection Property (Excel)
keywords: vbaxl10.chm133192
f1_keywords:
- vbaxl10.chm133192
ms.prod: excel
api_name:
- Excel.Application.PivotTableSelection
ms.assetid: e0a93c11-2e2f-23af-6cad-b4f22883128e
ms.date: 06/08/2017
---


# Application.PivotTableSelection Property (Excel)

 **True** if PivotTable reports use structured selection. Read/write **Boolean** .


## Syntax

 _expression_ . **PivotTableSelection**

 _expression_ A variable that represents an **Application** object.


## Example

This example enables structured selection mode and then sets the first PivotTable report on worksheet one to allow only data to be selected.


```vb
Application.PivotTableSelection = True 
Worksheets(1).PivotTables(1).SelectionMode = xlDataOnly
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

