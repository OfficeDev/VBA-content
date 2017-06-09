---
title: Workbook.PivotCaches Method (Excel)
keywords: vbaxl10.chm199124
f1_keywords:
- vbaxl10.chm199124
ms.prod: excel
api_name:
- Excel.Workbook.PivotCaches
ms.assetid: 0a2e7f10-c123-5c98-fb71-56868b9f8bde
ms.date: 06/08/2017
---


# Workbook.PivotCaches Method (Excel)

Returns a  **[PivotCaches](pivotcaches-object-excel.md)** collection that represents all the PivotTable caches in the specified workbook. Read-only.


## Syntax

 _expression_ . **PivotCaches**

 _expression_ A variable that represents a **Workbook** object.


### Return Value

PivotCaches


## Example

This example causes the PivotTable cache to update automatically each time the workbook is opened.


```vb
ActiveWorkbook.PivotCaches(1).RefreshOnFileOpen = True
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

