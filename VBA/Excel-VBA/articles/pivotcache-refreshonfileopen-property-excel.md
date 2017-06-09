---
title: PivotCache.RefreshOnFileOpen Property (Excel)
keywords: vbaxl10.chm227083
f1_keywords:
- vbaxl10.chm227083
ms.prod: excel
api_name:
- Excel.PivotCache.RefreshOnFileOpen
ms.assetid: aed513aa-b752-8b6e-0d6d-6fddab46df18
ms.date: 06/08/2017
---


# PivotCache.RefreshOnFileOpen Property (Excel)

 **True** if the PivotTable cache is automatically updated each time the workbook is opened. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **RefreshOnFileOpen**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks

Query tables and PivotTable reports are not automatically refreshed when you open the workbook by using the  **[Open](workbooks-open-method-excel.md)** method in Visual Basic. Use the **[Refresh](pivotcache-refresh-method-excel.md)** method to refresh the data after the workbook is open.


## Example

This example causes the PivotTable cache to automatically update each time the workbook is opened.


```vb
ActiveWorkbook.PivotCaches(1).RefreshOnFileOpen = True
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

