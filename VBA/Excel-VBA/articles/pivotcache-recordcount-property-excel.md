---
title: PivotCache.RecordCount Property (Excel)
keywords: vbaxl10.chm227079
f1_keywords:
- vbaxl10.chm227079
ms.prod: excel
api_name:
- Excel.PivotCache.RecordCount
ms.assetid: 5fcdcf2d-d52f-6ac1-ef09-8377fc5a1f4d
ms.date: 06/08/2017
---


# PivotCache.RecordCount Property (Excel)

Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only  **Long** .


## Syntax

 _expression_ . **RecordCount**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks

This property reflects the transient state of the cache at the time that it's queried. The cache can change between queries.


## Example

This example displays the number of cache records that contain "Kiwi" in the "Products" field.


```vb
MsgBox Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Product").PivotItems("Kiwi").RecordCount
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

