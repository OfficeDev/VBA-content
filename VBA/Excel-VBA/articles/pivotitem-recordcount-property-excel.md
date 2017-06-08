---
title: PivotItem.RecordCount Property (Excel)
keywords: vbaxl10.chm246088
f1_keywords:
- vbaxl10.chm246088
ms.prod: excel
api_name:
- Excel.PivotItem.RecordCount
ms.assetid: 2ba8ceff-5c9c-ed27-7b32-b9f9e7bd7ff0
ms.date: 06/08/2017
---


# PivotItem.RecordCount Property (Excel)

Returns the number of records in the PivotTable cache or the number of cache records that contain the specified item. Read-only  **Long** .


## Syntax

 _expression_ . **RecordCount**

 _expression_ A variable that represents a **PivotItem** object.


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


[PivotItem Object](pivotitem-object-excel.md)

