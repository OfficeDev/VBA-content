---
title: PivotTable.CacheIndex Property (Excel)
keywords: vbaxl10.chm235102
f1_keywords:
- vbaxl10.chm235102
ms.prod: excel
api_name:
- Excel.PivotTable.CacheIndex
ms.assetid: fe1a88b7-dfd0-e031-e739-0b5781de1c0d
ms.date: 06/08/2017
---


# PivotTable.CacheIndex Property (Excel)

Returns or sets the index number of the PivotTable cache. Read/write  **Long** .


## Syntax

 _expression_ . **CacheIndex**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

If you set the  **CacheIndex** property so that one PivotTable report uses the cache for a second PivotTable report, the first report's fields must be a valid subset of the fields in the second report.


## Example

This example sets the cache for the PivotTable report named "Pivot1" to the cache of the PivotTable report named "Pivot2."


```vb
Worksheets(1).PivotTables("Pivot1").CacheIndex = _ 
 Worksheets(1).PivotTables("Pivot2").CacheIndex
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

