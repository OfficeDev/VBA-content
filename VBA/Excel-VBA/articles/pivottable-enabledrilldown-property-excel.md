---
title: PivotTable.EnableDrilldown Property (Excel)
keywords: vbaxl10.chm235106
f1_keywords:
- vbaxl10.chm235106
ms.prod: excel
api_name:
- Excel.PivotTable.EnableDrilldown
ms.assetid: 329e6c74-6b23-eac8-2ffb-45696076c712
ms.date: 06/08/2017
---


# PivotTable.EnableDrilldown Property (Excel)

 **True** if drilldown is enabled. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **EnableDrilldown**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

Setting this property for a PivotTable report sets it for all fields in that report.

For OLAP data sources, the value is always  **True** .


## Example

This example disables drilldown for all fields in the first PivotTable report on worksheet one/.


```vb
Worksheets(1).PivotTables("Pivot1").EnableDrilldown = False
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

