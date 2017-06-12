---
title: PivotTable.PageFieldOrder Property (Excel)
keywords: vbaxl10.chm235119
f1_keywords:
- vbaxl10.chm235119
ms.prod: excel
api_name:
- Excel.PivotTable.PageFieldOrder
ms.assetid: 0c8a6473-f2ee-f357-b840-aaf61cee1fa0
ms.date: 06/08/2017
---


# PivotTable.PageFieldOrder Property (Excel)

Returns or sets the order in which page fields are added to the PivotTable report's layout. Can be one of the following  **[XlOrder](xlorder-enumeration-excel.md)** constants: **xlDownThenOver** or **xlOverThenDown** . The default constant is **xlDownThenOver** . Read/write **Long** .


## Syntax

 _expression_ . **PageFieldOrder**

 _expression_ A variable that represents a **PivotTable** object.


## Example

This example causes the PivotTable report to draw three page fields in a row before starting a new row.


```vb
With Worksheets(1).PivotTables("Pivot1") 
 .PageFieldOrder = xlOverThenDown 
 .PageFieldWrapCount = 3 
End With
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

