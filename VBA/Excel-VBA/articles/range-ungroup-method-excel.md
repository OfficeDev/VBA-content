---
title: Range.Ungroup Method (Excel)
keywords: vbaxl10.chm144212
f1_keywords:
- vbaxl10.chm144212
ms.prod: excel
api_name:
- Excel.Range.Ungroup
ms.assetid: ac20c780-1a8e-2709-13c4-a6ca8220fb0a
ms.date: 06/08/2017
---


# Range.Ungroup Method (Excel)

Promotes a range in an outline (that is, decreases its outline level). The specified range must be a row or column, or a range of rows or columns. If the range is in a PivotTable report, this method ungroups the items contained in the range.


## Syntax

 _expression_ . **Ungroup**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Remarks

If the active cell is in a field header of a parent field, all the groups in that field are ungrouped and the field is removed from the PivotTable report. When the last group in a parent field is ungrouped, the entire field is removed from the report.


## Example

This example ungroups the ORDER_DATE field.


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Set groupRange = pvtTable.PivotFields("ORDER_DATE").DataRange 
groupRange.Cells(1).Ungroup
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

