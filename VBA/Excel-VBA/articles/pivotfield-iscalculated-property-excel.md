---
title: PivotField.IsCalculated Property (Excel)
keywords: vbaxl10.chm240108
f1_keywords:
- vbaxl10.chm240108
ms.prod: excel
api_name:
- Excel.PivotField.IsCalculated
ms.assetid: 9f2f9856-c61f-d5b8-27bf-6511cac1e898
ms.date: 06/08/2017
---


# PivotField.IsCalculated Property (Excel)

 **True** if the PivotTable field is a calculated field or item. Read-only **Boolean** .


## Syntax

 _expression_ . **IsCalculated**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

For OLAP data sources, this property always returns  **False** .


## Example

This example disables the  **PivotTable Field** dialog box if the specified PivotTable report contains any calculated fields.


```vb
set pt = Worksheets(1).PivotTables("Pivot1") 
For Each fld in pt.PivotFields 
 If fld.IsCalculated Then pt.EnableFieldDialog = False 
Next
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

