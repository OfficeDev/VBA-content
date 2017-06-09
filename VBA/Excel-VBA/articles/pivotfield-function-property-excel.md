---
title: PivotField.Function Property (Excel)
keywords: vbaxl10.chm240081
f1_keywords:
- vbaxl10.chm240081
ms.prod: excel
api_name:
- Excel.PivotField.Function
ms.assetid: 855334f6-dd6d-c09f-7732-c621751374a9
ms.date: 06/08/2017
---


# PivotField.Function Property (Excel)

Returns or sets the function used to summarize the PivotTable field (data fields only). Read/write  **[XlConsolidationFunction](xlconsolidationfunction-enumeration-excel.md)** .


## Syntax

 _expression_ . **Function**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

For OLAP data sources, this property is read-only and always returns  **xlUnknown** . For other data sources, this property cannot be set to **xlUnknown** .


## Example

This example sets the Sum of 1994 field in the first PivotTable report on the active sheet to use the SUM function.


```vb
ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("Sum of 1994").Function = xlSum
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

