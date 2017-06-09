---
title: PivotTable.TotalsAnnotation Property (Excel)
keywords: vbaxl10.chm235136
f1_keywords:
- vbaxl10.chm235136
ms.prod: excel
api_name:
- Excel.PivotTable.TotalsAnnotation
ms.assetid: ce225526-f4b9-8b6a-0b19-21bea06cd728
ms.date: 06/08/2017
---


# PivotTable.TotalsAnnotation Property (Excel)

 **True** if an asterisk (\*) is displayed next to each subtotal and grand total value in the specified PivotTable report if the report is based on an OLAP data source. The default value is **True**. Read/write **Boolean**.


## Syntax

 _expression_ . **TotalsAnnotation**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

When this property is set to  **True**, the asterisk indicates that hidden items are included in the total. The asterisk appears regardless of whether any items in the report have been hidden.

For non-OLAP data sources, the value of this property is always  **False**.


## Example

This example turns off the asterisks in the first PivotTable report on the active worksheet.


```vb
ActiveSheet.PivotTables(1).TotalsAnnotation = False
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

