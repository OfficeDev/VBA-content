---
title: CalculatedFields Object (Excel)
keywords: vbaxl10.chm243072
f1_keywords:
- vbaxl10.chm243072
ms.prod: excel
api_name:
- Excel.CalculatedFields
ms.assetid: 6db4c889-f097-9a66-abc6-28f7f54f0478
ms.date: 06/08/2017
---


# CalculatedFields Object (Excel)

A collection of  **[PivotField](pivotfield-object-excel.md)** objects that represents all the calculated fields in the specified PivotTable report.


## Remarks

A report that contains Revenue and Expense fields could have a calculated field named "Profit" defined as the amount in the Revenue field minus the amount in the Expense field.

For OLAP data sources, you cannot set this collection, and it always returns  **Nothing** .

Use the  **[CalculatedFields](pivottable-calculatedfields-method-excel.md)** method to return the **CalculatedFields** collection .

Use  **CalculatedFields** ( _index_ ), where _index_ is specified field's name or index number, to return a single **PivotField** object from the **CalculatedFields** collection.


## Example

The following example deletes the calculated fields from the PivotTable report named "Pivot1"


```vb
For Each fld in _ 
 Worksheets(1).PivotTables("Pivot1").CalculatedFields 
 fld.Delete 
Next
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

