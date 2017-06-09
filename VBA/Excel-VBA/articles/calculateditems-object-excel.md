---
title: CalculatedItems Object (Excel)
keywords: vbaxl10.chm249072
f1_keywords:
- vbaxl10.chm249072
ms.prod: excel
api_name:
- Excel.CalculatedItems
ms.assetid: daad9732-6a20-d146-050e-da9e1c1e6f33
ms.date: 06/08/2017
---


# CalculatedItems Object (Excel)

A collection of  **[PivotItem](pivotitem-object-excel.md)** objects that represent all the calculated items in the specified PivotTable report.


## Remarks

A PivotTable report that contains January, February, and March items could have a calculated item named "FirstQuarter" defined as the sum of the amounts in January, February, and March.

Use the  **[CalculatedItems](pivotfield-calculateditems-method-excel.md)** method to return the **CalculatedItems** collection.

Use  **CalculatedFields** ( _index_ ), where _index_ is the name or index number of the field, to return a single **[PivotField](pivotfield-object-excel.md)** object from the **[CalculatedFields](calculatedfields-object-excel.md)** collection.


## Example

The following example creates a list of the calculated items in the first PivotTable report on worksheet one, along with their formulas.


```vb
Set pt = Worksheets(1).PivotTables(1) 
For Each ci In pt.PivotFields("Sales").CalculatedItems 
 r = r + 1 
 With Worksheets(2) 
 .Cells(r, 1).Value = ci.Name 
 .Cells(r, 2).Value = ci.Formula 
 End With 
Next
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

