---
title: PivotFormula Object (Excel)
keywords: vbaxl10.chm230072
f1_keywords:
- vbaxl10.chm230072
ms.prod: excel
api_name:
- Excel.PivotFormula
ms.assetid: 2955dad6-d686-1a83-ab56-76a00272c7e2
ms.date: 06/08/2017
---


# PivotFormula Object (Excel)

Represents a formula used to calculate results in a PivotTable report.


## Remarks

This object and its associated properties and methods aren't available for OLAP data sources because calculated fields and items aren't supported.


## Example

Use  **[PivotFormulas](pivottable-pivotformulas-property-excel.md)** ( _index_ ), where _index_ is the formula number or string on the left side of the formula, to return the **PivotFormula** object. The following example changes the index number for formula one in the first PivotTable report on the first worksheet so that this formula will be solved after formula two.


```vb
Worksheets(1).PivotTables(1).PivotFormulas(1).Index = 2
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


