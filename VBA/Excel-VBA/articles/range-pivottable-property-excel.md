---
title: Range.PivotTable Property (Excel)
keywords: vbaxl10.chm144177
f1_keywords:
- vbaxl10.chm144177
ms.prod: excel
api_name:
- Excel.Range.PivotTable
ms.assetid: ae3f77dc-5098-d60f-0afc-f4f01dbc33f0
ms.date: 06/08/2017
---


# Range.PivotTable Property (Excel)

Returns a  **[PivotTable](pivottable-object-excel.md)** object that represents the PivotTable report containing the upper-left corner of the specified range.


## Syntax

 _expression_ . **PivotTable**

 _expression_ A variable that represents a **Range** object.


## Example

This example sets the current page for the PivotTable report on Sheet1 to the page named "Canada."


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
pvtTable.PivotFields("Country").CurrentPage = "Canada"
```

This example determines the PivotTable report associated with the Sales chart on the active worksheet, and then it sets the page named "Oregon" as the current page for the PivotTable report.




```vb
Set objPT = _ 
 ActiveSheet.Charts("Sales").PivotLayout.PivotTable 
objPT.PivotFields("State").CurrentPageName = "Oregon"
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

