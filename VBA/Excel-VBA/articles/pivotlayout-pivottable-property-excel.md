---
title: PivotLayout.PivotTable Property (Excel)
keywords: vbaxl10.chm664082
f1_keywords:
- vbaxl10.chm664082
ms.prod: excel
api_name:
- Excel.PivotLayout.PivotTable
ms.assetid: b4393cb2-33d2-453b-81ef-4fada332539b
ms.date: 06/08/2017
---


# PivotLayout.PivotTable Property (Excel)

Returns a  **[PivotTable](pivottable-object-excel.md)** object that represents the PivotTable report associated with the PivotChart report.


## Syntax

 _expression_ . **PivotTable**

 _expression_ A variable that represents a **PivotLayout** object.


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


[PivotLayout Object](pivotlayout-object-excel.md)

