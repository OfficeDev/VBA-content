---
title: PivotField.ServerBased Property (Excel)
keywords: vbaxl10.chm240110
f1_keywords:
- vbaxl10.chm240110
ms.prod: excel
api_name:
- Excel.PivotField.ServerBased
ms.assetid: 8c97a617-e852-b21e-7acf-f0d31363adf3
ms.date: 06/08/2017
---


# PivotField.ServerBased Property (Excel)

 **True** if the data source for the specified PivotTable report is external and only the items matching the page field selection are retrieved. Read/write **Boolean** .


## Syntax

 _expression_ . **ServerBased**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

This property doesn't apply to OLAP data sources and is always  **False** .

When this property is  **True** , only records in the database that match the selected page field item are retrieved. From then on, whenever the user changes the page field selection, the newly selected page field item is passed to the query as a parameter, and the cache is refreshed.

This property cannot be set if any of the following conditions are true:




- The field is grouped.
    
- The data source isn't external.
    
- The cache is shared by two or more PivotTable reports.
    
- The field is a data type that cannot be server based (a memo field or an OLE object).
    



## Example

This example lists all the server-based page fields.


```vb
For Each fld in ActiveSheet.PivotTables(1).PageFields 
 If fld.ServerBased = True Then 
 r = r + 1 
 Worksheets(2).Cells(r, 1).Value = fld.Name 
 End If 
Next
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

