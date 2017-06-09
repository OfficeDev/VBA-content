---
title: Parameter.RefreshOnChange Property (Excel)
keywords: vbaxl10.chm523080
f1_keywords:
- vbaxl10.chm523080
ms.prod: excel
api_name:
- Excel.Parameter.RefreshOnChange
ms.assetid: 60e01ae1-82bd-e4eb-6ac7-805ffd05a811
ms.date: 06/08/2017
---


# Parameter.RefreshOnChange Property (Excel)

 **True** if the specified query table is refreshed whenever you change the parameter value of a parameter query. Read/write **Boolean** .


## Syntax

 _expression_ . **RefreshOnChange**

 _expression_ A variable that represents a **Parameter** object.


## Remarks

You can set this property to  **True** only if you use parameters of type **xlRange** and if the referenced parameter value is in a single cell. The refresh occurs when you change the value of the cell.


## Example

This example changes the SQL statement for the first query table on Sheet1. The clause "(ContactTitle=?)" indicates that the query is a parameter query, and the value of the title is set to the value of cell D4. The query table will be automatically refreshed whenever the value of this cell changes.


```vb
Set objQT = Worksheets("Sheet1").QueryTables(1) 
objQT.CommandText = "Select * From Customers Where (ContactTitle=?)" 
Set objParam1 = objQT.Parameters _ 
 .Add("Contact Title", xlParamTypeVarChar) 
objParam1.RefreshOnChange = True 
objParam1.SetParam xlRange, Range("D4")
```


## See also


#### Concepts


[Parameter Object](parameter-object-excel.md)

