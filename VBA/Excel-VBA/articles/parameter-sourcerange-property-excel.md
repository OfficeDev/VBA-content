---
title: Parameter.SourceRange Property (Excel)
keywords: vbaxl10.chm523077
f1_keywords:
- vbaxl10.chm523077
ms.prod: excel
api_name:
- Excel.Parameter.SourceRange
ms.assetid: 243ac075-24cc-549a-58fb-195d71dc6e68
ms.date: 06/08/2017
---


# Parameter.SourceRange Property (Excel)

Returns a  **Range** object that represents the cell that contains the value of the specified query parameter. Read-only.


## Syntax

 _expression_ . **SourceRange**

 _expression_ A variable that represents a **Parameter** object.


## Example

This example changes the value of the cell used as the source range for the query.


```vb
Set qt = Sheets("sheet1").QueryTables(1) 
Set param1 = qt.Parameters(1) 
Set r = param1.SourceRange 
r.Value = "New York" 
qt.Refresh
```


## See also


#### Concepts


[Parameter Object](parameter-object-excel.md)

