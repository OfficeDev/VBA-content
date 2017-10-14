---
title: QueryTable.CommandText Property (Excel)
keywords: vbaxl10.chm518113
f1_keywords:
- vbaxl10.chm518113
ms.prod: excel
api_name:
- Excel.QueryTable.CommandText
ms.assetid: 5f1f84f2-d613-17be-7b2e-3b6a3cc56002
ms.date: 06/08/2017
---


# QueryTable.CommandText Property (Excel)

Returns or sets the command string for the specified data source. Read/write  **Variant** .


## Syntax

 _expression_ . **CommandText**

 _expression_ An expression that returns a **QueryTable** object.


## Remarks

For OLE DB sources, the  **[CommandType](pivotcache-commandtype-property-excel.md)** property describes the value of the **CommandText** property.

For ODBC sources, setting the  **CommandText** causes the data to be refreshed.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **CommandText** property.

The sheet that contains the query table must be active to access this property.


## Example

This example sets the command string for the first query table's ODBC data source. Note that the command string is an SQL statement.


```vb
Set qtQtrResults = _ 
 Workbooks(1).Worksheets(1).QueryTables(1) 
With qtQtrResults 
 .CommandType = xlCmdSQL 
 .CommandText = _ 
 "Select ProductID From Products Where ProductID < 10" 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

