---
title: QueryTable.CommandType Property (Excel)
keywords: vbaxl10.chm518114
f1_keywords:
- vbaxl10.chm518114
ms.prod: excel
api_name:
- Excel.QueryTable.CommandType
ms.assetid: ed1b668c-a73c-0ee7-45ed-67a9d46921dd
ms.date: 06/08/2017
---


# QueryTable.CommandType Property (Excel)

Returns or sets one of the  **[XlCmdType](xlcmdtype-enumeration-excel.md)** constants listed in the following table in the remarks section. The constant that is returned or set describes the value of the **[CommandText](querytable-commandtext-property-excel.md)** property. The default value is **xlCmdSQL** . Read/write **XlCmdType** .


## Syntax

 _expression_ . **CommandType**

 _expression_ An expression that returns a **QueryTable** object.


## Remarks



| **XlCmdType** can be one of these **XlCmdType** constants.|
| **xlCmdCube** . Contains a cube name for an OLAP data source.|
| **xlCmdDefault** . Contains command text that the OLE DB provider understands.|
| **xlCmdSql** . Contains an SQL statement.|
| **xlCmdTable** . Contains a table name for accessing OLE DB data sources.|
You can set the  **CommandType** property only if the value of the **[QueryType](querytable-querytype-property-excel.md)** property for the query table or PivotTable cache is **xlOLEDBQuery** .

If the value of the  **CommandType** property is **xlCmdCube** , you cannot change this value if there is a PivotTable report associated with the query table.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **CommandType** property.


## Example

This example sets the command string for the first query table's ODBC data source. The command string is an SQL statement.


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

