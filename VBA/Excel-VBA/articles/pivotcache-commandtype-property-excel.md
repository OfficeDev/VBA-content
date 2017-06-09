---
title: PivotCache.CommandType Property (Excel)
keywords: vbaxl10.chm227088
f1_keywords:
- vbaxl10.chm227088
ms.prod: excel
api_name:
- Excel.PivotCache.CommandType
ms.assetid: bbe0ba26-efb9-428d-de2c-576116d92747
ms.date: 06/08/2017
---


# PivotCache.CommandType Property (Excel)

Returns or sets one of the  **[XlCmdType](xlcmdtype-enumeration-excel.md)** constants listed in the following table in the remarks section. The constant that is returned or set describes the value of the **[CommandText](pivotcache-commandtext-property-excel.md)** property. The default value is **xlCmdSQL** . Read/write **XlCmdType** .


## Syntax

 _expression_ . **CommandType**

 _expression_ An expression that returns a **PivotCache** object.


## Remarks



| **XlCmdType** can be one of these **XlCmdType** constants.|
| **xlCmdCube** . Contains a cube name for an OLAP data source.|
| **xlCmdDefault** . Contains command text that the OLE DB provider understands.|
| **xlCmdSql** . Contains an SQL statement.|
| **xlCmdTable** . Contains a table name for accessing OLE DB data sources.|
You can set the  **CommandType** property only if the value of the **[QueryType](pivotcache-querytype-property-excel.md)** property for the query table or PivotTable cache is **xlOLEDBQuery** .

If the value of the  **CommandType** property is **xlCmdCube** , you cannot change this value if there is a PivotTable report associated with the query table.


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


[PivotCache Object](pivotcache-object-excel.md)

