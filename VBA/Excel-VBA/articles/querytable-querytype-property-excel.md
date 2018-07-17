---
title: QueryTable.QueryType Property (Excel)
keywords: vbaxl10.chm518116
f1_keywords:
- vbaxl10.chm518116
ms.prod: excel
api_name:
- Excel.QueryTable.QueryType
ms.assetid: 7cf9ea40-62ea-7211-7832-31eceb44ed15
ms.date: 06/08/2017
---


# QueryTable.QueryType Property (Excel)

Indicates the type of query used by Microsoft Excel to populate the query table. Read-only  **[XlQueryType](xlquerytype-enumeration-excel.md)** .


## Syntax

 _expression_ . **QueryType**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks



| **XlQueryType** can be one of these **XlQueryType** constants.|
| **xlTextImport** . Based on a text file, for query tables only|
| **xlOLEDBQuery** . Based on an OLE DB query, including OLAP data sources|
| **xlWebQuery** . Based on a Web page, for query tables only|
| **xlADORecordset** . Based on an ADO recordset query|
| **xlDAORecordSet** . Based on a DAO recordset query, for query tables only|
| **xlODBCQuery** . Based on an ODBC data source|
You specify the data source in the prefix for the  **[Connection](querytable-connection-property-excel.md)** property's value.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **QueryType** property.


## Example

This example refreshes the first query table on the first worksheet if the table is based on a Web page.


```vb
Set qtQtrResults = _ 
 Workbooks(1).Worksheets(1).QueryTables(1) 
With qtQtrResults 
 if .QueryType = xlWebQuery Then 
 .Refresh 
 End If 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

