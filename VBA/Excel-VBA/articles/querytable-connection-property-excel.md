---
title: QueryTable.Connection Property (Excel)
keywords: vbaxl10.chm518087
f1_keywords:
- vbaxl10.chm518087
ms.prod: excel
api_name:
- Excel.QueryTable.Connection
ms.assetid: a576c5d2-113c-cbd0-1ad2-aa46591944de
ms.date: 06/08/2017
---


# QueryTable.Connection Property (Excel)

Returns or sets a string that contains one of the following: OLE DB settings that enable Microsoft Excel to connect to an OLE DB data source; ODBC settings that enable Microsoft Excel to connect to an ODBC data source; a URL that enables Microsoft Excel to connect to a Web data source; the path to and file name of a text file, or the path to and file name of a file that specifies a database or Web query. Read/write  **Variant** .


## Syntax

 _expression_ . **Connection**

 _expression_ An expression that returns a **QueryTable** object.


## Remarks

Setting the  **Connection** property doesn't immediately initiate the connection to the data source. You must use the **[Refresh](querytable-refresh-method-excel.md)** method to make the connection and retrieve the data.

For more information about the connection string syntax, see the  **[Add](querytables-add-method-excel.md)** method of the **QueryTables** collection.

Alternatively, you may choose to access a data source directly by using the Microsoft ActiveX Data Objects (ADO) library instead.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **Connection** property.


## Example

This example supplies new ODBC connection information for the first query table on the first worksheet.


```vb
Worksheets(1).QueryTables(1) _ 
 .Connection:="ODBC;DSN=96SalesData;UID=Rep21;PWD=NUyHwYQI;"
```

This example specifies a text file.




```vb
Worksheets(1).QueryTables(1) _ 
 Connection := "TEXT;C:\My Documents\19980331.txt"
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

