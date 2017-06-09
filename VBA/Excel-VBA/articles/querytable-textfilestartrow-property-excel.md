---
title: QueryTable.TextFileStartRow Property (Excel)
keywords: vbaxl10.chm518099
f1_keywords:
- vbaxl10.chm518099
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileStartRow
ms.assetid: 91b774d8-cf7b-354d-510e-a8561076532c
ms.date: 06/08/2017
---


# QueryTable.TextFileStartRow Property (Excel)

Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1. Read/write  **Long** .


## Syntax

 _expression_ . **TextFileStartRow**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when your query table is based on data from a text file (with the  **[QueryType](querytable-querytype-property-excel.md)** property set to **xlTextImport** ).

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileStartRow** property applies only to **QueryTable** objects.


## Example

This example sets row 5 as the starting row for text parsing in the query table on the first worksheet in the first workbook, and then it refreshes the query table.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFileParseType = xlDelimited 
 .TextFileStartRow = 5 
 .TextFileTabDelimiter = True 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

