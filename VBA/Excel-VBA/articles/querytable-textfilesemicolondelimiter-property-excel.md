---
title: QueryTable.TextFileSemicolonDelimiter Property (Excel)
keywords: vbaxl10.chm518104
f1_keywords:
- vbaxl10.chm518104
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileSemicolonDelimiter
ms.assetid: 61a4ea08-aadd-6cf5-b810-448fe00b68f5
ms.date: 06/08/2017
---


# QueryTable.TextFileSemicolonDelimiter Property (Excel)

 **True** if the semicolon is the delimiter when you import a text file into a query table, and if the value of the **[TextFileParseType](querytable-textfileparsetype-property-excel.md)** property is **xlDelimited** . The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **TextFileSemicolonDelimiter**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when your query table is based on data from a text file (with the  **[QueryType](querytable-querytype-property-excel.md)** property set to **xlTextImport** ).

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileSemicolonDelimiter** property applies only to **QueryTable** objects.


## Example

This example sets the semicolon to be the delimiter in the query table on the first worksheet in the first workbook and then refreshes the query table.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFileParseType = xlDelimited 
 .TextFileSemicolonDelimiter = True 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

