---
title: QueryTable.TextFileOtherDelimiter Property (Excel)
keywords: vbaxl10.chm518107
f1_keywords:
- vbaxl10.chm518107
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileOtherDelimiter
ms.assetid: e632984a-4316-4e65-754f-01a2c77d5cad
ms.date: 06/08/2017
---


# QueryTable.TextFileOtherDelimiter Property (Excel)

Returns or sets the character used as the delimiter when you import a text file into a query table. The default value is  **null** . Read/write **String** .


## Syntax

 _expression_ . **TextFileOtherDelimiter**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when your query table is based on data from a text file (with the  **[QueryType](querytable-querytype-property-excel.md)** property set to **xlTextImport** ), and only if the value of the **[TextFileParseType](querytable-textfileparsetype-property-excel.md)** property is **xlDelimited** .

If you specify more than one character in the string, only the first character is used.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileOtherDelimiter** property applies only to **QueryTable** objects.


## Example

This example sets the pound character (#) to be the delimiter for the query table on the first worksheet in the first workbook, and then it refreshes the query table.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFileParseType = xlDelimited 
 .TextFileOtherDelimiter = "#" 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

