---
title: QueryTable.TextFileTextQualifier Property (Excel)
keywords: vbaxl10.chm518101
f1_keywords:
- vbaxl10.chm518101
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileTextQualifier
ms.assetid: a8e6e8cd-4625-1538-b3cd-bf46395943f3
ms.date: 06/08/2017
---


# QueryTable.TextFileTextQualifier Property (Excel)

Returns or sets the text qualifier when you import a text file into a query table. The text qualifier specifies that the enclosed data is in text format. Read/write  **[XlTextQualifier](xltextqualifier-enumeration-excel.md)** .


## Syntax

 _expression_ . **TextFileTextQualifier**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks



| **XlTextQualifier** can be one of these **XlTextQualifier** constants.|
| **xlTextQualifierNone**|
| **xlTextQualifierDoubleQuote**_default_|
| **xlTextQualifierSingleQuote**|
Use this property only when your query table is based on data from a text file (with the  **[QueryType](querytable-querytype-property-excel.md)** property set to **xlTextImport** ).

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileTextQualifier** property applies only to **QueryTable** objects.


## Example

This example sets the single quotation mark character as the text qualifier for the query table on the first worksheet in the first workbook.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFileParseType = xlDelimited 
 .TextFileTextQualifier = xlTextQualifierSingleQuote 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

