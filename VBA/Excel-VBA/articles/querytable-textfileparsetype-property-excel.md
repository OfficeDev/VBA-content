---
title: QueryTable.TextFileParseType Property (Excel)
keywords: vbaxl10.chm518100
f1_keywords:
- vbaxl10.chm518100
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileParseType
ms.assetid: 58117c6a-bfe4-190b-ab72-1a26e961d25d
ms.date: 06/08/2017
---


# QueryTable.TextFileParseType Property (Excel)

Returns or sets the column format for the data in the text file that you're importing into a query table. Read/write  **[XlTextParsingType](xltextparsingtype-enumeration-excel.md)** .


## Syntax

 _expression_ . **TextFileParseType**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks



| **XlTextParsingType** can be one of these **XlTextParsingType** constants.|
| **xlFixedWidth** . Indicates that the data in the file is arranged in columns of fixed widths.|
| **xlDelimited**_default_ . Iindicates the file is delimited by delimiter characters|
Use this property only when your query table is based on data from a text file (with the  **[QueryType](querytable-querytype-property-excel.md)** property set to **xlTextImport** ).

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileParseType** property applies only to **QueryTable** objects.


## Example

This example imports a fixed-width text file into a new query table on the first worksheet in the first workbook. The first column in the text file is five characters wide and is imported as text. The second column is four characters wide and is skipped. The remainder of the text file is imported into the third column and has the General format applied to it.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1, 1)) 
With qtQtrResults 
 .TextFileParseType = xlFixedWidth 
 .TextFileFixedColumnWidths = Array(5, 4) 
 .TextFileColumnDataTypes = _ 
 Array(xlTextFormat, xlSkipColumn, xlGeneralFormat) 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

