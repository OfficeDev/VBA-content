---
title: QueryTable.TextFileFixedColumnWidths Property (Excel)
keywords: vbaxl10.chm518109
f1_keywords:
- vbaxl10.chm518109
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileFixedColumnWidths
ms.assetid: adfc63a2-3594-5b36-dccf-28a1cd99c84d
ms.date: 06/08/2017
---


# QueryTable.TextFileFixedColumnWidths Property (Excel)

Returns or sets an array of integers that correspond to the widths of the columns (in characters) in the text file that you're importing into a query table. Valid widths are from 1 through 32767 characters. Read/write  **Variant** .


## Syntax

 _expression_ . **TextFileFixedColumnWidths**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when your query table is based on data from a text file (with the  **[QueryType](querytable-querytype-property-excel.md)** property set to **xlTextImport** ), and only if the value of the **[TextFileParseType](querytable-textfileparsetype-property-excel.md)** property is **xlFixedWidth** .

You must specify a valid, nonnegative column width. If you specify columns that exceed the width of the text file, those values are ignored. If the width of the text file is greater than the total width of columns you specify, the balance of the text file is imported into an additional column.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileFixedColumnWidths** property applies only to **QueryTable** objects.


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

