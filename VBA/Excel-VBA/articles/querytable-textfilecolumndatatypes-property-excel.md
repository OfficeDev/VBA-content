---
title: QueryTable.TextFileColumnDataTypes Property (Excel)
keywords: vbaxl10.chm518108
f1_keywords:
- vbaxl10.chm518108
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileColumnDataTypes
ms.assetid: 05445aaf-df9c-5981-7803-639c80460906
ms.date: 06/08/2017
---


# QueryTable.TextFileColumnDataTypes Property (Excel)

Returns or sets an ordered array of constants that specify the data types applied to the corresponding columns in the text file that you are importing into a query table. The default constant for each column is  **xlGeneral** . Read/write **Variant** .


## Syntax

 _expression_ . **TextFileColumnDataTypes**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use the  **xlColumnDataType** constants listed in the following table to specify the column data types used or the actions taken during the data import.



|**Constant**|**Description**|
|:-----|:-----|
| **xlGeneralFormat**|General|
| **xlTextFormat**|Text|
| **xlSkipColumn**|Skip column|
| **xlDMYFormat**|Day-Month-Year date format|
| **xlDYMFormat**|Day-Year-Month date format|
| **xlEMDFormat**|EMD date|
| **xlMDYFormat**|Month-Day-Year date format|
| **xlMYDFormat**|Month-Year-Day date format|
| **xlYDMFormat**|Year-Day-Month date format|
| **xlYMDFormat**|Year-Month-Day date format|
Use this property only when your query table is based on data from a text file (with the  **[QueryType](querytable-querytype-property-excel.md)** property set to **xlTextImport** ).

If you specify more elements in the array that there are columns, those values are ignored.

Use  **xlEMDFormat** only if Chinese (Taiwan) language support is installed and selected. The **xlEMDFormat** constant specifies that Chinese (Taiwan) era dates are used.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileColumnDataTypes** property applies only to **QueryTable** objects.


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

