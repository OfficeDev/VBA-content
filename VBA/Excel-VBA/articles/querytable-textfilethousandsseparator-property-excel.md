---
title: QueryTable.TextFileThousandsSeparator Property (Excel)
keywords: vbaxl10.chm518119
f1_keywords:
- vbaxl10.chm518119
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileThousandsSeparator
ms.assetid: 8eaa9285-9936-2ba1-44d2-9745150b999e
ms.date: 06/08/2017
---


# QueryTable.TextFileThousandsSeparator Property (Excel)

Returns or sets the thousands separator character thatMicrosoft Excel uses when you import a text file into a query table. The default is the system thousands separator character. Read/write  **String** .


## Syntax

 _expression_ . **TextFileThousandsSeparator**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when your query table is based on data from a text file (with the  **[QueryTable](querytable-object-excel.md)** property set to **xlTextImport** ), especially when the file contains decimal and thousands separators that are different from those used on the computer, due to a different language setting being used.

The following table shows the results when you import text into Microsoft Excel using various separators. Numeric results are displayed in the rightmost column.



|**System decimal separator**|**System thousands separator**|**TextFileDecimalSeparator value**|**TextFileThousandsSeparator value**|**Text imported**|**Cell value (data type)**|
|:-----|:-----|:-----|:-----|:-----|:-----|
|Period|Comma|Comma|Period|123.123,45|123,123.45 (numeric)|
|Period|Comma|Comma|Comma|123.123,45|123.123,45 (text)|
|Comma|Period|Comma|Period|123,123.45|123,123.45 (numeric)|
|Period|Comma|Period|Comma|123 123.45|123 123.45 (text)|
|Period|Comma|Period|Space|123 123.45|123,123.45 (numeric)|
If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileThousandsSeparator** property applies only to **QueryTable** objects.


## Example

This example saves the original thousands separator and sets it to a period for the first query table on Sheet1, in preparation for importing a French text file (for example) into the U.S. English version of Microsoft Excel.


```
strDecSep = Worksheets("Sheet1").QueryTables(1) _ 
 .TextFileThousandsSeparator 
Worksheets("Sheet1").QueryTables(1) _ 
 .TextFileThousandsSeparator = "."
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

