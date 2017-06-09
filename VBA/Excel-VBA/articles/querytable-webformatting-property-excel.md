---
title: QueryTable.WebFormatting Property (Excel)
keywords: vbaxl10.chm518123
f1_keywords:
- vbaxl10.chm518123
ms.prod: excel
api_name:
- Excel.QueryTable.WebFormatting
ms.assetid: 3ba96959-1c50-8cc0-0025-b5006b1ad62c
ms.date: 06/08/2017
---


# QueryTable.WebFormatting Property (Excel)

Returns or sets a value that determines how much formatting from a Web page, if any, is applied when you import the page into a query table. Read/write  **[XlWebFormatting](xlwebformatting-enumeration-excel.md)** .


## Syntax

 _expression_ . **WebFormatting**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when the query table's  **[QueryType](querytable-querytype-property-excel.md)** property is set to **xlWebQuery** and the query returns an HTML document.



|XlWebFormatting can be one of these XlWebFormatting constants.|
| **xlWebFormattingAll**|
| **xlWebFormattingRTF**|
| **xlWebFormattingNone**_default_|
If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **WebFormatting** property applies only to **QueryTable** objects.


## Example

This example adds a new Web query table to the first worksheet in the first workbook, imports all of the Web page formatting applied to the data, and then refreshes the query table.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;http://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebFormatting = xlAll 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

