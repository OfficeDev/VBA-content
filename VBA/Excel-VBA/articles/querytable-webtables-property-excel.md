---
title: QueryTable.WebTables Property (Excel)
keywords: vbaxl10.chm518124
f1_keywords:
- vbaxl10.chm518124
ms.prod: excel
api_name:
- Excel.QueryTable.WebTables
ms.assetid: d60eb457-6276-2d86-bbd8-c2050b0695c9
ms.date: 06/08/2017
---


# QueryTable.WebTables Property (Excel)

Returns or sets a comma-delimited list of table names or table index numbers when you import a Web page into a query table. Read/write  **String** .


## Syntax

 _expression_ . **WebTables**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when the query table's  **[QueryType](querytable-querytype-property-excel.md)** property is set to **xlWebQuery** , the query returns an HTML document, and the value of the **[WebSelectionType](querytable-webselectiontype-property-excel.md)** property is **xlSpecifiedTables** .

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **WebTables** property applies only to **QueryTable** objects.


## Example

This example adds a new Web query table to the first worksheet in the first workbook and then imports data from the first and second tables in the Web page.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;http://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebFormatting = xlNone 
 .WebSelectionType = xlSpecifiedTables 
 .WebTables = "1,2" 
 .Refresh 
End With 

```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

