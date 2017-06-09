---
title: QueryTable.WebSingleBlockTextImport Property (Excel)
keywords: vbaxl10.chm518126
f1_keywords:
- vbaxl10.chm518126
ms.prod: excel
api_name:
- Excel.QueryTable.WebSingleBlockTextImport
ms.assetid: 044de013-a065-86a3-b910-d4dec0a761b8
ms.date: 06/08/2017
---


# QueryTable.WebSingleBlockTextImport Property (Excel)

 **True** if data from the HTML <PRE> tags in the specified Web page is processed all at once when you import the page into a query table. **False** if the data is imported in blocks of contiguous rows so that header rows will be recognized as such. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **WebSingleBlockTextImport**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when the query table's  **[QueryType](querytable-querytype-property-excel.md)** property is set to **xlWebQuery** and the query returns an HTML document.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **WebSingleBlockTextImport** property applies only to **QueryTable** objects.


## Example

This example adds a new Web query table to the first worksheet in the first workbook and and then imports all of the HTML <PRE> tag data all at once.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;http://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebSingleBlockTextImport = True 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

