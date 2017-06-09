---
title: QueryTable.WebConsecutiveDelimitersAsOne Property (Excel)
keywords: vbaxl10.chm518128
f1_keywords:
- vbaxl10.chm518128
ms.prod: excel
api_name:
- Excel.QueryTable.WebConsecutiveDelimitersAsOne
ms.assetid: cc10dd93-2574-7575-3326-1d2992f4c731
ms.date: 06/08/2017
---


# QueryTable.WebConsecutiveDelimitersAsOne Property (Excel)

 **True** if consecutive delimiters are treated as a single delimiter when you import data from HTML <PRE> tags in a Web page into a query table, and if the data is to be parsed into columns. **False** if you want to treat consecutive delimiters as multiple delimiters. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **WebConsecutiveDelimitersAsOne**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when the query table's  **[QueryType](querytable-querytype-property-excel.md)** property is set to **xlWebQuery** , the query returns an HTML document, and the **[WebPreFormattedTextToColumns](querytable-webpreformattedtexttocolumns-property-excel.md)** property is set to **True** .

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **WebConsecutiveDelimitersAsOne** property applies only to **QueryTable** objects.


## Example

This example sets the space character to be the delimiter in the query table on the first worksheet in the first workbook, and then it refreshes the query table. Consecutive spaces are treated as a single space.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;http://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebConsecutiveDelimitersAsOne = True 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

