---
title: QueryTable.WebDisableDateRecognition Property (Excel)
keywords: vbaxl10.chm518127
f1_keywords:
- vbaxl10.chm518127
ms.prod: excel
api_name:
- Excel.QueryTable.WebDisableDateRecognition
ms.assetid: 6db374e2-67b2-bf84-35d4-dd87494c0d81
ms.date: 06/08/2017
---


# QueryTable.WebDisableDateRecognition Property (Excel)

 **True** if data that resembles dates is parsed as text when you import a Web page into a query table. **False** if date recognition is used. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **WebDisableDateRecognition**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use this property only when the query table's  **[QueryType](querytable-querytype-property-excel.md)** property is set to **xlWebQuery** and the query returns an HTML document.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **WebDisableDateRecognition** property applies only to **QueryTable** objects.


## Example

This example turns off date recognition so that Web page data that resembles dates is imported as text. The example then refreshes the query table.


```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "URL;http://datasvr/98q1/19980331.htm", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .WebDisableDateRecognition = True 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

