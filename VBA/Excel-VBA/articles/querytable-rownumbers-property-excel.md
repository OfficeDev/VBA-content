---
title: QueryTable.RowNumbers Property (Excel)
keywords: vbaxl10.chm518075
f1_keywords:
- vbaxl10.chm518075
ms.prod: excel
api_name:
- Excel.QueryTable.RowNumbers
ms.assetid: e0e91e2a-f7b6-ef5b-8046-9e93a51395db
ms.date: 06/08/2017
---


# QueryTable.RowNumbers Property (Excel)

 **True** if row numbers are added as the first column of the specified query table. Read/write **Boolean** .


## Syntax

 _expression_ . **RowNumbers**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Setting this property to  **True** doesn't immediately cause row numbers to appear. The row numbers appear the next time the query table is refreshed, and they're reconfigured every time the query table is refreshed.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **RowNumbers** property.


## Example

This example adds row numbers and field names to the query table.


```vb
With Worksheets(1).QueryTables("ExternalData1") 
 .RowNumbers = True 
 .FieldNames = True 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

