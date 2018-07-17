---
title: QueryTable.FetchedRowOverflow Property (Excel)
keywords: vbaxl10.chm518080
f1_keywords:
- vbaxl10.chm518080
ms.prod: excel
api_name:
- Excel.QueryTable.FetchedRowOverflow
ms.assetid: 386aaf06-27d4-bfa1-cf5e-ac8c8bddef44
ms.date: 06/08/2017
---


# QueryTable.FetchedRowOverflow Property (Excel)

 **True** if the number of rows returned by the last use of the **[Refresh](querytable-refresh-method-excel.md)** method is greater than the number of rows available on the worksheet. Read-only **Boolean** .


## Syntax

 _expression_ . **FetchedRowOverflow**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **FetchedRowOverflow** property.


## Example

This example refreshes query table one. If the number of rows returned by the query exceeds the number of rows available on the worksheet, an error message is displayed.


```vb
With Worksheets(1).QueryTables(1) 
 .Refresh 
 If .FetchedRowOverflow Then 
 MsgBox "Query too large: please redefine." 
 End If 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

