---
title: QueryTable.Refreshing Property (Excel)
keywords: vbaxl10.chm518079
f1_keywords:
- vbaxl10.chm518079
ms.prod: excel
api_name:
- Excel.QueryTable.Refreshing
ms.assetid: 7b89fbec-3365-5b23-1b21-da3b0145d9bc
ms.date: 06/08/2017
---


# QueryTable.Refreshing Property (Excel)

 **True** if there is a background query in progress for the specified query table. Read only **Boolean** .


## Syntax

 _expression_ . **Refreshing**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Use the  **[CancelRefresh](querytable-cancelrefresh-method-excel.md)** method to cancel background queries.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **Refreshing** property.


## Example

This example displays a message box if there is a background query in progress for query table one.


```vb
With Worksheets(1).QueryTables(1) 
 If .Refreshing Then 
 MsgBox "Query is currently refreshing: please wait" 
 Else 
 .Refresh BackgroundQuery := False 
 .ResultRange.Select 
 End If 
End With 

```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

