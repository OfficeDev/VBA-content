---
title: QueryTable.AdjustColumnWidth Property (Excel)
keywords: vbaxl10.chm518112
f1_keywords:
- vbaxl10.chm518112
ms.prod: excel
api_name:
- Excel.QueryTable.AdjustColumnWidth
ms.assetid: 2901cc84-92d2-7021-2360-9c31dc1153b3
ms.date: 06/08/2017
---


# QueryTable.AdjustColumnWidth Property (Excel)

 **True** if the column widths are automatically adjusted for the best fit each time you refresh the specified query table. **False** if the column widths are not automatically adjusted with each refresh. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **AdjustColumnWidth**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

The maximum column width is two-thirds the width of the screen.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **AdjustColumnWidth** property.


## Example

This example turns off automatic column-width adjustment for the newly added query table on the first worksheet in the first workbook.


```vb
With Workbooks(1).Worksheets(1).QueryTables _ 
 .Add(Connection:= varDBConnStr, _ 
 Destination:=Range("B1"), _ 
 Sql:="Select Price From CurrentStocks " &; _ 
 "Where Symbol = 'MSFT'") 
 .AdjustColumnWidth = False 
 .Refresh 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

