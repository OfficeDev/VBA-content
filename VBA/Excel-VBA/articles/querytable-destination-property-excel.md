---
title: QueryTable.Destination Property (Excel)
keywords: vbaxl10.chm518086
f1_keywords:
- vbaxl10.chm518086
ms.prod: excel
api_name:
- Excel.QueryTable.Destination
ms.assetid: 11dc755d-1686-18e9-88df-b885328e8ef5
ms.date: 06/08/2017
---


# QueryTable.Destination Property (Excel)

Returns the cell in the upper-left corner of the query table destination range (the range where the resulting query table will be placed). The destination range must be on the worksheet that contains the  **QueryTable** object. Read-only **[Range](range-object-excel.md)** .


## Syntax

 _expression_ . **Destination**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **Destination** property.


## Example

This example scrolls through the active window until the upper-left corner of query table one is in the upper-left corner of the window.


```vb
Set d = Worksheets(1).QueryTables(1).Destination 
With ActiveWindow 
 .ScrollColumn = d.Column 
 .ScrollRow = d.Row 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

