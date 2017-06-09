---
title: QueryTable.FillAdjacentFormulas Property (Excel)
keywords: vbaxl10.chm518076
f1_keywords:
- vbaxl10.chm518076
ms.prod: excel
api_name:
- Excel.QueryTable.FillAdjacentFormulas
ms.assetid: 513a9218-a0b9-2bf6-ebac-1d9e7bb594df
ms.date: 06/08/2017
---


# QueryTable.FillAdjacentFormulas Property (Excel)

 **True** if formulas to the right of the specified query table are automatically updated whenever the query table is refreshed. Read/write **Boolean** .


## Syntax

 _expression_ . **FillAdjacentFormulas**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **FillAdjacentFormulas** property applies only to **QueryTable** objects.


## Example

This example sets query table one so that formulas to the right of it are automatically updated whenever the query table is refreshed.


```vb
Sheets("sheet1").QueryTables(1).FillAdjacentFormulas = True
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

