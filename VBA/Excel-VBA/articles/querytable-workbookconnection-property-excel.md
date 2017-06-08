---
title: QueryTable.WorkbookConnection Property (Excel)
keywords: vbaxl10.chm518138
f1_keywords:
- vbaxl10.chm518138
ms.prod: excel
api_name:
- Excel.QueryTable.WorkbookConnection
ms.assetid: d35d7bb6-5036-1dd9-46ff-e96127d3db09
ms.date: 06/08/2017
---


# QueryTable.WorkbookConnection Property (Excel)

Returns the  **[WorkbookConnection](workbookconnection-object-excel.md)** object that the query table uses. Read-only.


## Syntax

 _expression_ . **WorkbookConnection**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **WorkbookConnection** property applies only to **QueryTable** objects.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

