---
title: QueryTable.Parent Property (Excel)
keywords: vbaxl10.chm517075
f1_keywords:
- vbaxl10.chm517075
ms.prod: excel
api_name:
- Excel.QueryTable.Parent
ms.assetid: 6cf47be7-5e4a-31d0-0c11-e9506c052ecf
ms.date: 06/08/2017
---


# QueryTable.Parent Property (Excel)

Returns the parent object for the specified object. Read-only.


## Syntax

 _expression_ . **Parent**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **Parent** property.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

