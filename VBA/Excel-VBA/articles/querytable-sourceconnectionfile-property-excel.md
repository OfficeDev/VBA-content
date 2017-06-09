---
title: QueryTable.SourceConnectionFile Property (Excel)
keywords: vbaxl10.chm518131
f1_keywords:
- vbaxl10.chm518131
ms.prod: excel
api_name:
- Excel.QueryTable.SourceConnectionFile
ms.assetid: 2f7472a2-dbac-5dbb-ea27-1508211f001f
ms.date: 06/08/2017
---


# QueryTable.SourceConnectionFile Property (Excel)

Returns or sets a  **String** indicating the Microsoft Office Data Connection file or similar file that was used to create the QueryTable. Read/write.


## Syntax

 _expression_ . **SourceConnectionFile**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Data from Web queries or text queries is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object. You can use the **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **SourceConnectionFile** property.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **SourceConnectionFile** property.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

