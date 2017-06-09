---
title: QueryTable.TextFileVisualLayout Property (Excel)
keywords: vbaxl10.chm518137
f1_keywords:
- vbaxl10.chm518137
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileVisualLayout
ms.assetid: 13105ba8-945d-9e9b-f90c-9059e2ade9f1
ms.date: 06/08/2017
---


# QueryTable.TextFileVisualLayout Property (Excel)

Returns or sets a  **[XlTextVisualLayoutType](xltextvisuallayouttype-enumeration-excel.md)** enumeration that indicates whether the visual layout of the text being imported is left-to-right or right-to-left.


## Syntax

 _expression_ . **TextFileVisualLayout**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks



| **XlTextVisualLayoutType** can be one of the following **XlTextVisualLayoutType** constants.|
| **xlTextVisualLTR**|
| **xlTextVisualRTL**|
If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **TextFileVisualLayout** property applies only to **QueryTable** objects.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

