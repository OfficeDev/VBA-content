---
title: QueryTable.EnableRefresh Property (Excel)
keywords: vbaxl10.chm518084
f1_keywords:
- vbaxl10.chm518084
ms.prod: excel
api_name:
- Excel.QueryTable.EnableRefresh
ms.assetid: 79a0b628-b90d-1795-830f-e05bc6043517
ms.date: 06/08/2017
---


# QueryTable.EnableRefresh Property (Excel)

 **True** if the PivotTable cache or query table can be refreshed by the user. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **EnableRefresh**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

The  **[RefreshOnFileOpen](querytable-refreshonfileopen-property-excel.md)** property is ignored if the **EnableRefresh** property is set to **False** .

For OLAP data sources, setting this property to  **False** disables updates.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **EnableRefresh** property.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

