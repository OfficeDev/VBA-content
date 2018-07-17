---
title: QueryTable.RefreshPeriod Property (Excel)
keywords: vbaxl10.chm518120
f1_keywords:
- vbaxl10.chm518120
ms.prod: excel
api_name:
- Excel.QueryTable.RefreshPeriod
ms.assetid: 763c4793-9470-8c0e-3111-d0a0f02948b4
ms.date: 06/08/2017
---


# QueryTable.RefreshPeriod Property (Excel)

Returns or sets the number of minutes between refreshes. Read/write  **Long** .


## Syntax

 _expression_ . **RefreshPeriod**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

Setting the period to 0 (zero) disables automatic timed refreshes and is equivalent to setting this property to  **Null** .

The value of the  **RefreshPeriod** property can be an integer from 0 through 32767.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **RefreshPeriod** property.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

