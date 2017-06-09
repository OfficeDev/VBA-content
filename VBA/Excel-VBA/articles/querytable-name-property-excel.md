---
title: QueryTable.Name Property (Excel)
keywords: vbaxl10.chm518073
f1_keywords:
- vbaxl10.chm518073
ms.prod: excel
api_name:
- Excel.QueryTable.Name
ms.assetid: 56001390-df2e-b28a-6567-786453424f38
ms.date: 06/08/2017
---


# QueryTable.Name Property (Excel)

Returns or sets a  **String** value representing the name of the object.


## Syntax

 _expression_ . **Name**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

The following table shows example values of the  **Name** property and related properties given an OLAP data source with the unique name "[Europe].[France].[Paris]" and a non-OLAP data source with the item name "Paris".



|**Property**|**Value (OLAP data source)**|**Value (non-OLAP data source)**|
|:-----|:-----|:-----|
| **Caption**|Paris|Paris|
| **Name**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
| **SourceName**|[Europe].[France].[Paris] &nbsp;(read-only)|(Same as SQL property value, read-only)|
| **Value**|[Europe].[France].[Paris] &nbsp;(read-only)|Paris|
If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **Name** property applies only to **QueryTable** objects.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

