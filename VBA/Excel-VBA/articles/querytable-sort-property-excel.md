---
title: QueryTable.Sort Property (Excel)
keywords: vbaxl10.chm518139
f1_keywords:
- vbaxl10.chm518139
ms.prod: excel
api_name:
- Excel.QueryTable.Sort
ms.assetid: 92f268ef-507f-a565-be42-abea73c381a2
ms.date: 06/08/2017
---


# QueryTable.Sort Property (Excel)

Returns the sort criteria for the query table range. Read-only.


## Syntax

 _expression_ . **Sort**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If you import data using the user interface, data from Web queries or text queries is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from Web queries or text queries must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **QueryTable** property of the **ListObject** to access the **Sort** property.


## Example

This example refreshes the query table gets the sort criteria.


```vb
QueryTable.Refresh 
 
Dim var As Sort 
Set var = QueryTable.Sort
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

