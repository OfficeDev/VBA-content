---
title: QueryTable.FieldNames Property (Excel)
keywords: vbaxl10.chm518074
f1_keywords:
- vbaxl10.chm518074
ms.prod: excel
api_name:
- Excel.QueryTable.FieldNames
ms.assetid: ff7541cd-fa4d-6b1a-d8c3-0608cfc03b8d
ms.date: 06/08/2017
---


# QueryTable.FieldNames Property (Excel)

 **True** if field names from the data source appear as column headings for the returned data. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **FieldNames**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **FieldNames** property applies only to **QueryTable** objects.


## Example

This example sets query table one so that the field names don't appear in it.


```vb
Worksheets(1).QueryTables(1).FieldNames = False
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

