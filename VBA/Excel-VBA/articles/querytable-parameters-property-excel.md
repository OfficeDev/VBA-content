---
title: QueryTable.Parameters Property (Excel)
keywords: vbaxl10.chm518093
f1_keywords:
- vbaxl10.chm518093
ms.prod: excel
api_name:
- Excel.QueryTable.Parameters
ms.assetid: d82f0ef7-9e3a-b9e5-9b9f-d402fb7a573e
ms.date: 06/08/2017
---


# QueryTable.Parameters Property (Excel)

Returns a  **[Parameters](parameters-object-excel.md)** collection that represents the query table parameters. Read-only.


## Syntax

 _expression_ . **Parameters**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

You can use the  **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **Parameters** property.


## Example

This example returns the  **Parameters** collection from an existing parameter query. If the first parameter uses the character data type, the user is instructed to enter characters only in the prompt dialog box.


```vb
With Sheets("sheet1").QueryTables(1).Parameters(1) 
 If .DataType = xlParamTypeVarChar Then 
 .SetParam xlPrompt, "Enter a character only" 
 End If 
End With
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

