---
title: QueryTable.WebDisableRedirections Property (Excel)
keywords: vbaxl10.chm518129
f1_keywords:
- vbaxl10.chm518129
ms.prod: excel
api_name:
- Excel.QueryTable.WebDisableRedirections
ms.assetid: 36aec986-de9c-2c7e-a07c-ae77d75d4c7c
ms.date: 06/08/2017
---


# QueryTable.WebDisableRedirections Property (Excel)

 **True** if Web query redirections are disabled for a **QueryTable** object. The default value is **False** . Read/write **Boolean** .


## Syntax

 _expression_ . **WebDisableRedirections**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **WebDisableRedirections** property applies only to **QueryTable** objects.


## Example

In this example, Microsoft Excel determines the settings of Web query redirections for the first worksheet in the workbook. This example assumes a  **QueryTable** object exists on the first worksheet, otherwise a run-time error will occur.


```vb
Sub CheckWebQuerySetting() 
 Dim wksSheet As Worksheet 
 Set wksSheet = Application.ActiveSheet 
 MsgBox wksSheet.QueryTables(1).WebDisableRedirections 
End Sub
```


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

