---
title: Application.OnUndo Method (Excel)
keywords: vbaxl10.chm133185
f1_keywords:
- vbaxl10.chm133185
ms.prod: excel
api_name:
- Excel.Application.OnUndo
ms.assetid: 12e59bbb-e134-3728-7c8d-629dcda0e908
ms.date: 06/08/2017
---


# Application.OnUndo Method (Excel)

Sets the text of the  **Undo** command and the name of the procedure that?s run if you choose the **Undo** command after running the procedure that sets this property.


## Syntax

 _expression_ . **OnUndo**( **_Text_** , **_Procedure_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text that appears with the  **Undo** command.|
| _Procedure_|Required| **String**|The name of the procedure that?s run when you choose the  **Undo** command.|

## Remarks

If a procedure doesn?t use the  **OnUndo** method, the **Undo** command is disabled.

The procedure must use the  **[OnRepeat](application-onrepeat-method-excel.md)** and **OnUndo** methods last, to prevent the repeat and undo procedures from being overwritten by subsequent actions in the procedure.


## Example

This example sets the repeat and undo procedures.


```vb
Application.OnRepeat "Repeat VB Procedure", _ 
 "Book1.xls!My_Repeat_Sub" 
Application.OnUndo "Undo VB Procedure", _ 
 "Book1.xls!My_Undo_Sub"
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

