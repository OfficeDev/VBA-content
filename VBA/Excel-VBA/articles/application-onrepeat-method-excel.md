---
title: Application.OnRepeat Method (Excel)
keywords: vbaxl10.chm133181
f1_keywords:
- vbaxl10.chm133181
ms.prod: excel
api_name:
- Excel.Application.OnRepeat
ms.assetid: 7d535e14-c779-af87-60eb-68ec8e651459
ms.date: 06/08/2017
---


# Application.OnRepeat Method (Excel)

Sets the  **Repeat** item and the name of the procedure that will run if you choose the **Repeat** command after running the procedure that sets this property.


## Syntax

 _expression_ . **OnRepeat**( **_Text_** , **_Procedure_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text that appears with the  **Repeat** command.|
| _Procedure_|Required| **String**|The name of the procedure that will be run when you choose the  **Repeat** command.|

## Remarks

If a procedure doesn?t use the  **OnRepeat** method, the **Repeat** command repeats procedure that was run most recently.

The procedure must use the  **OnRepeat** and **OnUndo** methods last, to prevent the repeat and undo procedures from being overwritten by subsequent actions in the procedure.


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

