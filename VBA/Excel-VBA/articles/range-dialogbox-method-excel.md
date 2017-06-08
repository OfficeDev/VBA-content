---
title: Range.DialogBox Method (Excel)
keywords: vbaxl10.chm144117
f1_keywords:
- vbaxl10.chm144117
ms.prod: excel
api_name:
- Excel.Range.DialogBox
ms.assetid: d2d4a677-bd6a-910d-ff53-f95585f40925
ms.date: 06/08/2017
---


# Range.DialogBox Method (Excel)

Displays a dialog box defined by a dialog box definition table on a Microsoft Excel 4.0 macro sheet. Returns the number of the chosen control, or returns  **False** if the user clicks the **Cancel** button.


## Syntax

 _expression_ . **DialogBox**

 _expression_ A variable that represents a **Range** object.


### Return Value

Variant


## Remarks

 The **Range** must refer to a dialog box definition table on a Microsoft Excel 4.0 macro sheet.


## Example

This example runs a Microsoft Excel 4.0 dialog box and then displays the return value in a message box. The  `dialogRange` variable refers to the dialog box definition table on the Microsoft Excel 4.0 macro sheet named "Macro1."


```vb
Set dialogRange = Excel4MacroSheets("Macro1").Range("myDialogBox") 
result = dialogRange.DialogBox 
MsgBox result
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

