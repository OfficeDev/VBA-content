---
title: Validation.ShowInput Property (Excel)
keywords: vbaxl10.chm532088
f1_keywords:
- vbaxl10.chm532088
ms.prod: excel
api_name:
- Excel.Validation.ShowInput
ms.assetid: 8760c403-c982-4cbd-6211-cb8b1c83bc91
ms.date: 06/08/2017
---


# Validation.ShowInput Property (Excel)

 **True** if the data validation input message will be displayed whenever the user selects a cell in the data validation range. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowInput**

 _expression_ A variable that represents a **Validation** object.


## Example

This example adds data validation to cell A10. The input value must be from 5 through 10; if the user types invalid data, an error message is displayed but no input message is displayed.


```vb
With Worksheets(1).Range("A10").Validation 
 .Add Type:=xlValidateWholeNumber, _ 
 AlertStyle:=xlValidAlertStop, _ 
 Operator:=xlBetween, Formula1:="5", _ 
 Formula2:="10" 
 .ErrorMessage = "value must be between 5 and 10" 
 .ShowInput = False 
 .ShowError = True 
End With
```


## See also


#### Concepts


[Validation Object](validation-object-excel.md)

