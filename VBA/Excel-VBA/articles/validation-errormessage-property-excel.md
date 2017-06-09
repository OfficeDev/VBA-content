---
title: Validation.ErrorMessage Property (Excel)
keywords: vbaxl10.chm532079
f1_keywords:
- vbaxl10.chm532079
ms.prod: excel
api_name:
- Excel.Validation.ErrorMessage
ms.assetid: e5708bb8-7929-9e69-f020-567c4f6fc67d
ms.date: 06/08/2017
---


# Validation.ErrorMessage Property (Excel)

Returns or sets the data validation error message. Read/write  **String** .


## Syntax

 _expression_ . **ErrorMessage**

 _expression_ A variable that represents a **Validation** object.


## Example

This example adds data validation to cell E5 and specifies both the input and error messages.


```vb
With Range("e5").Validation 
 .Add Type:=xlValidateWholeNumber, _ 
 AlertStyle:= xlValidAlertStop, _ 
 Operator:=xlBetween, Formula1:="5", Formula2:="10" 
 .InputTitle = "Integers" 
 .ErrorTitle = "Integers" 
 .InputMessage = "Enter an integer from five to ten" 
 .ErrorMessage = "You must enter a number from five to ten" 
End With
```


## See also


#### Concepts


[Validation Object](validation-object-excel.md)

