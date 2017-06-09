---
title: Validation.InputMessage Property (Excel)
keywords: vbaxl10.chm532081
f1_keywords:
- vbaxl10.chm532081
ms.prod: excel
api_name:
- Excel.Validation.InputMessage
ms.assetid: cef219c7-4fb2-128c-b091-170f63f70a98
ms.date: 06/08/2017
---


# Validation.InputMessage Property (Excel)

Returns or sets the data validation input message. Read/write  **String** .


## Syntax

 _expression_ . **InputMessage**

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

