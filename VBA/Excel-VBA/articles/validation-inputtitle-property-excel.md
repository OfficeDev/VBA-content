---
title: Validation.InputTitle Property (Excel)
keywords: vbaxl10.chm532082
f1_keywords:
- vbaxl10.chm532082
ms.prod: excel
api_name:
- Excel.Validation.InputTitle
ms.assetid: 77e6bb8b-1fc2-084c-69b7-31b07f8145d3
ms.date: 06/08/2017
---


# Validation.InputTitle Property (Excel)

Returns or sets the title of the data-validation input dialog box. Read/write  **String**. Limited to 32 characters.


## Syntax

 _expression_ . **InputTitle**

 _expression_ A variable that represents a **Validation** object.


## Example

This example turns on data validation for cell E5.


```vb
With Range("e5").Validation 
 .Add xlValidateWholeNumber, _ 
 xlValidAlertInformation, xlBetween, "5", "10" 
 .InputTitle = "Integers" 
 .ErrorTitle = "Integers" 
 .InputMessage = "Enter an integer from five to ten" 
 .ErrorMessage = "You must enter a number from five to ten" 
End With
```


## See also


#### Concepts


[Validation Object](validation-object-excel.md)

