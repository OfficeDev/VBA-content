---
title: Validation.ShowError Property (Excel)
keywords: vbaxl10.chm532087
f1_keywords:
- vbaxl10.chm532087
ms.prod: excel
api_name:
- Excel.Validation.ShowError
ms.assetid: 19f7e431-6a6a-d8ed-98fe-c931cfb95498
ms.date: 06/08/2017
---


# Validation.ShowError Property (Excel)

 **True** if the data validation error message will be displayed whenever the user enters invalid data. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowError**

 _expression_ A variable that represents a **Validation** object.


## Example

This example adds data validation to cell A10 on worksheet one. The input value must be from 5 through 10; if the user types invalid data, an error message is displayed but no input message is displayed.


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

