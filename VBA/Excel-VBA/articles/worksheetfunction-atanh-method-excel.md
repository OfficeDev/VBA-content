---
title: WorksheetFunction.Atanh Method (Excel)
keywords: vbaxl10.chm137169
f1_keywords:
- vbaxl10.chm137169
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Atanh
ms.assetid: 4399ebf8-5eff-9ec0-421e-1fe3f5fdc5c1
ms.date: 06/08/2017
---


# WorksheetFunction.Atanh Method (Excel)

Returns the inverse hyperbolic tangent of a number. Number must be between -1 and 1 (excluding -1 and 1). 


## Syntax

 _expression_ . **Atanh**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Any real number between 1 and -1.|

### Return Value

Double


## Remarks

The inverse hyperbolic tangent is the value whose hyperbolic tangent is  _Arg1_ , so Atanh(Tanh(number)) equals _Arg1_ .


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

