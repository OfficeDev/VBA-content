---
title: WorksheetFunction.Delta Method (Excel)
keywords: vbaxl10.chm137295
f1_keywords:
- vbaxl10.chm137295
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Delta
ms.assetid: a8698aa3-88cf-fe5f-be57-f01daddfa4fd
ms.date: 06/08/2017
---


# WorksheetFunction.Delta Method (Excel)

Tests whether two values are equal. Returns 1 if number1 = number2; returns 0 otherwise.


## Syntax

 _expression_ . **Delta**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number1 - the first number.|
| _Arg2_|Optional| **Variant**|Number2 - the second number. If omitted, number2 is assumed to be zero.|

### Return Value

Double


## Remarks

 Use this function to filter a set of values. For example, by summing several DELTA functions you calculate the count of equal pairs. This function is also known as the Kronecker Delta function.


- If number1 is nonnumeric, DELTA returns the #VALUE! error value.
    
- If number2 is nonnumeric, DELTA returns the #VALUE! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

