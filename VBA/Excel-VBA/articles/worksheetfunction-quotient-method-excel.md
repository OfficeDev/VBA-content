---
title: WorksheetFunction.Quotient Method (Excel)
keywords: vbaxl10.chm137294
f1_keywords:
- vbaxl10.chm137294
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Quotient
ms.assetid: 33a057f8-dbb7-0f0e-fabd-ebdd4d471159
ms.date: 06/08/2017
---


# WorksheetFunction.Quotient Method (Excel)

Returns the integer portion of a division. Use this function when you want to discard the remainder of a division.


## Syntax

 _expression_ . **Quotient**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Numerator - the dividend.|
| _Arg2_|Required| **Variant**|Denominator - the divisor.|

### Return Value

Double


## Remarks

If either argument is nonnumeric, QUOTIENT returns the #VALUE! error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

