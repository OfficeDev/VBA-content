---
title: WorksheetFunction.DollarFr Method (Excel)
keywords: vbaxl10.chm137320
f1_keywords:
- vbaxl10.chm137320
ms.prod: excel
api_name:
- Excel.WorksheetFunction.DollarFr
ms.assetid: a024cc74-605f-7ac5-77f9-7368f8b22f8c
ms.date: 06/08/2017
---


# WorksheetFunction.DollarFr Method (Excel)

Converts a dollar price expressed as a decimal number into a dollar price expressed as a fraction. Use DOLLARFR to convert decimal numbers to fractional dollar numbers, such as securities prices.


## Syntax

 _expression_ . **DollarFr**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Decimal_dollar - is a decimal number.|
| _Arg2_|Required| **Variant**|Fraction - the integer to use in the denominator of a fraction.|

### Return Value

Double


## Remarks




- If fraction is not an integer, it is truncated.
    
- If fraction is less than 0, DOLLARFR returns the #NUM! error value.
    
- If fraction is 0, DOLLARFR returns the #DIV/0! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

