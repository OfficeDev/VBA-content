---
title: WorksheetFunction.DollarDe Method (Excel)
keywords: vbaxl10.chm137319
f1_keywords:
- vbaxl10.chm137319
ms.prod: excel
api_name:
- Excel.WorksheetFunction.DollarDe
ms.assetid: 626462e2-3415-1552-eb7e-8f7bb5346852
ms.date: 06/08/2017
---


# WorksheetFunction.DollarDe Method (Excel)

Converts a dollar price expressed as a fraction into a dollar price expressed as a decimal number. Use DOLLARDE to convert fractional dollar numbers, such as securities prices, to decimal numbers.


## Syntax

 _expression_ . **DollarDe**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Fractional_dollar - is a number expressed as a fraction.|
| _Arg2_|Required| **Variant**|Fraction - the integer to use in the denominator of the fraction.|

### Return Value

Double


## Remarks




- If fraction is not an integer, it is truncated.
    
- If fraction is less than 0, DOLLARDE returns the #NUM! error value.
    
- If fraction is 0, DOLLARDE returns the #DIV/0! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

