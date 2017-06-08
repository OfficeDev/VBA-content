---
title: WorksheetFunction.RoundDown Method (Excel)
keywords: vbaxl10.chm137158
f1_keywords:
- vbaxl10.chm137158
ms.prod: excel
api_name:
- Excel.WorksheetFunction.RoundDown
ms.assetid: 44b334b1-39cf-3be1-bc57-02864c29a995
ms.date: 06/08/2017
---


# WorksheetFunction.RoundDown Method (Excel)

Rounds a number down, toward zero.


## Syntax

 _expression_ . **RoundDown**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - any real number that you want rounded down.|
| _Arg2_|Required| **Double**|Num_digits - the number of digits to which you want to round number.|

### Return Value

Double


## Remarks




- ROUNDDOWN behaves like ROUND, except that it always rounds a number down.
    
- If num_digits is greater than 0 (zero), then number is rounded down to the specified number of decimal places.
    
- If num_digits is 0, then number is rounded down to the nearest integer.
    
- If num_digits is less than 0, then number is rounded down to the left of the decimal point.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

