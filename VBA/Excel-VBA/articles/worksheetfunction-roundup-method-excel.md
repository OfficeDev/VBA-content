---
title: WorksheetFunction.RoundUp Method (Excel)
keywords: vbaxl10.chm137157
f1_keywords:
- vbaxl10.chm137157
ms.prod: excel
api_name:
- Excel.WorksheetFunction.RoundUp
ms.assetid: daff9e6a-5ed8-b502-24c1-c4ffe01d2d0f
ms.date: 06/08/2017
---


# WorksheetFunction.RoundUp Method (Excel)

Rounds a number up, away from 0 (zero).


## Syntax

 _expression_ . **RoundUp**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - any real number that you want rounded up.|
| _Arg2_|Required| **Double**|Num_digits - the number of digits to which you want to round number.|

### Return Value

Double


## Remarks




- ROUNDUP behaves like ROUND, except that it always rounds a number up.
    
- If num_digits is greater than 0 (zero), then number is rounded up to the specified number of decimal places.
    
- If num_digits is 0, then number is rounded up to the nearest integer.
    
- If num_digits is less than 0, then number is rounded up to the left of the decimal point.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

