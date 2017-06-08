---
title: WorksheetFunction.NPer Method (Excel)
keywords: vbaxl10.chm137109
f1_keywords:
- vbaxl10.chm137109
ms.prod: excel
api_name:
- Excel.WorksheetFunction.NPer
ms.assetid: ea610791-bed5-d2d3-6405-6372f46e28d8
ms.date: 06/08/2017
---


# WorksheetFunction.NPer Method (Excel)

Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate.


## Syntax

 _expression_ . **NPer**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Rate - the interest rate per period.|
| _Arg2_|Required| **Double**|Pmt - the payment made each period; it cannot change over the life of the annuity. Typically, pmt contains principal and interest but no other fees or taxes.|
| _Arg3_|Required| **Double**|Pv - the present value, or the lump-sum amount that a series of future payments is worth right now.|
| _Arg4_|Optional| **Variant**|Fv - the future value, or a cash balance you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (the future value of a loan, for example, is 0).|
| _Arg5_|Optional| **Variant**|Type - the number 0 or 1 and indicates when payments are due.|

### Return Value

Double


## Remarks





|**Set type equal to**|**If payments are due**|
|:-----|:-----|
|0 or omitted|At the end of the period|
|1|At the beginning of the period|

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

