---
title: WorksheetFunction.Rate Method (Excel)
keywords: vbaxl10.chm137111
f1_keywords:
- vbaxl10.chm137111
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Rate
ms.assetid: 5b412b46-d54a-a36a-a309-c819f2671185
ms.date: 06/08/2017
---


# WorksheetFunction.Rate Method (Excel)

Returns the interest rate per period of an annuity. RATE is calculated by iteration and can have zero or more solutions. If the successive results of RATE do not converge to within 0.0000001 after 20 iterations, RATE returns the #NUM! error value.


## Syntax

 _expression_ . **Rate**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Nper - the total number of payment periods in an annuity.|
| _Arg2_|Required| **Double**|Pmt - the payment made each period and cannot change over the life of the annuity. Typically, pmt includes principal and interest but no other fees or taxes. If pmt is omitted, you must include the fv argument.|
| _Arg3_|Required| **Double**|Pv - the present value ? the total amount that a series of future payments is worth now.|
| _Arg4_|Optional| **Variant**|Fv - the future value, or a cash balance you want to attain after the last payment is made. If fv is omitted, it is assumed to be 0 (the future value of a loan, for example, is 0).|
| _Arg5_|Optional| **Variant**|Type - the number 0 or 1 and indicates when payments are due.|
| _Arg6_|Optional| **Variant**|Guess - your guess for what the rate will be.|

### Return Value

Double


## Remarks

For a complete description of the arguments nper, pmt, pv, fv, and type, see PV.



|**Set type equal to**|**If payments are due**|
|:-----|:-----|
|0 or omitted|At the end of the period|
|1|At the beginning of the period|

- If you omit guess, it is assumed to be 10 percent.
    
- If RATE does not converge, try different values for guess. RATE usually converges if guess is between 0 and 1.
    
Make sure that you are consistent about the units you use for specifying guess and nper. If you make monthly payments on a four-year loan at 12 percent annual interest, use 12%/12 for guess and 4*12 for nper. If you make annual payments on the same loan, use 12% for guess and 4 for nper.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

