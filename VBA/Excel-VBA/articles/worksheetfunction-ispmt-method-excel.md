---
title: WorksheetFunction.Ispmt Method (Excel)
keywords: vbaxl10.chm137244
f1_keywords:
- vbaxl10.chm137244
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Ispmt
ms.assetid: e728944b-f15e-623b-08a4-97d45d3b8473
ms.date: 06/08/2017
---


# WorksheetFunction.Ispmt Method (Excel)

Calculates the interest paid during a specific period of an investment. This function is provided for compatibility with Lotus 1-2-3.


## Syntax

 _expression_ . **Ispmt**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Rate - the interest rate for the investment.|
| _Arg2_|Required| **Double**|Per - the period for which you want to find the interest, and must be between 1 and nper.|
| _Arg3_|Required| **Double**|Nper - the total number of payment periods for the investment.|
| _Arg4_|Required| **Double**|Pv - the present value of the investment. For a loan, pv is the loan amount.|

### Return Value

Double


## Remarks




- Make sure that you are consistent about the units you use for specifying rate and nper. If you make monthly payments on a four-year loan at an annual interest rate of 12 percent, use 12%/12 for rate and 4*12 for nper. If you make annual payments on the same loan, use 12% for rate and 4 for nper.
    
- For all the arguments, the cash you pay out, such as deposits to savings or other withdrawals, is represented by negative numbers; the cash you receive, such as dividend checks and other deposits, is represented by positive numbers.
    
- For additional information about financial functions, see the PV function.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

