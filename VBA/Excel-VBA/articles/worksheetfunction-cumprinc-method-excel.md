---
title: WorksheetFunction.CumPrinc Method (Excel)
keywords: vbaxl10.chm137323
f1_keywords:
- vbaxl10.chm137323
ms.prod: excel
api_name:
- Excel.WorksheetFunction.CumPrinc
ms.assetid: 6e561b97-97e2-11d8-0240-86fe374044ca
ms.date: 06/08/2017
---


# WorksheetFunction.CumPrinc Method (Excel)

Returns the cumulative principal paid on a loan between start_period and end_period.


## Syntax

 _expression_ . **CumPrinc**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The interest rate.|
| _Arg2_|Required| **Variant**| The total number of payment periods.|
| _Arg3_|Required| **Variant**| The present value.|
| _Arg4_|Required| **Variant**|The first period in the calculation. Payment periods are numbered beginning with 1.|
| _Arg5_|Required| **Variant**|The last period in the calculation.|
| _Arg6_|Required| **Variant**|The timing of the payment.|

### Return Value

Double


## Remarks

The following tables lists values used in  _Arg6_ .



|**Type**|**Timing**|
|:-----|:-----|
|0 (zero)|Payment at the end of the period|
|1|Payment at the beginning of the period|

- Make sure that you are consistent about the units you use for specifying rate and nper. If you make monthly payments on a four-year loan at an annual interest rate of 12 percent, use 12%/12 for rate and 4*12 for  _Arg2_ . If you make annual payments on the same loan, use 12% for rate and 4 for _Arg2_ .
    
-  _Arg2_ , _Arg4_ , _Arg5_ , and type are truncated to integers.
    
- If rate ? 0,  _Arg2_ ? 0, or _Arg3_ ? 0, CumPrinc generates an error.
    
- If  _Arg4_ < 1, _Arg5_ < 1, or _Arg4_ > _Arg5_ , CumPrinc generates an error.
    
- If  _Arg6_ is any number other than 0 or 1, CumPrinc generates an error.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

