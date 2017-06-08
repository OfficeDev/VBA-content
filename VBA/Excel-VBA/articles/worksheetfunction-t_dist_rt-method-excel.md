---
title: WorksheetFunction.T_Dist_RT Method (Excel)
keywords: vbaxl10.chm137385
f1_keywords:
- vbaxl10.chm137385
ms.prod: excel
api_name:
- Excel.WorksheetFunction.T_Dist_RT
ms.assetid: 2f512dbc-09bc-c14c-c5eb-c7283afb0147
ms.date: 06/08/2017
---


# WorksheetFunction.T_Dist_RT Method (Excel)

Returns the right-tailed Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are to be computed. The t-distribution is used in the hypothesis testing of small sample data sets. Use this function in place of a table of critical values for the t-distribution.


## Syntax

 _expression_ . **T_Dist_RT**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The numeric value at which to evaluate the distribution.|
| _Arg2_|Required| **Double**|Degrees_freedom - An integer that indicates the number of degrees of freedom.|

### Return Value

Double


## Remarks




- If any argument is non-numeric, T_DIST_RT returns the #VALUE! error value.
    
- If degrees_freedom < 1, T_DIST_RT returns the #NUM! error value.
    
- The degrees_freedom and tails arguments are truncated to integers.
    
- If tails is any value other than 1 or 2, T_DIST_RT returns the #NUM! error value.
    
- If x < 0, then T_DIST_RT returns the #NUM! error value.
    
- If tails = 1, T_DIST_RT is calculated as T_DIST_RT = P( X>x ), where X is a random variable that follows the t-distribution. If tails = 2, T_DIST_RT is calculated as T_DIST_RT = P(|X| > x) = P(X > x or X < -x).
    
- Because x < 0 is not allowed, to use T_DIST_RT when x < 0, note that T_DIST_RT(-x,df) = 1 ? T_DIST_RT(x,df) = P(X > -x) and T_DIST_2T(-x,df) = T_DIST_2T(x df) = P(|X| > x).
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

