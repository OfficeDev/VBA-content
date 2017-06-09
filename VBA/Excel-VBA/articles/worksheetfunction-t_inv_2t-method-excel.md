---
title: WorksheetFunction.T_Inv_2T Method (Excel)
keywords: vbaxl10.chm137387
f1_keywords:
- vbaxl10.chm137387
ms.prod: excel
api_name:
- Excel.WorksheetFunction.T_Inv_2T
ms.assetid: 5edc686a-e205-23a4-f4b8-4fabef3c9c49
ms.date: 06/08/2017
---


# WorksheetFunction.T_Inv_2T Method (Excel)

Returns the t-value of the Student t-distribution as a function of the probability and the degrees of freedom.


## Syntax

 _expression_ . **T_Inv_2T**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - The probability associated with the two-tailed Student t-distribution.|
| _Arg2_|Required| **Double**|Degrees_freedom - The number of degrees of freedom with which to characterize the distribution.|

### Return Value

Double


## Remarks




- If either argument is non-numeric, T_INV_2T returns the #VALUE! error value.
    
- If probability < 0 or if probability > 1, T_INV_2T returns the #NUM! error value.
    
- If degrees_freedom is not an integer, it is truncated.
    
- If degrees_freedom < 1, T_INV_2T returns the #NUM! error value.
    
- T_INV_2T returns that value t, such that P(|X| > t) = probability where X is a random variable that follows the t-distribution and P(|X| > t) = P(X < -t or X > t).
    
- A one-tailed t-value can be returned by replacing probability with 2*probability. For a probability of 0.05 and degrees of freedom of 10, the two-tailed value is calculated with T_INV_2T(0.05,10), which returns 2.28139. The one-tailed value for the same probability and degrees of freedom can be calculated with T_INV_2T(2*0.05,10), which returns 1.812462. Given a value for probability, T_INV_2T seeks that value x such that T_DIST_RT(x, degrees_freedom, 2) = probability. Thus, precision of T_INV_2T depends on precision of T_DIST_RT. T_INV_2T uses an iterative search technique. If the search has not converged after 100 iterations, the function returns the #N/A error value. 
    
     **Note**   In some tables, probability is described as (1-p).

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

