---
title: WorksheetFunction.TInv Method (Excel)
keywords: vbaxl10.chm137236
f1_keywords:
- vbaxl10.chm137236
ms.prod: excel
api_name:
- Excel.WorksheetFunction.TInv
ms.assetid: a336dfb7-cc7c-5e67-dd36-9e4d5e96f850
ms.date: 06/08/2017
---


# WorksheetFunction.TInv Method (Excel)

Returns the t-value of the Student's t-distribution as a function of the probability and the degrees of freedom.


## 


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.For more information about the new functions, see the [T_Inv](worksheetfunction-t_inv-method-excel.md) and[T_Inv_2T](worksheetfunction-t_inv_2t-method-excel.md) methods.


## Syntax

 _expression_ . **TInv**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - the probability associated with the two-tailed Student's t-distribution.|
| _Arg2_|Required| **Double**|Degrees_freedom - the number of degrees of freedom with which to characterize the distribution.|

### Return Value

Double


## Remarks




- If either argument is nonnumeric, TINV returns the #VALUE! error value.
    
- If probability < 0 or if probability > 1, TINV returns the #NUM! error value.
    
- If degrees_freedom is not an integer, it is truncated.
    
- If degrees_freedom < 1, TINV returns the #NUM! error value.
    
- TINV returns that value t, such that P(|X| > t) = probability where X is a random variable that follows the t-distribution and P(|X| > t) = P(X < -t or X > t).
    
- A one-tailed t-value can be returned by replacing probability with 2*probability. For a probability of 0.05 and degrees of freedom of 10, the two-tailed value is calculated with TINV(0.05,10), which returns 2.228139. The one-tailed value for the same probability and degrees of freedom can be calculated with TINV(2*0.05,10), which returns 1.812462. Given a value for probability, TINV seeks that value x such that TDIST(x, degrees_freedom, 2) = probability. Thus, precision of TINV depends on precision of TDIST. 
    
     **Note**   In some tables, probability is described as (1-p).

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

