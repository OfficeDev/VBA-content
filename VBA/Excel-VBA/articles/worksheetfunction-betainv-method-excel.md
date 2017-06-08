---
title: WorksheetFunction.BetaInv Method (Excel)
keywords: vbaxl10.chm137176
f1_keywords:
- vbaxl10.chm137176
ms.prod: excel
api_name:
- Excel.WorksheetFunction.BetaInv
ms.assetid: 13588c71-8075-7145-915b-fd46251a3395
ms.date: 06/08/2017
---


# WorksheetFunction.BetaInv Method (Excel)

Returns the inverse of the cumulative distribution function for a specified beta distribution. That is, if probability = BetaDist(x,...), then BetaInv(probability,...) = x.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [Beta_Inv](worksheetfunction-beta_inv-method-excel.md) method.

## Syntax

 _expression_ . **BetaInv**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|A probability associated with the beta distribution.|
| _Arg2_|Required| **Double**|The Alpha parameter of the distribution.|
| _Arg3_|Required| **Double**|The Beta parameter the distribution.|
| _Arg4_|Optional| **Variant**|An optional lower bound to the interval of x.|
| _Arg5_|Optional| **Variant**|An optional upper bound to the interval of x.|

### Return Value

Double


## Remarks

 The beta distribution can be used in project planning to model probable completion times given an expected completion time and variability.


- If any argument is nonnumeric, BetaInv generates an error value.
    
- If alpha ? 0 or beta ? 0, BetaInv generates an error value.
    
- If probability ? 0 or probability > 1, BetaInv generates an error value.
    
- If you omit values for A and B, BetaInv uses the standard cumulative beta distribution, so that A = 0 and B = 1.
    
Given a value for probability, BetaInv seeks that value x such that BetaDist(x, alpha, beta, A, B) = probability. Thus, precision of BetaInv depends on precision of BetaDist. BetaInv uses an iterative search technique. If the search has not converged after 100 iterations, the function generates an error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

