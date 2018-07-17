---
title: WorksheetFunction.GammaInv Method (Excel)
keywords: vbaxl10.chm137191
f1_keywords:
- vbaxl10.chm137191
ms.prod: excel
api_name:
- Excel.WorksheetFunction.GammaInv
ms.assetid: 7b0e95f4-dd58-50f2-89ec-22bfa932766f
ms.date: 06/08/2017
---


# WorksheetFunction.GammaInv Method (Excel)

Returns the inverse of the gamma cumulative distribution. If p = GAMMADIST(x,...), then GAMMAINV(p,...) = x.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [Gamma_Inv](worksheetfunction-gammainv-method-excel.md) method.

## Syntax

 _expression_ . **GammaInv**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - the probability associated with the gamma distribution.|
| _Arg2_|Required| **Double**|Alpha - a parameter to the distribution.|
| _Arg3_|Required| **Double**|Beta - a parameter to the distribution. If beta = 1, GAMMAINV returns the standard gamma distribution.|

### Return Value

Double


## Remarks

You can use this function to study a variable whose distribution may be skewed.


- If any argument is text, GAMMAINV returns the #VALUE! error value.
    
- If probability < 0 or probability > 1, GAMMAINV returns the #NUM! error value.
    
- If alpha ? 0 or if beta ? 0, GAMMAINV returns the #NUM! error value.
    
Given a value for probability, GAMMAINV seeks that value x such that GAMMADIST(x, alpha, beta, TRUE) = probability. Thus, precision of GAMMAINV depends on precision of GAMMADIST. GAMMAINV uses an iterative search technique. If the search has not converged after 64 iterations, the function returns the #N/A error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

