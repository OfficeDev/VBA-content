---
title: WorksheetFunction.ChiInv Method (Excel)
keywords: vbaxl10.chm137179
f1_keywords:
- vbaxl10.chm137179
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ChiInv
ms.assetid: 10b89d77-bc9f-80b0-dc31-f90c50f7e580
ms.date: 06/08/2017
---


# WorksheetFunction.ChiInv Method (Excel)

Returns the inverse of the one-tailed probability of the chi-squared distribution.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new functions, see the [ChiSq_Inv_RT](worksheetfunction-chisq_inv_rt-method-excel.md) and[ChiSq_Inv](worksheetfunction-chisq_inv-method-excel.md) methods.

## Syntax

 _expression_ . **ChiInv**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|A probability associated with the chi-squared distribution.|
| _Arg2_|Required| **Double**|The number of degrees of freedom.|

### Return Value

Double


## Remarks

 If probability = ChiDist(x,...), then ChiInv(probability,...) = x. Use this function to compare observed results with expected ones in order to decide whether your original hypothesis is valid.


- If either argument is nonnumeric, ChiInv generates an error.
    
- If probability < 0 or probability > 1, ChiInv generates an error.
    
- If degrees_freedom is not an integer, it is truncated.
    
- If degrees_freedom < 1 or degrees_freedom ? 10^10, ChiInv generates an error.
    
Given a value for probability, ChiInv seeks that value x such that ChiDist(x, degrees_freedom) = probability. Thus, precision of ChiInv depends on precision of ChiDist. ChiInv uses an iterative search technique. If the search has not converged after 64 iterations, the function generates an error.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

