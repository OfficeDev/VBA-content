---
title: WorksheetFunction.ChiSq_Inv_RT Method (Excel)
keywords: vbaxl10.chm137401
f1_keywords:
- vbaxl10.chm137401
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ChiSq_Inv_RT
ms.assetid: 4c92ac86-6f3b-6bdb-cae9-5790db659e2a
ms.date: 06/08/2017
---


# WorksheetFunction.ChiSq_Inv_RT Method (Excel)

Returns the inverse of the right-tailed probability of the chi-squared distribution.


## Syntax

 _expression_ . **ChiSq_Inv_RT**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|A probability associated with the chi-squared distribution.|
| _Arg2_|Required| **Double**|The number of degrees of freedom.|

### Return Value

Double


## Remarks

 If probability = ChiSq_Dist_RT(x,...), then ChiSq_Inv_RT(probability,...) = x. Use this function to compare observed results with expected ones in order to decide whether your original hypothesis is valid:


- If either argument is nonnumeric, ChiSq_Inv_RT generates an error.
    
- If probability < 0 or probability > 1, ChiSq_Inv_RT generates an error.
    
- If degrees_freedom is not an integer, it is truncated.
    
- If degrees_freedom < 1 or degrees_freedom ? 10^10, ChiSq_Inv_RT generates an error.
    
Given a value for probability, ChiSq_Inv_RT seeks that value x such that ChiSq_Dist_RT(x, degrees_freedom) = probability. Thus, precision of ChiSq_Inv_RT depends on precision of ChiSq_Dist_RT. ChiSq_Inv_RT uses an iterative search technique. If the search has not converged after 64 iterations, the function generates an error.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

