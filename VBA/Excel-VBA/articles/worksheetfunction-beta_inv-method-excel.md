---
title: WorksheetFunction.Beta_Inv Method (Excel)
keywords: vbaxl10.chm137397
f1_keywords:
- vbaxl10.chm137397
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Beta_Inv
ms.assetid: f652b2b8-a966-1b1e-bfcd-1554923c1740
ms.date: 06/08/2017
---


# WorksheetFunction.Beta_Inv Method (Excel)

Returns the inverse of the cumulative distribution function for a specified beta distribution. That is, if probability = Beta_Dist(x,...), then Beta_Inv(probability,...) = x.


## Syntax

 _expression_ . **Beta_Inv**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** )

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

 The beta distribution can be used in project planning to model probable completion times given an expected completion time and variability:


- If any argument is nonnumeric, Beta_Inv generates an error value.
    
- If alpha ? 0 or beta ? 0, Beta_Inv generates an error value.
    
- If probability ? 0 or probability > 1, Beta_Inv generates an error value.
    
- If you omit values for A and B (lower and upper bound), Beta_Inv uses the standard cumulative beta distribution, so that A = 0 and B = 1.
    
Given a value for probability, Beta_Inv seeks that value x such that Beta_Dist(x, alpha, beta, TRUE, A, B) = probability. Thus, precision of Beta_Inv depends on precision of Beta_Dist. Beta_Inv uses an iterative search technique.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

