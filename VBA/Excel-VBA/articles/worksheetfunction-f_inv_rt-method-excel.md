---
title: WorksheetFunction.F_Inv_RT Method (Excel)
keywords: vbaxl10.chm137405
f1_keywords:
- vbaxl10.chm137405
ms.prod: excel
api_name:
- Excel.WorksheetFunction.F_Inv_RT
ms.assetid: 0852b011-ec06-ac01-cc94-993f379270bf
ms.date: 06/08/2017
---


# WorksheetFunction.F_Inv_RT Method (Excel)

Returns the inverse of the right-tailed F probability distribution. If p = F_DIST_RT(x,...), then F_INV_RT(p,...) = x.


## Syntax

 _expression_ . **F_Inv_RT**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - a probability associated with the F cumulative distribution.|
| _Arg2_|Required| **Double**|Degrees_freedom1 - the numerator degrees of freedom.|
| _Arg3_|Required| **Double**|Degrees_freedom2 - the denominator degrees of freedom.|

### Return Value

Double


## Remarks

The F distribution can be used in an F-test that compares the degree of variability in two data sets. For example, you can analyze income distributions in the United States and Canada to determine whether the two countries have a similar degree of income diversity:


- If any argument is nonnumeric, F_INV_RT returns the #VALUE! error value.
    
- If probability < 0 or probability > 1, F_INV_RT returns the #NUM! error value.
    
- If degrees_freedom1 or degrees_freedom2 is not an integer, it is truncated.
    
- If degrees_freedom1 < 1 or degrees_freedom1 ? 10^10, F_INV_RT returns the #NUM! error value.
    
- If degrees_freedom2 < 1 or degrees_freedom2 ? 10^10, F_INV_RT returns the #NUM! error value.
    
F_INV_RT can be used to return critical values from the F distribution. For example, the output of an ANOVA calculation often includes data for the F statistic, F probability, and F critical value at the 0.05 significance level. To return the critical value of F, use the significance level as the probability argument to F_INV_RT.

Given a value for probability, F_INV_RT seeks that value x such that F_DIST_RT(x, degrees_freedom1, degrees_freedom2) = probability. Thus, precision of F_INV_RT depends on precision of F_DIST_RT. F_INV_RT uses an iterative search technique. If the search has not converged after 64 iterations, the function returns the #N/A error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

