---
title: WorksheetFunction.Beta_Dist Method (Excel)
keywords: vbaxl10.chm137396
f1_keywords:
- vbaxl10.chm137396
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Beta_Dist
ms.assetid: f691e4b0-3021-6a7e-3306-af7b5cb3720b
ms.date: 06/08/2017
---


# WorksheetFunction.Beta_Dist Method (Excel)

Returns the beta cumulative distribution function.


## Syntax

 _expression_ . **Beta_Dist**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The value between A and B at which to evaluate the function.|
| _Arg2_|Required| **Double**|The Alpha parameter of the distribution.|
| _Arg3_|Required| **Double**|The Beta parameter of the distribution.|
| _Arg4_|Optional| **Variant**|Cumulative - a logical value that determines the form of the function. If cumulative is TRUE, BETA.DIST returns the cumulative distribution function; if FALSE, it returns the probability density function.|
| _Arg5_|Optional| **Variant**|An optional lower bound to the interval of x.|
| _Arg6_|Optional| **Variant**|An optional upper bound to the interval of x.|

### Return Value

Double


## Remarks

The beta distribution is commonly used to study variation in the percentage of something across samples, such as the fraction of the day people spend watching television:


- If any argument is nonnumeric, Beta_Dist returns the #VALUE! error value.
    
- If alpha ? 0 or beta ? 0, Beta_Dist generates an error value.
    
- If x < A, x > B, or A = B, Beta_Dist generates an error value.
    
- If you omit values for A and B (lower and upper bound), Beta_Dist uses the standard cumulative beta distribution, so that A = 0 and B = 1.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

