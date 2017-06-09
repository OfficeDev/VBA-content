---
title: WorksheetFunction.ChiSq_Dist Method (Excel)
keywords: vbaxl10.chm137398
f1_keywords:
- vbaxl10.chm137398
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ChiSq_Dist
ms.assetid: be655878-fdb2-7b04-0a9b-6d39652b7e77
ms.date: 06/08/2017
---


# WorksheetFunction.ChiSq_Dist Method (Excel)

Returns the chi-squared distribution.


## Syntax

 _expression_ . **ChiSq_Dist**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The value at which you want to evaluate the distribution.|
| _Arg2_|Required| **Double**|Deg_freedom - The number of degrees of freedom.|
| _Arg3_|Optional| **Variant**|Cumulative - A logical value that determines the form of the function. If cumulative is TRUE, CHISQ_DIST returns the cumulative distribution function; if FALSE, it returns the probability density function. |

### Return Value

Double


## Remarks




- If any argument is nonnumeric, CHISQ_DIST returns the #VALUE! error value. 
    
- If x is negative, CHISQ_DIST returns the #NUM! error value. 
    
- If deg_freedom is not an integer, it is truncated. 
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

