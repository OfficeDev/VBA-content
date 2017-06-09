---
title: WorksheetFunction.F_Dist_RT Method (Excel)
keywords: vbaxl10.chm137403
f1_keywords:
- vbaxl10.chm137403
ms.prod: excel
api_name:
- Excel.WorksheetFunction.F_Dist_RT
ms.assetid: 307f9afa-3e15-edce-cabb-dd96b351cdab
ms.date: 06/08/2017
---


# WorksheetFunction.F_Dist_RT Method (Excel)

Returns the right-tailed F probability distribution. You can use this function to determine whether two data sets have different degrees of diversity. For example, you can examine the test scores of men and women entering high school and determine if the variability in the females is different from that found in the males.


## Syntax

 _expression_ . **F_Dist_RT**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Degrees_freedom1 - the numerator degrees of freedom.|
| _Arg3_|Required| **Double**|Degrees_freedom2 - the denominator degrees of freedom.|

### Return Value

Double


## Remarks




- If any argument is nonnumeric, F_DIST_RT returns the #VALUE! error value.
    
- If x is negative, F_DIST_RT returns the #NUM! error value.
    
- If degrees_freedom1 or degrees_freedom2 is not an integer, it is truncated.
    
- If degrees_freedom1 < 1 or degrees_freedom1 ? 10^10, F_DIST_RT returns the #NUM! error value.
    
- If degrees_freedom2 < 1 or degrees_freedom2 ? 10^10, F_DIST_RT returns the #NUM! error value.
    
- F_DIST_RT is calculated as F_DIST_RT=P( F>x ), where F is a random variable that has an F distribution with degrees_freedom1 and degrees_freedom2 degrees of freedom.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

