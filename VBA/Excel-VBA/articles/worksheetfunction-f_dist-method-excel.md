---
title: WorksheetFunction.F_Dist Method (Excel)
keywords: vbaxl10.chm137402
f1_keywords:
- vbaxl10.chm137402
ms.prod: excel
api_name:
- Excel.WorksheetFunction.F_Dist
ms.assetid: 7b18fd63-120f-fddf-a20a-00d4182778a5
ms.date: 06/08/2017
---


# WorksheetFunction.F_Dist Method (Excel)

Returns the F probability distribution.


## Syntax

 _expression_ . **F_Dist**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The value at which to evaluate the function.|
| _Arg2_|Required| **Double**|Deg_freedom1 - The numerator degrees of freedom.|
| _Arg3_|Required| **Double**|Deg_freedom2 - The denominator degrees of freedom.|
| _Arg4_|Optional| **Variant**|Cumulative - A logical value that determines the form of the function. If cumulative is TRUE, F_DIST returns the cumulative distribution function; if FALSE, it returns the probability density function.|

### Return Value

Double


## Remarks




- If any argument is nonnumeric, F_DIST returns the #VALUE! error value.
    
- If x is negative, F_DIST returns the #NUM! error value.
    
- If deg_freedom1 or deg_freedom2 is not an integer, it is truncated.
    
-  If deg_freedom1 < 1, F_DIST returns the #NUM! error value.
    
- If deg_freedom < 1, F_DIST returns the #NUM! error value.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

