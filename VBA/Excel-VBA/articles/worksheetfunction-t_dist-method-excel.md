---
title: WorksheetFunction.T_Dist Method (Excel)
keywords: vbaxl10.chm137383
f1_keywords:
- vbaxl10.chm137383
ms.prod: excel
api_name:
- Excel.WorksheetFunction.T_Dist
ms.assetid: a6b7ad29-d00f-f779-9531-4d05bc216036
ms.date: 06/08/2017
---


# WorksheetFunction.T_Dist Method (Excel)

Returns a Student t-distribution where a numeric value (x) is a calculated value of t for which the Percentage Points are computed.


## Syntax

 _expression_ . **T_Dist**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The numeric value at which to evaluate the distribution.|
| _Arg2_|Required| **Double**|Deg_freedom - An integer that indicates the number of degrees of freedom.|
| _Arg3_|Required| **Boolean**|Cumulative - A logical value that determines the form of the function. If cumulative is TRUE, T_DIST returns the cumulative distribution function; if FALSE, it returns the probability density function.|

### Return Value

Double


## Remarks




- If any argument is nonnumeric, T_DIST returns the #VALUE! error value.
    
- If deg_freedom < 1, T_DIST returns the #NUM! error value.
    
- If x < 0, then T_DIST returns the #NUM! error value.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

