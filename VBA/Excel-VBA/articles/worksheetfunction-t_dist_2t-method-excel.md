---
title: WorksheetFunction.T_Dist_2T Method (Excel)
keywords: vbaxl10.chm137384
f1_keywords:
- vbaxl10.chm137384
ms.prod: excel
api_name:
- Excel.WorksheetFunction.T_Dist_2T
ms.assetid: e4927634-d94c-5bcc-7bef-ad35a315bc69
ms.date: 06/08/2017
---


# WorksheetFunction.T_Dist_2T Method (Excel)

Returns the two-tailed Student t-distribution.


## Syntax

 _expression_ . **T_Dist_2T**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The numeric value at which to evaluate the distribution.|
| _Arg2_|Required| **Double**|Deg_freedom - An integer that indicates the number of degrees of freedom.|

### Return Value

Double


## Remarks




- If any argument is non-numeric, T_DIST_2T returns the #VALUE! error value.
    
- If deg_freedom < 1, T_DIST_2T returns the #NUM! error value.
    
- If x < 0, then T_DIST_2T returns the #NUM! error value.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

