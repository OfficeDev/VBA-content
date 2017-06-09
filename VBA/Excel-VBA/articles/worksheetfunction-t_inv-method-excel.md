---
title: WorksheetFunction.T_Inv Method (Excel)
keywords: vbaxl10.chm137386
f1_keywords:
- vbaxl10.chm137386
ms.prod: excel
api_name:
- Excel.WorksheetFunction.T_Inv
ms.assetid: 0104e8a3-0beb-69bb-d9b5-20c319d740f6
ms.date: 06/08/2017
---


# WorksheetFunction.T_Inv Method (Excel)

Returns the left-tailed inverse of the Student t-distribution.


## Syntax

 _expression_ . **T_Dist_2T**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - The probability associated with the Student t-distribution.|
| _Arg2_|Required| **Double**|Deg_freedom - The number of degrees of freedom with which to characterize the distribution.|

### Return Value

Double


## Remarks




- If either argument is non-numeric, T_INV returns the #VALUE! error value.
    
- If probability < 0 or if probability > 1, T_INV returns the #NUM! error value.
    
- If deg_freedom is not an integer, it is truncated.
    
- If deg_freedom < 1, T_INV returns the #NUM! error value.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

