---
title: WorksheetFunction.Norm_S_Inv Method (Excel)
keywords: vbaxl10.chm137411
f1_keywords:
- vbaxl10.chm137411
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Norm_S_Inv
ms.assetid: 731c1354-2f2e-8fa8-3ced-576dd4d3ce1c
ms.date: 06/08/2017
---


# WorksheetFunction.Norm_S_Inv Method (Excel)

Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of 0 (zero) and a standard deviation of one.


## Syntax

 _expression_ . **Norm_S_Inv**( **_Arg1_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - A probability corresponding to the normal distribution.|

### Return Value

Double


## Remarks


- If probability is non-numeric, NORM_S_INV returns the #VALUE! error value.
    
- If probability < 0 or if probability > 1, NORM_S_INV returns the #NUM! error value.
    
Given a value for probability, NORM_S_INV seeks that value z such that NORM_S_DIST(z) = probability. Thus, precision of NORM_S_INV depends on precision of NORM_S_DIST. NORM_S_INV uses an iterative search technique. If the search has not converged after 100 iterations, the function returns the #N/A error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

