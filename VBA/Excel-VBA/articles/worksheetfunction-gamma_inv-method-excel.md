---
title: WorksheetFunction.Gamma_Inv Method (Excel)
keywords: vbaxl10.chm137367
f1_keywords:
- vbaxl10.chm137367
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Gamma_Inv
ms.assetid: a13d812f-9e27-e5e0-0226-7b0f5c666a91
ms.date: 06/08/2017
---


# WorksheetFunction.Gamma_Inv Method (Excel)

Returns the inverse of the gamma cumulative distribution. If p = GAMMA_DIST(x,...), then GAMMA_INV(p,...) = x.


## Syntax

 _expression_ . **Gamma_Inv**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - the probability associated with the gamma distribution.|
| _Arg2_|Required| **Double**|Alpha - a parameter to the distribution.|
| _Arg3_|Required| **Double**|Beta - a parameter to the distribution. If beta = 1, GAMMA_INV returns the standard gamma distribution.|

### Return Value

Double


## Remarks

You can use this function to study a variable whose distribution may be skewed:


- If any argument is text, GAMMA_INV returns the #VALUE! error value.
    
- If probability < 0 or probability > 1, GAMMA_INV returns the #NUM! error value.
    
- If alpha ? 0 or if beta ? 0, GAMMA_INV returns the #NUM! error value.
    
Given a value for probability, GAMMA_INV seeks that value x such that GAMMA_DIST(x, alpha, beta, TRUE) = probability. Thus, precision of GAMMA_INV depends on precision of GAMMA_DIST. GAMMA_INV uses an iterative search technique. If the search has not converged after 100 iterations, the function returns the #N/A error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

