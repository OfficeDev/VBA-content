---
title: WorksheetFunction.Norm_Inv Method (Excel)
keywords: vbaxl10.chm137371
f1_keywords:
- vbaxl10.chm137371
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Norm_Inv
ms.assetid: 0069b45f-629d-6212-18da-6954be00181f
ms.date: 06/08/2017
---


# WorksheetFunction.Norm_Inv Method (Excel)

Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation.


## Syntax

 _expression_ . **Norm_Inv**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - A probability corresponding to the normal distribution.|
| _Arg2_|Required| **Double**|Mean - The arithmetic mean of the distribution.|
| _Arg3_|Required| **Double**|Standard_dev - The standard deviation of the distribution.|

### Return Value

Double


## Remarks


- If any argument is non-numeric, NORM_INV returns the #VALUE! error value.
    
- If probability <= 0 or if probability >= 1, NORM_INV returns the #NUM! error value.
    
- If standard_dev ? 0, NORM_INV returns the #NUM! error value.
    
- If mean = 0 and standard_dev = 1, NORM_INV uses the standard normal distribution (see NORM_S_INV).
    
Given a value for probability, NORM_INV seeks that value x such that NORM_DIST(x, mean, standard_dev, TRUE) = probability. Thus, precision of NORM_INV depends on precision of NORM_DIST.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

