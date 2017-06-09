---
title: WorksheetFunction.NormSInv Method (Excel)
keywords: vbaxl10.chm137200
f1_keywords:
- vbaxl10.chm137200
ms.prod: excel
api_name:
- Excel.WorksheetFunction.NormSInv
ms.assetid: 88b209e4-3dc0-7c21-e175-55c1f133919e
ms.date: 06/08/2017
---


# WorksheetFunction.NormSInv Method (Excel)

Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of zero and a standard deviation of one.


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.For more information about the new function, see the [Norm_S_Inv](worksheetfunction-norm_s_inv-method-excel.md) method.


## Syntax

 _expression_ . **NormSInv**( **_Arg1_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - a probability corresponding to the normal distribution.|

### Return Value

Double


## Remarks


- If probability is nonnumeric, NORMSINV returns the #VALUE! error value.
    
- If probability <= 0 or if probability >= 1, NORMSINV returns the #NUM! error value.
    
Given a value for probability, NORMSINV seeks that value z such that NORMSDIST(z) = probability. Thus, precision of NORMSINV depends on precision of NORMSDIST.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

