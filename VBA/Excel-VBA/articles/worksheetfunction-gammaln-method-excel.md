---
title: WorksheetFunction.GammaLn Method (Excel)
keywords: vbaxl10.chm137175
f1_keywords:
- vbaxl10.chm137175
ms.prod: excel
api_name:
- Excel.WorksheetFunction.GammaLn
ms.assetid: 89dbd9e8-cd88-405d-8f88-351b4dc39f02
ms.date: 06/08/2017
---


# WorksheetFunction.GammaLn Method (Excel)

Returns the natural logarithm of the gamma function, ?(x).


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new function, see the [GammaLn_Precise](worksheetfunction-gammaln_precise-method-excel.md) method.

## Syntax

 _expression_ . **GammaLn**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value for which you want to calculate GAMMALN.|

### Return Value

Double


## Remarks




- If x is nonnumeric, GAMMALN returns the #VALUE! error value.
    
- If x ? 0, GAMMALN returns the #NUM! error value.
    
- The number e raised to the GAMMALN(i) power, where i is an integer, returns the same result as (i - 1)!.
    
- GAMMALN is calculated as follows:
![Formula](images/awfgamm1_ZA06051143.gif)where: 
![Formula](images/awfgamm2_ZA06051144.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

