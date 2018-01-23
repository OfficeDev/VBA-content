---
title: WorksheetFunction.LogNorm_Inv Method (Excel)
keywords: vbaxl10.chm137408
f1_keywords:
- vbaxl10.chm137408
ms.prod: excel
api_name:
- Excel.WorksheetFunction.LogNorm_Inv
ms.assetid: d8a3c416-c2c4-dc57-e1f0-1d05e9fec2a1
ms.date: 06/08/2017
---


# WorksheetFunction.LogNorm_Inv Method (Excel)

Returns the inverse of the lognormal cumulative distribution function. Use the lognormal distribution to analyze logarithmically transformed data.


## Syntax

 _expression_ . **LogNorm_Inv**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Probability - A probability associated with the lognormal distribution.|
| _Arg2_|Required| **Double**|Mean - The mean of ln(x).|
| _Arg3_|Required| **Double**|Standard_dev - The standard deviation of ln(x).|

### Return Value

Double


## Remarks




- If any argument is nonnumeric, LOGNORM_INV returns the #VALUE! error value.
    
- If probability <= 0 or probability >= 1, LOGNORM_INV returns the #NUM! error value.
    
- If standard_dev <= 0, LOGNORM_INV returns the #NUM! error value.
    
- The inverse of the lognormal distribution function is:
![Inverse of the lognormal distribution function.](images/LOGNORM_INV_ZA10390997.jpg)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

