---
title: WorksheetFunction.Standardize Method (Excel)
keywords: vbaxl10.chm137201
f1_keywords:
- vbaxl10.chm137201
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Standardize
ms.assetid: b268e2f8-e206-37a6-93a1-fdff7b88d4db
ms.date: 06/08/2017
---


# WorksheetFunction.Standardize Method (Excel)

Returns a normalized value from a distribution characterized by mean and standard_dev.


## Syntax

 _expression_ . **Standardize**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value you want to normalize.|
| _Arg2_|Required| **Double**|Mean - the arithmetic mean of the distribution.|
| _Arg3_|Required| **Double**|Standard_dev - the standard deviation of the distribution.|

### Return Value

Double


## Remarks




- If standard_dev ? 0, STANDARDIZE returns the #NUM! error value.
    
- The equation for the normalized value is:
![Formula](images/awfstand_ZA06051247.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

