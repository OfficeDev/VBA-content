---
title: WorksheetFunction.SeriesSum Method (Excel)
keywords: vbaxl10.chm137291
f1_keywords:
- vbaxl10.chm137291
ms.prod: excel
api_name:
- Excel.WorksheetFunction.SeriesSum
ms.assetid: 096faaa8-4bd3-fd61-4442-b29785a93c7c
ms.date: 06/08/2017
---


# WorksheetFunction.SeriesSum Method (Excel)

Returns the sum of a power series based on the formula:
![Formula](images/awfsrssm_ZA06051246.gif)




## Syntax

 _expression_ . **SeriesSum**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|X - the input value to the power series.|
| _Arg2_|Required| **Variant**|N - the initial power to which you want to raise x.|
| _Arg3_|Required| **Variant**|M - the step by which to increase n for each term in the series.|
| _Arg4_|Required| **Variant**|Coefficients - a set of coefficients by which each successive power of x is multiplied. The number of values in coefficients determines the number of terms in the power series. For example, if there are three values in coefficients, then there will be three terms in the power series.|

### Return Value

Double


## Remarks

If any argument is nonnumeric, SERIESSUM returns the #VALUE! error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

