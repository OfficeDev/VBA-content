---
title: WorksheetFunction.PercentRank_Exc Method (Excel)
keywords: vbaxl10.chm137374
f1_keywords:
- vbaxl10.chm137374
ms.prod: excel
api_name:
- Excel.WorksheetFunction.PercentRank_Exc
ms.assetid: 7d887f36-769c-2d02-c1cf-321d84a2bb56
ms.date: 06/08/2017
---


# WorksheetFunction.PercentRank_Exc Method (Excel)

Returns the rank of a value in a data set as a percentage (0..1, exclusive) of the data set.


## Syntax

 _expression_ . **PercentRank_Exc**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array - The array or range of data with numeric values that defines relative standing.|
| _Arg2_|Required| **Double**|X - The value for which you want to know the rank.|
| _Arg3_|Optional| **Variant**|Significance - A value that identifies the number of significant digits for the returned percentage value. If omitted, PERCENTRANK.EXC uses three digits (0.xxx).|

### Return Value

Double


## Remarks




- If array is empty, PERCENTRANK_EXC returns the #NUM! error value.
    
- If significance < 1, PERCENTRANK_EXC returns the #NUM! error value.
    
- If x does not match one of the values in array, PERCENTRANK_EXC interpolates to return the correct percentage rank.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

