---
title: WorksheetFunction.SumProduct Method (Excel)
keywords: vbaxl10.chm137163
f1_keywords:
- vbaxl10.chm137163
ms.prod: excel
api_name:
- Excel.WorksheetFunction.SumProduct
ms.assetid: 26562c80-1575-3019-f98c-9c974a9b863f
ms.date: 06/08/2017
---


# WorksheetFunction.SumProduct Method (Excel)

Multiplies corresponding components in the given arrays, and returns the sum of those products.


## Syntax

 _expression_ . **SumProduct**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Array1, array2, array3, ... - 2 to 30 arrays whose components you want to multiply and then add.|

### Return Value

Double


## Remarks




- The array arguments must have the same dimensions. If they do not, SUMPRODUCT returns the #VALUE! error value.
    
- SUMPRODUCT treats array entries that are not numeric as if they were zeros.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

