---
title: WorksheetFunction.Max Method (Excel)
keywords: vbaxl10.chm137080
f1_keywords:
- vbaxl10.chm137080
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Max
ms.assetid: f0b2df1d-3b0e-2387-fa91-f8bf8cb6c4da
ms.date: 06/08/2017
---


# WorksheetFunction.Max Method (Excel)

Returns the largest value in a set of values.


## Syntax

 _expression_ . **Max**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|Number1, number2, ... - 1 to 30 numbers for which you want to find the maximum value.|

### Return Value

Double


## Remarks




- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted. 
    
- If an argument is an array or reference, only numbers in that array or reference are used. Empty cells, logical values, or text in the array or reference are ignored.
    
- If the arguments contain no numbers, MAX returns 0 (zero).
    
- Arguments that are error values or text that cannot be translated into numbers cause errors. 
    
- If you want to include logical values and text representations of numbers in a reference as part of the calculation, use the MAXA function.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

