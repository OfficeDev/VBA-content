---
title: WorksheetFunction.SumSq Method (Excel)
keywords: vbaxl10.chm137225
f1_keywords:
- vbaxl10.chm137225
ms.prod: excel
api_name:
- Excel.WorksheetFunction.SumSq
ms.assetid: 63e68e24-459a-d8bb-21b2-e9905a3c14ff
ms.date: 06/08/2017
---


# WorksheetFunction.SumSq Method (Excel)

Returns the sum of the squares of the arguments.


## Syntax

 _expression_ . **SumSq**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number1, number2, ... - 1 to 30 arguments for which you want the sum of the squares. You can also use a single array or a reference to an array instead of arguments separated by commas.|

### Return Value

Double


## Remarks




- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Numbers, logical values, and text representations of numbers that you type directly into the list of arguments are counted. 
    
- If an argument is an array or reference, only numbers in that array or reference are counted. Empty cells, logical values, text, or error values in the array or reference are ignored. 
    
- Arguments that are error values or text that cannot be translated into numbers cause errors. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

