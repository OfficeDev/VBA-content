---
title: WorksheetFunction.Mode_Mult Method (Excel)
keywords: vbaxl10.chm137368
f1_keywords:
- vbaxl10.chm137368
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Mode_Mult
ms.assetid: 13dfb3a9-2b9d-21de-29df-b3bc79b8fb59
ms.date: 06/08/2017
---


# WorksheetFunction.Mode_Mult Method (Excel)

Returns a vertical array of the most frequently occurring, or repetitive, values in an array or range of data.


## Syntax

 _expression_ . **Mode_Mult**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number1 - The first number argument for which you want to calculate the mode.|
| _Arg2 - Arg30_|Optional| **Variant**|Number2 - Number30 - Number arguments from 2 to 30 for which you want to calculate the mode. You can also use a single array or a reference to an array instead of arguments separated by commas.|

### Return Value

Variant


## Remarks




- Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- If the data set contains no duplicate data points, MODE_MULT returns the #N/A error value.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

