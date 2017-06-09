---
title: WorksheetFunction.Or Method (Excel)
keywords: vbaxl10.chm137093
f1_keywords:
- vbaxl10.chm137093
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Or
ms.assetid: 2e77bb7a-5393-2d54-c669-0c1f58a0bdfd
ms.date: 06/08/2017
---


# WorksheetFunction.Or Method (Excel)

Returns TRUE if any argument is TRUE; returns FALSE if all arguments are FALSE.


## Syntax

 _expression_ . **Or**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Logical1,logical2,... - 1 to 30 conditions you want to test that can be either TRUE or FALSE.|

### Return Value

Boolean


## Remarks




- The arguments must evaluate to logical values such as TRUE or FALSE, or in arrays or references that contain logical values.
    
- If an array or reference argument contains text or empty cells, those values are ignored.
    
- If the specified range contains no logical values, OR returns the #VALUE! error value.
    
- You can use an OR array formula to see if a value occurs in an array. To enter an array formula, press CTRL+SHIFT+ENTER.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

