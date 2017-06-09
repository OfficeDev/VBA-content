---
title: WorksheetFunction.Subtotal Method (Excel)
keywords: vbaxl10.chm137240
f1_keywords:
- vbaxl10.chm137240
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Subtotal
ms.assetid: ec854287-1b12-8195-6b30-9101140d642e
ms.date: 06/08/2017
---


# WorksheetFunction.Subtotal Method (Excel)

Creates subtotals. 


## Syntax

 _expression_ . **Subtotal**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|A number indicating the aggregation function to be used by the method.|
| _Arg2_|Required| **Range**|The first  **[Range](range-object-excel.md)** object for which a subtotal is to be calculated.|
| _Arg3-Arg30_|Optional| **Variant**|Subsequent  **[Range](range-object-excel.md)** objects for which a subtotal is to be calculated.|

### Return Value

A Double value that represents the subtotal.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

