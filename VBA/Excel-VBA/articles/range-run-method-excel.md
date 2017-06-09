---
title: Range.Run Method (Excel)
keywords: vbaxl10.chm144192
f1_keywords:
- vbaxl10.chm144192
ms.prod: excel
api_name:
- Excel.Range.Run
ms.assetid: b7a0480a-9f10-8aad-6592-3cbde72720cd
ms.date: 06/08/2017
---


# Range.Run Method (Excel)

Runs the Microsoft Excel macro at this location. The range must be on a macro sheet.


## Syntax

 _expression_ . **Run**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg30_|Optional| **Variant**|The arguments that should be passed to the function.|

### Return Value

Variant


## Remarks

You cannot use named arguments with this method. Arguments must be passed by position.

The  **Run** method returns whatever the called macro returns. Objects passed as arguments to the macro are converted to values (by applying the **Value** property to the object). This means that you cannot pass objects to macros by using the **Run** method.


## See also


#### Concepts


[Range Object](range-object-excel.md)

