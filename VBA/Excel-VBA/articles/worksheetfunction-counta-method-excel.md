---
title: WorksheetFunction.CountA Method (Excel)
keywords: vbaxl10.chm137142
f1_keywords:
- vbaxl10.chm137142
ms.prod: excel
api_name:
- Excel.WorksheetFunction.CountA
ms.assetid: b3d8662b-a886-daf8-2ce0-763017fbcd94
ms.date: 06/08/2017
---


# WorksheetFunction.CountA Method (Excel)

Counts the number of cells that are not empty and the values within the list of arguments.


## Syntax

 _expression_ . **CountA**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|1 to 30 arguments representing the values you want to count.|

### Return Value

Double


## Remarks

 Use CountA to count the number of cells that contain data in a range or array.


- A value is any type of information, including error values and empty text (""). A value does not include empty cells. 
    
- If an argument is an array or reference, only values in that array or reference are used. Empty cells and text values in the array or reference are ignored.
    
- If you do not need to count logical values, text, or error values, use the Count function.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

