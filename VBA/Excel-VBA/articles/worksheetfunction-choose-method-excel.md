---
title: WorksheetFunction.Choose Method (Excel)
keywords: vbaxl10.chm137121
f1_keywords:
- vbaxl10.chm137121
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Choose
ms.assetid: f4031f31-2647-80fd-8458-c84f29d95e63
ms.date: 06/08/2017
---


# WorksheetFunction.Choose Method (Excel)

Uses  _Arg1_ as the index to return a value from the list of value arguments.


## Syntax

 _expression_ . **Choose**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Specifies which value argument is selected.  _Arg1_ must be a number between 1 and 29, or a formula or reference to a cell containing a number between 1 and 29.|
| _Arg2 - Arg30_|Required| **Variant**|1 to 29 value arguments from which Choose selects a value or an action to perform based on  _Arg1_. The arguments can be numbers, cell references, defined names, formulas, functions, or text.|

### Return Value

Variant


## Remarks




- If  _Arg1_ is 1, Choose returns value1; if it is 2, Choose returns value2; and so on.
    
- If  _Arg1_ is less than 1 or greater than the number of the last value in the list, Choose generates an error.
    
- If  _Arg1_ is a fraction, it is truncated to the lowest integer before being used.
    

- If  _Arg1_ is an array, every value is evaluated when Choose is evaluated.
    
- The value arguments to Choose can be range references as well as single values. For example, the formula:=SUM(Choose(2,A1:A10,B1:B10,C1:C10))evaluates to:=SUM(B1:B10)which then returns a value based on the values in the range B1:B10.The Choose function is evaluated first, returning the reference B1:B10. The SUM function is then evaluated using B1:B10, the result of the Choose function, as its argument.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

