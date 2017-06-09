---
title: WorksheetFunction.Rept Method (Excel)
keywords: vbaxl10.chm137091
f1_keywords:
- vbaxl10.chm137091
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Rept
ms.assetid: acf1bf30-3722-79f3-c3ab-42c3f14aa435
ms.date: 06/08/2017
---


# WorksheetFunction.Rept Method (Excel)

Repeats text a given number of times. Use REPT to fill a cell with a number of instances of a text string.


## Syntax

 _expression_ . **Rept**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Text - the text you want to repeat.|
| _Arg2_|Required| **Double**|Number_times - a positive number specifying the number of times to repeat text.|

### Return Value

String


## Remarks




- If number_times is 0 (zero), REPT returns "" (empty text).
    
- If number_times is not an integer, it is truncated.
    
- The result of the REPT function cannot be longer than 32,767 characters, or REPT returns #VALUE!.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

