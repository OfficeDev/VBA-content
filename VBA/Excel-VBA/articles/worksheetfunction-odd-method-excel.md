---
title: WorksheetFunction.Odd Method (Excel)
keywords: vbaxl10.chm137202
f1_keywords:
- vbaxl10.chm137202
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Odd
ms.assetid: 28a30d51-ba7b-f7b4-55a5-39b85f7f4cd7
ms.date: 06/08/2017
---


# WorksheetFunction.Odd Method (Excel)

Returns number rounded up to the nearest odd integer.


## Syntax

 _expression_ . **Odd**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the value to round.|

### Return Value

Double


## Remarks




- If number is nonnumeric, ODD returns the #VALUE! error value.
    
- Regardless of the sign of number, a value is rounded up when adjusted away from zero. If number is an odd integer, no rounding occurs.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

