---
title: WorksheetFunction.Floor_Precise Method (Excel)
keywords: vbaxl10.chm137420
f1_keywords:
- vbaxl10.chm137420
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Floor_Precise
ms.assetid: 003159fa-9397-a648-67aa-5751c93e3c92
ms.date: 06/08/2017
---


# WorksheetFunction.Floor_Precise Method (Excel)

Rounds the specified number to the nearest multiple of significance.


## Syntax

 _expression_ . **Floor_Precise**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the numeric value you want to round.|
| _Arg2_|Optional| **Variant**|Significance - the multiple to which you want to round.|

### Return Value

Double


## Remarks

Depending on the sign of the number and significance arguments, the  **Floor_Precise** method rounds either away from or towards zero.



|**Sign ( _Arg1_ / _Arg2_ )**|**Rounding**|
|:-----|:-----|
|-/-|Rounds away from zero.|
|+/+|Rounds toward zero.|
|-/+|Rounds away from zero.|
|+/-|Rounds toward zero.|

- If either argument is nonnumeric, the  **Floor_Precise** method generates an error.
    
-  If number is an exact multiple of significance, no rounding occurs.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

