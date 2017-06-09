---
title: WorksheetFunction.Acos Method (Excel)
keywords: vbaxl10.chm137120
f1_keywords:
- vbaxl10.chm137120
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Acos
ms.assetid: 76954fdf-5aa0-de8d-1f7c-4184ebc472f4
ms.date: 06/08/2017
---


# WorksheetFunction.Acos Method (Excel)

Returns the arccosine, or inverse cosine, of a number. The arccosine is the angle whose cosine is  _Arg1_. The returned angle is given in radians in the range 0 (zero) to pi.


## Syntax

 _expression_ . **Acos**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The cosine of the angle you want and must be from -1 to 1.|

### Return Value

Double


## Remarks

If you want to convert the result from radians to degrees, multiply it by 180/PI() or use the [Degrees](worksheetfunction-degrees-method-excel.md) method.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

