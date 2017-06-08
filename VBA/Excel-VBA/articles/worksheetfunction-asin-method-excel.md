---
title: WorksheetFunction.Asin Method (Excel)
keywords: vbaxl10.chm137119
f1_keywords:
- vbaxl10.chm137119
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Asin
ms.assetid: 24195cf6-d762-169d-fb7d-aa15dfbfd152
ms.date: 06/08/2017
---


# WorksheetFunction.Asin Method (Excel)

Returns the arcsine, or inverse sine, of a number. The arcsine is the angle whose sine is  _Arg1_. The returned angle is given in radians in the range -pi/2 to pi/2.


## Syntax

 _expression_ . **Asin**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The sine of the angle you want and must be from -1 to 1.|

### Return Value

Double


## Remarks

To express the arcsine in degrees, multiply the result by 180/PI( ) or use the [Degrees](worksheetfunction-degrees-method-excel.md) method.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

