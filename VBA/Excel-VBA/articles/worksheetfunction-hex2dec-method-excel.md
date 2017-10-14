---
title: WorksheetFunction.Hex2Dec Method (Excel)
keywords: vbaxl10.chm137262
f1_keywords:
- vbaxl10.chm137262
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Hex2Dec
ms.assetid: e2e0614c-583e-8a1f-b852-683c119d5a5a
ms.date: 06/08/2017
---


# WorksheetFunction.Hex2Dec Method (Excel)

Converts a hexadecimal number to decimal.


## Syntax

 _expression_ . **Hex2Dec**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the hexadecimal number you want to convert. Number cannot contain more than 10 characters (40 bits). The most significant bit of number is the sign bit. The remaining 39 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|

### Return Value

String


## Remarks

If number is not a valid hexadecimal number, HEX2DEC returns the #NUM! error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

