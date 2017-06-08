---
title: WorksheetFunction.Oct2Dec Method (Excel)
keywords: vbaxl10.chm137269
f1_keywords:
- vbaxl10.chm137269
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Oct2Dec
ms.assetid: 08694db9-136b-9bfe-7939-436f4773bffb
ms.date: 06/08/2017
---


# WorksheetFunction.Oct2Dec Method (Excel)

Converts an octal number to decimal.


## Syntax

 _expression_ . **Oct2Dec**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the octal number you want to convert. Number may not contain more than 10 octal characters (30 bits). The most significant bit of number is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|

### Return Value

String


## Remarks

If number is not a valid octal number, OCT2DEC returns the #NUM! error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

