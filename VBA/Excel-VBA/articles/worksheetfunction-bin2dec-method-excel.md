---
title: WorksheetFunction.Bin2Dec Method (Excel)
keywords: vbaxl10.chm137270
f1_keywords:
- vbaxl10.chm137270
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Bin2Dec
ms.assetid: 05a212f7-8330-002f-8bbc-f54550d1276e
ms.date: 06/08/2017
---


# WorksheetFunction.Bin2Dec Method (Excel)

Converts a binary number to decimal.


## Syntax

 _expression_ . **Bin2Dec**( **_Arg1_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The binary number you want to convert. Number cannot contain more than 10 characters (10 bits). The most significant bit of number is the sign bit. The remaining 9 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|

### Return Value

String


## Remarks

If number is not a valid binary number, or if number contains more than 10 characters (10 bits), Bin2Dec generates an error value.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

