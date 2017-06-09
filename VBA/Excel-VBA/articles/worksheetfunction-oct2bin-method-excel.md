---
title: WorksheetFunction.Oct2Bin Method (Excel)
keywords: vbaxl10.chm137267
f1_keywords:
- vbaxl10.chm137267
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Oct2Bin
ms.assetid: a11c26e2-1320-f76f-547e-fa9e0ac20087
ms.date: 06/08/2017
---


# WorksheetFunction.Oct2Bin Method (Excel)

Converts an octal number to binary.


## Syntax

 _expression_ . **Oct2Bin**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the octal number you want to convert. Number may not contain more than 10 characters. The most significant bit of number is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, OCT2BIN uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

### Return Value

String


## Remarks




- If number is negative, OCT2BIN ignores places and returns a 10-character binary number.
    
- If number is negative, it cannot be less than 7777777000, and if number is positive, it cannot be greater than 777.
    
- If number is not a valid octal number, OCT2BIN returns the #NUM! error value.
    
- If OCT2BIN requires more than places characters, it returns the #NUM! error value.
    
- If places is not an integer, it is truncated.
    
- If places is nonnumeric, OCT2BIN returns the #VALUE! error value.
    
- If places is negative, OCT2BIN returns the #NUM! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

