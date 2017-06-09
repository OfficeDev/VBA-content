---
title: WorksheetFunction.Hex2Bin Method (Excel)
keywords: vbaxl10.chm137261
f1_keywords:
- vbaxl10.chm137261
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Hex2Bin
ms.assetid: 373a8eb3-7f60-f03e-04f8-ebb5f0de47f6
ms.date: 06/08/2017
---


# WorksheetFunction.Hex2Bin Method (Excel)

Converts a hexadecimal number to binary.


## Syntax

 _expression_ . **Hex2Bin**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the hexadecimal number you want to convert. Number cannot contain more than 10 characters. The most significant bit of number is the sign bit (40th bit from the right). The remaining 9 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, HEX2BIN uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

### Return Value

String


## Remarks




- If number is negative, HEX2BIN ignores places and returns a 10-character binary number.
    
- If number is negative, it cannot be less than FFFFFFFE00, and if number is positive, it cannot be greater than 1FF.
    
- If number is not a valid hexadecimal number, HEX2BIN returns the #NUM! error value.
    
- If HEX2BIN requires more than places characters, it returns the #NUM! error value.
    
- If places is not an integer, it is truncated.
    
- If places is nonnumeric, HEX2BIN returns the #VALUE! error value.
    
- If places is negative, HEX2BIN returns the #NUM! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

