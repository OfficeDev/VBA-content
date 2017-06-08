---
title: WorksheetFunction.Dec2Hex Method (Excel)
keywords: vbaxl10.chm137265
f1_keywords:
- vbaxl10.chm137265
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Dec2Hex
ms.assetid: 32e8f754-9d67-1b99-08d3-1eee27237369
ms.date: 06/08/2017
---


# WorksheetFunction.Dec2Hex Method (Excel)

Converts a decimal number to hexadecimal.


## Syntax

 _expression_ . **Dec2Hex**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the decimal integer you want to convert. If number is negative, places is ignored and DEC2HEX returns a 10-character (40-bit) hexadecimal number in which the most significant bit is the sign bit. The remaining 39 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, DEC2HEX uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

### Return Value

String


## Remarks




- If number < -549,755,813,888 or if number > 549,755,813,887, DEC2HEX returns the #NUM! error value.
    
- If number is nonnumeric, DEC2HEX returns the #VALUE! error value.
    
- If DEC2HEX requires more than places characters, it returns the #NUM! error value.
    
- If places is not an integer, it is truncated.
    
- If places is nonnumeric, DEC2HEX returns the #VALUE! error value.
    
- If places is negative, DEC2HEX returns the #NUM! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

