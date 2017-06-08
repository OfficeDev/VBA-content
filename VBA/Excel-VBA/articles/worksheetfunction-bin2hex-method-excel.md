---
title: WorksheetFunction.Bin2Hex Method (Excel)
keywords: vbaxl10.chm137272
f1_keywords:
- vbaxl10.chm137272
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Bin2Hex
ms.assetid: dc1dbd75-89c5-47de-386f-121e23c4b294
ms.date: 06/08/2017
---


# WorksheetFunction.Bin2Hex Method (Excel)

Converts a binary number to hexadecimal.


## Syntax

 _expression_ . **Bin2Hex**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The binary number you want to convert. Number cannot contain more than 10 characters (10 bits). The most significant bit of number is the sign bit. The remaining 9 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|The number of characters to use. If places is omitted, Bin2Hex uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

### Return Value

String


## Remarks




- If number is not a valid binary number, or if number contains more than 10 characters (10 bits), Bin2Hex generates an error.
    
- If number is negative, Bin2Hex ignores places and returns a 10-character hexadecimal number.
    
- If Bin2Hex requires more than places characters, it generates an error.
    
- If places is not an integer, it is truncated.
    
- If places is nonnumeric, Bin2Hex generates an error.
    
- If places is negative, Bin2Hex generates an error.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

