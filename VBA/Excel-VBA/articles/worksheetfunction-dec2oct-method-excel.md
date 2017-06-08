---
title: WorksheetFunction.Dec2Oct Method (Excel)
keywords: vbaxl10.chm137266
f1_keywords:
- vbaxl10.chm137266
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Dec2Oct
ms.assetid: 2aac7d4d-57ef-0d8f-1432-62e98ddc1c41
ms.date: 06/08/2017
---


# WorksheetFunction.Dec2Oct Method (Excel)

Converts a decimal number to octal.


## Syntax

 _expression_ . **Dec2Oct**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Number - the decimal integer you want to convert. If number is negative, places is ignored and DEC2OCT returns a 10-character (30-bit) octal number in which the most significant bit is the sign bit. The remaining 29 bits are magnitude bits. Negative numbers are represented using two's-complement notation.|
| _Arg2_|Optional| **Variant**|Places - the number of characters to use. If places is omitted, DEC2OCT uses the minimum number of characters necessary. Places is useful for padding the return value with leading 0s (zeros).|

### Return Value

String


## Remarks




- If number < -536,870,912 or if number > 536,870,911, DEC2OCT returns the #NUM! error value.
    
- If number is nonnumeric, DEC2OCT returns the #VALUE! error value.
    
- If DEC2OCT requires more than places characters, it returns the #NUM! error value.
    
- If places is not an integer, it is truncated.
    
- If places is nonnumeric, DEC2OCT returns the #VALUE! error value.
    
- If places is negative, DEC2OCT returns the #NUM! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

