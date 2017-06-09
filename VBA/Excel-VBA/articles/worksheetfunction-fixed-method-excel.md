---
title: WorksheetFunction.Fixed Method (Excel)
keywords: vbaxl10.chm137084
f1_keywords:
- vbaxl10.chm137084
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Fixed
ms.assetid: befc65b2-0216-dbd7-e376-edbcbfe532c5
ms.date: 06/08/2017
---


# WorksheetFunction.Fixed Method (Excel)

Rounds a number to the specified number of decimals, formats the number in decimal format using a period and commas, and returns the result as text.


## Syntax

 _expression_ . **Fixed**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the number you want to round and convert to text.|
| _Arg2_|Optional| **Variant**|Decimals - the number of digits to the right of the decimal point.|
| _Arg3_|Optional| **Variant**|No_commas - a logical value that, if TRUE, prevents FIXED from including commas in the returned text.|

### Return Value

String


## Remarks




- Numbers in Microsoft Excel can never have more than 15 significant digits, but decimals can be as large as 127.
    
- If decimals is negative, number is rounded to the left of the decimal point.
    
- If you omit decimals, it is assumed to be 2.
    
- If no_commas is FALSE or omitted, then the returned text includes commas as usual.
    
- The major difference between formatting a cell containing a number with the  **Cells** command ( **Format** menu) and formatting a number directly with the FIXED function is that FIXED converts its result to text. A number formatted with the **Cells** command is still a number.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

