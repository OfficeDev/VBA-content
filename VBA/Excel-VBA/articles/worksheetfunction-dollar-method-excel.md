---
title: WorksheetFunction.Dollar Method (Excel)
keywords: vbaxl10.chm137083
f1_keywords:
- vbaxl10.chm137083
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Dollar
ms.assetid: 246988c8-568a-640b-affb-fd1cd8907889
ms.date: 06/08/2017
---


# WorksheetFunction.Dollar Method (Excel)

The function described in this Help topic converts a number to text format and applies a currency symbol. The name of the function (and the symbol that it applies) depends upon your language settings.


## Syntax

 _expression_ . **Dollar**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - a number, a reference to a cell containing a number, or a formula that evaluates to a number.|
| _Arg2_|Optional| **Variant**|Decimals - the number of digits to the right of the decimal point. If decimals is negative, number is rounded to the left of the decimal point. If you omit decimals, it is assumed to be 2.|

### Return Value

String


## Remarks

The major difference between formatting a cell that contains a number with the  **Cells** command ( **Format** menu) and formatting a number directly with the DOLLAR function is that DOLLAR converts its result to text. A number formatted with the **Cells** command is still a number. You can continue to use numbers formatted with DOLLAR in formulas, because Microsoft Excel converts numbers entered as text values to numbers when it calculates.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

