---
title: WorksheetFunction.AccrInt Method (Excel)
keywords: vbaxl10.chm137345
f1_keywords:
- vbaxl10.chm137345
ms.prod: excel
api_name:
- Excel.WorksheetFunction.AccrInt
ms.assetid: 17444208-5141-3ffe-1802-b19be0defc52
ms.date: 06/08/2017
---


# WorksheetFunction.AccrInt Method (Excel)

Returns the accrued interest for a security that pays periodic interest.


## Syntax

 _expression_ . **AccrInt**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Issue date - Security's issue date.|
| _Arg2_|Required| **Variant**|First interest - Security's first interest date.|
| _Arg3_|Required| **Variant**|Settlement - Security's settlement date|
| _Arg4_|Required| **Variant**|Rate - Security's annual coupon rate.|
| _Arg5_|Required| **Variant**|Par - Security's par value.|
| _Arg6_|Required| **Variant**|Frequency - Number of coupon payments per year.|
| _Arg7_|Optional| **Variant**|Basis - The type of day count basis to use.|

### Return Value

Double


## Remarks


 **Important**  Dates should be entered using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.

The following table describes the values that can be used for  _Arg5_ .



|**Basis**|**Day count basis**|
|:-----|:-----|
|0 or omitted|US (NASD) 30/360|
|1|Actual/actual|
|2|Actual/360|
|3|Actual/365|
|4|European 30/360|

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

