---
title: WorksheetFunction.AmorLinc Method (Excel)
keywords: vbaxl10.chm137343
f1_keywords:
- vbaxl10.chm137343
ms.prod: excel
api_name:
- Excel.WorksheetFunction.AmorLinc
ms.assetid: 9daa4b32-2364-fcfc-13e8-c3e7689700d4
ms.date: 06/08/2017
---


# WorksheetFunction.AmorLinc Method (Excel)

Returns the depreciation for each accounting period. This function is provided for the French accounting system.


## Syntax

 _expression_ . **AmorLinc**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The cost of the asset.|
| _Arg2_|Required| **Variant**|The date of the purchase of the asset.|
| _Arg3_|Required| **Variant**|The date of the end of the first period.|
| _Arg4_|Required| **Variant**|The salvage value at the end of the life of the asset.|
| _Arg5_|Required| **Variant**|The period.|
| _Arg6_|Required| **Variant**|The rate of depreciation.|
| _Arg7_|Optional| **Variant**|The year basis to be used.|

### Return Value

Double


## Remarks

If an asset is purchased in the middle of the accounting period, the prorated depreciation is taken into account.

The following table describes values used for  _Arg7_ .



|**Basis**|**Date system**|
|:-----|:-----|
|0 or omitted|360 days (NASD method)|
|1|Actual|
|3|365 days in a year|
|4|360 days in a year (European method)|

 **Important**  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.

Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default. 


 **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

