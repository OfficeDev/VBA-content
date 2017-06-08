---
title: WorksheetFunction.AmorDegrc Method (Excel)
keywords: vbaxl10.chm137342
f1_keywords:
- vbaxl10.chm137342
ms.prod: excel
api_name:
- Excel.WorksheetFunction.AmorDegrc
ms.assetid: 1abf060a-e68f-1155-28c3-3d257fd4fbcf
ms.date: 06/08/2017
---


# WorksheetFunction.AmorDegrc Method (Excel)

Returns the depreciation for each accounting period. This function is provided for the French accounting system.


## Syntax

 _expression_ . **AmorDegrc**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** )

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

If an asset is purchased in the middle of the accounting period, the prorated depreciation is taken into account. The method is similar to [AmorLinc](worksheetfunction-amorlinc-method-excel.md), except that a depreciation coefficient is applied in the calculation depending on the life of the assets.

The following table describes the values used in  _Arg7_ .



|**Basis**|**Date system**|
|:-----|:-----|
|0 or omitted|360 days (NASD method)|
|1|Actual|
|3|365 days in a year|
|4|360 days in a year (European method)|

- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- This function will return the depreciation until the last period of the life of the assets or until the cumulated value of depreciation is greater than the cost of the assets minus the salvage value.
    
- The depreciation coefficients are:
    

|**Life of assets (1/rate)**|**Depreciation coefficient**|
|:-----|:-----|
|Between 3 and 4 years|1.5|
|Between 5 and 6 years|2|
|More than 6 years|2.5|
- The depreciation rate will grow to 50 percent for the period preceding the last period and will grow to 100 percent for the last period.
    
- If the life of assets is between 0 (zero) and 1, 1 and 2, 2 and 3, or 4 and 5, the #NUM! error value is returned.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

