---
title: WorksheetFunction.Vdb Method (Excel)
keywords: vbaxl10.chm137161
f1_keywords:
- vbaxl10.chm137161
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Vdb
ms.assetid: 601a57eb-56da-c3e5-4e6c-3029202c317d
ms.date: 06/08/2017
---


# WorksheetFunction.Vdb Method (Excel)

Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify. VDB stands for variable declining balance.


## Syntax

 _expression_ . **Vdb**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Cost - the initial cost of the asset.|
| _Arg2_|Required| **Double**|Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset). This value can be 0.|
| _Arg3_|Required| **Double**|Life - the number of periods over which the asset is depreciated (sometimes called the useful life of the asset).|
| _Arg4_|Required| **Double**|Start_period - the starting period for which you want to calculate the depreciation. Start_period must use the same units as life.|
| _Arg5_|Required| **Double**|End_period - the ending period for which you want to calculate the depreciation. End_period must use the same units as life.|
| _Arg6_|Optional| **Variant**|Factor - the rate at which the balance declines. If factor is omitted, it is assumed to be 2 (the double-declining balance method). Change factor if you do not want to use the double-declining balance method. For a description of the double-declining balance method, see DDB.|
| _Arg7_|Optional| **Variant**|No_switch - a logical value specifying whether to switch to straight-line depreciation when depreciation is greater than the declining balance calculation.|

### Return Value

Double


## Remarks


- If no_switch is TRUE, Microsoft Excel does not switch to straight-line depreciation even when the depreciation is greater than the declining balance calculation.
    
- If no_switch is FALSE or omitted, Excel switches to straight-line depreciation when depreciation is greater than the declining balance calculation.
    
All arguments except no_switch must be positive numbers.


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

