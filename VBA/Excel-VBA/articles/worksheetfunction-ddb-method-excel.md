---
title: WorksheetFunction.Ddb Method (Excel)
keywords: vbaxl10.chm137135
f1_keywords:
- vbaxl10.chm137135
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Ddb
ms.assetid: 7514f3b3-ca21-ec3f-28c5-f34281fc1a1f
ms.date: 06/08/2017
---


# WorksheetFunction.Ddb Method (Excel)

Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify.


## Syntax

 _expression_ . **Ddb**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Cost - the initial cost of the asset.|
| _Arg2_|Required| **Double**|Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset). This value can be 0.|
| _Arg3_|Required| **Double**|Life - the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset).|
| _Arg4_|Required| **Double**|Period - the period for which you want to calculate the depreciation. Period must use the same units as life.|
| _Arg5_|Optional| **Variant**|Factor - the rate at which the balance declines. If factor is omitted, it is assumed to be 2 (the double-declining balance method).|

### Return Value

Double


## Remarks


 **Important**  All five arguments must be positive numbers.


- The double-declining balance method computes depreciation at an accelerated rate. Depreciation is highest in the first period and decreases in successive periods. DDB uses the following formula to calculate depreciation for a period: `Min( (cost - total depreciation from prior periods) * (factor/life), (cost - salvage - total depreciation from prior periods) )`
    
- Change factor if you do not want to use the double-declining balance method.
    
- Use the VDB function if you want to switch to the straight-line depreciation method when depreciation is greater than the declining balance calculation.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

