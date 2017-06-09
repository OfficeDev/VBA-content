---
title: WorksheetFunction.Syd Method (Excel)
keywords: vbaxl10.chm137134
f1_keywords:
- vbaxl10.chm137134
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Syd
ms.assetid: 5d63250b-5002-b159-e1b8-f47752b85e73
ms.date: 06/08/2017
---


# WorksheetFunction.Syd Method (Excel)

Returns the sum-of-years' digits depreciation of an asset for a specified period.


## Syntax

 _expression_ . **Syd**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Cost - the initial cost of the asset.|
| _Arg2_|Required| **Double**|Salvage - the value at the end of the depreciation (sometimes called the salvage value of the asset).|
| _Arg3_|Required| **Double**|Life - the number of periods over which the asset is depreciated (sometimes called the useful life of the asset).|
| _Arg4_|Required| **Double**|Per - the period and must use the same units as life.|

### Return Value

Double


## Remarks




- SYD is calculated as follows:
![Formula](images/awfsyd_ZA06051253.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

