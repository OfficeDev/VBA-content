---
title: WorksheetFunction.Rank_Avg Method (Excel)
keywords: vbaxl10.chm137379
f1_keywords:
- vbaxl10.chm137379
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Rank_Avg
ms.assetid: fd1c60c7-9a37-95b8-12d9-d1d7a42c650a
ms.date: 06/08/2017
---


# WorksheetFunction.Rank_Avg Method (Excel)

Returns the rank of a number in a list of numbers; that is its size relative to other values in the list. If more than one value has the same rank, the average rank is returned.


## Syntax

 _expression_ . **Rank_Avg**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - The number whose rank you want to find.|
| _Arg2_|Required| **Range**|Ref - An array of, or a reference to, a list of numbers. Non-numeric values in reference are ignored.|
| _Arg3_|Optional| **Variant**|Order - A number that specifies how to rank number. If the order is 0 (zero) or omitted, Microsoft Excel ranks the number as if the reference was a list sorted in descending order. If the order is any non-zero value, Microsoft Excel ranks number as if the reference were a list sorted in ascending order.|

### Return Value

Double


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

