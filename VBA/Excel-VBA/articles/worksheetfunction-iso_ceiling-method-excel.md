---
title: WorksheetFunction.ISO_Ceiling Method (Excel)
keywords: vbaxl10.chm137393
f1_keywords:
- vbaxl10.chm137393
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ISO_Ceiling
ms.assetid: e7011c98-0165-a333-6b99-b455913e8575
ms.date: 06/08/2017
---


# WorksheetFunction.ISO_Ceiling Method (Excel)

Returns a number that is rounded up to the nearest integer or to the nearest multiple of significance.


## Syntax

 _expression_ . **ISO_Ceiling**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - The value to be rounded.|
| _Arg2_|Optional| **Variant**|Significance - The optional multiple to which number is to be rounded. If significance is omitted, its default value is 1.<table><tr><th>**Note**</th></tr><tr><td>The absolute value of the multiple is used, so that the ISO_CEILING function returns the mathematical ceiling irrespective of the signs of number and significance.</td></tr></table>|

### Return Value

Double


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

