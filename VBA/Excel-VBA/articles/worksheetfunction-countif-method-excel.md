---
title: WorksheetFunction.CountIf Method (Excel)
keywords: vbaxl10.chm137242
f1_keywords:
- vbaxl10.chm137242
ms.prod: excel
api_name:
- Excel.WorksheetFunction.CountIf
ms.assetid: d0251b63-cc9e-a58c-1862-adbd58004126
ms.date: 06/08/2017
---


# WorksheetFunction.CountIf Method (Excel)

Counts the number of cells within a range that meet the given criteria.


## Syntax

 _expression_ . **CountIf**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|The range of cells from which you want to count cells.|
| _Arg2_|Required| **Variant**|The criteria in the form of a number, expression, cell reference, or text that defines which cells will be counted. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.|

### Return Value

Double


## Remarks




- You can use the wildcard characters, question mark (?) and asterisk (*), for the criteria. A question mark matches any single character; an asterisk matches any sequence of characters. If you want to find an actual question mark or asterisk, type a tilde (~) before the character.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

