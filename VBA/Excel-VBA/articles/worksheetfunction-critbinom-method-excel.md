---
title: WorksheetFunction.CritBinom Method (Excel)
keywords: vbaxl10.chm137182
f1_keywords:
- vbaxl10.chm137182
ms.prod: excel
api_name:
- Excel.WorksheetFunction.CritBinom
ms.assetid: df9bb77f-b3b5-3e2b-d0b1-f42aabe9c14a
ms.date: 06/08/2017
---


# WorksheetFunction.CritBinom Method (Excel)

Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value.


## Syntax

 _expression_ . **CritBinom**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|The number of Bernoulli trials.|
| _Arg2_|Required| **Double**|The probability of a success on each trial.|
| _Arg3_|Required| **Double**|The criterion value.|

### Return Value

Double


## Remarks

 Use this function for quality assurance applications. For example, use CritBinom to determine the greatest number of defective parts that are allowed to come off an assembly line run without rejecting the entire lot.


- If any argument is nonnumeric, CritBinom generates an error.
    
- If trials is not an integer, it is truncated.
    
- If trials < 0, CritBinom generates an error.
    
- If probability_s is < 0 or probability_s > 1, CritBinom generates an error.
    
- If alpha < 0 or alpha > 1, CritBinom generates an error.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

