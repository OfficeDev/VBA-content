---
title: WorksheetFunction.Rank Method (Excel)
keywords: vbaxl10.chm137159
f1_keywords:
- vbaxl10.chm137159
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Rank
ms.assetid: e75cabc4-1d97-b8fd-4e7d-3b12ab6a53c5
ms.date: 06/08/2017
---


# WorksheetFunction.Rank Method (Excel)

Returns the rank of a number in a list of numbers. The rank of a number is its size relative to other values in a list. (If you were to sort the list, the rank of the number would be its position.)


## 


 **Important**  This function has been replaced with one or more new functions that may provide improved accuracy and whose names better reflect their usage. This function is still available for compatibility with earlier versions of Excel. However, if backward compatibility is not required, you should consider using the new functions from now on, because they more accurately describe their functionality.

For more information about the new functions, see the [Rank_Eq](worksheetfunction-rank_eq-method-excel.md) and[Rank_Avg](worksheetfunction-rank_avg-method-excel.md) methods.


## Syntax

 _expression_ . **Rank**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - the number whose rank you want to find.|
| _Arg2_|Required| **Range**|Ref - an array of, or a reference to, a list of numbers. Nonnumeric values in ref are ignored.|
| _Arg3_|Optional| **Variant**|Order - a number specifying how to rank number.|

### Return Value

Double


## Remarks




- If order is 0 (zero) or omitted, Microsoft Excel ranks number as if ref were a list sorted in descending order.
    
- If order is any nonzero value, Microsoft Excel ranks number as if ref were a list sorted in ascending order.
    

- RANK gives duplicate numbers the same rank. However, the presence of duplicate numbers affects the ranks of subsequent numbers. For example, in a list of integers sorted in ascending order, if the number 10 appears twice and has a rank of 5, then 11 would have a rank of 7 (no number would have a rank of 6). 
    
- For some purposes one might want to use a definition of rank that takes ties into account. In the previous example, one would want a revised rank of 5.5 for the number 10. This can be done by adding the following correction factor to the value returned by RANK. This correction factor is appropriate both for the case where rank is computed in descending order (order = 0 or omitted) or ascending order (order = nonzero value). Correction factor for tied ranks=[COUNT(ref) + 1 ? RANK(number, ref, 0) ? RANK(number, ref, 1)]/2. In the following example, RANK(A2,A1:A5,1) equals 3. The correction factor is (5 + 1 ? 2 ? 3)/2 = 0.5 and the revised rank that takes ties into account is 3 + 0.5 = 3.5. If number occurs only once in ref, the correction factor will be 0, since RANK would not have to be adjusted for a tie. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

