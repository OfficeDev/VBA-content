---
title: WorksheetFunction.Rank_Eq Method (Excel)
keywords: vbaxl10.chm137380
f1_keywords:
- vbaxl10.chm137380
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Rank_Eq
ms.assetid: 8c2d2544-a948-7b38-e489-803cb6616066
ms.date: 06/08/2017
---


# WorksheetFunction.Rank_Eq Method (Excel)

Returns the rank of a number in a list of numbers. The rank of a number is its size relative to other values in a list. (If you were to sort the list, the rank of the number would be its position.)


## Syntax

 _expression_ . **Rank_Eq**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Number - The number whose rank you want to find.|
| _Arg2_|Required| **Range**|Ref - An array of, or a reference to, a list of numbers. Non-numeric values in reference are ignored.|
| _Arg3_|Optional| **Variant**|Order - A number that specifies how to rank the number.|

### Return Value

Double


## Remarks




- If the order is 0 (zero) or omitted, Microsoft Excel ranks the number as if the reference was a list sorted in descending order.
    
- If the order is any non-zero value, Microsoft Excel ranks the number as if the reference was a list sorted in ascending order.
    

- RANK_EQ gives duplicate numbers the same rank. However, the presence of duplicate numbers affects the ranks of subsequent numbers. For example, in a list of integers sorted in ascending order, if the number 10 appears twice and has a rank of 5, then 11 would have a rank of 7 (no number would have a rank of 6). 
    
- For some purposes you might want to use a definition of rank that takes ties into account. In the previous example, you would want a revised rank of 5.5 for the number 10. To do this, add the following correction factor to the value returned by RANK_EQ. This correction factor is appropriate both for the case where rank is computed in descending order (order = 0 or omitted) or ascending order (order = nonzero value). Correction factor for tied ranks=[COUNT(ref) + 1 ? RANK_EQ(number, ref, 0) ? RANK_EQ(number, ref, 1)]/2. In the following example, RANK_EQ(A2,A1:A5,1) equals 3. The correction factor is (5 + 1 ? 2 ? 3)/2 = 0.5 and the revised rank that takes ties into account is 3 + 0.5 = 3.5. If number occurs only once in ref, the correction factor will be 0, since RANK_EQ would not have to be adjusted for a tie. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

