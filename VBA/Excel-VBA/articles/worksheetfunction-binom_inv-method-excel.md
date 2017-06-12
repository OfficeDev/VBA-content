---
title: WorksheetFunction.Binom_Inv Method (Excel)
keywords: vbaxl10.chm137415
f1_keywords:
- vbaxl10.chm137415
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Binom_Inv
ms.assetid: 30af29b2-fc97-656b-d703-905caf7fcbb5
ms.date: 06/08/2017
---


# WorksheetFunction.Binom_Inv Method (Excel)

Returns the inverse of the individual term binomial distribution probability.


## Syntax

 _expression_ . **Binom_Inv**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Trials - the number of Bernoulli trials.|
| _Arg2_|Required| **Double**|Probability_s - the probability of a success on each trial.|
| _Arg3_|Required| **Double**|Alpha - the criterion value.|

### Return Value

Double


## Remarks




- If Trials, Probability_s, or Alpha is nonnumeric, the  **Binom_Inv** method generates an error.
    
- If Trials is not an integer, it is truncated.
    
- If Trials < 0, the  **Binom_Inv** method generates an error.
    
- If Probability_s < 0 or Probability_s > 1, the  **Binom_Inv** method generates an error.
    
- If Alpha < 0 or Alpha > 1, the  **Binom_Inv** method generates an error.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

