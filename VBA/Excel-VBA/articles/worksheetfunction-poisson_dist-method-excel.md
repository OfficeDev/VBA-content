---
title: WorksheetFunction.Poisson_Dist Method (Excel)
keywords: vbaxl10.chm137376
f1_keywords:
- vbaxl10.chm137376
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Poisson_Dist
ms.assetid: 338193e2-6b52-417a-97b9-d6ba12a1275e
ms.date: 06/08/2017
---


# WorksheetFunction.Poisson_Dist Method (Excel)

Returns the Poisson distribution. A common application of the Poisson distribution is predicting the number of events over a specific time, such as the number of cars arriving at a toll plaza in one minute.


## Syntax

 _expression_ . **Poisson_Dist**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - The number of events.|
| _Arg2_|Required| **Double**|Mean - The expected numeric value.|
| _Arg3_|Required| **Boolean**|Cumulative - A logical value that determines the form of the probability distribution returned. If cumulative is TRUE, POISSON_DIST returns the cumulative Poisson probability that the number of random events occurring will be between 0 (zero) and x inclusive; if FALSE, it returns the Poisson probability mass function that the number of events occurring will be exactly x.|

### Return Value

Double


## Remarks




- If x is not an integer, it is truncated.
    
- If x or mean is nonnumeric, POISSON_DIST returns the #VALUE! error value.
    
- If x < 0, POISSON_DIST returns the #NUM! error value.
    
- If mean ? 0, POISSON_DIST returns the #NUM! error value.
    
- POISSON_DIST is calculated as follows. For cumulative = FALSE: 
![POISSON_DIST equation for cumulative= FALSE](images/POISSON_DIST_FALSE_ZA10390998.jpg)For cumulative = TRUE: 
![POISSON_DIST equation for cumulative= TRUE](images/POISSON_DIST_TRUE_ZA10390999.jpg)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

