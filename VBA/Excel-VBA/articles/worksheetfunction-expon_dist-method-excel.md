---
title: WorksheetFunction.Expon_Dist Method (Excel)
keywords: vbaxl10.chm137365
f1_keywords:
- vbaxl10.chm137365
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Expon_Dist
ms.assetid: 19627dab-1c33-2348-389e-18a76604b237
ms.date: 06/08/2017
---


# WorksheetFunction.Expon_Dist Method (Excel)

Returns the exponential distribution. Use EXPON_DIST to model the time between events, such as how long an automated bank teller takes to deliver cash. For example, you can use EXPON_DIST to determine the probability that the process takes at most 1 minute.


## Syntax

 _expression_ . **Expon_Dist**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|X - the value of the function.|
| _Arg2_|Required| **Double**|Lambda - the parameter value.|
| _Arg3_|Required| **Boolean**|Cumulative - a logical value that indicates which form of the exponential function to provide. If cumulative is TRUE, EXPONDIST returns the cumulative distribution function; if FALSE, it returns the probability density function.|

### Return Value

Double


## Remarks




- If x or lambda is nonnumeric, EXPON_DIST returns the #VALUE! error value.
    
- If x < 0, EXPON_DIST returns the #NUM! error value.
    
- If lambda ? 0, EXPON_DIST returns the #NUM! error value.
    
- The equation for the probability density function is:
![Formula](images/awfxpnd1_ZA06051267.gif)


    
- The equation for the cumulative distribution function is:
![Formula](images/awfxpnd2_ZA06051268.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

