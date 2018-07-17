---
title: WorksheetFunction.Nominal Method (Excel)
keywords: vbaxl10.chm137321
f1_keywords:
- vbaxl10.chm137321
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Nominal
ms.assetid: 4ba61f10-233b-400b-76e1-90147fd7f503
ms.date: 06/08/2017
---


# WorksheetFunction.Nominal Method (Excel)

Returns the nominal annual interest rate, given the effective rate and the number of compounding periods per year.


## Syntax

 _expression_ . **Nominal**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Effect_rate - the effective interest rate.|
| _Arg2_|Required| **Variant**|Npery - the number of compounding periods per year.|

### Return Value

Double


## Remarks




- Npery is truncated to an integer.
    
- If either argument is nonnumeric, NOMINAL returns the #VALUE! error value.
    
- If effect_rate ? 0 or if npery < 1, NOMINAL returns the #NUM! error value.
    
- NOMINAL is related to EFFECT as shown in the following equation:
![Formula](images/awfnomnl_ZA06051211.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

