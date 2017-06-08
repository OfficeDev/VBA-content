---
title: WorksheetFunction.MIrr Method (Excel)
keywords: vbaxl10.chm137112
f1_keywords:
- vbaxl10.chm137112
ms.prod: excel
api_name:
- Excel.WorksheetFunction.MIrr
ms.assetid: 5c11a445-0b5a-ce7f-d881-e5f85cdf648a
ms.date: 06/08/2017
---


# WorksheetFunction.MIrr Method (Excel)

Returns the modified internal rate of return for a series of periodic cash flows. MIRR considers both the cost of the investment and the interest received on reinvestment of cash.


## Syntax

 _expression_ . **MIrr**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Values - an array or a reference to cells that contain numbers. These numbers represent a series of payments (negative values) and income (positive values) occurring at regular periods.|
| _Arg2_|Required| **Double**|Finance_rate - the interest rate you pay on the money used in the cash flows.|
| _Arg3_|Required| **Double**|Reinvest_rate - the interest rate you receive on the cash flows as you reinvest them.|

### Return Value

Double


## Remarks




- Values must contain at least one positive value and one negative value to calculate the modified internal rate of return. Otherwise, MIRR returns the #DIV/0! error value.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    

- MIRR uses the order of values to interpret the order of cash flows. Be sure to enter your payment and income values in the sequence you want and with the correct signs (positive values for cash received, negative values for cash paid).
    
- If n is the number of cash flows in values, frate is the finance_rate, and rrate is the reinvest_rate, then the formula for MIRR is:
![Formula](images/awfmirr_ZA06051207.gif)


    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

