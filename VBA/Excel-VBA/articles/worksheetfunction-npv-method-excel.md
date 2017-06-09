---
title: WorksheetFunction.Npv Method (Excel)
keywords: vbaxl10.chm137081
f1_keywords:
- vbaxl10.chm137081
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Npv
ms.assetid: c191e00d-20e1-1648-efe9-73fab00f28db
ms.date: 06/08/2017
---


# WorksheetFunction.Npv Method (Excel)

Calculates the net present value of an investment by using a discount rate and a series of future payments (negative values) and income (positive values).


## Syntax

 _expression_ . **Npv**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Rate - the rate of discount over the length of one period.|
| _Arg2 - Arg30_|Required| **Variant**|Value1, value2, ... - 1 to 29 arguments representing the payments and income.|

### Return Value

Double


## Remarks




- Value1, value2, ... must be equally spaced in time and occur at the end of each period.
    
- NPV uses the order of value1, value2, ... to interpret the order of cash flows. Be sure to enter your payment and income values in the correct sequence.
    
- Arguments that are numbers, empty cells, logical values, or text representations of numbers are counted; arguments that are error values or text that cannot be translated into numbers are ignored.
    
- If an argument is an array or reference, only numbers in that array or reference are counted. Empty cells, logical values, text, or error values in the array or reference are ignored.
    

- The NPV investment begins one period before the date of the value1 cash flow and ends with the last cash flow in the list. The NPV calculation is based on future cash flows. If your first cash flow occurs at the beginning of the first period, the first value must be added to the NPV result, not included in the values arguments. For more information, see the examples below.
    
- If n is the number of cash flows in the list of values, the formula for NPV is:
![Formula](images/awfnpv_ZA06051212.gif)


    
- NPV is similar to the PV function (present value). The primary difference between PV and NPV is that PV allows cash flows to begin either at the end or at the beginning of the period. Unlike the variable NPV cash flow values, PV cash flows must be constant throughout the investment. For information about annuities and financial functions, see PV.
    
- NPV is also related to the IRR function (internal rate of return). IRR is the rate for which NPV equals zero: NPV(IRR(...), ...) = 0.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

