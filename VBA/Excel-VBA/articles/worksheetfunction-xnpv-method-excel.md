---
title: WorksheetFunction.Xnpv Method (Excel)
keywords: vbaxl10.chm137307
f1_keywords:
- vbaxl10.chm137307
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Xnpv
ms.assetid: db61e7a8-70c2-9e32-48dd-adddcbc886b6
ms.date: 06/08/2017
---


# WorksheetFunction.Xnpv Method (Excel)

Returns the net present value for a schedule of cash flows that is not necessarily periodic. Read/write  **Double** .


## Syntax

 _expression_ . **Xnpv**( **_Arg1_** , **_Arg2_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|A series of cash flows that corresponds to a schedule of payments in dates. The first payment is optional and corresponds to a cost or payment that occurs at the beginning of the investment.|
| _Arg2_|Required| **Variant**|A schedule of payment dates that corresponds to the cash flow payments. The first payment date indicates the beginning of the schedule of payments. All other dates must be later than this date, but they may occur in any order.|

### Return Value

Double


## Remarks

To calculate the net present value for a series of cash flows that is periodic, use the  **[Npv](worksheetfunction-npv-method-excel.md)** method.


 **Important**  The  **Xnpv** method does not provide a parameter that corresponds to the _rate_ argument required by the corresponding **XNPV** function (=XNPV( **_rate_** , _values_ , _dates_ )). To work around this limitation in VBA code, instead of using the **Xpnv** method, call the **XNPV** function by using the **[Evaluate](application-evaluate-method-excel.md)** method as shown in the following example.


## Example

The following example returns the net present value for an investment with the above cost and returns. The cash flows are discounted at 9 percent. (2086.6476 or 2086.65)


```vb
Dim npv As Double 
npv = Application.Evaluate("=XNPV(.09,A2:A6,B2:B6)")
```


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

