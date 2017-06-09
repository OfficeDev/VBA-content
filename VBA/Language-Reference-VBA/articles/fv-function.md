---
title: FV Function
keywords: vblr6.chm1009280
f1_keywords:
- vblr6.chm1009280
ms.prod: office
ms.assetid: 9f77a5f2-77a9-ae4a-4ef0-c27136fcbd63
ms.date: 06/08/2017
---


# FV Function



Returns a [Double](vbe-glossary.md) specifying the future value of an annuity based on periodic, fixed payments and a fixed interest rate.
 **Syntax**
 **FV( _rate_**, **_nper_**, **_pmt_** [, **_pv_** [, **_type_** ]] **)**
The  **FV** function has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_rate_**|Required.  **Double** specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.|
|**_nper_**|Required. [Integer](vbe-glossary.md) specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.|
|**_pmt_**|Required.  **Double** specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.|
|**_pv_**|Optional. [Variant](vbe-glossary.md) specifying present value (or lump sum) of a series of future payments. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you will make. If omitted, 0 is assumed.|
|**_type_**|Optional.  **Variant** specifying when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.|
 **Remarks**
An annuity is a series of fixed cash payments made over a period of time. An annuity can be a loan (such as a home mortgage) or an investment (such as a monthly savings plan).
The  **_rate_** and **_nper_**[arguments](vbe-glossary.md) must be calculated using payment periods expressed in the same units. For example, if **_rate_** is calculated using months, **_nper_** must also be calculated using months.
For all arguments, cash paid out (such as deposits to savings) is represented by negative numbers; cash received (such as dividend checks) is represented by positive numbers.

## Example

This example uses the  **FV** function to return the future value of an investment given the percentage rate that accrues per period ( `APR / 12`), the total number of payments ( `TotPmts`), the payment ( `Payment`), the current value of the investment ( `PVal`), and a number that indicates whether the payment is made at the beginning or end of the payment period ( `PayType`). Note that because  `Payment` represents cash paid out, it's a negative number.


```vb
Dim Fmt, Payment, APR, TotPmts, PayType, PVal, FVal
Const ENDPERIOD = 0, BEGINPERIOD = 1    ' When payments are made.
Fmt = "###,###,##0.00"    ' Define money format.
Payment = InputBox("How much do you plan to save each month?")
APR = InputBox("Enter the expected interest annual percentage rate.")
If APR > 1 Then APR = APR / 100    ' Ensure proper form.
TotPmts = InputBox("For how many months do you expect to save?")
PayType = MsgBox("Do you make payments at the end of month?", vbYesNo)
If PayType = vbNo Then PayType = BEGINPERIOD Else PayType = ENDPERIOD
PVal = InputBox("How much is in this savings account now?")
FVal = FV(APR / 12, TotPmts, -Payment, -PVal, PayType)
MsgBox "Your savings will be worth " &; Format(FVal, Fmt) &; "."
```


