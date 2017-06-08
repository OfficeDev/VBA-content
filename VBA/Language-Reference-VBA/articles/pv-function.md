---
title: PV Function
keywords: vblr6.chm1009287
f1_keywords:
- vblr6.chm1009287
ms.prod: office
ms.assetid: b09d617d-675f-68b9-5b57-7134bed9040d
ms.date: 06/08/2017
---


# PV Function



Returns a [Double](vbe-glossary.md) specifying the present value of an annuity based on periodic, fixed payments to be paid in the future and a fixed interest rate.
 **Syntax**
 **PV( _rate_**, **_nper_**, **_pmt_** [, **_fv_** [, **_type_** ]] **)**
The  **PV** function has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_rate_**|Required.  **Double** specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.|
|**_nper_**|Required. [Integer](vbe-glossary.md) specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.|
|**_pmt_**|Required.  **Double** specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.|
|**_fv_**|Optional. [Variant](vbe-glossary.md) specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, then $50,000 is the future value. If omitted, 0 is assumed.|
|**_type_**|Optional.  **Variant** specifying when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.|
 **Remarks**
An annuity is a series of fixed cash payments made over a period of time. An annuity can be a loan (such as a home mortgage) or an investment (such as a monthly savings plan).
The  **_rate_** and **_nper_**[arguments](vbe-glossary.md) must be calculated using payment periods expressed in the same units. For example, if **_rate_** is calculated using months, **_nper_** must also be calculated using months.
For all arguments, cash paid out (such as deposits to savings) is represented by negative numbers; cash received (such as dividend checks) is represented by positive numbers.

## Example

In this example, the  **PV** function returns the present value of an $1,000,000 annuity that will provide $50,000 a year for the next 20 years. Provided are the expected annual percentage rate ( `APR`), the total number of payments ( `TotPmts`), the amount of each payment ( `YrIncome`), the total future value of the investment ( `FVal`), and a number that indicates whether each payment is made at the beginning or end of the payment period ( `PayType`).Note that  `YrIncome` is a negative number because it represents cash paid out from the annuity each year.


```vb
Dim Fmt, APR, TotPmts, YrIncome, FVal, PayType, PVal
Const ENDPERIOD = 0, BEGINPERIOD = 1    ' When payments are made.
Fmt = "###,##0.00"    ' Define money format.
APR = .0825    ' Annual percentage rate.
TotPmts = 20    ' Total number of payments.
YrIncome = 50000    ' Yearly income.
FVal = 1000000    ' Future value.
PayType = BEGINPERIOD    ' Payment at beginning of month.
PVal = PV(APR, TotPmts, -YrIncome, FVal, PayType)
MsgBox "The present value is " &; Format(PVal, Fmt) &; "."

```


