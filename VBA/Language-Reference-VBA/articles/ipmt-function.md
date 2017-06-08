---
title: IPmt Function
keywords: vblr6.chm1009281
f1_keywords:
- vblr6.chm1009281
ms.prod: office
ms.assetid: ee1db73e-9121-8451-8780-d4681f1f42dd
ms.date: 06/08/2017
---


# IPmt Function



Returns a [Double](vbe-glossary.md) specifying the interest payment for a given period of an annuity based on periodic, fixed payments and a fixed interest rate.
 **Syntax**
 **IPmt( _rate_**, **_per_**, **_nper_**, **_pv_** [, **_fv_** [, **_type_** ]] **)**
The  **IPmt** function has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_rate_**|Required.  **Double** specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.|
|**_per_**|Required.  **Double** specifying payment period in the range 1 through **_nper_**.|
|**_nper_**|Required.  **Double** specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.|
|**_pv_**|Required.  **Double** specifying present value, or value today, of a series of future payments or receipts. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you will make.|
|**_fv_**|Optional. [Variant](vbe-glossary.md) specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, then $50,000 is the future value. If omitted, 0 is assumed.|
|**_type_**|Optional.  **Variant** specifying when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.|
 **Remarks**
An annuity is a series of fixed cash payments made over a period of time. An annuity can be a loan (such as a home mortgage) or an investment (such as a monthly savings plan).
The  **_rate_** and **_nper_**[arguments](vbe-glossary.md) must be calculated using payment periods expressed in the same units. For example, if **_rate_** is calculated using months, **_nper_** must also be calculated using months.
For all arguments, cash paid out (such as deposits to savings) is represented by negative numbers; cash received (such as dividend checks) is represented by positive numbers.

## Example

This example uses the  **IPmt** function to calculate how much of a payment is interest when all the payments are of equal value. Given are the interest percentage rate per period ( `APR / 12`), the payment period for which the interest portion is desired ( `Period`), the total number of payments ( `TotPmts`), the present value or principal of the loan ( `PVal`), the future value of the loan ( `FVal`), and a number that indicates whether the payment is due at the beginning or end of the payment period ( `PayType`).


```VB.net
Dim FVal, Fmt, PVal, APR, TotPmts, PayType, Period, IntPmt, TotInt, Msg
Const ENDPERIOD = 0, BEGINPERIOD = 1    ' When payments are made.
FVal = 0    ' Usually 0 for a loan.
Fmt = "###,###,##0.00"    ' Define money format.
PVal = InputBox("How much do you want to borrow?")
APR = InputBox("What is the annual percentage rate of your loan?")
If APR > 1 Then APR = APR / 100    ' Ensure proper form.
TotPmts = InputBox("How many monthly payments?")
PayType = MsgBox("Do you make payments at end of the month?", vbYesNo)
If PayType = vbNo Then PayType = BEGINPERIOD Else PayType = ENDPERIOD
For Period = 1 To TotPmts    ' Total all interest.
    IntPmt = IPmt(APR / 12, Period, TotPmts, -PVal, FVal, PayType)
    TotInt = TotInt + IntPmt
Next Period
Msg = "You'll pay a total of " &; Format(TotInt, Fmt) 
Msg = Msg &; " in interest for this loan."
MsgBox Msg    ' Display results.

```


