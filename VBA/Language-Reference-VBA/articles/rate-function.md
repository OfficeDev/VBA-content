---
title: Rate Function
keywords: vblr6.chm1009288
f1_keywords:
- vblr6.chm1009288
ms.prod: office
ms.assetid: fa2c01bd-e717-c199-00b1-e2e56ec86b01
ms.date: 06/08/2017
---


# Rate Function



Returns a [Double](vbe-glossary.md) specifying the interest rate per period for an annuity.
 **Syntax**
 **Rate( _nper_**, **_pmt_**, **_pv_** [, **_fv_** [, **_type_** [, **_guess_** ]]] **)**
The  **Rate** function has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_nper_**|Required.  **Double** specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.|
|**_pmt_**|Required.  **Double** specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.|
|**_pv_**|Required.  **Double** specifying present value, or value today, of a series of future payments or receipts. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you will make.|
|**_fv_**|Optional. [Variant](vbe-glossary.md) specifying future value or cash balance you want after you make the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, then $50,000 is the future value. If omitted, 0 is assumed.|
|**_type_**|Optional.  **Variant** specifying a number indicating when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.|
|**_guess_**|Optional.  **Variant** specifying value you estimate will be returned by **Rate**. If omitted, **_guess_** is 0.1 (10 percent).|
 **Remarks**
An annuity is a series of fixed cash payments made over a period of time. An annuity can be a loan (such as a home mortgage) or an investment (such as a monthly savings plan).
For all [arguments](vbe-glossary.md), cash paid out (such as deposits to savings) is represented by negative numbers; cash received (such as dividend checks) is represented by positive numbers.
 **Rate** is calculated by iteration. Starting with the value of **_guess_**, **Rate** cycles through the calculation until the result is accurate to within 0.00001 percent. If **Rate** can't find a result after 20 tries, it fails. If your guess is 10 percent and **Rate** fails, try a different value for **_guess_**.

## Example

This example uses the  **Rate** function to calculate the interest rate of a loan given the total number of payments ( `TotPmts`), the amount of the loan payment ( `Payment`), the present value or principal of the loan ( `PVal`), the future value of the loan ( `FVal`), a number that indicates whether the payment is due at the beginning or end of the payment period ( `PayType`), and an approximation of the expected interest rate ( `Guess`).


```vb
Dim Fmt, FVal, Guess, PVal, Payment, TotPmts, PayType, APR
Const ENDPERIOD = 0, BEGINPERIOD = 1    ' When payments are made.
Fmt = "##0.00"    ' Define percentage format.
FVal = 0    ' Usually 0 for a loan.
Guess = .1    ' Guess of 10 percent.
PVal = InputBox("How much did you borrow?")
Payment = InputBox("What's your monthly payment?")
TotPmts = InputBox("How many monthly payments do you have to make?")
PayType = MsgBox("Do you make payments at the end of the month?", _
vbYesNo)
If PayType = vbNo Then PayType = BEGINPERIOD Else PayType = ENDPERIOD
APR = (Rate(TotPmts, -Payment, PVal, FVal, PayType, Guess) * 12) * 100
MsgBox "Your interest rate is " &; Format(CInt(APR), Fmt) &; " percent."

```


