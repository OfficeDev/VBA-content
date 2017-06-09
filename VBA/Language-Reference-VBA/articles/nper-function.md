---
title: NPer Function
keywords: vblr6.chm1009446
f1_keywords:
- vblr6.chm1009446
ms.prod: office
ms.assetid: c4043308-390e-f8bd-800b-ca00e99214d1
ms.date: 06/08/2017
---


# NPer Function



Returns a [Double](vbe-glossary.md) specifying the number of periods for an annuity based on periodic, fixed payments and a fixed interest rate.
 **Syntax**
 **NPer( _rate_**, **_pmt_**, **_pv_** [, **_fv_** [, **_type_** ]] **)**
The  **NPer** function has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_rate_**|Required.  **Double** specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.|
|**_pmt_**|Required.  **Double** specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.|
|**_pv_**|Required.  **Double** specifying present value, or value today, of a series of future payments or receipts. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you will make.|
|**_fv_**|Optional. [Variant](vbe-glossary.md) specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, then $50,000 is the future value. If omitted, 0 is assumed.|
|**_type_**|Optional.  **Variant** specifying when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.|
 **Remarks**
An annuity is a series of fixed cash payments made over a period of time. An annuity can be a loan (such as a home mortgage) or an investment (such as a monthly savings plan).
For all [arguments](vbe-glossary.md), cash paid out (such as deposits to savings) is represented by negative numbers; cash received (such as dividend checks) is represented by positive numbers.

## Example

This example uses the  **NPer** function to return the number of periods during which payments must be made to pay off a loan whose value is contained in `PVal`. Also provided are the interest percentage rate per period ( `APR / 12`), the payment ( `Payment`), the future value of the loan ( `FVal`), and a number that indicates whether the payment is due at the beginning or end of the payment period ( `PayType`).


```vb
Dim FVal, PVal, APR, Payment, PayType, TotPmts
Const ENDPERIOD = 0, BEGINPERIOD = 1    ' When payments are made.
FVal = 0    ' Usually 0 for a loan.
PVal = InputBox("How much do you want to borrow?")
APR = InputBox("What is the annual percentage rate of your loan?")
If APR > 1 Then APR = APR / 100    ' Ensure proper form.
Payment = InputBox("How much do you want to pay each month?")
PayType = MsgBox("Do you make payments at the end of month?", vbYesNo)
If PayType = vbNo Then PayType = BEGINPERIOD Else PayType = ENDPERIOD
TotPmts = NPer(APR / 12, -Payment, PVal, FVal, PayType)
If Int(TotPmts) <> TotPmts Then TotPmts = Int(TotPmts) + 1
MsgBox "It will take you " &; TotPmts &; " months to pay off your loan."
```


