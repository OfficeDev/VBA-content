---
title: WorksheetFunction.IntRate Method (Excel)
keywords: vbaxl10.chm137310
f1_keywords:
- vbaxl10.chm137310
ms.prod: excel
api_name:
- Excel.WorksheetFunction.IntRate
ms.assetid: cf5c96e2-6f5e-dcaa-7682-fd925c76d2c6
ms.date: 06/08/2017
---


# WorksheetFunction.IntRate Method (Excel)

Returns the interest rate for a fully invested security.


## Syntax

 _expression_ . **IntRate**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Settlement - the security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.|
| _Arg2_|Required| **Variant**|Maturity - the security's maturity date. The maturity date is the date when the security expires.|
| _Arg3_|Required| **Variant**|Investment - the amount invested in the security.|
| _Arg4_|Required| **Variant**|Redemption - the amount to be received at maturity.|
| _Arg5_|Optional| **Variant**|Basis - the type of day count basis to use.|

### Return Value

Double


## Remarks



|**Basis**|**Day count basis**|
|:-----|:-----|
|0 or omitted|US (NASD) 30/360|
|1|Actual/actual|
|2|Actual/360|
|3|Actual/365|
|4|European 30/360|

- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- The settlement date is the date a buyer purchases a coupon, such as a bond. The maturity date is the date when a coupon expires. For example, suppose a 30-year bond is issued on January 1, 2008, and is purchased by a buyer six months later. The issue date would be January 1, 2008, the settlement date would be July 1, 2008, and the maturity date would be January 1, 2038, which is 30 years after the January 1, 2008, issue date.
    
- Settlement, maturity, and basis are truncated to integers.
    
- If settlement or maturity is not a valid date, INTRATE returns the #VALUE! error value.
    
- If investment ? 0 or if redemption ? 0, INTRATE returns the #NUM! error value.
    
- If basis < 0 or if basis > 4, INTRATE returns the #NUM! error value.
    
- If settlement ? maturity, INTRATE returns the #NUM! error value.
    
- INTRATE is calculated as follows:
![Formula](images/awfintrt_ZA06051176.gif)where: B = number of days in a year, depending on the year basis. DIM = number of days from settlement to maturity. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

