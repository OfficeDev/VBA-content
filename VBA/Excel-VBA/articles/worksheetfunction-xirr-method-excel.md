---
title: WorksheetFunction.Xirr Method (Excel)
keywords: vbaxl10.chm137306
f1_keywords:
- vbaxl10.chm137306
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Xirr
ms.assetid: ac3b11b1-501a-1585-5c60-6e82167522aa
ms.date: 06/08/2017
---


# WorksheetFunction.Xirr Method (Excel)

Returns the internal rate of return for a schedule of cash flows that is not necessarily periodic. To calculate the internal rate of return for a series of periodic cash flows, use the IRR function.


## Syntax

 _expression_ . **Xirr**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Values - a series of cash flows that corresponds to a schedule of payments in dates. The first payment is optional and corresponds to a cost or payment that occurs at the beginning of the investment. If the first value is a cost or payment, it must be a negative value. All succeeding payments are discounted based on a 365-day year. The series of values must contain at least one positive and one negative value.|
| _Arg2_|Required| **Variant**|Dates - a schedule of payment dates that corresponds to the cash flow payments. The first payment date indicates the beginning of the schedule of payments. All other dates must be later than this date, but they may occur in any order. Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.|
| _Arg3_|Optional| **Variant**|Guess - a number that you guess is close to the result of XIRR.|

### Return Value

Double


## Remarks




- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- Numbers in dates are truncated to integers.
    
- XIRR expects at least one positive cash flow and one negative cash flow; otherwise, XIRR returns the #NUM! error value.
    
- If any number in dates is not a valid date, XIRR returns the #VALUE! error value.
    
- If any number in dates precedes the starting date, XIRR returns the #NUM! error value.
    
- If values and dates contain a different number of values, XIRR returns the #NUM! error value.
    
- In most cases you do not need to provide guess for the XIRR calculation. If omitted, guess is assumed to be 0.1 (10 percent).
    
- XIRR is closely related to XNPV, the net present value function. The rate of return calculated by XIRR is the interest rate corresponding to XNPV = 0.
    
- Excel uses an iterative technique for calculating XIRR. Using a changing rate (starting with guess), XIRR cycles through the calculation until the result is accurate within 0.000001 percent. If XIRR can't find a result that works after 100 tries, the #NUM! error value is returned. The rate is changed until:
![Formula](images/awfxirr_ZA06051264.gif)where: di = the ith, or last, payment date. d1 = the 0th payment date. Pi = the ith, or last, payment. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

