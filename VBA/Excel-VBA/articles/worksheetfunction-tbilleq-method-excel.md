---
title: WorksheetFunction.TBillEq Method (Excel)
keywords: vbaxl10.chm137315
f1_keywords:
- vbaxl10.chm137315
ms.prod: excel
api_name:
- Excel.WorksheetFunction.TBillEq
ms.assetid: 4b52fbb3-5d25-3fae-cdf8-ec3d406ce787
ms.date: 06/08/2017
---


# WorksheetFunction.TBillEq Method (Excel)

Returns the bond-equivalent yield for a Treasury bill.


## Syntax

 _expression_ . **TBillEq**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Settlement - the Treasury bill's settlement date. The security settlement date is the date after the issue date when the Treasury bill is traded to the buyer.|
| _Arg2_|Required| **Variant**|Maturity - the Treasury bill's maturity date. The maturity date is the date when the Treasury bill expires.|
| _Arg3_|Optional| **Variant**|Discount - the Treasury bill's discount rate.|

### Return Value

Double


## Remarks


 **Important**  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.


- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- Settlement and maturity are truncated to integers.
    
- If settlement or maturity is not a valid date, TBILLEQ returns the #VALUE! error value.
    
- If discount ? 0, TBILLEQ returns the #NUM! error value.
    
- If settlement > maturity, or if maturity is more than one year after settlement, TBILLEQ returns the #NUM! error value.
    
- TBILLEQ is calculated as TBILLEQ = (365 x rate)/(360-(rate x DSM)), where DSM is the number of days between settlement and maturity computed according to the 360 days per year basis.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

