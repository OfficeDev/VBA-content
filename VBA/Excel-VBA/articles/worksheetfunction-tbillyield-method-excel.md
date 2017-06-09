---
title: WorksheetFunction.TBillYield Method (Excel)
keywords: vbaxl10.chm137317
f1_keywords:
- vbaxl10.chm137317
ms.prod: excel
api_name:
- Excel.WorksheetFunction.TBillYield
ms.assetid: 00827b10-a295-ef06-8947-fd9769bc1db5
ms.date: 06/08/2017
---


# WorksheetFunction.TBillYield Method (Excel)

Returns the yield for a Treasury bill.


## Syntax

 _expression_ . **TBillYield**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Settlement - the Treasury bill's settlement date. The security settlement date is the date after the issue date when the Treasury bill is traded to the buyer.|
| _Arg2_|Required| **Variant**|Maturity - the Treasury bill's maturity date. The maturity date is the date when the Treasury bill expires.|
| _Arg3_|Optional| **Variant**|Pr - the Treasury bill's price per $100 face value.|

### Return Value

Double


## Remarks


 **Important**  Dates should be entered by using the DATE function, or as results of other formulas or functions. For example, use DATE(2008,5,23) for the 23rd day of May, 2008. Problems can occur if dates are entered as text.


- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- Settlement and maturity are truncated to integers.
    
- If settlement or maturity is not a valid date, TBILLYIELD returns the #VALUE! error value.
    
- If pr ? 0, TBILLYIELD returns the #NUM! error value.
    
- If settlement ? maturity, or if maturity is more than one year after settlement, TBILLYIELD returns the #NUM! error value.
    
- TBILLYIELD is calculated as follows:
![Formula](images/awftbyd_ZA06051257.gif)where: DSM = number of days from settlement to maturity, excluding any maturity date that is more than one calendar year after the settlement date. 
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

