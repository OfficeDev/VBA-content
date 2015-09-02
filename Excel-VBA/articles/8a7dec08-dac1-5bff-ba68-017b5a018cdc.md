
# WorksheetFunction.CoupDayBs Method (Excel)

 **Last modified:** July 28, 2015

Returns the number of days from the beginning of the coupon period to the settlement date.

## Syntax

 _expression_. **CoupDayBs**( **_Arg1_**,  **_Arg2_**,  **_Arg3_**,  **_Arg4_**)

 _expression_A variable that represents a  **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Arg1|Required| **Variant**|The security's settlement date. The security settlement date is the date after the issue date when the security is traded to the buyer.|
|Arg2|Required| **Variant**|The security's maturity date. The maturity date is the date when the security expires.|
|Arg3|Required| **Variant**|The number of coupon payments per year. For annual payments, frequency = 1; for semiannual, frequency = 2; for quarterly, frequency = 4.|
|Arg4|Optional| **Variant**|The type of day count basis to use.|

### Return Value

Double


## Remarks

The following table contains the list of values for  _Arg4_.



|**Basis**|**Day count basis**|
|:-----|:-----|
|0 or omitted|US (NASD) 30/360|
|1|Actual/actual|
|2|Actual/360|
|3|Actual/365|
|4|European 30/360|

- Microsoft Excel stores dates as sequential serial numbers so they can be used in calculations. By default, January 1, 1900 is serial number 1, and January 1, 2008 is serial number 39448 because it is 39,448 days after January 1, 1900. Microsoft Excel for the Macintosh uses a different date system as its default.
    
     **Note**  Visual Basic for Applications (VBA) calculates serial dates differently than Excel. In VBA, serial number 1 is December 31, 1899, rather than January 1, 1900. 
- The settlement date is the date a buyer purchases a coupon, such as a bond. The maturity date is the date when a coupon expires. For example, suppose a 30-year bond is issued on January 1, 2008, and is purchased by a buyer six months later. The issue date would be January 1, 2008, the settlement date would be July 1, 2008, and the maturity date would be January 1, 2038, 30 years after the January 1, 2008, issue date.
    
- All arguments are truncated to integers.
    
- If settlement or maturity is not a valid date, CoupDayBs generates an error.
    
- If frequency is any number other than 1, 2, or 4, CoupDayBs generates an error.
    
- If basis < 0 or if basis > 4, CoupDayBs generates an error.
    
- If settlement â‰¥ maturity, CoupDayBs generates an error.
    

## See also


#### Concepts


 [WorksheetFunction Object](7b1d5639-363d-632c-2cf0-2232562646b6.md)
#### Other resources


 [WorksheetFunction Object Members](6811ca87-4b53-0bff-88c9-30bf7497879a.md)
