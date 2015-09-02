
# WorksheetFunction.BesselY Method (Excel)

 **Last modified:** July 28, 2015

Returns the Bessel function, which is also called the Weber function or the Neumann function.

## Syntax

 _expression_. **BesselY**( **_Arg1_**,  **_Arg2_**)

 _expression_A variable that represents a  **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Arg1|Required| **Variant**|The value at which to evaluate the function.|
|Arg2|Required| **Variant**|The order of the function. If n is not an integer, it is truncated.|

### Return Value

Double


## Remarks




- If x is nonnumeric, BesselY generates an error value.
    
- If n is nonnumeric, BesselY generates an error value.
    
- If n < 0, BesselY generates an error value.
    
- The n-th order Bessel function of the variable x is:
![](../images/awfbsly1_ZA06051118.gif)


    

## See also


#### Concepts


 [WorksheetFunction Object](7b1d5639-363d-632c-2cf0-2232562646b6.md)
#### Other resources


 [WorksheetFunction Object Members](6811ca87-4b53-0bff-88c9-30bf7497879a.md)
