
# WorksheetFunction.Effect Method (Excel)

 **Last modified:** July 28, 2015

Returns the effective annual interest rate, given the nominal annual interest rate and the number of compounding periods per year.

## Syntax

 _expression_. **Effect**( **_Arg1_**,  **_Arg2_**)

 _expression_A variable that represents a  **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Arg1|Required| **Variant**|Nominal_rate - the nominal interest rate.|
|Arg2|Required| **Variant**|Npery - the number of compounding periods per year.|

### Return Value

Double


## Remarks




- Npery is truncated to an integer.
    
- If either argument is nonnumeric, EFFECT returns the #VALUE! error value.
    
- If nominal_rate â‰¤ 0 or if npery < 1, EFFECT returns the #NUM! error value.
    
- EFFECT is calculated as follows:
![](../images/awfefect_ZA06051135.gif)


    

## See also


#### Concepts


 [WorksheetFunction Object](7b1d5639-363d-632c-2cf0-2232562646b6.md)
#### Other resources


 [WorksheetFunction Object Members](6811ca87-4b53-0bff-88c9-30bf7497879a.md)
