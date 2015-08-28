
# WorksheetFunction.Delta Method (Excel)

 **Last modified:** July 28, 2015

Tests whether two values are equal. Returns 1 if number1 = number2; returns 0 otherwise.

## Syntax

 _expression_. **Delta**( **_Arg1_**,  **_Arg2_**)

 _expression_A variable that represents a  **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Arg1|Required| **Variant**|Number1 - the first number.|
|Arg2|Optional| **Variant**|Number2 - the second number. If omitted, number2 is assumed to be zero.|

### Return Value

Double


## Remarks

 Use this function to filter a set of values. For example, by summing several DELTA functions you calculate the count of equal pairs. This function is also known as the Kronecker Delta function.


- If number1 is nonnumeric, DELTA returns the #VALUE! error value.
    
- If number2 is nonnumeric, DELTA returns the #VALUE! error value.
    

## See also


#### Concepts


 [WorksheetFunction Object](7b1d5639-363d-632c-2cf0-2232562646b6.md)
#### Other resources


 [WorksheetFunction Object Members](6811ca87-4b53-0bff-88c9-30bf7497879a.md)
