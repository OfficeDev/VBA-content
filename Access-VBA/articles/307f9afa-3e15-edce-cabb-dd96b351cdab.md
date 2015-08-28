
# WorksheetFunction.F_Dist_RT Method (Excel)

 **Last modified:** July 28, 2015

Returns the right-tailed F probability distribution. You can use this function to determine whether two data sets have different degrees of diversity. For example, you can examine the test scores of men and women entering high school and determine if the variability in the females is different from that found in the males.

## Syntax

 _expression_. **F_Dist_RT**( **_Arg1_**,  **_Arg2_**,  **_Arg3_**)

 _expression_A variable that represents a  ** [WorksheetFunction](7b1d5639-363d-632c-2cf0-2232562646b6.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Arg1|Required| **Double**|X - the value at which to evaluate the function.|
|Arg2|Required| **Double**|Degrees_freedom1 - the numerator degrees of freedom.|
|Arg3|Required| **Double**|Degrees_freedom2 - the denominator degrees of freedom.|

### Return Value

Double


## Remarks




- If any argument is nonnumeric, F_DIST_RT returns the #VALUE! error value.
    
- If x is negative, F_DIST_RT returns the #NUM! error value.
    
- If degrees_freedom1 or degrees_freedom2 is not an integer, it is truncated.
    
- If degrees_freedom1 < 1 or degrees_freedom1 â‰¥ 10^10, F_DIST_RT returns the #NUM! error value.
    
- If degrees_freedom2 < 1 or degrees_freedom2 â‰¥ 10^10, F_DIST_RT returns the #NUM! error value.
    
- F_DIST_RT is calculated as F_DIST_RT=P( F>x ), where F is a random variable that has an F distribution with degrees_freedom1 and degrees_freedom2 degrees of freedom.
    

## See also


#### Concepts


 [WorksheetFunction Object](7b1d5639-363d-632c-2cf0-2232562646b6.md)
#### Other resources


 [WorksheetFunction Object Members](6811ca87-4b53-0bff-88c9-30bf7497879a.md)
