
# WorksheetFunction.Fixed Method (Excel)

 **Last modified:** July 28, 2015

Rounds a number to the specified number of decimals, formats the number in decimal format using a period and commas, and returns the result as text.

## Syntax

 _expression_. **Fixed**( **_Arg1_**,  **_Arg2_**,  **_Arg3_**)

 _expression_A variable that represents a  **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Arg1|Required| **Double**|Number - the number you want to round and convert to text.|
|Arg2|Optional| **Variant**|Decimals - the number of digits to the right of the decimal point.|
|Arg3|Optional| **Variant**|No_commas - a logical value that, if TRUE, prevents FIXED from including commas in the returned text.|

### Return Value

String


## Remarks




- Numbers in Microsoft Excel can never have more than 15 significant digits, but decimals can be as large as 127.
    
- If decimals is negative, number is rounded to the left of the decimal point.
    
- If you omit decimals, it is assumed to be 2.
    
- If no_commas is FALSE or omitted, then the returned text includes commas as usual.
    
- The major difference between formatting a cell containing a number with the  **Cells** command ( **Format** menu) and formatting a number directly with the FIXED function is that FIXED converts its result to text. A number formatted with the **Cells** command is still a number.
    

## See also


#### Concepts


 [WorksheetFunction Object](7b1d5639-363d-632c-2cf0-2232562646b6.md)
#### Other resources


 [WorksheetFunction Object Members](6811ca87-4b53-0bff-88c9-30bf7497879a.md)
