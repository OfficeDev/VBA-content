
# ColorScale.ModifyAppliesToRange Method (Excel)

 **Last modified:** July 28, 2015

Sets the cell range to which this formatting rule applies.

## Syntax

 _expression_. **ModifyAppliesToRange**( **_Range_**)

 _expression_A variable that represents a  **ColorScale** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Range|Required| **Range**|The range to which this formatting rule will be applied.|

## Remarks

The range must be in the A1 reference style and be entirely contained within the sheet that is the parent of the  ** [FormatConditions](2486d4b4-605c-76d8-132a-694c0c600a81.md)** collection. It can include the range operator (a colon), the intersection operator (a space), or the union operator (a comma). Dollar signs can also be used but they are ignored.

You can also use a local defined name in any part of the range, but the name must be in the language of the macro.


## See also


#### Concepts


 [ColorScale Object](3982b041-9178-7a45-7453-c88963501a3c.md)
#### Other resources


 [ColorScale Object Members](e14df078-3af6-a32e-d66f-3410b7bdb4d4.md)
