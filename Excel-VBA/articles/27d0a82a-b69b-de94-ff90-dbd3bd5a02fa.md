
# FormatCondition.Priority Property (Excel)

 **Last modified:** July 28, 2015

Returns or sets the priority value of the conditional formatting rule. The priority determines the order of evaluation when multiple conditional formatting rules exist in a worksheet.

## Syntax

 _expression_. **Priority**

 _expression_A variable that represents a  **FormatCondition** object.


## Remarks

When setting the priority, the value must be a positive integer between 1 and the total number of conditional formatting rules on the worksheet. The priority must be a unique value for all rules on the worksheet, so changing the priority for the specified conditional formatting rule may cause the priority value of the other rules on the worksheet to be shifted.


## See also


#### Concepts


 [FormatCondition Object](38a2bca9-9b28-3ef2-8c7a-4d35a27229ec.md)
#### Other resources


 [FormatCondition Object Members](8f4bebce-0bf4-03de-62f0-4454ea699c5f.md)
