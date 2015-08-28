
# ConditionValue.Value Property (Excel)

 **Last modified:** July 28, 2015

Returns or sets the shortest bar or longest bar threshold value for a data bar conditional format. Read/write  **Variant**.

## Syntax

 _expression_. **Value**

 _expression_A variable that represents a  **ConditionValue** object.


## Remarks

You can set the value only if the  ** [ConditionValue.Type](20467063-f402-4e7f-42ba-581b61b83a15.md)** property for the conditional format is set to one of the following constants: **xlConditionValueNumber**,  **xlConditionValuePercent**,  **xlConditionValuePercentile**, or  **xlConditionValueFormula**.

If the threshold type is a formula, you can set the formula as a  **String**. The formula must return a single number.


## See also


#### Concepts


 [ConditionValue Object](a39335db-4e0a-66aa-393b-3aa7e5268c00.md)
#### Other resources


 [ConditionValue Object Members](59e72c1f-3e56-294b-408a-de7aba0ed331.md)
