
# Task.Baseline4DurationText Property (Project)

 **Last modified:** July 28, 2015

Gets or sets a string representation of the baseline duration of a task. Read/write  **String**.

## Syntax

 _expression_. **Baseline4DurationText**

 _expression_An expression that returns a  **Task** object.


## Remarks

The  **Baseline4DurationText** property is used for manually scheduled tasks. When you convert a manually scheduled task to an auto-scheduled task, if **Baseline4DurationText** has any value, you should convert the value to a date for the **Baseline4Duration** property.

