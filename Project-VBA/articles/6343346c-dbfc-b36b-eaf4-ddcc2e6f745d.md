
# WeekDays.Count Property (Project)

 **Last modified:** July 28, 2015

Gets the number of items in the  **WeekDays** collection. Read-only **Integer**.

## Syntax

 _expression_. **Count**

 _expression_A variable that represents a  **WeekDays** object.


## Example

The following example shows there are seven days in the week for the calendar of the specified resource.


```
Debug.Print ActiveProject.Resources(1).Calendar.WorkWeeks(1).WeekDays.Count
```


## See also


#### Concepts


 [WeekDays Collection Object](757437a0-e2ff-0027-f044-87d1cb357f62.md)
