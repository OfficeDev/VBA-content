
# TimeScaleValue.Clear Method (Project)

 **Last modified:** July 28, 2015

Clears the value of a timescaled data item.

## Syntax

 _expression_. **Clear**

 _expression_A variable that represents a  **TimeScaleValue** object.


## Example

The following example schedules a half-day of work on Fridays by creating an 8 A.M. to noon shift and removing the second and third shifts.


```
Sub HalfDayFridays() 
 With ActiveProject.Calendar.Weekdays(pjFriday) 
 .Shift1.Start = #8:00:00 AM# 
 .Shift1.Finish = #12:00:00 PM# 
 .Shift2.Clear 
 .Shift3.Clear 
 End With 
End Sub
```

