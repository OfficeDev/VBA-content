---
title: TimeScaleValue.Clear Method (Project)
ms.prod: project-server
api_name:
- Project.TimeScaleValue.Clear
ms.assetid: 3ed3a584-5496-cdf4-eafa-e0ecdd01edfd
ms.date: 06/08/2017
---


# TimeScaleValue.Clear Method (Project)

Clears the value of a timescaled data item.


## Syntax

 _expression_. **Clear**

 _expression_ A variable that represents a **TimeScaleValue** object.


## Example

The following example schedules a half-day of work on Fridays by creating an 8 A.M. to noon shift and removing the second and third shifts.


```vb
Sub HalfDayFridays() 
 With ActiveProject.Calendar.Weekdays(pjFriday) 
 .Shift1.Start = #8:00:00 AM# 
 .Shift1.Finish = #12:00:00 PM# 
 .Shift2.Clear 
 .Shift3.Clear 
 End With 
End Sub
```


