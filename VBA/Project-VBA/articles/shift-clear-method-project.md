---
title: Shift.Clear Method (Project)
ms.prod: project-server
api_name:
- Project.Shift.Clear
ms.assetid: 89243732-8c83-ba1e-01ff-fdbfa4d4c4d2
ms.date: 06/08/2017
---


# Shift.Clear Method (Project)

Clears the start and finish times of a work shift.


## Syntax

 _expression_. **Clear**

 _expression_ A variable that represents a **Shift** object.


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


